Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports AttributeListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.AttributeListSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports ParameterSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ParameterSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)
            Public Overrides Function VisitBracketedParameterList(ByVal node As CSS.BracketedParameterListSyntax) As VisualBasicSyntaxNode
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = node.Parameters.GetSeparators
                Dim OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken).WithConvertedTriviaFrom(node.OpenBracketToken)
                Dim CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken)
                Dim Items As New List(Of ParameterSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = node.Parameters.Count - 1
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.ParameterSyntax = node.Parameters(i)
                    Dim ItemWithTrivia As ParameterSyntax = DirectCast(e.Accept(Me).WithConvertedTrailingTriviaFrom(e), ParameterSyntax)
                    Dim Item As ParameterSyntax = ItemWithTrivia.WithoutTrivia
                    Items.Add(ItemWithTrivia)
                    If SeparatorCount > i Then
                        If Items.Last.ContainsEOLTrivia Then
                            Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                        End If
                    End If
                Next
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                RestructureNodesAndSeparators(OpenParenToken, Items, Separators, CloseParenToken, TrailingTriviaList)
                Return SyntaxFactory.ParameterList(OpenParenToken, SyntaxFactory.SeparatedList(Items, Separators), CloseParenToken).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitParameter(ByVal node As CSS.ParameterSyntax) As VisualBasicSyntaxNode
                Dim returnType As TypeSyntax = DirectCast(node.Type?.Accept(Me), TypeSyntax)
                If returnType IsNot Nothing Then
                    If returnType.ToString.StartsWith("[") Then

                        Dim TReturnType As TypeSyntax = SyntaxFactory.ParseTypeName(returnType.ToString.Substring(1).Replace("]", "")).WithTriviaFrom(returnType)
                        If Not TReturnType.IsMissing Then
                            returnType = TReturnType
                        End If
                    End If
                    If node.HasLeadingTrivia AndAlso node.Type.GetLeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                        returnType = returnType.With({SyntaxFactory.Space}, ConvertTrivia(node.Type.GetLeadingTrivia))
                    End If
                End If
                Dim [default] As EqualsValueSyntax = Nothing
                If node.[Default] IsNot Nothing Then
                    [default] = SyntaxFactory.EqualsValue(DirectCast(node.[Default]?.Value.Accept(Me), ExpressionSyntax))
                End If

                Dim newAttributes As AttributeListSyntax()
                Dim modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.Local)
                If (modifiers.Count = 0 AndAlso returnType IsNot Nothing) OrElse node.Modifiers.Any(CS.SyntaxKind.ThisKeyword) Then
                    modifiers = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.ByValKeyword))
                    newAttributes = Array.Empty(Of AttributeListSyntax)
                ElseIf node.Modifiers.Any(CS.SyntaxKind.OutKeyword) Then
                    newAttributes = {SyntaxFactory.AttributeList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Attribute(SyntaxFactory.ParseTypeName("Out"))))}
                Else
                    newAttributes = Array.Empty(Of AttributeListSyntax)
                End If

                If [default] IsNot Nothing Then
                    modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.OptionalKeyword))
                End If

                Dim id As SyntaxToken = GenerateSafeVBToken(id:=node.Identifier, IsQualifiedName:=False)
                Dim OriginalAttributeList As IEnumerable(Of AttributeListSyntax) = node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax))
                Dim OriginalAttributeListHasOut As Boolean = OriginalAttributeList.Count > 0 AndAlso OriginalAttributeList(0).Attributes(0).Name.ToString = "Out"
                Dim AttributeLists As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(If(OriginalAttributeListHasOut, OriginalAttributeList, newAttributes.Concat(OriginalAttributeList)))
                Dim Identifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(id)
                Dim AsClause As SimpleAsClauseSyntax = If(returnType Is Nothing, Nothing, SyntaxFactory.SimpleAsClause(returnType))
                Dim parameterSyntax1 As ParameterSyntax = SyntaxFactory.Parameter(attributeLists:=AttributeLists,
                                                                                  modifiers:=modifiers,
                                                                                  identifier:=Identifier,
                                                                                  asClause:=AsClause,
                                                                                  [default]:=[default]).WithLeadingTrivia(SyntaxFactory.ElasticSpace)
                Return parameterSyntax1
            End Function
            Public Overrides Function VisitParameterList(ByVal node As CSS.ParameterListSyntax) As VisualBasicSyntaxNode
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = node.Parameters.GetSeparators
                Dim OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken).WithConvertedTriviaFrom(node.OpenParenToken)
                Dim CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken).WithConvertedTriviaFrom(node.CloseParenToken)
                Dim Items As New List(Of ParameterSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = node.Parameters.Count - 1
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.ParameterSyntax = node.Parameters(i)
                    Dim ItemWithTrivia As ParameterSyntax = DirectCast(e.Accept(Me), ParameterSyntax)
                    If ItemWithTrivia.Modifiers.Count = 1 And ItemWithTrivia.Modifiers(0).ToString = "ByVal" Then
                        ItemWithTrivia = ItemWithTrivia.WithModifiers(New SyntaxTokenList)
                    End If
                    Items.Add(ItemWithTrivia)
                    If SeparatorCount > i Then
                        Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                    End If
                Next
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                RestructureNodesAndSeparators(OpenParenToken, Items, Separators, CloseParenToken, TrailingTriviaList)
                Return SyntaxFactory.ParameterList(OpenParenToken, SyntaxFactory.SeparatedList(Items, Separators), CloseParenToken)
            End Function
        End Class
    End Class
End Namespace