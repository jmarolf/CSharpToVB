Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports ArgumentSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArgumentSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports IdentifierNameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.IdentifierNameSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)
            Private Function VisitCSArguments(CS_OpenToken As SyntaxToken, CS_VisitorArguments As SeparatedSyntaxList(Of CSS.ArgumentSyntax), CS_CloseToken As SyntaxToken) As VisualBasicSyntaxNode
                If CS_VisitorArguments.Count = 0 Then
                    Return SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList(CS_VisitorArguments.[Select](Function(a As CSS.ArgumentSyntax) DirectCast(a.Accept(Me), ArgumentSyntax))))
                End If
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = CS_VisitorArguments.GetSeparators
                Dim NodeList As New List(Of ArgumentSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = CS_VisitorArguments.Count - 1
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.ArgumentSyntax = CS_VisitorArguments(i)
                    Dim ArgumentSyntaxNode As ArgumentSyntax = DirectCast(e.Accept(Me), ArgumentSyntax)
                    NodeList.Add(ArgumentSyntaxNode)
                    If SeparatorCount > i Then
                        Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                    End If
                Next
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                Dim OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken).WithConvertedTriviaFrom(CS_OpenToken)
                Dim CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken).WithConvertedTriviaFrom(CS_CloseToken)
                RestructureNodesAndSeparators(OpenParenToken, NodeList, Separators, CloseParenToken, TrailingTriviaList)
                If TrailingTriviaList.Count > 0 Then
                    Stop
                End If
                Return SyntaxFactory.ArgumentList(
                                                  OpenParenToken,
                                                  SyntaxFactory.SeparatedList(NodeList, Separators),
                                                  CloseParenToken
                                                  ).WithTrailingTrivia(TrailingTriviaList)
            End Function

            Public Overrides Function VisitArgument(ByVal node As CSS.ArgumentSyntax) As VisualBasicSyntaxNode
                Dim name As NameColonEqualsSyntax = Nothing
                Dim NodeExpression As CSS.ExpressionSyntax = node.Expression
                Dim ArgumentWithTrivia As ExpressionSyntax = Nothing
                Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
                Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                Try
                    ArgumentWithTrivia = DirectCast(NodeExpression.Accept(Me), ExpressionSyntax)
                    If TypeOf node.Parent Is BracketedArgumentListSyntax Then
                        Dim _Typeinfo As TypeInfo = ModelExtensions.GetTypeInfo(mSemanticModel, NodeExpression)
                        If _Typeinfo.ConvertedType IsNot _Typeinfo.Type Then
                            If _Typeinfo.Type?.SpecialType = SpecialType.System_Char Then '
                                ArgumentWithTrivia = SyntaxFactory.ParseExpression($"ChrW({ArgumentWithTrivia.WithoutTrivia.ToString})").WithTriviaFrom(ArgumentWithTrivia)
                            End If
                        End If
                    End If
                    If node.NameColon IsNot Nothing Then
                        name = SyntaxFactory.NameColonEquals(DirectCast(node.NameColon.Name.Accept(Me), IdentifierNameSyntax))
                        Dim NameWithOutColon As String = name.Name.ToString.Replace(":=", "")
                        If NameWithOutColon.EndsWith("_Renamed") Then
                            name = SyntaxFactory.NameColonEquals(SyntaxFactory.IdentifierName(NameWithOutColon.Replace("_Renamed", "")))
                        End If
                    End If

                    If ArgumentWithTrivia.HasLeadingTrivia Then
                        For Each trivia As SyntaxTrivia In ArgumentWithTrivia.GetLeadingTrivia
                            Select Case trivia.Kind
                                Case SyntaxKind.WhitespaceTrivia
                                    NewLeadingTrivia.Add(trivia)
                                Case SyntaxKind.EndOfLineTrivia
                                    ' Possibly need to Ignore TODO
                                    NewLeadingTrivia.Add(trivia)
                                Case SyntaxKind.CommentTrivia
                                    NewTrailingTrivia.Add(trivia)
                                Case Else
                                    Stop
                            End Select
                        Next
                    End If
                    NewTrailingTrivia.AddRange(ArgumentWithTrivia.GetTrailingTrivia)
                Catch ex As Exception
                    Stop
                End Try
                ArgumentWithTrivia = ArgumentWithTrivia.WithLeadingTrivia(NewLeadingTrivia).WithTrailingTrivia(SyntaxFactory.ElasticSpace)
                Return SyntaxFactory.SimpleArgument(name, ArgumentWithTrivia).WithTrailingTrivia(NewTrailingTrivia)
            End Function

            Public Overrides Function VisitArgumentList(ByVal node As CSS.ArgumentListSyntax) As VisualBasicSyntaxNode
                Dim CS_OpenParenToken As SyntaxToken = node.OpenParenToken
                Dim CS_VisitorArguments As SeparatedSyntaxList(Of CSS.ArgumentSyntax) = node.Arguments
                Dim CS_CloseParenToken As SyntaxToken = node.CloseParenToken

                Return VisitCSArguments(CS_OpenParenToken, CS_VisitorArguments, CS_CloseParenToken)
            End Function

            Public Overrides Function VisitBracketedArgumentList(ByVal node As CSS.BracketedArgumentListSyntax) As VisualBasicSyntaxNode
                Dim CS_OpenBracketToken As SyntaxToken = node.OpenBracketToken
                Dim CS_VisitorArguments As SeparatedSyntaxList(Of CSS.ArgumentSyntax) = node.Arguments
                Dim CS_CloseBracketToken As SyntaxToken = node.CloseBracketToken

                Return VisitCSArguments(CS_OpenBracketToken, CS_VisitorArguments, CS_CloseBracketToken)
            End Function
            Public Overrides Function VisitOmittedTypeArgument(ByVal node As CSS.OmittedTypeArgumentSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.ParseTypeName("").WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitTypeArgumentList(ByVal node As CSS.TypeArgumentListSyntax) As VisualBasicSyntaxNode
                Dim CS_VisitorArguments As SeparatedSyntaxList(Of CSS.TypeSyntax) = node.Arguments

                If CS_VisitorArguments.Count = 0 Then
                    Throw UnreachableException
                    Return SyntaxFactory.TypeArgumentList(SyntaxFactory.SeparatedList(CS_VisitorArguments.[Select](Function(a As CSS.TypeSyntax) DirectCast(a.Accept(Me), TypeSyntax))))
                End If
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = CS_VisitorArguments.GetSeparators
                Dim NodeList As New List(Of TypeSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = CS_VisitorArguments.Count - 1
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.TypeSyntax = CS_VisitorArguments(i)
                    Dim TypeSyntaxNode As TypeSyntax = DirectCast(e.Accept(Me), TypeSyntax)
                    NodeList.Add(TypeSyntaxNode)
                    If SeparatorCount > i Then
                        Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                    End If
                Next
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                Dim OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken).WithConvertedTriviaFrom(node.LessThanToken)
                Dim CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken).WithConvertedTriviaFrom(node.GreaterThanToken)
                RestructureNodesAndSeparators(OpenParenToken, NodeList, Separators, CloseParenToken, TrailingTriviaList)
                If TrailingTriviaList.Count > 0 Then
                    Stop
                End If
                Return SyntaxFactory.TypeArgumentList(
                                                  OpenParenToken, OfKeyword,
                                                  SyntaxFactory.SeparatedList(NodeList, Separators),
                                                  CloseParenToken
                                                  ).WithTrailingTrivia(TrailingTriviaList)
            End Function
        End Class
    End Class
End Namespace
