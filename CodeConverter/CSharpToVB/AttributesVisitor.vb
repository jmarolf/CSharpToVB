Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports ArgumentListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArgumentListSyntax
Imports ArgumentSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArgumentSyntax
Imports AttributeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.AttributeSyntax
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
            Public Overrides Function VisitAttribute(ByVal node As CSS.AttributeSyntax) As VisualBasicSyntaxNode
                Dim list As CSS.AttributeListSyntax = DirectCast(node.Parent, CSS.AttributeListSyntax)
                Return SyntaxFactory.Attribute(DirectCast(list.Target?.Accept(Me), AttributeTargetSyntax), DirectCast(node.Name.Accept(Me).WithConvertedTriviaFrom(node.Name), TypeSyntax), DirectCast(node.ArgumentList?.Accept(Me), ArgumentListSyntax))
            End Function
            Public Overrides Function VisitAttributeArgument(ByVal node As CSS.AttributeArgumentSyntax) As VisualBasicSyntaxNode
                Dim name As NameColonEqualsSyntax = Nothing
                If node.NameColon IsNot Nothing Then
                    name = SyntaxFactory.NameColonEquals(DirectCast(node.NameColon.Name.Accept(Me), IdentifierNameSyntax))
                    ' HACK for VB Error
                    If name.ToString = "[error]:=" Then
                        name = Nothing
                    End If
                ElseIf node.NameEquals IsNot Nothing Then
                    name = SyntaxFactory.NameColonEquals(DirectCast(node.NameEquals.Name.Accept(Me), IdentifierNameSyntax))
                End If

                Dim value As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                Return SyntaxFactory.SimpleArgument(name, value).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitAttributeArgumentList(ByVal node As CSS.AttributeArgumentListSyntax) As VisualBasicSyntaxNode
                Dim ArgumentNodes As New List(Of ArgumentSyntax)
                For i As Integer = 0 To node.Arguments.Count - 1
                    Dim a As CSS.AttributeArgumentSyntax = node.Arguments(i)
                    ArgumentNodes.Add(DirectCast(a.Accept(Me), ArgumentSyntax))
                Next
                Return SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList(ArgumentNodes)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitAttributeList(ByVal node As CSS.AttributeListSyntax) As VisualBasicSyntaxNode
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = node.Attributes.GetSeparators
                Dim LessThanToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.LessThanToken).WithConvertedTriviaFrom(node.OpenBracketToken)
                Dim GreaterThenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.GreaterThanToken).WithConvertedTriviaFrom(node.CloseBracketToken)
                Dim AttributeList As New List(Of AttributeSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = node.Attributes.Count - 1
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.AttributeSyntax = node.Attributes(i)
                    AttributeList.Add(DirectCast(e.Accept(Me), AttributeSyntax))
                    If SeparatorCount > i Then
                        Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                    End If
                Next
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                RestructureNodesAndSeparators(LessThanToken, AttributeList, Separators, GreaterThenToken, TrailingTriviaList, AllowLeadingComments:=True)
                Dim Attributes1 As SeparatedSyntaxList(Of AttributeSyntax) = SyntaxFactory.SeparatedList(AttributeList, Separators)
                Return SyntaxFactory.AttributeList(LessThanToken, Attributes1, GreaterThenToken)
            End Function

            Public Overrides Function VisitAttributeTargetSpecifier(ByVal node As CSS.AttributeTargetSpecifierSyntax) As VisualBasicSyntaxNode
                Dim id As SyntaxToken
                Select Case CS.CSharpExtensions.Kind(node.Identifier)
                    Case CS.SyntaxKind.AssemblyKeyword
                        id = SyntaxFactory.Token(SyntaxKind.AssemblyKeyword)
                    Case CS.SyntaxKind.ReturnKeyword
                        ' Not necessary, return attributes are moved by ConvertAndSplitAttributes.
                        Return Nothing
                    Case Else
                        Throw New NotSupportedException()
                End Select

                Return SyntaxFactory.AttributeTarget(id).WithConvertedTriviaFrom(node)
            End Function

        End Class
    End Class
End Namespace