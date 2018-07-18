Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports NameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.NameSyntax
Imports SimpleNameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.SimpleNameSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports TypeArgumentListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeArgumentListSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)
            Private Function WrapTypedNameIfNecessary(ByVal name As ExpressionSyntax, ByVal originalName As CSS.ExpressionSyntax) As VisualBasicSyntaxNode
                If TypeOf originalName Is CSS.InvocationExpressionSyntax Then
                    Return name
                End If
                Dim OriginalNameParent As SyntaxNode = originalName.Parent
                If TypeOf OriginalNameParent Is CSS.ArgumentSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.AttributeSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.ConstantPatternSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.InvocationExpressionSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.MemberAccessExpressionSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.MemberBindingExpressionSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.NameSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.ObjectCreationExpressionSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.QualifiedNameSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.TupleElementSyntax OrElse
                   TypeOf OriginalNameParent Is CSS.TypeArgumentListSyntax Then
                    Return name
                End If
                'If TypeOf OriginalNameParent Is CSS.InvocationExpressionSyntax Then
                '    Return name
                'End If
                Dim OriginalNameParentArgumentList As SyntaxNode = OriginalNameParent.Parent
                If OriginalNameParentArgumentList IsNot Nothing AndAlso TypeOf OriginalNameParentArgumentList Is CSS.ArgumentListSyntax Then
                    Dim OriginalNameParentArgumentListParentInvocationExpression As SyntaxNode = OriginalNameParentArgumentList.Parent
                    If OriginalNameParentArgumentListParentInvocationExpression IsNot Nothing AndAlso TypeOf OriginalNameParentArgumentListParentInvocationExpression Is CSS.InvocationExpressionSyntax Then
                        Dim expression As CSS.InvocationExpressionSyntax = DirectCast(OriginalNameParentArgumentListParentInvocationExpression, CSS.InvocationExpressionSyntax)
                        If TypeOf expression?.Expression Is CSS.IdentifierNameSyntax Then
                            If DirectCast(expression.Expression, CSS.IdentifierNameSyntax).Identifier.ToString = "nameof" Then
                                Return name
                            End If
                        End If
                    End If
                End If

                Dim symbolInfo As SymbolInfo = ModelExtensions.GetSymbolInfo(mSemanticModel, originalName)
                Dim symbol As ISymbol = If(symbolInfo.Symbol, symbolInfo.CandidateSymbols.FirstOrDefault())
                If symbol?.Kind = SymbolKind.Method Then
                    Return SyntaxFactory.AddressOfExpression(name)
                End If
                Return name
            End Function
            Public Overrides Function VisitAliasQualifiedName(ByVal node As CSS.AliasQualifiedNameSyntax) As VisualBasicSyntaxNode
                Return WrapTypedNameIfNecessary(SyntaxFactory.QualifiedName(DirectCast(node.[Alias].Accept(Me), NameSyntax), DirectCast(node.Name.Accept(Me), SimpleNameSyntax)), node).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitGenericName(ByVal node As CSS.GenericNameSyntax) As VisualBasicSyntaxNode
                Dim TypeArgumentList As TypeArgumentListSyntax = DirectCast(node.TypeArgumentList.Accept(Me), TypeArgumentListSyntax)
                Return WrapTypedNameIfNecessary(SyntaxFactory.GenericName(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False), TypeArgumentList), node)
            End Function

            Public Overrides Function VisitIdentifierName(ByVal node As CSS.IdentifierNameSyntax) As VisualBasicSyntaxNode
                Dim OriginalNameParent As SyntaxNode = node.Parent
                If TypeOf OriginalNameParent Is CSS.MemberAccessExpressionSyntax OrElse
                    TypeOf OriginalNameParent Is CSS.MemberBindingExpressionSyntax OrElse
                    OriginalNameParent.IsKind(CSharp.SyntaxKind.AsExpression) OrElse
                    (TypeOf OriginalNameParent Is CSS.NameEqualsSyntax AndAlso TypeOf OriginalNameParent.Parent Is CSS.AnonymousObjectMemberDeclaratorSyntax) Then
                    If TypeOf OriginalNameParent Is CSS.MemberAccessExpressionSyntax Then
                        Dim ParentAsMemberAccessExpression As CSS.MemberAccessExpressionSyntax = DirectCast(OriginalNameParent, CSS.MemberAccessExpressionSyntax)
                        If ParentAsMemberAccessExpression.Expression.IsKind(CS.SyntaxKind.IdentifierName) Then
                            Dim IdentifierExpression As CSS.IdentifierNameSyntax = DirectCast(ParentAsMemberAccessExpression.Expression, IdentifierNameSyntax)
                            If IdentifierExpression.Identifier.ToString = node.Identifier.ToString Then
                                Return WrapTypedNameIfNecessary(SyntaxFactory.IdentifierName(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False)), node)
                            End If
                        End If
                    End If
                    Return WrapTypedNameIfNecessary(SyntaxFactory.IdentifierName(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=True)), node)
                End If
                If TypeOf OriginalNameParent Is CSS.DeclarationExpressionSyntax Then
                    If node.ToString = "var" Then
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ObjectKeyword))
                    End If
                    Return ConvertToType(node.ToString)
                End If
                If TypeOf OriginalNameParent Is CSS.UsingDirectiveSyntax OrElse
                    OriginalNameParent.IsKind(CS.SyntaxKind.TypeArgumentList, CS.SyntaxKind.SimpleBaseType) Then
                    Return WrapTypedNameIfNecessary(SyntaxFactory.IdentifierName(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=True, IsTypeName:=True)), node)
                End If
                ' The trivial on node reflects the wrong place on the file as order is switched so don't convert trivia here
                Return WrapTypedNameIfNecessary(SyntaxFactory.IdentifierName(GenerateSafeVBToken(node.Identifier, OriginalNameParent.IsKind(CS.SyntaxKind.QualifiedName))), node)
            End Function
            Public Overrides Function VisitQualifiedName(ByVal node As CSS.QualifiedNameSyntax) As VisualBasicSyntaxNode
                Return WrapTypedNameIfNecessary(SyntaxFactory.QualifiedName(DirectCast(node.Left.Accept(Me), NameSyntax), DirectCast(node.Right.Accept(Me), SimpleNameSyntax)), node)
            End Function
        End Class
    End Class
End Namespace