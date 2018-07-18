Imports Microsoft.CodeAnalysis

Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax

Public Module StatementFinder

    Public Function GetStatementwithIssues(node As CS.CSharpSyntaxNode, Optional ByRef IsField As Boolean = False) As CS.CSharpSyntaxNode

        Dim StatementWithIssue As CS.CSharpSyntaxNode = node.FirstAncestorOrSelf(Of CSS.StatementSyntax)
        If StatementWithIssue IsNot Nothing Then
            Dim StatementWithIssueParent As SyntaxNode = StatementWithIssue.Parent
            While StatementWithIssueParent.IsKind(CS.SyntaxKind.ElseClause)
                StatementWithIssue = StatementWithIssue.Parent.FirstAncestorOrSelf(Of CSS.StatementSyntax)
                StatementWithIssueParent = StatementWithIssue.Parent
            End While
            Return StatementWithIssue
        End If

        StatementWithIssue = node.FirstAncestorOrSelf(Of CSS.MethodDeclarationSyntax)
        If StatementWithIssue IsNot Nothing Then
            Return StatementWithIssue
        End If
        StatementWithIssue = node.FirstAncestorOrSelf(Of CSS.ConstructorDeclarationSyntax)
        If StatementWithIssue IsNot Nothing Then
            Return StatementWithIssue
        End If
        StatementWithIssue = node.FirstAncestorOrSelf(Of CSS.PropertyDeclarationSyntax)
        If StatementWithIssue IsNot Nothing Then
            Return StatementWithIssue
        End If
        Dim FieldWithIssue As CSS.FieldDeclarationSyntax = node.FirstAncestorOrSelf(Of CSS.FieldDeclarationSyntax)
        If FieldWithIssue IsNot Nothing Then
            IsField = True
            Return FieldWithIssue
        End If
        Throw UnexpectedValue($"Can't find parent 'statement' of {node.ToString}")
    End Function

End Module