Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Module ListExtensions
    <Extension>
    Public Function ContainsName(ImportList As List(Of ImportsStatementSyntax), ImportName As String) As Boolean
        For Each ImportToCheck As ImportsStatementSyntax In ImportList
            For Each Clause As ImportsClauseSyntax In ImportToCheck.ImportsClauses
                If Clause.ToString = ImportName Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

End Module
