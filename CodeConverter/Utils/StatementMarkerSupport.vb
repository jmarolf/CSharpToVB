Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax

Public Module StatementMarker
    Private FieldDictionary As New Dictionary(Of CS.CSharpSyntaxNode, Integer)
    Private FieldSupportTupleList As New List(Of (Index As Integer, Trivia As SyntaxTriviaList))
    Private NextIndex As Integer = 0
    Private StatementDictionary As New Dictionary(Of CS.CSharpSyntaxNode, Integer)
    Private StatementSupportTupleList As New List(Of (Index As Integer, Statement As VisualBasic.VisualBasicSyntaxNode, RemoveStatement As Boolean))

    ''' <summary>
    ''' Add a marker so we can add a statement higher up in the result tree
    ''' </summary>
    ''' <param name="Node">The C# statement above which we can add the statements we need</param>
    ''' <param name="Statement">The Statement we want to add above the Node</param>
    ''' <param name="RemoveStatement">If True we will replace the Node Statement with the new statement(s)
    ''' otherwise we just add the statement BEFORE the node</param>
    ''' <param name="AllowDuplicates">True if we can put do multiple replacements</param>
    <Extension>
    Public Sub AddMarker(Node As CS.CSharpSyntaxNode, Statement As VisualBasic.VisualBasicSyntaxNode, RemoveStatement As Boolean, AllowDuplicates As Boolean)
        If StatementDictionary.ContainsKey(Node) Then
            If Not AllowDuplicates Then
                Return
            End If
        Else
            StatementDictionary.Add(Node, NextIndex)
            NextIndex += 1
        End If
        Dim Index As Integer = StatementDictionary(Node)
        Dim Found As Boolean = False
        For Each t As (Index As Integer, Statement As VisualBasic.VisualBasicSyntaxNode, RemoveStatement As Boolean) In StatementSupportTupleList
            If t.Index = Index AndAlso t.Statement.ToString = Statement.ToString AndAlso t.RemoveStatement = RemoveStatement Then
                Found = True
            End If
        Next
        If Found Then
            Return
        End If
        StatementSupportTupleList.Add((Index, Statement, RemoveStatement))
    End Sub

    <Extension>
    Public Sub AddMarker(Node As CSS.FieldDeclarationSyntax, Trivia As SyntaxTriviaList, AllowDuplicates As Boolean)
        If FieldDictionary.ContainsKey(Node) Then
            If Not AllowDuplicates Then
                Return
            End If
        Else
            FieldDictionary.Add(Node, NextIndex)
            NextIndex += 1
        End If
        Dim Index As Integer = FieldDictionary(Node)
        FieldSupportTupleList.Add((Index, Trivia))
    End Sub

    Public Function AddSpecialCommentToField(node As CSS.FieldDeclarationSyntax, FieldDeclaration As FieldDeclarationSyntax) As FieldDeclarationSyntax
        If Not FieldDictionary.ContainsKey(node) Then
            Return FieldDeclaration
        End If
        Dim LeadingTrivia As New List(Of SyntaxTrivia)
        Dim Index As Integer = FieldDictionary(node)
        For Each Field As (Index As Integer, Trivia As SyntaxTriviaList) In FieldSupportTupleList
            If Field.Index = Index Then
                LeadingTrivia.AddRange(Field.Trivia)
            End If
        Next
        FieldDictionary.Remove(node)
        If FieldDictionary.Count = 0 Then
            FieldSupportTupleList.Clear()
        End If
        LeadingTrivia.AddRange(FieldDeclaration.GetLeadingTrivia)
        Return FieldDeclaration.WithLeadingTrivia(LeadingTrivia)
    End Function

    Public Function GetMarkerErrorMessage() As String
        Dim builder As New System.Text.StringBuilder()
        builder.Append($" Marker Error StatementDictionary.Count = {StatementDictionary.Count}, FieldDictionary.Count > {FieldDictionary.Count}{vbCrLf}")
        For Each statement As CS.CSharpSyntaxNode In StatementDictionary.Keys
            builder.Append(statement.ToFullString)
        Next
        For Each statement As CS.CSharpSyntaxNode In FieldDictionary.Keys
            builder.Append(statement.ToFullString)
        Next
        Return builder.ToString()
    End Function

    Public Function MarkerError() As Boolean
        Return StatementDictionary.Count > 0 Or FieldDictionary.Count > 0
    End Function

    Public Function ReplaceStatementWithMarkedStatement(node As CS.CSharpSyntaxNode, Statements As SyntaxList(Of StatementSyntax)) As SyntaxList(Of StatementSyntax)
        Dim NewNodesList As New SyntaxList(Of StatementSyntax)
        Dim RemoveStatement As Boolean = False
        If Not StatementDictionary.ContainsKey(node) Then
            Return Statements
        End If
        Dim Index As Integer = StatementDictionary(node)

        For Each StatementTuple As (Index As Integer, Statement As StatementSyntax, RemoveStatement As Boolean) In StatementSupportTupleList
            If StatementTuple.Index = Index Then
                NewNodesList = NewNodesList.Add(StatementTuple.Statement)
                RemoveStatement = RemoveStatement Or StatementTuple.RemoveStatement
            End If
        Next
        StatementDictionary.Remove(node)
        If StatementDictionary.Count = 0 Then
            StatementSupportTupleList.Clear()
        End If
        If Not RemoveStatement Then
            NewNodesList = NewNodesList.AddRange(Statements)
        End If
        Return NewNodesList
    End Function

End Module