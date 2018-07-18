Imports System.Runtime.CompilerServices
Imports System.Text

Imports IVisualBasicCode.CodeConverter.Util
Imports IVisualBasicCode.CodeConverter.VB
Imports IVisualBasicCode.CodeConverter.VB.CSharpConverter

Imports Microsoft.CodeAnalysis

Imports CS = Microsoft.CodeAnalysis.CSharp
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory

Public Module TriviaListSupport
    Const vbDoubleQuote As String = """"
    Private ReadOnly mSemanticModel As SemanticModel
    Private ReadOnly mTargetDocument As Document
    Public VB_EOLTrivia As SyntaxTrivia = SyntaxFactory.EndOfLineTrivia(vbCrLf)

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="FullString"></param>
    ''' <returns></returns>
    ''' <remarks>Added by PC</remarks>
    Private Function ConvertSourceTextToTriviaList(FullString As String, Optional LeadingComment As String = "") As SyntaxTriviaList
        Dim NewTrivia As New SyntaxTriviaList
        If LeadingComment.IsNotEmptyNullOrWhitespace Then
            NewTrivia = NewTrivia.Add(SyntaxFactory.CommentTrivia($"' {LeadingComment}"))
            NewTrivia = NewTrivia.Add(VB_EOLTrivia)
        End If
        Dim sb As New StringBuilder
        For Each chr As String In FullString
            If chr.IsNewLine Then
                If sb.Length > 0 Then
                    NewTrivia = NewTrivia.Add(SyntaxFactory.CommentTrivia($"' {sb.ToString}"))
                    NewTrivia = NewTrivia.Add(VB_EOLTrivia)
                    sb.Clear()
                End If
            ElseIf chr = vbTab Then
                sb.Append("    ")
            Else
                sb.Append(chr)
            End If
        Next
        If sb.Length > 0 Then
            NewTrivia = NewTrivia.Add(SyntaxFactory.CommentTrivia($"' {sb.ToString}"))
            NewTrivia = NewTrivia.Add(VB_EOLTrivia)
        End If

        Return NewTrivia
    End Function

    Private Function ConvertDocumentCommentExternalTrivia(t As SyntaxTrivia) As SyntaxTrivia
        If t.IsKind(CS.SyntaxKind.DocumentationCommentExteriorTrivia) Then
            Return SyntaxFactory.DocumentationCommentExteriorTrivia(t.ToString.Replace("///", "'''"))
        End If

        Throw UnreachableException()
    End Function

    <Extension>
    Public Function ConvertNodeToMultipleLineComment(Of T As SyntaxNode)(node As T, Optional LeadingComment As String = "") As IEnumerable(Of SyntaxTrivia)
        If node Is Nothing Then
            Return New SyntaxTriviaList
        End If
        Return ConvertSourceTextToTriviaList(node.ToFullString, LeadingComment)
    End Function

    Public Function CreateVBDocumentCommentFromCSharpComment(singleLineDocumentationComment As CS.Syntax.DocumentationCommentTriviaSyntax) As VisualBasic.Syntax.DocumentationCommentTriviaSyntax

        Dim walker As New XMLVisitor(mSemanticModel, New NodesVisitor(mSemanticModel, mTargetDocument))
        walker.Visit(singleLineDocumentationComment)

        Dim xmlNodes As New List(Of VisualBasic.Syntax.XmlNodeSyntax)
        For Each node As CS.Syntax.XmlNodeSyntax In singleLineDocumentationComment.Content
            Try
                xmlNodes.Add(DirectCast(node.Accept(walker), VisualBasic.Syntax.XmlNodeSyntax))
            Catch ex As Exception
                Stop
                Throw
            End Try
        Next
        Return SyntaxFactory.DocumentationComment(xmlNodes.ToArray)
    End Function

    Public Function TranslateTokenList(ChildTokens As IEnumerable(Of SyntaxToken)) As SyntaxTokenList
        Dim NewTokenList As New SyntaxTokenList
        For Each token As SyntaxToken In ChildTokens
            Dim NewLeadingTriviaList As New SyntaxTriviaList
            Dim NewTrailingTriviaList As New SyntaxTriviaList
            If token.HasLeadingTrivia Then
                For Each t As SyntaxTrivia In token.LeadingTrivia
                    Dim FullString As String = t.ToFullString
                    If FullString.Length > 3 AndAlso FullString.StartsWith("///") Then
                        Stop
                    End If
                    NewLeadingTriviaList = NewLeadingTriviaList.Add(ConvertDocumentCommentExternalTrivia(t))
                Next
            End If
            If token.HasTrailingTrivia Then
                Throw UnexpectedValue($"XMLToken '{token.ToFullString}' should not have trailing Trivia")
            End If
            Select Case token.RawKind
                Case CS.SyntaxKind.XmlTextLiteralToken
                    NewTokenList = NewTokenList.Add(SyntaxFactory.XmlTextLiteralToken(leadingTrivia:=NewLeadingTriviaList, text:=token.Text, value:=token.ValueText, trailingTrivia:=NewTrailingTriviaList))
                Case CS.SyntaxKind.XmlTextLiteralNewLineToken
                    NewTokenList = NewTokenList.Add(SyntaxFactory.XmlTextNewLine(text:=vbCrLf, value:=vbCrLf, leading:=NewLeadingTriviaList, trailing:=NewTrailingTriviaList))
                Case CS.SyntaxKind.XmlEntityLiteralToken
                    NewTokenList = NewTokenList.Add(SyntaxFactory.XmlEntityLiteralToken(leadingTrivia:=NewLeadingTriviaList, text:=token.Text, value:=token.Text.ToString, trailingTrivia:=NewTrailingTriviaList))
                Case Else
                    Stop
            End Select
        Next
        Return NewTokenList
    End Function

End Module