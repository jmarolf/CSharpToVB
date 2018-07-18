Option Explicit On
Option Infer Off
Option Strict On

Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Imports CS = Microsoft.CodeAnalysis.CSharp

Namespace IVisualBasicCode.CodeConverter.Util
    Friend Module SyntaxTriviaExtensions

        <Extension>
        Private Function IsCommentOrDirectiveTrivia(t As SyntaxTrivia) As Boolean
            If t.IsCommentTrivia Then
                Return True
            End If
            If t.IsDirective Then
                Return True
            End If
            Return False
        End Function

        <Extension>
        Private Function IsCommentTrivia(t As SyntaxTrivia) As Boolean
            If t.IsKind(SyntaxKind.CommentTrivia) Then
                Return True
            End If
            If t.MatchesKind(CS.SyntaxKind.SingleLineCommentTrivia, CS.SyntaxKind.MultiLineCommentTrivia) Then
                Return True
            End If
            Return False
        End Function

        <Extension()>
        Public Function [Do](Of T)(ByVal source As IEnumerable(Of T), ByVal action As Action(Of T)) As IEnumerable(Of T)
            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If

            If action Is Nothing Then
                Throw New ArgumentNullException(NameOf(action))
            End If

            ' perf optimization. try to not use enumerator if possible
            Dim list As IList(Of T) = TryCast(source, IList(Of T))
            If list IsNot Nothing Then
                Dim i As Integer = 0
                Dim count As Integer = list.Count
                Do While i < count
                    action(list(i))
                    i += 1
                Loop
            Else
                For Each value As T In source
                    action(value)
                Next value
            End If

            Return source
        End Function

        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function AsString(ByVal trivia As IEnumerable(Of SyntaxTrivia)) As String
            If trivia.Any() Then
                Dim sb As New StringBuilder()
                trivia.Select(Function(t As SyntaxTrivia) t.ToFullString()).Do(Function(s As String) sb.Append(s))
                Return sb.ToString()
            Else
                Return String.Empty
            End If
        End Function
        <Extension>
        Public Function CollectAndConvertCommentTrivia(TriviaList As SyntaxTriviaList) As List(Of SyntaxTrivia)
            Dim CommentTrivia As New List(Of SyntaxTrivia)
            For Each t As SyntaxTrivia In TriviaList
                If t.IsKind(SyntaxKind.CommentTrivia) Then
                    CommentTrivia.Add(t)
                    CommentTrivia.Add(SyntaxFactory.ElasticSpace)
                ElseIf t.IsComment Then
                    CommentTrivia.AddRange(ConvertTrivia({t}))
                    CommentTrivia.Add(SyntaxFactory.ElasticSpace)

                End If
            Next
            Return CommentTrivia
        End Function

        <Extension>
        Public Function ContainsCommentOrDirectiveTrivia(TriviaList As List(Of SyntaxTrivia)) As Boolean
            If TriviaList.Count = 0 Then
                Return False
            End If
            For Each t As SyntaxTrivia In TriviaList
                If t.IsWhitespaceOrEndOfLine Then
                    Continue For
                End If
                If t.Kind = SyntaxKind.None Then
                    Continue For
                End If
                If t.IsCommentOrDirectiveTrivia Then
                    Return True
                End If
                Stop
            Next
            Return False
        End Function

        ''' <summary>
        ''' Syntax Trivia in any Language
        ''' </summary>
        ''' <param name="TriviaList"></param>
        ''' <returns>True if any Trivia is a Comment or a Directive</returns>
        <Extension>
        Public Function ContainsCommentOrDirectiveTrivia(TriviaList As SyntaxTriviaList) As Boolean
            If TriviaList.Count = 0 Then Return False
            For Each t As SyntaxTrivia In TriviaList
                If t.IsWhitespaceOrEndOfLine Then
                    Continue For
                End If
                If t.RawKind = 0 Then
                    Continue For
                End If
                If t.IsCommentOrDirectiveTrivia Then
                    Return True
                End If
                Stop
            Next
            Return False
        End Function

        ''' <summary>
        '''
        ''' </summary>
        ''' <param name="node"></param>
        ''' <returns>True if any Trivia is a Comment or a Directive</returns>
        <Extension>
        Public Function ContainsCommentOrDirectiveTrivia(node As VisualBasicSyntaxNode) As Boolean
            Dim CurrentToken As SyntaxToken = node.GetFirstToken
            While CurrentToken <> Nothing
                If CurrentToken.LeadingTrivia.ContainsCommentOrDirectiveTrivia OrElse CurrentToken.TrailingTrivia.ContainsCommentOrDirectiveTrivia Then
                    Return True
                End If
                CurrentToken = CurrentToken.GetNextToken
            End While

            Return False
        End Function

        <Extension>
        Public Function ContainsEOLTrivia(node As VisualBasicSyntaxNode) As Boolean
            If Not node.HasTrailingTrivia Then
                Return False
            End If
            Dim TriviaList As SyntaxTriviaList = node.GetTrailingTrivia
            For Each t As SyntaxTrivia In TriviaList
                If t.IsEndOfLine Then
                    Return True
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function ContainsEOLTrivia(Token As SyntaxToken) As Boolean
            If Not Token.HasTrailingTrivia Then
                Return False
            End If
            Dim TriviaList As SyntaxTriviaList = Token.TrailingTrivia
            For Each t As SyntaxTrivia In TriviaList
                If t.IsEndOfLine Then
                    Return True
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function ContainsEOLTrivia(TriviaList As SyntaxTriviaList) As Boolean
            For Each t As SyntaxTrivia In TriviaList
                If t.IsEndOfLine Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function ExtractComments(ListOfTrivia As IEnumerable(Of SyntaxTrivia), Leading As Boolean) As IEnumerable(Of SyntaxTrivia)
            Dim CommentTrivia As New List(Of SyntaxTrivia)
            Dim FoundEOL As Boolean = Leading
            For Each t As SyntaxTrivia In ListOfTrivia
                Select Case t.RawKind
                    Case SyntaxKind.CommentTrivia,
                            SyntaxKind.DocumentationCommentExteriorTrivia,
                            SyntaxKind.EmptyStatement,
                            SyntaxKind.BadDirectiveTrivia,
                            SyntaxKind.ConstDirectiveTrivia,
                            SyntaxKind.DisabledTextTrivia,
                            SyntaxKind.DisableWarningDirectiveTrivia,
                            SyntaxKind.ElseDirectiveTrivia,
                            SyntaxKind.ElseIfDirectiveTrivia,
                            SyntaxKind.EnableWarningDirectiveTrivia,
                            SyntaxKind.EndExternalSourceDirectiveTrivia,
                            SyntaxKind.EndIfDirectiveTrivia,
                            SyntaxKind.EndRegionDirectiveTrivia,
                            SyntaxKind.ExternalChecksumDirectiveTrivia,
                            SyntaxKind.ExternalSourceDirectiveTrivia,
                            SyntaxKind.IfDirectiveTrivia,
                            SyntaxKind.ReferenceDirectiveTrivia,
                            SyntaxKind.RegionDirectiveTrivia
                        If FoundEOL Then
                            CommentTrivia.Add(VB_EOLTrivia)
                        End If
                        CommentTrivia.Add(t)
                        FoundEOL = True
                    Case SyntaxKind.EndOfLineTrivia
                        FoundEOL = True
                    Case SyntaxKind.WhitespaceTrivia
                        'ignore
                    Case Else
                        Stop
                End Select
            Next
            If FoundEOL Then
                CommentTrivia.Add(VB_EOLTrivia)
            End If
            If CommentTrivia.Count = 1 Then
                CommentTrivia.Clear()
            End If
            Return CommentTrivia
        End Function

        <Extension>
        Public Function FullWidth(ByVal trivia As SyntaxTrivia) As Integer
            Return trivia.FullSpan.Length
        End Function

        <Extension>
        Public Function GetCommentText(ByVal trivia As SyntaxTrivia) As String
            Dim commentText As String = trivia.ToString()
            If trivia.IsKind(CS.SyntaxKind.SingleLineCommentTrivia) Then
                If commentText.StartsWith("//") Then
                    commentText = commentText.Substring(2)
                End If
                Return commentText.TrimStart(Nothing)
            ElseIf trivia.IsKind(VisualBasic.SyntaxKind.CommentTrivia) Then
                If commentText.StartsWith("'") Then
                    commentText = commentText.Substring(1)
                End If
                Return commentText
            ElseIf trivia.IsKind(CS.SyntaxKind.MultiLineCommentTrivia) Then
                Dim textBuilder As New StringBuilder()

                If commentText.EndsWith("*/") Then
                    commentText = commentText.Substring(0, commentText.Length - 2)
                End If

                If commentText.StartsWith("/*") Then
                    commentText = commentText.Substring(2)
                End If

                commentText = commentText.Trim()

                Dim newLine As String = Environment.NewLine
                Dim lines As String() = commentText.Split({newLine}, StringSplitOptions.None)
                For Each line As String In lines
                    Dim trimmedLine As String = line.Trim()

                    ' Note: we trim leading '*' characters in multi-line comments.
                    ' If the '*' was intentional, sorry, it's gone.
                    If trimmedLine.StartsWith("*") Then
                        trimmedLine = trimmedLine.TrimStart("*"c)
                        trimmedLine = trimmedLine.TrimStart(Nothing)
                    End If

                    textBuilder.AppendLine(trimmedLine)
                Next line

                ' remove last line break
                textBuilder.Remove(textBuilder.Length - newLine.Length, newLine.Length)
                Return textBuilder.ToString()
            Else
                Throw New InvalidOperationException()
            End If
        End Function

        <Extension>
        Public Function GetPreviousTrivia(ByVal trivia As SyntaxTrivia, ByVal syntaxTree As SyntaxTree, ByVal cancellationToken As CancellationToken, Optional ByVal FindInsideTrivia As Boolean = False) As SyntaxTrivia
            Dim span As Text.TextSpan = trivia.FullSpan
            If span.Start = 0 Then
                Return Nothing
            End If

            Return syntaxTree.GetRoot(cancellationToken).FindTrivia(position:=span.Start - 1, findInsideTrivia:=FindInsideTrivia)
        End Function

        <Extension>
        Public Function IsComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsSingleLineComment OrElse trivia.IsMultiLineComment
        End Function

        <Extension>
        Public Function IsCompleteMultiLineComment(ByVal trivia As SyntaxTrivia) As Boolean
            If trivia.IsKind(CS.SyntaxKind.MultiLineCommentTrivia) Then
                Return False
            End If

            Dim text As String = trivia.ToFullString()
            Return text.Length >= 4 AndAlso text(text.Length - 1) = "/"c AndAlso text(text.Length - 2) = "*"c
        End Function

        <Extension>
        Public Function IsDocComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsSingleLineDocComment() OrElse trivia.IsMultiLineDocComment()
        End Function

        <Extension>
        Public Function IsDocumentationCommentTrivia(t As SyntaxTrivia) As Boolean
            If t.IsKind(VisualBasic.SyntaxKind.DocumentationCommentTrivia) Then
                Return True
            End If
            Return False
        End Function

        <Extension>
        Public Function IsElastic(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.HasAnnotation(SyntaxAnnotation.ElasticAnnotation)
        End Function

        <Extension>
        Public Function IsEndOfLine(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(CS.SyntaxKind.EndOfLineTrivia) OrElse
                trivia.IsKind(VisualBasic.SyntaxKind.EndOfLineTrivia)
        End Function

        <Extension>
        Public Function IsMultiLineComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(CS.SyntaxKind.MultiLineCommentTrivia)
        End Function

        <Extension>
        Public Function IsMultiLineDocComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(CS.SyntaxKind.MultiLineDocumentationCommentTrivia)
        End Function

        <Extension>
        Public Function IsRegularComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsSingleLineComment() OrElse trivia.IsMultiLineComment()
        End Function

        <Extension>
        Public Function IsRegularOrDocComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsSingleLineComment() OrElse trivia.IsMultiLineComment() OrElse trivia.IsDocComment()
        End Function

        <Extension>
        Public Function IsSingleLineComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(CS.SyntaxKind.SingleLineCommentTrivia) OrElse trivia.IsKind(VisualBasic.SyntaxKind.CommentTrivia)
        End Function

        <Extension>
        Public Function IsSingleLineDocComment(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(CS.SyntaxKind.SingleLineDocumentationCommentTrivia)
        End Function

        <Extension>
        Public Function IsSkippedTokensTrivia(t As SyntaxTrivia) As Boolean
            If t.IsKind(VisualBasic.SyntaxKind.DocumentationCommentTrivia) Then
                Return True
            End If
            Return False
        End Function

        <Extension>
        Public Function IsWhitespaceOrEndOfLine(ByVal trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(CS.SyntaxKind.WhitespaceTrivia) OrElse
                trivia.IsKind(CS.SyntaxKind.EndOfLineTrivia) OrElse
                trivia.IsKind(SyntaxKind.EndOfLineTrivia) OrElse
                trivia.IsKind(SyntaxKind.WhitespaceTrivia)
        End Function

        ' VB
        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function MatchesKind(ByVal trivia As SyntaxTrivia, ByVal kind As SyntaxKind) As Boolean
            Return trivia.IsKind(kind)
        End Function

        ' VB
        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function MatchesKind(ByVal trivia As SyntaxTrivia, ByVal kind1 As SyntaxKind, ByVal kind2 As SyntaxKind) As Boolean
            Return trivia.IsKind(kind1) OrElse trivia.IsKind(kind2)
        End Function

        ' VB
        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function MatchesKind(ByVal trivia As SyntaxTrivia, ParamArray ByVal kinds() As SyntaxKind) As Boolean
            For Each kind As SyntaxKind In kinds
                If trivia.IsKind(kind) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ' C#
        <Extension>
        Public Function MatchesKind(ByVal trivia As SyntaxTrivia, ByVal kind As CS.SyntaxKind) As Boolean
            Return trivia.IsKind(kind)
        End Function

        ' C#
        <Extension>
        Public Function MatchesKind(ByVal trivia As SyntaxTrivia, ByVal kind1 As CS.SyntaxKind, ByVal kind2 As CS.SyntaxKind) As Boolean
            Return trivia.IsKind(kind1) OrElse trivia.IsKind(kind2)
        End Function

        ' C#
        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function MatchesKind(ByVal trivia As SyntaxTrivia, ParamArray ByVal kinds() As CS.SyntaxKind) As Boolean
            For Each kind As CS.SyntaxKind In kinds
                If trivia.IsKind(kind) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Move internal trivia to end
        ''' </summary>
        ''' <param name="node"></param>
        ''' <returns></returns>
        <Extension>
        Public Function RestructureCommentTrivia(Of T As SyntaxNode)(node As T) As T
            Dim CurrentToken As SyntaxToken = node.GetFirstToken
            Dim CommentTrailingTrivia As New List(Of SyntaxTrivia)
            While CurrentToken <> Nothing
                If CurrentToken.LeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                    Dim NewLeadingTrivia As New SyntaxTriviaList
                    For Each trivia As SyntaxTrivia In CurrentToken.LeadingTrivia
                        If trivia.IsKind(SyntaxKind.CommentTrivia) Then
                            CommentTrailingTrivia.Add(trivia)
                        Else
                            If Not trivia.IsKind(SyntaxKind.EndOfLineTrivia) Then
                                NewLeadingTrivia = NewLeadingTrivia.Add(trivia)
                            End If
                        End If
                    Next
                    node = node.ReplaceToken(CurrentToken, CurrentToken.WithLeadingTrivia(NewLeadingTrivia))
                End If
                If CurrentToken.TrailingTrivia.ContainsCommentOrDirectiveTrivia Then
                    Dim NewTrailingTrivia As New SyntaxTriviaList
                    For Each trivia As SyntaxTrivia In CurrentToken.TrailingTrivia
                        If trivia.IsKind(SyntaxKind.CommentTrivia) Then
                            CommentTrailingTrivia.Add(trivia)
                        Else
                            If Not trivia.IsKind(SyntaxKind.EndOfLineTrivia) Then
                                NewTrailingTrivia = NewTrailingTrivia.Add(trivia)
                            End If
                        End If
                    Next
                    node = node.ReplaceToken(CurrentToken, CurrentToken.WithTrailingTrivia(NewTrailingTrivia))
                End If
                CurrentToken = CurrentToken.GetNextToken
            End While

            Return node.WithTrailingTrivia(CommentTrailingTrivia)
        End Function

        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function ToSyntaxTriviaList(l As List(Of SyntaxTrivia)) As SyntaxTriviaList
            Dim NewSyntaxTriviaList As New SyntaxTriviaList
            Return NewSyntaxTriviaList.AddRange(l)
        End Function

    End Module
End Namespace