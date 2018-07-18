Imports IVisualBasicCode.CodeConverter.Util

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Public Module RestuructureSeparatedLists
    Public IgnoredIfDepth As Integer = 0

    Private Function DirectiveNotAllowedHere(Trivia As SyntaxTrivia) As SyntaxTrivia
        Return SyntaxFactory.CommentTrivia($" ' VB does not allow directives here, original statement: {Trivia.ToFullString.Replace("  ", " ").WithoutNewLines}")
    End Function

    Private Function RestructureTrivia(Node As VisualBasicSyntaxNode, ByRef Separator As SyntaxToken, ByRef MovedItemLeadingTrivia As List(Of SyntaxTrivia), ByRef MovedTrailingTrivia As List(Of SyntaxTrivia)) As VisualBasicSyntaxNode
        Dim ItemLeadingTriviaList As New List(Of SyntaxTrivia)
        ItemLeadingTriviaList.AddRange(MovedItemLeadingTrivia)
        MovedItemLeadingTrivia.Clear()

        Dim ItemTrailingTriviaList As New List(Of SyntaxTrivia)

        Dim SeparatorTrailingTrivia As New List(Of SyntaxTrivia)
        SeparatorTrailingTrivia.AddRange(MovedTrailingTrivia)
        MovedTrailingTrivia.Clear()

        Dim FoundSeparatorTrailingEOL As Boolean = False
        Dim FoundItemTrailingEOL As Boolean = False
        Dim IsMultiLineLambda As Boolean
        If Node.IsKind(SyntaxKind.SimpleArgument) Then
            Dim Expression As Syntax.ExpressionSyntax = DirectCast(Node, Syntax.ArgumentSyntax).GetExpression
            IsMultiLineLambda = If(Expression IsNot Nothing AndAlso Expression.IsKind(SyntaxKind.MultiLineFunctionLambdaExpression, SyntaxKind.MultiLineSubLambdaExpression), True, False)
        Else
            IsMultiLineLambda = False
        End If
        If Node.HasLeadingTrivia Then
            For Each t As SyntaxTrivia In Node.GetLeadingTrivia
                Select Case t.Kind
                    Case SyntaxKind.CommentTrivia
                        SeparatorTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                        SeparatorTrailingTrivia.Add(t)
                        FoundSeparatorTrailingEOL = True
                    Case SyntaxKind.EndOfLineTrivia
                        If IsMultiLineLambda Then
                            MovedItemLeadingTrivia.Add(VB_EOLTrivia)
                        Else
                            ItemLeadingTriviaList.Clear()
                        End If
                    Case SyntaxKind.WhitespaceTrivia
                        ItemLeadingTriviaList.Add(t)
                    Case SyntaxKind.IfDirectiveTrivia
                        ' Not allowed here so notify and comment out
                        SeparatorTrailingTrivia.Add(DirectiveNotAllowedHere(t))
                    Case Else
                        Stop
                End Select
            Next
        End If
        If Node.HasTrailingTrivia Then
            For Each t As SyntaxTrivia In Node.GetTrailingTrivia
                Select Case t.Kind
                    Case SyntaxKind.CommentTrivia
                        SeparatorTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                        SeparatorTrailingTrivia.Add(t)
                        FoundSeparatorTrailingEOL = True
                        FoundItemTrailingEOL = False
                    Case SyntaxKind.EndOfLineTrivia
                        FoundItemTrailingEOL = True
                    Case SyntaxKind.WhitespaceTrivia
                        ItemTrailingTriviaList.Add(t)
                    Case SyntaxKind.SkippedTokensTrivia
                    Case SyntaxKind.EnableWarningDirectiveTrivia, SyntaxKind.DisableWarningDirectiveTrivia
                        SeparatorTrailingTrivia.Add(DirectiveNotAllowedHere(t))
                    Case Else
                        Stop
                End Select
            Next
        End If
        If Separator = Nothing Then
            If SeparatorTrailingTrivia.Count > 0 Then
                If FoundItemTrailingEOL Then
                    ItemTrailingTriviaList.AddRange(SeparatorTrailingTrivia)
                    ItemTrailingTriviaList.Add(VB_EOLTrivia)
                Else
                    If SeparatorTrailingTrivia.ContainsCommentOrDirectiveTrivia Then
                        If Not FoundSeparatorTrailingEOL Then
                            ItemTrailingTriviaList.AddRange(SeparatorTrailingTrivia)
                        End If
                        ItemTrailingTriviaList.Add(VB_EOLTrivia)
                        SeparatorTrailingTrivia.Clear()
                    End If
                End If
            End If
            Return Node.With(ItemLeadingTriviaList, ItemTrailingTriviaList)
        End If
        FoundSeparatorTrailingEOL = False
        If Separator.HasLeadingTrivia Then
            Dim SeparatorLeadingTrivia As New List(Of SyntaxTrivia)
            For Each t As SyntaxTrivia In Separator.LeadingTrivia
                Select Case t.Kind
                    Case SyntaxKind.CommentTrivia
                        SeparatorTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                        SeparatorTrailingTrivia.Add(t)
                        FoundSeparatorTrailingEOL = True
                    Case SyntaxKind.EndOfLineTrivia
                        FoundSeparatorTrailingEOL = True
                    Case SyntaxKind.WhitespaceTrivia
                        SeparatorLeadingTrivia.Add(t)
                    Case Else
                        Stop
                End Select
            Next
            Separator = Separator.WithLeadingTrivia(SeparatorLeadingTrivia)
        End If
        If Separator.HasTrailingTrivia Then
            Dim TrailingTrivia As SyntaxTriviaList = Separator.TrailingTrivia
            For i As Integer = 0 To TrailingTrivia.Count - 1
                Dim t As SyntaxTrivia = TrailingTrivia(i)
                Select Case t.Kind
                    Case SyntaxKind.CommentTrivia
                        If i < TrailingTrivia.Count - 1 AndAlso TrailingTrivia(i + 1).IsKind(SyntaxKind.EndOfLineTrivia) Then
                            SeparatorTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                            SeparatorTrailingTrivia.Add(t)
                        Else
                            MovedTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                            MovedTrailingTrivia.Add(t)
                        End If
                    Case SyntaxKind.EndOfLineTrivia
                        FoundSeparatorTrailingEOL = True
                    Case SyntaxKind.WhitespaceTrivia
                        If FoundSeparatorTrailingEOL Then
                            MovedItemLeadingTrivia.Add(t)
                        Else
                            SeparatorTrailingTrivia.Add(t)
                        End If
                    Case Else
                        Stop
                End Select
            Next
        End If
        If FoundSeparatorTrailingEOL OrElse SeparatorTrailingTrivia.ContainsCommentOrDirectiveTrivia Then
            SeparatorTrailingTrivia.Add(VB_EOLTrivia)
        End If
        Separator = Separator.WithTrailingTrivia(SeparatorTrailingTrivia)
        Return Node.With(ItemLeadingTriviaList, ItemTrailingTriviaList)
    End Function

    Public Sub RestructureNodesAndSeparators(Of T As VisualBasicSyntaxNode)(ByRef OpenToken As SyntaxToken, ByRef Items As List(Of T), ByRef Separators As List(Of SyntaxToken), ByRef CloseToken As SyntaxToken, ByRef FinalTrailingTriviaList As List(Of SyntaxTrivia), Optional AllowLeadingComments As Boolean = False)
        Dim TokenLeadingTrivia As New List(Of SyntaxTrivia)
        Dim TokenTrailingTrivia As New List(Of SyntaxTrivia)
        TokenTrailingTrivia.AddRange(FinalTrailingTriviaList)
        FinalTrailingTriviaList.Clear()
        Dim ItemLeadingTriviaList As New List(Of SyntaxTrivia)
        Dim FoundEOL As Boolean = False
        If AllowLeadingComments Then
            TokenLeadingTrivia.AddRange(OpenToken.LeadingTrivia)
        Else
            If OpenToken.HasLeadingTrivia Then
                Dim LeadingTrivia As SyntaxTriviaList = OpenToken.LeadingTrivia
                For i As Integer = 0 To LeadingTrivia.Count - 1
                    Dim Trivia As SyntaxTrivia = LeadingTrivia(i)
                    Select Case Trivia.Kind
                        Case SyntaxKind.CommentTrivia
                            FinalTrailingTriviaList.Add(Trivia)
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            TokenLeadingTrivia.Add(Trivia)
                        Case SyntaxKind.CommentTrivia,
                            SyntaxKind.DocumentationCommentExteriorTrivia,
                            SyntaxKind.EmptyStatement,
                            SyntaxKind.BadDirectiveTrivia,
                            SyntaxKind.ConstDirectiveTrivia,
                            SyntaxKind.DisableWarningDirectiveTrivia,
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
                            TokenLeadingTrivia.Add(Trivia)
                        Case Else
                            Stop
                    End Select
                Next
            End If
            If FoundEOL Then
                TokenLeadingTrivia.Add(VB_EOLTrivia)
                FoundEOL = False
            End If
        End If
        If OpenToken.HasTrailingTrivia Then
            For Each Trivia As SyntaxTrivia In OpenToken.TrailingTrivia
                Select Case Trivia.Kind
                    Case SyntaxKind.CommentTrivia
                        FinalTrailingTriviaList.Add(Trivia)
                    Case SyntaxKind.EndOfLineTrivia
                        FoundEOL = True
                    Case SyntaxKind.WhitespaceTrivia
                        TokenLeadingTrivia.Add(Trivia)
                    Case Else
                        Stop
                End Select
            Next
        End If
        If FoundEOL Then
            TokenTrailingTrivia.Add(VB_EOLTrivia)
            FoundEOL = False
        End If
        OpenToken = OpenToken.With(TokenLeadingTrivia, TokenTrailingTrivia)
        TokenLeadingTrivia.Clear()
        TokenTrailingTrivia.Clear()
        Dim NeedTrailingEOL As Boolean = False
        Dim LastTrailingTrivia As New List(Of SyntaxTrivia)
        If CloseToken.HasLeadingTrivia Then
            Dim i As Integer
            Dim CloseTokenLeadingTrivia As SyntaxTriviaList = CloseToken.LeadingTrivia
            For i = 0 To CloseTokenLeadingTrivia.Count - 1
                Dim Trivia As SyntaxTrivia = CloseTokenLeadingTrivia(i)
                Select Case Trivia.Kind
                    Case SyntaxKind.CommentTrivia
                        TokenTrailingTrivia.Add(Trivia)
                        NeedTrailingEOL = True
                    Case SyntaxKind.EndOfLineTrivia
                        ' ignore
                    Case SyntaxKind.WhitespaceTrivia
                        TokenLeadingTrivia.Add(Trivia)
                    Case SyntaxKind.IfDirectiveTrivia
                        LastTrailingTrivia.Add(DirectiveNotAllowedHere(Trivia))
                    Case SyntaxKind.DisabledTextTrivia, SyntaxKind.EndIfDirectiveTrivia
                        LastTrailingTrivia.Add(SyntaxFactory.CommentTrivia(Trivia.ToFullString.Replace("  ", " ").WithoutNewLines))
                    Case Else
                        Stop
                End Select
            Next
        End If
        ' For now ignore
        'If FoundEOL Then
        '    TokenLeadingTrivia.Add(VB_EOLTrivia)
        '    FoundEOL = False
        'End If
        FoundEOL = NeedTrailingEOL
        If CloseToken.HasTrailingTrivia Then
            For Each Trivia As SyntaxTrivia In CloseToken.TrailingTrivia
                Select Case Trivia.Kind
                    Case SyntaxKind.CommentTrivia
                        TokenTrailingTrivia.Add(Trivia)
                        FoundEOL = True
                    Case SyntaxKind.EndOfLineTrivia
                        FoundEOL = True
                    Case SyntaxKind.WhitespaceTrivia
                        TokenLeadingTrivia.Add(Trivia)
                    Case Else
                        Stop
                End Select
            Next
        End If
        If FoundEOL Then
            TokenTrailingTrivia.Add(VB_EOLTrivia)
        End If
        CloseToken = CloseToken.With(TokenLeadingTrivia, TokenTrailingTrivia)
        FoundEOL = False
        If Items.Count > 0 Then
            Dim SeparatorCount As Integer = Items.Count - 1
            For i As Integer = 0 To Items.Count - 1
                Dim ItemWithTrivia As VisualBasicSyntaxNode = Items(i)
                If SeparatorCount > i Then
                    Items(i) = DirectCast(RestructureTrivia(Items(i), Separators(i), ItemLeadingTriviaList, FinalTrailingTriviaList), T)
                    If i = 0 Then
                        OpenToken = OpenToken.WithAppendedTrailingTrivia(ItemLeadingTriviaList)
                    Else
                        Separators(i - 1) = Separators(i - 1).WithAppendedTrailingTrivia(ItemLeadingTriviaList)
                    End If
                    ItemLeadingTriviaList.Clear()
                Else
                    LastTrailingTrivia.AddRange(FinalTrailingTriviaList)
                    FinalTrailingTriviaList.Clear()
                    Items(i) = DirectCast(RestructureTrivia(Items(i), Nothing, ItemLeadingTriviaList, LastTrailingTrivia), T)
                    If ItemLeadingTriviaList.Any Then
                        If Items.Count = 1 Then
                            OpenToken = OpenToken.WithAppendedTrailingTrivia(ItemLeadingTriviaList)
                        Else
                            Stop
                        End If
                    End If
                End If
            Next
            If Items.Last.ContainsEOLTrivia Then
                If FinalTrailingTriviaList.Any AndAlso FinalTrailingTriviaList.Last = VB_EOLTrivia Then
                    Items(Items.Count - 1) = Items.Last.WithTrailingTrivia(FinalTrailingTriviaList)
                    FinalTrailingTriviaList.Clear()
                End If
            End If
            If FinalTrailingTriviaList.Any Then
                Stop
            End If
        End If

    End Sub
End Module