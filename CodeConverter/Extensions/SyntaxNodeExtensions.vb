Option Explicit On
Option Infer Off
Option Strict On

Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices
Imports System.Text

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.CSharp.Syntax

Imports CS = Microsoft.CodeAnalysis.CSharp

Imports VBSyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind

Namespace IVisualBasicCode.CodeConverter.Util
    Partial Friend Module SyntaxNodeExtensions
        Friend Const DefaultEOL As String = vbCrLf

        <Extension>
        Private Function ConvertTrivia(t As SyntaxTrivia) As SyntaxTrivia
            If t.IsKind(SyntaxKind.None) Then
                Return Nothing
            End If

            If t.IsKind(SyntaxKind.List) Then
                Stop
            End If
            ' C# -> VB
            If t.IsKind(CS.SyntaxKind.WhitespaceTrivia) Then
                Return VisualBasic.SyntaxFactory.WhitespaceTrivia(t.ToString)
            End If

            If t.IsKind(CS.SyntaxKind.EndOfLineTrivia) Then
                If t.ToString.IsNewLine Then
                    Return VB_EOLTrivia
                Else
                    Stop
                End If
            End If

            If t.IsKind(CS.SyntaxKind.SingleLineCommentTrivia) Then
                If t.ToFullString.EndsWith("*/") Then
                    Return VisualBasic.SyntaxFactory.CommentTrivia($"'{ReplaceLeadingSlashes(t.ToFullString.Substring(2, t.ToFullString.Length - 4))}")
                End If
                Return VisualBasic.SyntaxFactory.CommentTrivia($"'{ReplaceLeadingSlashes(t.ToFullString.Substring(2))}")
            End If

            If t.IsKind(CS.SyntaxKind.MultiLineCommentTrivia) Then
                If t.ToFullString.Contains(vbLf) AndAlso Not t.ToFullString.EndsWith(vbLf) Then
                    Throw UnexpectedValue(NameOf(CS.SyntaxKind.MultiLineCommentTrivia))
                End If
                If t.ToFullString.EndsWith("*/") Then
                    Return VisualBasic.SyntaxFactory.CommentTrivia($"'{t.ToFullString.Substring(2, t.ToFullString.Length - 4)}")
                End If
                Return VisualBasic.SyntaxFactory.CommentTrivia($"'{t.ToFullString.Substring(2)}")
            End If

            If t.IsKind(CS.SyntaxKind.SingleLineDocumentationCommentTrivia) Then
                Dim VBSingleLineDocumentationCommentTrivia As VisualBasic.Syntax.DocumentationCommentTriviaSyntax = CreateVBDocumentCommentFromCSharpComment(DirectCast(t.GetStructure, CS.Syntax.DocumentationCommentTriviaSyntax))
                Return VisualBasic.SyntaxFactory.Trivia(VBSingleLineDocumentationCommentTrivia)
            End If

            If t.MatchesKind(CS.SyntaxKind.DocumentationCommentExteriorTrivia) Then
                Return VisualBasic.SyntaxFactory.SyntaxTrivia(VBSyntaxKind.CommentTrivia, "'''")
            End If

            If t.MatchesKind(CS.SyntaxKind.EndOfDocumentationCommentToken) Then
                Stop
                Dim sld As StructuredTriviaSyntax = DirectCast(t.GetStructure, StructuredTriviaSyntax)
            End If

            If t.MatchesKind(CS.SyntaxKind.MultiLineDocumentationCommentTrivia) Then
                Stop
                Dim sld As StructuredTriviaSyntax = DirectCast(t.GetStructure, StructuredTriviaSyntax)
            End If

            If t.IsKind(CS.SyntaxKind.RegionDirectiveTrivia) Then
                Dim NameString As String = $"{t.ToString.Replace("#region ", "").Replace("# region", "").Replace("""", "").Trim}"
                NameString = $"""{NameString}"""
                Dim RegionDirectiveTriviaNode As VisualBasic.Syntax.RegionDirectiveTriviaSyntax =
                       VisualBasic.SyntaxFactory.RegionDirectiveTrivia(
                             VisualBasic.SyntaxFactory.Token(VBSyntaxKind.HashToken),
                             VisualBasic.SyntaxFactory.Token(VBSyntaxKind.RegionKeyword),
                             VisualBasic.SyntaxFactory.StringLiteralToken($"{NameString}", $"{NameString}"))
                Return VisualBasic.SyntaxFactory.Trivia(RegionDirectiveTriviaNode)
            End If

            If t.IsKind(CS.SyntaxKind.EndRegionDirectiveTrivia) Then
                Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.EndRegionDirectiveTrivia())
            End If

            If t.IsKind(CS.SyntaxKind.IfDirectiveTrivia) Then
                Dim HashIF As CS.Syntax.IfDirectiveTriviaSyntax = DirectCast(t.GetStructure, CS.Syntax.IfDirectiveTriviaSyntax)
                If t.Token.Parent?.AncestorsAndSelf.OfType(Of CS.Syntax.InitializerExpressionSyntax).Any Then
                    IgnoredIfDepth += 1
                End If
                Dim Expression1 As String = HashIF.Condition.ToString.
                        Replace("!", "Not ").
                        Replace("==", "=").
                        Replace("!=", "<>").
                        Replace("&&", "And").
                        Replace("||", "Or").
                        Replace("  ", " ").
                        Replace("false", "False").
                        Replace("true", "True")

                Dim condition As VisualBasic.Syntax.ExpressionSyntax = VisualBasic.SyntaxFactory.ParseExpression(Expression1)
                Dim IfOrElseIfKeyword As SyntaxToken = VisualBasic.SyntaxFactory.Token(VBSyntaxKind.IfKeyword)
                Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.IfDirectiveTrivia(VBSyntaxKind.IfDirectiveTrivia, IfOrElseIfKeyword, condition))
            End If

            If t.IsKind(CS.SyntaxKind.ElseDirectiveTrivia) Then
                Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.ElseDirectiveTrivia.NormalizeWhitespace)
            End If
            If t.IsKind(CS.SyntaxKind.EndIfDirectiveTrivia) Then
                If IgnoredIfDepth > 0 Then
                    IgnoredIfDepth -= 1
                    Return VisualBasic.SyntaxFactory.CommentTrivia($"'VB does not allow directives here, original statement {t.ToFullString.WithoutNewLines}")
                End If
                Dim HashEndIf As StructuredTriviaSyntax = DirectCast(t.GetStructure, CS.Syntax.EndIfDirectiveTriviaSyntax)
                Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.EndIfDirectiveTrivia.NormalizeWhitespace)
            End If
            If t.IsKind(CS.SyntaxKind.PragmaWarningDirectiveTrivia) Then
                Dim StructuredTriva As CS.Syntax.PragmaWarningDirectiveTriviaSyntax = DirectCast(t.GetStructure, PragmaWarningDirectiveTriviaSyntax)
                Dim ErrorList As New List(Of VisualBasic.Syntax.IdentifierNameSyntax)
                Dim TrialingTrivia As IEnumerable(Of SyntaxTrivia) = Nothing
                For Each i As ExpressionSyntax In StructuredTriva.ErrorCodes
                    Dim ErrorCode As String = i.ToString
                    If ErrorCode.IsInteger Then
                        ErrorCode = $"CS_{ErrorCode}"
                    End If
                    ErrorList.Add(VisualBasic.SyntaxFactory.IdentifierName(ErrorCode))
                    TrialingTrivia = ConvertTrivia(i.GetTrailingTrivia)
                Next
#Disable Warning CC0013 ' Use Ternary operator.
                If StructuredTriva.DisableOrRestoreKeyword.ToString = "disable" Then
#Enable Warning CC0013 ' Use Ternary operator.
                    Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.DisableWarningDirectiveTrivia(ErrorList.ToArray).WithTrailingTrivia(TrialingTrivia).WithPrependedLeadingTrivia(SyntaxFactory.Comment($"' The value of the warning needs to be manually translated"), VB_EOLTrivia).NormalizeWhitespace)
                Else
                    Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.EnableWarningDirectiveTrivia(ErrorList.ToArray).WithTrailingTrivia(TrialingTrivia).WithPrependedLeadingTrivia(SyntaxFactory.Comment($"' The value of the warning needs to be manually translated"), VB_EOLTrivia).NormalizeWhitespace)
                End If
            End If
            If t.IsKind(CS.SyntaxKind.DisabledTextTrivia) Then
                If IgnoredIfDepth > 0 Then
                    Return VisualBasic.SyntaxFactory.DisabledTextTrivia(t.ToString.WithoutNewLines)
                End If
                Return VisualBasic.SyntaxFactory.DisabledTextTrivia(t.ToString.WithoutNewLines)
            End If
            Dim CSSyntaxKinds As Dictionary(Of SyntaxKind, VBSyntaxKind) = New Dictionary(Of SyntaxKind, VBSyntaxKind) From {
                        {CS.SyntaxKind.SkippedTokensTrivia, VBSyntaxKind.SkippedTokensTrivia},
                        {CS.SyntaxKind.DefineDirectiveTrivia, VBSyntaxKind.ConstDirectiveTrivia},
                        {CS.SyntaxKind.WarningDirectiveTrivia, VBSyntaxKind.DisableWarningDirectiveTrivia},
                        {CS.SyntaxKind.ReferenceDirectiveTrivia, VBSyntaxKind.ReferenceDirectiveTrivia},
                        {CS.SyntaxKind.BadDirectiveTrivia, VBSyntaxKind.BadDirectiveTrivia},
                        {CS.SyntaxKind.ConflictMarkerTrivia, VBSyntaxKind.ConflictMarkerTrivia},
                        {CS.SyntaxKind.LoadDirectiveTrivia, VBSyntaxKind.ExternalSourceDirectiveTrivia},
                        {CS.SyntaxKind.LineDirectiveTrivia, VBSyntaxKind.ExternalChecksumDirectiveTrivia}
               }
            If t.IsKind(CS.SyntaxKind.SkippedTokensTrivia) Then
                Dim ConvertedKind As KeyValuePair(Of SyntaxKind, VBSyntaxKind)? = CSSyntaxKinds.FirstOrNullable(Function(kvp As KeyValuePair(Of SyntaxKind, VBSyntaxKind)) t.IsKind(kvp.Key))
                Return If(ConvertedKind.HasValue, VisualBasic.SyntaxFactory.CommentTrivia($"/* TODO: Error Skipped {ConvertedKind.Value.Key}"), Nothing)
            End If

            Select Case t.Kind
                Case CS.SyntaxKind.DefineDirectiveTrivia
                    Dim DefineDirective As DefineDirectiveTriviaSyntax = DirectCast(t.GetStructure, DefineDirectiveTriviaSyntax)
                    Dim Name As SyntaxToken = VisualBasic.SyntaxFactory.Identifier(DefineDirective.Name.ValueText)
                    Dim value As VisualBasic.Syntax.ExpressionSyntax = VisualBasic.SyntaxFactory.TrueLiteralExpression(VisualBasic.SyntaxFactory.Token(VBSyntaxKind.TrueKeyword))
                    Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.ConstDirectiveTrivia(Name, value).WithConvertedTriviaFrom(DefineDirective))
                Case CS.SyntaxKind.ElifDirectiveTrivia
                    Dim ELIfDirective As ElifDirectiveTriviaSyntax = DirectCast(t.GetStructure, ElifDirectiveTriviaSyntax)
                    If t.Token.Parent.AncestorsAndSelf.OfType(Of CS.Syntax.InitializerExpressionSyntax).Any Then
                        IgnoredIfDepth += 1
                    End If
                    Dim Expression1 As String = ELIfDirective.Condition.ToString.
                        Replace("!", "Not ").
                        Replace("==", "=").
                        Replace("!=", "<>").
                        Replace("&&", "And").
                        Replace("||", "Or").
                        Replace("  ", " ").
                        Replace("false", "False").
                        Replace("true", "True")

                    Dim condition As VisualBasic.Syntax.ExpressionSyntax = VisualBasic.SyntaxFactory.ParseExpression(Expression1)
                    Dim IfOrElseIfKeyword As SyntaxToken = VisualBasic.SyntaxFactory.Token(VBSyntaxKind.ElseIfKeyword)
                    Return VisualBasic.SyntaxFactory.Trivia(VisualBasic.SyntaxFactory.ElseIfDirectiveTrivia(IfOrElseIfKeyword, condition).WithConvertedTriviaFrom(ELIfDirective))
                Case SyntaxKind.LineDirectiveTrivia
                    Return VisualBasic.SyntaxFactory.CommentTrivia($"' TODO: Check VB does not support Line Directive Trivia, Original Directive {t.ToString}")
                Case SyntaxKind.ErrorDirectiveTrivia
                    Return VisualBasic.SyntaxFactory.CommentTrivia($"' TODO: Check VB does not support Error Directive Trivia, Original Directive {t.ToString}")
                Case Else
                    Stop
            End Select
            Throw New NotImplementedException($"t.Kind({t.Kind}) Is unknown")

        End Function

        Private Function RemoveLeadingSpacesStar(line As String) As String
            Dim NewStringBuilder As New StringBuilder
            Dim SkipSpace As Boolean = True
            Dim SkipStar As Boolean = True
            For Each c As String In line
                Select Case c
                    Case " "
                        If SkipSpace Then
                            Continue For
                        End If
                        NewStringBuilder.Append(c)
                    Case "*"
                        If SkipStar Then
                            SkipSpace = False
                            SkipStar = False
                            Continue For
                        End If
                        NewStringBuilder.Append(c)
                    Case Else
                        SkipSpace = False
                        SkipStar = False
                        NewStringBuilder.Append(c)
                End Select
            Next
            Return NewStringBuilder.ToString
        End Function

        Private Function ReplaceLeadingSlashes(CommentTriviaBody As String) As String
            For i As Integer = 0 To CommentTriviaBody.Length - 1
                If CommentTriviaBody.Substring(i, 1) = "/" Then
                    CommentTriviaBody = CommentTriviaBody.Remove(i, 1).Insert(i, "'")
                Else
                    Exit For
                End If
            Next
            Return CommentTriviaBody
        End Function

        <Extension()>
        Public Function [With](Of T As SyntaxNode)(ByVal node As T, ByVal leadingTrivia As IEnumerable(Of SyntaxTrivia), ByVal trailingTrivia As IEnumerable(Of SyntaxTrivia)) As T
            Return node.WithLeadingTrivia(leadingTrivia).WithTrailingTrivia(trailingTrivia)
        End Function

        <Extension>
        Public Function ConvertTrivia(triviaToConvert As IReadOnlyCollection(Of SyntaxTrivia)) As IEnumerable(Of SyntaxTrivia)
            Dim TriviaList As New List(Of SyntaxTrivia)
            Dim LastTriviaIsIfDirective As Boolean = False
            Try
                For Each t As SyntaxTrivia In triviaToConvert
                    If t.IsKind(SyntaxKind.MultiLineCommentTrivia) Then
                        Dim Lines() As String = t.ToFullString.Substring(2).Split(CType(vbLf, Char))
                        For Each line As String In Lines
                            If line.EndsWith("*/") Then
                                TriviaList.Add(VisualBasic.SyntaxFactory.CommentTrivia($"' {RemoveLeadingSpacesStar(line.Substring(0, line.Length - 2))}"))
                            Else
                                TriviaList.Add(VisualBasic.SyntaxFactory.CommentTrivia($"' {RemoveLeadingSpacesStar(line)}"))
                            End If
                        Next
                    Else
                        Dim ConvertedTrivia As SyntaxTrivia = ConvertTrivia(t)
                        If ConvertedTrivia = Nothing Then
                            Continue For
                        End If
                        If ConvertedTrivia.IsKind(VBSyntaxKind.IfDirectiveTrivia) OrElse ConvertedTrivia.IsKind(VisualBasic.SyntaxKind.ElseDirectiveTrivia) Then
                            LastTriviaIsIfDirective = True
                        ElseIf ConvertedTrivia.IsKind(VBSyntaxKind.DisabledTextTrivia) Then
                            TriviaList.Add(VB_EOLTrivia)
                            LastTriviaIsIfDirective = False
                        Else
                            LastTriviaIsIfDirective = False
                        End If
                        TriviaList.Add(ConvertedTrivia)
                    End If
                Next
            Catch ex As Exception
                Stop
                Throw
            End Try
            Return TriviaList
        End Function

        <Extension()>
        Public Function GetAncestor(Of TNode As SyntaxNode)(ByVal node As SyntaxNode) As TNode
            If node Is Nothing Then
                Return Nothing
            End If

            Return node.GetAncestors(Of TNode)().FirstOrDefault()
        End Function

        <Extension()>
        Public Iterator Function GetAncestors(Of TNode As SyntaxNode)(ByVal node As SyntaxNode) As IEnumerable(Of TNode)
            Dim current As SyntaxNode = node.Parent
            While current IsNot Nothing
                If TypeOf current Is TNode Then
                    Yield DirectCast(current, TNode)
                End If

                current = If(TypeOf current Is IStructuredTriviaSyntax, (DirectCast(current, IStructuredTriviaSyntax)).ParentTrivia.Token.Parent, current.Parent)
            End While
        End Function

        <Extension()>
        Public Iterator Function GetAncestorsOrThis(Of TNode As SyntaxNode)(ByVal node As SyntaxNode) As IEnumerable(Of TNode)
            Dim current As SyntaxNode = node
            While current IsNot Nothing
                If TypeOf current Is TNode Then
                    Yield DirectCast(current, TNode)
                End If

                current = If(TypeOf current Is IStructuredTriviaSyntax, (DirectCast(current, IStructuredTriviaSyntax)).ParentTrivia.Token.Parent, current.Parent)
            End While
        End Function

        <Extension()>
        Public Function GetBraces(ByVal node As SyntaxNode) As ValueTuple(Of SyntaxToken, SyntaxToken)
            Dim namespaceNode As NamespaceDeclarationSyntax = TryCast(node, NamespaceDeclarationSyntax)
            If namespaceNode IsNot Nothing Then
                Return ValueTuple.Create(namespaceNode.OpenBraceToken, namespaceNode.CloseBraceToken)
            End If

            Dim baseTypeNode As BaseTypeDeclarationSyntax = TryCast(node, BaseTypeDeclarationSyntax)
            If baseTypeNode IsNot Nothing Then
                Return ValueTuple.Create(baseTypeNode.OpenBraceToken, baseTypeNode.CloseBraceToken)
            End If

            Dim accessorListNode As AccessorListSyntax = TryCast(node, AccessorListSyntax)
            If accessorListNode IsNot Nothing Then
                Return ValueTuple.Create(accessorListNode.OpenBraceToken, accessorListNode.CloseBraceToken)
            End If

            Dim blockNode As BlockSyntax = TryCast(node, BlockSyntax)
            If blockNode IsNot Nothing Then
                Return ValueTuple.Create(blockNode.OpenBraceToken, blockNode.CloseBraceToken)
            End If

            Dim switchStatementNode As SwitchStatementSyntax = TryCast(node, SwitchStatementSyntax)
            If switchStatementNode IsNot Nothing Then
                Return ValueTuple.Create(switchStatementNode.OpenBraceToken, switchStatementNode.CloseBraceToken)
            End If

            Dim anonymousObjectCreationExpression As AnonymousObjectCreationExpressionSyntax = TryCast(node, AnonymousObjectCreationExpressionSyntax)
            If anonymousObjectCreationExpression IsNot Nothing Then
                Return ValueTuple.Create(anonymousObjectCreationExpression.OpenBraceToken, anonymousObjectCreationExpression.CloseBraceToken)
            End If

            Dim initializeExpressionNode As InitializerExpressionSyntax = TryCast(node, InitializerExpressionSyntax)
            If initializeExpressionNode IsNot Nothing Then
                Return ValueTuple.Create(initializeExpressionNode.OpenBraceToken, initializeExpressionNode.CloseBraceToken)
            End If

            Return New ValueTuple(Of SyntaxToken, SyntaxToken)()
        End Function

        <Extension()>
        Public Function IsKind(ByVal node As SyntaxNode, ByVal kind1 As SyntaxKind, ByVal kind2 As SyntaxKind) As Boolean
            If node Is Nothing Then
                Return False
            End If

            Dim csharpKind As SyntaxKind = node.Kind()
            Return csharpKind = kind1 OrElse csharpKind = kind2
        End Function

        <Extension()>
        Public Function IsKind(ByVal node As SyntaxNode, ByVal kind1 As SyntaxKind, ByVal kind2 As SyntaxKind, ByVal kind3 As SyntaxKind, ByVal kind4 As SyntaxKind) As Boolean
            If node Is Nothing Then
                Return False
            End If

            Dim csharpKind As SyntaxKind = node.Kind()
            Return csharpKind = kind1 OrElse csharpKind = kind2 OrElse csharpKind = kind3 OrElse csharpKind = kind4
        End Function

        <Extension>
        Public Function ParentHasSameTrailingTrivia(otherNode As SyntaxNode) As Boolean
            If otherNode.Parent Is Nothing Then
                Return False
            End If
            Return otherNode.Parent.GetLastToken() = otherNode.GetLastToken()
        End Function

        <Extension()>
        Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(ByVal node As T, ParamArray trivia As SyntaxTrivia()) As T
            If trivia.Length = 0 Then
                Return node
            End If

            Return node.WithAppendedTrailingTrivia(DirectCast(trivia, IEnumerable(Of SyntaxTrivia)))
        End Function

        <Extension()>
        Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(ByVal node As T, ByVal trivia As SyntaxTriviaList) As T
            If trivia.Count = 0 Then
                Return node
            End If
            If node Is Nothing Then
                Return Nothing
            End If

            Return node.WithTrailingTrivia(node.GetTrailingTrivia().Concat(trivia))
        End Function

        <Extension()>
        Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(ByVal node As T, ByVal trivia As IEnumerable(Of SyntaxTrivia)) As T
            If node Is Nothing Then
                Return Nothing
            End If
            Return node.WithAppendedTrailingTrivia(trivia.ToSyntaxTriviaList())
        End Function

        <Extension()>
        Public Function WithPrependedLeadingTrivia(Of T As SyntaxNode)(ByVal node As T, ParamArray trivia As SyntaxTrivia()) As T
            If trivia.Length = 0 Then
                Return node
            End If

            Return node.WithPrependedLeadingTrivia(DirectCast(trivia, IEnumerable(Of SyntaxTrivia)))
        End Function

        <Extension()>
        Public Function WithPrependedLeadingTrivia(Of T As SyntaxNode)(ByVal node As T, ByVal trivia As SyntaxTriviaList) As T
            If trivia.Count = 0 Then
                Return node
            End If

            Return node.WithLeadingTrivia(trivia.Concat(node.GetLeadingTrivia()))
        End Function

        <Extension()>
        Public Function WithPrependedLeadingTrivia(Of T As SyntaxNode)(ByVal node As T, ByVal trivia As IEnumerable(Of SyntaxTrivia)) As T
            Return node.WithPrependedLeadingTrivia(trivia.ToSyntaxTriviaList())
        End Function

        <Extension>
        Public Function WithRemovedTrailingEOLTrivia(Of T As SyntaxNode)(node As T) As T
            If Not node.HasTrailingTrivia Then
                Return node
            End If

            Dim NodeTrailingTrivia As SyntaxTriviaList = node.GetTrailingTrivia
            If NodeTrailingTrivia.ContainsEOLTrivia Then
                Dim NewTriviaList As New List(Of SyntaxTrivia)
                For Each Trivia As SyntaxTrivia In NodeTrailingTrivia
                    If Trivia.IsKind(VisualBasic.SyntaxKind.EndOfLineTrivia) Then
                        Continue For
                    End If
                    NewTriviaList.Add(Trivia)
                Next
                Return node.WithTrailingTrivia(NewTriviaList)
            Else
                Return node
            End If
        End Function

#Region "WithConvertedTriviaFrom"
        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function WithConvertedTriviaFrom(Of TSyntax As XmlNodeSyntax)(node As TSyntax, otherNode As XmlNodeSyntax) As TSyntax
            If otherNode Is Nothing Then
                Return node
            End If
            If otherNode.HasLeadingTrivia Then
                node = node.WithLeadingTrivia(ConvertTrivia(otherNode.GetLeadingTrivia()))
            End If
            If Not otherNode.HasTrailingTrivia OrElse ParentHasSameTrailingTrivia(otherNode) Then
                Return node
            End If
            Return node.WithTrailingTrivia(ConvertTrivia(otherNode.GetTrailingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedTriviaFrom(Of T As SyntaxNode)(node As T, otherNode As SyntaxNode) As T
            If otherNode Is Nothing Then
                Return node
            End If
            If otherNode.HasLeadingTrivia Then
                node = node.WithLeadingTrivia(ConvertTrivia(otherNode.GetLeadingTrivia()))
            End If
            If Not otherNode.HasTrailingTrivia OrElse ParentHasSameTrailingTrivia(otherNode) Then
                Return node
            End If
            Return node.WithTrailingTrivia(ConvertTrivia(otherNode.GetTrailingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedTriviaFrom(Of T As SyntaxNode)(node As T, otherToken As SyntaxToken) As T
            If otherToken.HasLeadingTrivia Then
                node = node.WithLeadingTrivia(ConvertTrivia(otherToken.LeadingTrivia))
            End If
            If Not otherToken.HasTrailingTrivia Then
                Return node
            End If
            Return node.WithTrailingTrivia(ConvertTrivia(otherToken.TrailingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedTriviaFrom(Token As SyntaxToken, otherNode As SyntaxNode) As SyntaxToken
            If otherNode.HasLeadingTrivia Then
                Token = Token.WithLeadingTrivia(ConvertTrivia(otherNode.GetLeadingTrivia))
            End If
            If Not otherNode.HasTrailingTrivia OrElse ParentHasSameTrailingTrivia(otherNode) Then
                Return Token
            End If
            Return Token.WithTrailingTrivia(ConvertTrivia(otherNode.GetTrailingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedTriviaFrom(Token As SyntaxToken, otherToken As SyntaxToken) As SyntaxToken
            Try

                If otherToken.HasLeadingTrivia Then
                    Dim LeadingTrivia As New List(Of SyntaxTrivia)
                    LeadingTrivia.AddRange(ConvertTrivia(otherToken.LeadingTrivia()))
                    Token = Token.WithLeadingTrivia(LeadingTrivia)
                End If
                Return Token.WithTrailingTrivia(ConvertTrivia(otherToken.TrailingTrivia()))
            Catch ex As Exception
                Stop
            End Try
        End Function

        ''' <summary>
        ''' Allows for swapping trivia usually for cast or declarations
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="node"></param>
        ''' <param name="LeadingNode"></param>
        ''' <param name="TrailingNode"></param>
        ''' <returns></returns>
        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function WithConvertedTriviaFrom(Of T As SyntaxNode)(node As T, LeadingNode As SyntaxNode, TrailingNode As SyntaxNode) As T
            Return node.WithConvertedLeadingTriviaFrom(LeadingNode).WithConvertedTrailingTriviaFrom(TrailingNode)
        End Function

#End Region

#Region "WithConvertedLeadingTriviaFrom"

        <Extension>
        Public Function WithConvertedLeadingTriviaFrom(Of T As SyntaxNode)(node As T, otherNode As SyntaxNode) As T
            If otherNode Is Nothing OrElse Not otherNode.HasLeadingTrivia Then
                Return node
            End If
            Return node.WithLeadingTrivia(ConvertTrivia(otherNode.GetLeadingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedLeadingTriviaFrom(Of T As SyntaxNode)(node As T, otherToken As SyntaxToken) As T
            If Not otherToken.HasLeadingTrivia Then
                Return node
            End If
            Return node.WithLeadingTrivia(ConvertTrivia(otherToken.LeadingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedLeadingTriviaFrom(node As SyntaxToken, otherToken As SyntaxToken) As SyntaxToken
            If Not otherToken.HasLeadingTrivia Then
                Return node
            End If
            Return node.WithLeadingTrivia(ConvertTrivia(otherToken.LeadingTrivia()))
        End Function

#End Region

#Region "WithConvertedTrailingTriviaFrom"

        <Extension>
        Public Function WithConvertedTrailingTriviaFrom(Of T As SyntaxNode)(node As T, otherNode As SyntaxNode) As T
            If otherNode Is Nothing OrElse Not otherNode.HasTrailingTrivia Then
                Return node
            End If
            If ParentHasSameTrailingTrivia(otherNode) Then
                Return node
            End If
            Return node.WithTrailingTrivia(ConvertTrivia(otherNode.GetTrailingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedTrailingTriviaFrom(Of T As SyntaxNode)(node As T, otherToken As SyntaxToken) As T
            If Not otherToken.HasTrailingTrivia Then
                Return node
            End If
            Return node.WithTrailingTrivia(ConvertTrivia(otherToken.TrailingTrivia()))
        End Function

        <Extension>
        Public Function WithConvertedTrailingTriviaFrom(Token As SyntaxToken, otherToken As SyntaxToken) As SyntaxToken
            Return Token.WithTrailingTrivia(ConvertTrivia(otherToken.TrailingTrivia()))
        End Function

#End Region

    End Module
End Namespace