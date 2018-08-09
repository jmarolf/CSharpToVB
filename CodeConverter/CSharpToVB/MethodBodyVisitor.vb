Option Explicit On
Option Infer Off
Option Strict On
Imports System.Runtime.InteropServices
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
Imports LiteralExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.LiteralExpressionSyntax
Imports QualifiedNameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.QualifiedNameSyntax
Imports StatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.StatementSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax
Imports UsingStatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.UsingStatementSyntax
Imports VariableDeclaratorSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.VariableDeclaratorSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public NotInheritable Class CSharpConverter

        Friend Class MethodBodyVisitor
            Inherits CS.CSharpSyntaxVisitor(Of SyntaxList(Of StatementSyntax))

            Private ReadOnly LiteralExpression_1 As ExpressionSyntax = SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(1))
            Private ReadOnly mSemanticModel As SemanticModel
            Private mBlockInfo As Stack(Of BlockInfo) = New Stack(Of BlockInfo)()
            Private mNodesVisitor As NodesVisitor
            ' currently only works with switch blocks

            Private switchCount As Integer = 0
            Public Sub New(ByVal semanticModel As SemanticModel, ByVal nodesVisitor As NodesVisitor)
                mSemanticModel = semanticModel
                mNodesVisitor = nodesVisitor
            End Sub

            Public Property IsInterator As Boolean

            Private Shared Function GetPossibleEventName(ByVal expression As CSS.ExpressionSyntax) As String
                Dim ident As CSS.IdentifierNameSyntax = TryCast(expression, CSS.IdentifierNameSyntax)
                If ident IsNot Nothing Then Return ident.Identifier.Text
                Dim fre As CSS.MemberAccessExpressionSyntax = TryCast(expression, CSS.MemberAccessExpressionSyntax)
                If fre IsNot Nothing AndAlso fre.Expression.IsKind(CS.SyntaxKind.ThisExpression) Then Return fre.Name.Identifier.Text
                Return Nothing
            End Function

            Private Shared Function IsSimpleStatement(ByVal statement As CSS.StatementSyntax) As Boolean
                Return TypeOf statement Is CSS.ExpressionStatementSyntax OrElse
                    TypeOf statement Is CSS.BreakStatementSyntax OrElse
                    TypeOf statement Is CSS.ContinueStatementSyntax OrElse
                    TypeOf statement Is CSS.ReturnStatementSyntax OrElse
                    TypeOf statement Is CSS.YieldStatementSyntax OrElse
                    TypeOf statement Is CSS.ThrowStatementSyntax
            End Function

            ''' <summary>
            ''' Used for moving Trivia between Condition and Then in a If Statement
            ''' </summary>
            ''' <param name="OriginalTrailingTrivia"></param>
            ''' <param name="ThenTrailingTrivia"></param>
            ''' <param name="ConditionPreservedTrailingTrivia"></param>
            Private Shared Sub RestructureConditionThenTrailingTrivia(OriginalTrailingTrivia As IEnumerable(Of SyntaxTrivia), ByRef ThenTrailingTrivia As List(Of SyntaxTrivia), ByRef ConditionPreservedTrailingTrivia As List(Of SyntaxTrivia))
                For Each t As SyntaxTrivia In OriginalTrailingTrivia
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            ThenTrailingTrivia.Add(t)
                        Case SyntaxKind.WhitespaceTrivia
                            ConditionPreservedTrailingTrivia.Add(t)
                        Case SyntaxKind.EndOfLineTrivia, SyntaxKind.None
                            ' No return allowed after condition before Then
                        Case Else
                            Stop
                    End Select
                Next
            End Sub

            ''' <summary>
            ''' Used for moving comment Trivia to end of statement
            ''' </summary>
            ''' <param name="OriginalTrailingTrivia"></param>
            ''' <param name="FoundEOL"></param>
            ''' <param name="NewTrailingTrivia"></param>
            ''' <returns></returns>
            Private Shared Function RestructureTrailingTrivia(OriginalTrailingTrivia As IEnumerable(Of SyntaxTrivia), FoundEOL As Boolean, ByRef NewTrailingTrivia As List(Of SyntaxTrivia)) As Boolean
                For Each t As SyntaxTrivia In OriginalTrailingTrivia
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            NewTrailingTrivia.Add(t)
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            NewTrailingTrivia.Add(t)
                    End Select
                Next

                Return FoundEOL
            End Function

            ''' <summary>
            '''
            ''' </summary>
            ''' <param name="nodes"></param>
            ''' <param name="comment"></param>
            ''' <returns></returns>
            ''' <remarks>PC added ' to start of CommentTrivia</remarks>
            Private Shared Function WrapInComment(ByVal nodes As SyntaxList(Of StatementSyntax), NodeWithComments As CSS.StatementSyntax, ByVal comment As String) As SyntaxList(Of StatementSyntax)
                If nodes.Count > 0 Then
                    nodes = nodes.Replace(nodes(0), nodes(0).WithConvertedTriviaFrom(NodeWithComments).WithPrependedLeadingTrivia(SyntaxFactory.CommentTrivia($"' BEGIN TODO: {comment}")))
                    nodes = nodes.Add(SyntaxFactory.EmptyStatement.WithLeadingTrivia(VB_EOLTrivia, SyntaxFactory.CommentTrivia($"' END TODO: {comment}")))
                End If

                Return nodes
            End Function

            Private Iterator Function AddLabels(ByVal blocks As CaseBlockSyntax(), ByVal gotoLabels As List(Of VisualBasicSyntaxNode)) As IEnumerable(Of CaseBlockSyntax)
                For Each _block As CaseBlockSyntax In blocks
                    Dim block As CaseBlockSyntax = _block
                    For Each caseClause As CaseClauseSyntax In block.CaseStatement.Cases
                        Dim expression As VisualBasicSyntaxNode = If(TypeOf caseClause Is ElseCaseClauseSyntax, DirectCast(caseClause, VisualBasicSyntaxNode), (DirectCast(caseClause, SimpleCaseClauseSyntax)).Value)
                        If gotoLabels.Any(Function(label As VisualBasicSyntaxNode) label.IsEquivalentTo(expression)) Then block = block.WithStatements(block.Statements.Insert(0, SyntaxFactory.LabelStatement(MakeGotoSwitchLabel(expression))))
                    Next

                    Yield block
                Next
            End Function

            Private Sub CollectElseBlocks(ByVal node As CSS.IfStatementSyntax, ByVal elseIfBlocks As List(Of ElseIfBlockSyntax), ByRef elseBlock As ElseBlockSyntax, ByRef OpenBraceTrailingTrivia As List(Of SyntaxTrivia), ByRef CloseBraceLeadingTrivia As List(Of SyntaxTrivia))
                If node.[Else] Is Nothing Then
                    Return
                End If

                If TypeOf node.[Else].Statement Is CSS.IfStatementSyntax Then
                    Dim [elseIf] As CSS.IfStatementSyntax = DirectCast(node.[Else].Statement, CSS.IfStatementSyntax)
                    Dim ElseIFKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ElseIfKeyword).WithConvertedTriviaFrom(node.Else)
                    Dim NewThenTrailingTrivia As New List(Of SyntaxTrivia)
                    Dim Condition As ExpressionSyntax = DirectCast([elseIf].Condition.Accept(mNodesVisitor), ExpressionSyntax)
                    NewThenTrailingTrivia.AddRange(Condition.GetTrailingTrivia)
                    If node.CloseParenToken.HasLeadingTrivia AndAlso node.CloseParenToken.LeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                        NewThenTrailingTrivia.AddRange(ConvertTrivia(node.CloseParenToken.LeadingTrivia))
                    End If
                    NewThenTrailingTrivia.AddRange(ConvertTrivia([elseIf].GetLeadingTrivia))
                    Dim ThenKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ThenKeyword).WithTrailingTrivia(NewThenTrailingTrivia)
                    Dim ElseIfStatement As ElseIfStatementSyntax = SyntaxFactory.ElseIfStatement(ElseIFKeyword, Condition.WithTrailingTrivia(SyntaxFactory.Space), ThenKeyword)
                    Dim ElseIfBlock As ElseIfBlockSyntax = SyntaxFactory.ElseIfBlock(ElseIfStatement, ConvertBlock([elseIf].Statement, OpenBraceTrailingTrivia, CloseBraceLeadingTrivia))
                    elseIfBlocks.Add(ElseIfBlock)
                    CollectElseBlocks([elseIf], elseIfBlocks, elseBlock, OpenBraceTrailingTrivia, CloseBraceLeadingTrivia)
                Else
                    Dim Statements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.[Else].Statement, OpenBraceTrailingTrivia, CloseBraceLeadingTrivia)
                    Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                    If node.Else.Statement.GetBraces.Item1.HasTrailingTrivia Then
                        TrailingTriviaList.AddRange(ConvertTrivia(node.Else.Statement.GetBraces.Item1.TrailingTrivia))
                    End If
                    Dim ElseStatement As ElseStatementSyntax = SyntaxFactory.ElseStatement(SyntaxFactory.Token(SyntaxKind.ElseKeyword)).WithTrailingTrivia(TrailingTriviaList)
                    elseBlock = SyntaxFactory.ElseBlock(ElseStatement, Statements).WithPrependedLeadingTrivia(OpenBraceTrailingTrivia).WithAppendedTrailingTrivia(CloseBraceLeadingTrivia)
                    OpenBraceTrailingTrivia.Clear()
                    CloseBraceLeadingTrivia.Clear()
                End If
            End Sub

            Private Function ConvertBlock(ByVal node As CSS.StatementSyntax, ByRef OpenBraceTrailiningTrivia As List(Of SyntaxTrivia), ByRef CloseBraceLeadingTrivia As List(Of SyntaxTrivia)) As SyntaxList(Of StatementSyntax)
                Dim Braces As (OpenBrace As SyntaxToken, CloseBrace As SyntaxToken) = node.GetBraces
                Dim OpenBrace As SyntaxToken = If(Braces.OpenBrace = Nothing, New SyntaxToken, Braces.OpenBrace)
                OpenBraceTrailiningTrivia.AddRange(ExtractComments(ConvertTrivia(OpenBrace.TrailingTrivia), Leading:=False))
                Dim CloseBrace As SyntaxToken = If(Braces.CloseBrace = Nothing, New SyntaxToken, Braces.CloseBrace)
                CloseBraceLeadingTrivia.AddRange(collection:=ExtractComments(ListOfTrivia:=ConvertTrivia(CloseBrace.LeadingTrivia), Leading:=True))
                If TypeOf node Is CSS.BlockSyntax Then
                    Dim NodeBlock As CSS.BlockSyntax = DirectCast(node, CSS.BlockSyntax)
                    Dim Statements As IEnumerable(Of StatementSyntax) = NodeBlock.Statements.Where(Function(s As CSS.StatementSyntax) Not (TypeOf s Is CSS.EmptyStatementSyntax)).SelectMany(Function(s As CSS.StatementSyntax) s.Accept(Me))
                    Dim StatementList As SyntaxList(Of StatementSyntax)
                    StatementList = SyntaxFactory.List(Statements)
                    If Statements.Count = 0 Then
                        StatementList = StatementList.Add(SyntaxFactory.EmptyStatement)
                    End If
                    Return StatementList
                End If

                If TypeOf node Is CSS.EmptyStatementSyntax Then
                    Return SyntaxFactory.List(Of StatementSyntax)()
                End If

                Return node.Accept(Me)
            End Function
            Private Function ConvertCatchClause(ByVal index As Integer, ByVal catchClause As CSS.CatchClauseSyntax) As CatchBlockSyntax
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim statements As SyntaxList(Of StatementSyntax) = ConvertBlock(catchClause.Block, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)

                If catchClause.Declaration Is Nothing Then
                    Return SyntaxFactory.CatchBlock(SyntaxFactory.CatchStatement().WithTrailingTrivia(VB_EOLTrivia).WithAppendedTrailingTrivia(ClosingBraceLeadingTrivia), statements)
                End If
                If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                    statements = statements.Replace(statements(0), statements(0).WithPrependedLeadingTrivia(OpenBraceTrailingTrivia))
                    Dim laststatement As Integer = statements.Count - 1
                    statements = statements.Replace(statements(laststatement), statements(laststatement).WithAppendedTrailingTrivia(ClosingBraceLeadingTrivia))
                End If
                Dim type As TypeSyntax = DirectCast(catchClause.Declaration.Type.Accept(mNodesVisitor), TypeSyntax)
                Dim simpleTypeName As String
                simpleTypeName = If(TypeOf type Is QualifiedNameSyntax, (DirectCast(type, QualifiedNameSyntax)).Right.ToString(), type.ToString())
                Dim identifier As SyntaxToken = If(catchClause.Declaration.Identifier.IsKind(CS.SyntaxKind.None),
                                                        SyntaxFactory.Identifier($"__unused{simpleTypeName}{index + 1}__"),
                                                        GenerateSafeVBToken(catchClause.Declaration.Identifier, IsQualifiedName:=False))
                Dim WhenClause As Syntax.CatchFilterClauseSyntax = If(catchClause.Filter Is Nothing, Nothing, SyntaxFactory.CatchFilterClause(filter:=DirectCast(catchClause.Filter.FilterExpression.Accept(mNodesVisitor), ExpressionSyntax)))
                Dim CatchStatement As CatchStatementSyntax = SyntaxFactory.CatchStatement(
                                                        identifierName:=SyntaxFactory.IdentifierName(identifier),
                                                        asClause:=SyntaxFactory.SimpleAsClause(type),
                                                        whenClause:=WhenClause).WithConvertedLeadingTriviaFrom(catchClause)
                If Not CatchStatement.HasTrailingTrivia Then
                    CatchStatement = CatchStatement.WithTrailingTrivia(VB_EOLTrivia)
                ElseIf CatchStatement.GetTrailingTrivia.Last <> VB_EOLTrivia Then
                    CatchStatement = CatchStatement.WithTrailingTrivia(VB_EOLTrivia)
                End If
                Return SyntaxFactory.CatchBlock(CatchStatement, statements)
            End Function

            Private Function ConvertForToSimpleForNext(ByVal node As CSS.ForStatementSyntax, <Out> ByRef block As StatementSyntax) As Boolean
                block = Nothing
                Dim hasVariable As Boolean = node.Declaration IsNot Nothing AndAlso node.Declaration.Variables.Count = 1
                If Not hasVariable AndAlso node.Initializers.Count <> 1 Then
                    Return False
                End If
                If node.Incrementors.Count <> 1 Then
                    Return False
                End If
                Dim Incrementors As VisualBasicSyntaxNode = node.Incrementors.FirstOrDefault()?.Accept(mNodesVisitor)
                Dim iterator As AssignmentStatementSyntax = TryCast(Incrementors, AssignmentStatementSyntax)
                If iterator Is Nothing OrElse Not iterator.IsKind(SyntaxKind.AddAssignmentStatement, SyntaxKind.SubtractAssignmentStatement) Then
                    Return False
                End If
                Dim iteratorIdentifier As IdentifierNameSyntax = TryCast(iterator.Left, IdentifierNameSyntax)
                If iteratorIdentifier Is Nothing Then
                    Return False
                End If
                Dim stepExpression As LiteralExpressionSyntax = TryCast(iterator.Right, LiteralExpressionSyntax)
                If stepExpression Is Nothing OrElse Not (TypeOf stepExpression.Token.Value Is Integer) Then Return False
                Dim [step] As Integer = CInt(stepExpression.Token.Value)
                If SyntaxTokenExtensions.IsKind(iterator.OperatorToken, SyntaxKind.MinusToken) Then
                    [step] = -[step]
                End If
                Dim condition As CSS.BinaryExpressionSyntax = TryCast(node.Condition, CSS.BinaryExpressionSyntax)
                If condition Is Nothing OrElse Not (TypeOf condition.Left Is CSS.IdentifierNameSyntax) Then
                    Return False
                End If
                If (DirectCast(condition.Left, CSS.IdentifierNameSyntax)).Identifier.IsEquivalentTo(iteratorIdentifier.Identifier) Then
                    Return False
                End If
                Dim [end] As ExpressionSyntax
                If iterator.IsKind(SyntaxKind.SubtractAssignmentStatement) Then
                    If condition.IsKind(CS.SyntaxKind.GreaterThanOrEqualExpression) OrElse condition.IsKind(CS.SyntaxKind.NotEqualsExpression) Then
                        [end] = DirectCast(condition.Right.Accept(mNodesVisitor), ExpressionSyntax)
                    ElseIf condition.IsKind(CS.SyntaxKind.GreaterThanExpression) Then
                        [end] = SyntaxFactory.BinaryExpression(SyntaxKind.AddExpression, DirectCast(condition.Right.Accept(mNodesVisitor), ExpressionSyntax), SyntaxFactory.Token(SyntaxKind.PlusToken), LiteralExpression_1)
                    Else
                        Return False
                    End If
                Else
                    If condition.IsKind(CS.SyntaxKind.LessThanOrEqualExpression) OrElse condition.IsKind(CS.SyntaxKind.NotEqualsExpression) Then
                        [end] = DirectCast(condition.Right.Accept(mNodesVisitor), ExpressionSyntax)
                    ElseIf condition.IsKind(CS.SyntaxKind.LessThanExpression) Then
                        [end] = SyntaxFactory.BinaryExpression(SyntaxKind.SubtractExpression, DirectCast(condition.Right.Accept(mNodesVisitor), ExpressionSyntax), SyntaxFactory.Token(SyntaxKind.MinusToken), LiteralExpression_1)
                    Else
                        Return False
                    End If
                End If

                Dim variable As VisualBasicSyntaxNode
                Dim start As ExpressionSyntax
                If hasVariable Then
                    Dim v As CSS.VariableDeclaratorSyntax = node.Declaration.Variables(0)
                    start = DirectCast(v.Initializer?.Value.Accept(mNodesVisitor), ExpressionSyntax)
                    If start Is Nothing Then
                        Return False
                    End If
                    variable = SyntaxFactory.VariableDeclarator(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(id:=v.Identifier, IsQualifiedName:=False))), asClause:=If(node.Declaration.Type.IsVar, Nothing, SyntaxFactory.SimpleAsClause(type:=DirectCast(node.Declaration.Type.Accept(mNodesVisitor), TypeSyntax))), initializer:=Nothing)
                Else
                    Dim initializer As CSS.AssignmentExpressionSyntax = TryCast(node.Initializers.FirstOrDefault(), CSS.AssignmentExpressionSyntax)
                    If initializer Is Nothing OrElse Not initializer.IsKind(CS.SyntaxKind.SimpleAssignmentExpression) Then
                        Return False
                    End If
                    If Not (TypeOf initializer.Left Is CSS.IdentifierNameSyntax) Then
                        Return False
                    End If
                    If (DirectCast(initializer.Left, CSS.IdentifierNameSyntax)).Identifier.IsEquivalentTo(iteratorIdentifier.Identifier) Then
                        Return False
                    End If
                    variable = initializer.Left.Accept(mNodesVisitor)
                    start = DirectCast(initializer.Right.Accept(mNodesVisitor), ExpressionSyntax)
                End If
                Dim ForStatementTrailingTrivia As New List(Of SyntaxTrivia)
                If node.CloseParenToken.HasTrailingTrivia Then
                    ForStatementTrailingTrivia.AddRange(ConvertTrivia(node.CloseParenToken.TrailingTrivia))
                End If
                Dim StatementFirstToken As SyntaxToken = node.Statement.GetFirstToken
                If StatementFirstToken.IsKind(CS.SyntaxKind.OpenBraceToken) AndAlso
                   (StatementFirstToken.LeadingTrivia.ContainsCommentOrDirectiveTrivia OrElse
                    StatementFirstToken.TrailingTrivia.ContainsCommentOrDirectiveTrivia) Then
                    Stop
                End If
                Dim ForKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ForKeyword).WithConvertedLeadingTriviaFrom(node.ForKeyword)
                Dim EqualsToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.EqualsToken)
                Dim ToKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ToKeyword)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim Statements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                block = SyntaxFactory.ForBlock(SyntaxFactory.ForStatement(forKeyword:=ForKeyword,
                                                                           controlVariable:=variable,
                                                                           equalsToken:=EqualsToken,
                                                                           fromValue:=start,
                                                                           toKeyword:=ToKeyword,
                                                                           toValue:=[end],
                                                                           stepClause:=If([step] = 1, Nothing, SyntaxFactory.ForStepClause(stepValue:=GetLiteralExpression([step], Token:=New SyntaxToken))
                                                                           )).WithTrailingTrivia(ForStatementTrailingTrivia),
                                                statements:=Statements,
                                                nextStatement:=SyntaxFactory.NextStatement().WithLeadingTrivia(OpenBraceTrailingTrivia).WithTrailingTrivia(ClosingBraceLeadingTrivia)
                                                )
                Return True
            End Function

            Private Function ConvertSingleExpression(ByVal node As CSS.ExpressionSyntax) As StatementSyntax
                Dim exprNode As VisualBasicSyntaxNode = node.Accept(mNodesVisitor)
                Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
                NewLeadingTrivia.AddRange(exprNode.GetLeadingTrivia)
                Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                NewTrailingTrivia.AddRange(exprNode.GetTrailingTrivia)
                NewTrailingTrivia.AddRange(node.GetTrailingTrivia)
                exprNode = exprNode.WithoutTrivia
                If Not (TypeOf exprNode Is StatementSyntax) Then
                    If TypeOf exprNode Is Syntax.ObjectCreationExpressionSyntax Then
                        Dim ModifiersList As SyntaxTokenList = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                        Dim variable As VariableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(SyntaxFactory.Identifier(GetUniqueVariableNameInScope(node, "tempVar", mSemanticModel)))), SyntaxFactory.AsNewClause(DirectCast(exprNode, NewExpressionSyntax)), initializer:=Nothing)

                        Dim SeparatedListOfvariableDeclarations As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(variable)
                        exprNode = SyntaxFactory.LocalDeclarationStatement(ModifiersList, SeparatedListOfvariableDeclarations)
                    ElseIf TypeOf exprNode Is Syntax.InvocationExpressionSyntax Then
                        exprNode = If(exprNode.GetFirstToken.IsKind(SyntaxKind.NewKeyword), SyntaxFactory.CallStatement(DirectCast(exprNode, ExpressionSyntax)), DirectCast(SyntaxFactory.ExpressionStatement(DirectCast(exprNode, ExpressionSyntax)), VisualBasicSyntaxNode))
                    Else
                        exprNode = SyntaxFactory.ExpressionStatement(DirectCast(exprNode, ExpressionSyntax))
                    End If
                End If
                Return DirectCast(exprNode, StatementSyntax).WithLeadingTrivia(NewLeadingTrivia).WithTrailingTrivia(NewTrailingTrivia)
            End Function

            Private Function ConvertSwitchLabel(ByVal label As CSS.CaseSwitchLabelSyntax) As CaseClauseSyntax
                Return SyntaxFactory.SimpleCaseClause(DirectCast(label.Value.Accept(mNodesVisitor), ExpressionSyntax))
            End Function

            Private Function ConvertSwitchSection(ByVal section As CSS.SwitchSectionSyntax) As CaseBlockSyntax
                Dim NewDimStatements As New List(Of Syntax.StatementSyntax)
                If section.Labels.OfType(Of CSS.DefaultSwitchLabelSyntax)().Any() Then
                    Return SyntaxFactory.CaseElseBlock(SyntaxFactory.CaseElseStatement(SyntaxFactory.ElseCaseClause()), ConvertSwitchSectionBlock(section, NewDimStatements))
                End If
                Dim LabelList As New List(Of CaseClauseSyntax)
                Dim LabelTrivia As New List(Of SyntaxTrivia)
                Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
                NewLeadingTrivia.AddRange(ConvertTrivia(section.GetLeadingTrivia))
                For Each CaseLabel As CSS.SwitchLabelSyntax In section.Labels
                    If TypeOf CaseLabel Is CSS.CaseSwitchLabelSyntax Then
                        LabelList.Add(ConvertSwitchLabel(DirectCast(CaseLabel, CaseSwitchLabelSyntax)))
                        LabelTrivia.AddRange(CaseLabel.GetTrailingTrivia.CollectAndConvertCommentTrivia)
                        Continue For
                    End If
                    If TypeOf CaseLabel Is CSS.CasePatternSwitchLabelSyntax Then
                        Dim PatternLabel As CSS.CasePatternSwitchLabelSyntax = DirectCast(CaseLabel, CasePatternSwitchLabelSyntax)
                        If TypeOf PatternLabel.Pattern Is ConstantPatternSyntax Then
                            Dim ConstantPattern As ConstantPatternSyntax = DirectCast(PatternLabel.Pattern, ConstantPatternSyntax).WithLeadingTrivia()
                            LabelList.Add(SyntaxFactory.SimpleCaseClause(DirectCast(ConstantPattern.Expression.Accept(mNodesVisitor), ExpressionSyntax)))
                            LabelTrivia.AddRange(CaseLabel.GetTrailingTrivia.CollectAndConvertCommentTrivia)
                        ElseIf TypeOf PatternLabel.Pattern Is DeclarationPatternSyntax Then
                            Dim Pattern As DeclarationPatternSyntax = DirectCast(PatternLabel.Pattern, DeclarationPatternSyntax)
                            Dim Type As TypeSyntax = DirectCast(Pattern.Type.Accept(mNodesVisitor), TypeSyntax)
                            Dim Identifier As SyntaxToken
                            If TypeOf Pattern.Designation Is SingleVariableDesignationSyntax Then
                                Identifier = GenerateSafeVBToken(DirectCast(Pattern.Designation, SingleVariableDesignationSyntax).Identifier, IsQualifiedName:=False)
                            ElseIf TypeOf Pattern.Designation Is DiscardDesignationSyntax Then
                                Identifier = SyntaxFactory.Identifier(GetUniqueVariableNameInScope(section, "tempVar", mSemanticModel))
                            Else
                                Stop
                            End If
                            Dim Modifiers As SyntaxTokenList
                            Modifiers = Modifiers.Add(SyntaxFactory.Token(SyntaxKind.DimKeyword))

                            If TypeOf Pattern.Designation IsNot DiscardDesignationSyntax Then
                                Dim SeparatedSyntaxList As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(Identifier))
                                Dim SwitchExpression As ExpressionSyntax = DirectCast(DirectCast(section.Parent, CSS.SwitchStatementSyntax).Expression.Accept(mNodesVisitor), ExpressionSyntax)
                                Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(SyntaxFactory.CTypeExpression(SwitchExpression, Type))
                                Dim variable As VariableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(SeparatedSyntaxList, SyntaxFactory.SimpleAsClause(Type), Initializer)
                                Dim Declarators As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(variable)

                                NewDimStatements.Add(SyntaxFactory.LocalDeclarationStatement(Modifiers, Declarators))
                                LabelList.Add(SyntaxFactory.SimpleCaseClause(DirectCast(Pattern.Designation.Accept(mNodesVisitor), ExpressionSyntax)))
                            Else
                                LabelList.Add(SyntaxFactory.SimpleCaseClause(DirectCast(Pattern.Type.Accept(mNodesVisitor), ExpressionSyntax)))
                            End If
                            If PatternLabel.WhenClause IsNot Nothing Then
                                NewLeadingTrivia.AddRange(PatternLabel.ConvertNodeToMultipleLineComment("TODO TASK: VB has no equivalent to the C# 'when' clause in 'case' statements, original lines are below, no translation was attempted"))
                            End If
                        End If
                    End If
                Next
                Dim CaseStatement As CaseStatementSyntax = SyntaxFactory.CaseStatement(SyntaxFactory.SeparatedList(LabelList)).With(NewLeadingTrivia, LabelTrivia)
                Return SyntaxFactory.CaseBlock(CaseStatement, ConvertSwitchSectionBlock(section, NewDimStatements))
            End Function

            Private Function ConvertSwitchSectionBlock(ByVal section As CSS.SwitchSectionSyntax, Statements As List(Of Syntax.StatementSyntax)) As SyntaxList(Of StatementSyntax)
                Dim lastStatement As CSS.StatementSyntax = section.Statements.LastOrDefault()
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                For Each s As CSS.StatementSyntax In section.Statements
                    If s Is lastStatement AndAlso TypeOf s Is CSS.BreakStatementSyntax Then Continue For
                    Statements.AddRange(ConvertBlock(s, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia))
                    If OpenBraceTrailingTrivia.Count > 0 Then
                        Stop
                    End If
                    If ClosingBraceLeadingTrivia.Count > 0 Then
                        Statements(Statements.Count - 1) = Statements.Last.WithAppendedTrailingTrivia(ClosingBraceLeadingTrivia)
                    End If
                Next
                Return SyntaxFactory.List(Statements)
            End Function

            Private Function MakeGotoSwitchLabel(ByVal expression As VisualBasicSyntaxNode) As String
                Dim expressionText As String = If(TypeOf expression Is ElseCaseClauseSyntax, "Default", expression.ToString())
                Return $"_Select{switchCount}_Case{expressionText.Replace(".", "_").Replace("""", "").Replace("[", "").Replace("]", "").Replace(" ", NameOf(Space))}"
            End Function

            Private Function TryConvertRaiseEvent(ByVal node As CSS.IfStatementSyntax, <Out> ByRef name As IdentifierNameSyntax, ByVal arguments As List(Of ArgumentSyntax)) As Boolean
                name = Nothing
                Dim condition As CSS.ExpressionSyntax = node.Condition
                While TypeOf condition Is CSS.ParenthesizedExpressionSyntax
                    condition = (DirectCast(condition, CSS.ParenthesizedExpressionSyntax)).Expression
                End While

                If Not (TypeOf condition Is CSS.BinaryExpressionSyntax) Then Return False
                Dim be As CSS.BinaryExpressionSyntax = DirectCast(condition, CSS.BinaryExpressionSyntax)
                If Not be.IsKind(CS.SyntaxKind.NotEqualsExpression) OrElse
                    (Not be.Left.IsKind(CS.SyntaxKind.NullLiteralExpression) AndAlso
                    Not be.Right.IsKind(CS.SyntaxKind.NullLiteralExpression)) Then
                    Return False
                End If
                Dim singleStatement As CSS.ExpressionStatementSyntax
                If TypeOf node.Statement Is CSS.BlockSyntax Then
                    Dim block As CSS.BlockSyntax = (DirectCast(node.Statement, CSS.BlockSyntax))
                    If block.Statements.Count <> 1 Then Return False
                    singleStatement = TryCast(block.Statements(0), CSS.ExpressionStatementSyntax)
                Else
                    singleStatement = TryCast(node.Statement, CSS.ExpressionStatementSyntax)
                End If

                If singleStatement Is Nothing OrElse Not (TypeOf singleStatement.Expression Is CSS.InvocationExpressionSyntax) Then Return False
                Dim possibleEventName As String = If(GetPossibleEventName(be.Left), GetPossibleEventName(be.Right))
                If possibleEventName Is Nothing Then Return False
                Dim invocation As CSS.InvocationExpressionSyntax = DirectCast(singleStatement.Expression, CSS.InvocationExpressionSyntax)
                Dim invocationName As String = GetPossibleEventName(invocation.Expression)
                If possibleEventName <> invocationName Then Return False
                name = SyntaxFactory.IdentifierName(possibleEventName)
                arguments.AddRange(invocation.ArgumentList.Arguments.[Select](Function(a As CSS.ArgumentSyntax) DirectCast(a.Accept(mNodesVisitor), ArgumentSyntax)))
                Return True
            End Function

            Private Function WillConvertToFor(ByVal node As CSS.ForStatementSyntax) As Boolean
                Dim hasVariable As Boolean = node.Declaration IsNot Nothing AndAlso node.Declaration.Variables.Count = 1
                If Not hasVariable AndAlso node.Initializers.Count <> 1 Then
                    Return False
                End If
                If node.Incrementors.Count <> 1 Then
                    Return False
                End If
                Dim Incrementors As VisualBasicSyntaxNode = node.Incrementors.FirstOrDefault()?.Accept(mNodesVisitor)
                Dim iterator As AssignmentStatementSyntax = TryCast(Incrementors, AssignmentStatementSyntax)
                If iterator Is Nothing OrElse Not iterator.IsKind(SyntaxKind.AddAssignmentStatement, SyntaxKind.SubtractAssignmentStatement) Then
                    Return False
                End If
                Dim iteratorIdentifier As IdentifierNameSyntax = TryCast(iterator.Left, IdentifierNameSyntax)
                If iteratorIdentifier Is Nothing Then
                    Return False
                End If
                Dim stepExpression As LiteralExpressionSyntax = TryCast(iterator.Right, LiteralExpressionSyntax)
                If stepExpression Is Nothing OrElse Not (TypeOf stepExpression.Token.Value Is Integer) Then
                    Return False
                End If
                Dim condition As CSS.BinaryExpressionSyntax = TryCast(node.Condition, CSS.BinaryExpressionSyntax)
                If condition Is Nothing OrElse Not (TypeOf condition.Left Is CSS.IdentifierNameSyntax) Then
                    Return False
                End If
                If (DirectCast(condition.Left, CSS.IdentifierNameSyntax)).Identifier.IsEquivalentTo(iteratorIdentifier.Identifier) Then
                    Return False
                End If
                If iterator.IsKind(SyntaxKind.SubtractAssignmentStatement) Then
                    If condition.IsKind(CS.SyntaxKind.GreaterThanOrEqualExpression) OrElse condition.IsKind(CS.SyntaxKind.NotEqualsExpression) Then
                    ElseIf condition.IsKind(CS.SyntaxKind.GreaterThanExpression) Then
                    Else
                        Return False
                    End If
                Else
                    If condition.IsKind(CS.SyntaxKind.LessThanOrEqualExpression) OrElse condition.IsKind(CS.SyntaxKind.NotEqualsExpression) Then
                    ElseIf condition.IsKind(CS.SyntaxKind.LessThanExpression) Then
                    Else
                        Return False
                    End If
                End If

                Dim start As ExpressionSyntax
                If hasVariable Then
                    Dim v As CSS.VariableDeclaratorSyntax = node.Declaration.Variables(0)
                    start = DirectCast(v.Initializer?.Value.Accept(mNodesVisitor), ExpressionSyntax)
                    If start Is Nothing Then
                        Return False
                    End If
                Else
                    Dim initializer As CSS.AssignmentExpressionSyntax = TryCast(node.Initializers.FirstOrDefault(), CSS.AssignmentExpressionSyntax)
                    If initializer Is Nothing OrElse Not initializer.IsKind(CS.SyntaxKind.SimpleAssignmentExpression) Then
                        Return False
                    End If
                    If Not (TypeOf initializer.Left Is CSS.IdentifierNameSyntax) Then
                        Return False
                    End If
                    If (DirectCast(initializer.Left, CSS.IdentifierNameSyntax)).Identifier.IsEquivalentTo(iteratorIdentifier.Identifier) Then
                        Return False
                    End If
                End If
                Return True
            End Function

            Public Shared Function GetUniqueVariableNameInScope(node As CS.CSharpSyntaxNode, variableNameBase As String, lSemanticModel As SemanticModel) As String
                Dim mWithBlockTempVariableNames As New List(Of String)
                Dim reservedNames As IEnumerable(Of String) = mWithBlockTempVariableNames.Concat(node.DescendantNodesAndSelf().SelectMany(Function(lSyntaxNode As SyntaxNode) lSemanticModel.LookupSymbols(lSyntaxNode.SpanStart).[Select](Function(s As ISymbol) s.Name))).Distinct
                Dim UniqueVariableName As String = EnsureUniqueness(variableNameBase, reservedNames, isCaseSensitive:=False)
                UsedIdentifiers.Add(UniqueVariableName, value:=New SymbolTableEntry(_Name:=UniqueVariableName, _IsType:=False))
                Return UniqueVariableName
            End Function
            Public Overrides Function DefaultVisit(ByVal node As SyntaxNode) As SyntaxList(Of StatementSyntax)
                Throw New NotImplementedException(node.[GetType]().ToString & " not implemented!")
            End Function

            Public Overrides Function VisitBlock(node As BlockSyntax) As SyntaxList(Of StatementSyntax)
                Dim ListOfStatements As SyntaxList(Of StatementSyntax) = SyntaxFactory.List(node.Statements.Where(Function(s As CSS.StatementSyntax) Not (TypeOf s Is CSS.EmptyStatementSyntax)).SelectMany(Function(s As CSS.StatementSyntax) s.Accept(Me)))
                If node.OpenBraceToken.HasLeadingTrivia OrElse node.OpenBraceToken.HasTrailingTrivia Then
                    ListOfStatements = ListOfStatements.Insert(0, SyntaxFactory.EmptyStatement.WithConvertedTriviaFrom(node.OpenBraceToken))
                End If
                If node.CloseBraceToken.HasLeadingTrivia OrElse node.OpenBraceToken.HasTrailingTrivia Then
                    ListOfStatements = ListOfStatements.Add(SyntaxFactory.EmptyStatement.WithConvertedTriviaFrom(node.CloseBraceToken))
                End If
                Return ListOfStatements
            End Function

            Public Overrides Function VisitBreakStatement(ByVal node As CSS.BreakStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim statementKind As SyntaxKind = SyntaxKind.None
                Dim keywordKind As SyntaxKind = SyntaxKind.None
                For Each stmt As CSS.StatementSyntax In node.GetAncestors(Of CSS.StatementSyntax)()
                    If TypeOf stmt Is CSS.DoStatementSyntax Then
                        statementKind = SyntaxKind.ExitDoStatement
                        keywordKind = SyntaxKind.DoKeyword
                        Exit For
                    End If

                    If TypeOf stmt Is CSS.WhileStatementSyntax Then
                        statementKind = SyntaxKind.ExitWhileStatement
                        keywordKind = SyntaxKind.WhileKeyword
                        Exit For
                    End If

                    If TypeOf stmt Is CSS.ForStatementSyntax Then
                        If WillConvertToFor(DirectCast(stmt, CSS.ForStatementSyntax)) Then
                            statementKind = SyntaxKind.ExitForStatement
                            keywordKind = SyntaxKind.ForKeyword
                            Exit For
                        Else
                            statementKind = SyntaxKind.ExitWhileStatement
                            keywordKind = SyntaxKind.WhileKeyword
                            Exit For
                        End If
                    ElseIf TypeOf stmt Is CSS.ForEachStatementSyntax Then
                        statementKind = SyntaxKind.ExitForStatement
                        keywordKind = SyntaxKind.ForKeyword
                        Exit For
                    End If

                    If TypeOf stmt Is CSS.SwitchStatementSyntax Then
                        statementKind = SyntaxKind.ExitSelectStatement
                        keywordKind = SyntaxKind.SelectKeyword
                        Exit For
                    End If
                Next

                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.ExitStatement(statementKind, SyntaxFactory.Token(keywordKind)).WithConvertedTriviaFrom(node))
            End Function

            Public Overrides Function VisitCheckedStatement(ByVal node As CSS.CheckedStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                If node.Keyword.ValueText = "checked" Then
                    Return WrapInComment(ConvertBlock(node.Block, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia), node, "Visual Basic Default Is checked math, check that this works for you!")
                End If
                Return WrapInComment(ConvertBlock(node.Block, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia), node, "Visual Basic does Not support unchecked statements!")
            End Function

            Public Overrides Function VisitContinueStatement(ByVal node As CSS.ContinueStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim statementKind As SyntaxKind = SyntaxKind.None
                Dim keywordKind As SyntaxKind = SyntaxKind.None
                For Each stmt As CSS.StatementSyntax In node.GetAncestors(Of CSS.StatementSyntax)()
                    If TypeOf stmt Is CSS.DoStatementSyntax Then
                        statementKind = SyntaxKind.ContinueDoStatement
                        keywordKind = SyntaxKind.DoKeyword
                        Exit For
                    End If

                    If TypeOf stmt Is CSS.WhileStatementSyntax Then
                        statementKind = SyntaxKind.ContinueWhileStatement
                        keywordKind = SyntaxKind.WhileKeyword
                        Exit For
                    End If

                    If TypeOf stmt Is CSS.ForStatementSyntax OrElse TypeOf stmt Is CSS.ForEachStatementSyntax Then
                        statementKind = SyntaxKind.ContinueForStatement
                        keywordKind = SyntaxKind.ForKeyword
                        Exit For
                    End If
                Next

                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.ContinueStatement(statementKind, SyntaxFactory.Token(keywordKind)).WithConvertedTriviaFrom(node))
            End Function

            Public Overrides Function VisitDeclarationExpression(node As DeclarationExpressionSyntax) As SyntaxList(Of StatementSyntax)
                Return MyBase.VisitDeclarationExpression(node)
            End Function

            Public Overrides Function VisitDoStatement(ByVal node As CSS.DoStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim condition As ExpressionSyntax = DirectCast(node.Condition.Accept(mNodesVisitor), ExpressionSyntax)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim stmt As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                    Stop
                End If
                Dim DoStatement As Syntax.DoStatementSyntax = SyntaxFactory.DoStatement(SyntaxKind.SimpleDoStatement)
                Dim LoopStatement As LoopStatementSyntax = SyntaxFactory.LoopStatement(SyntaxKind.LoopWhileStatement, SyntaxFactory.WhileClause(condition))
                Dim block As DoLoopBlockSyntax = SyntaxFactory.DoLoopWhileBlock(DoStatement, stmt, LoopStatement)
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(block)
            End Function

            Public Overrides Function VisitEmptyStatement(node As CSS.EmptyStatementSyntax) As SyntaxList(Of StatementSyntax)
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.EmptyStatement().WithConvertedTriviaFrom(node))
            End Function

            Public Overrides Function VisitExpressionStatement(ByVal node As CSS.ExpressionStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim TrailingTrivia As New List(Of SyntaxTrivia)
                Dim Modifiers As SyntaxTokenList = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                Dim statementSyntax1 As StatementSyntax
                Dim IdentifierString As String

                If node.GetFirstToken.IsKind(CS.SyntaxKind.NewKeyword) AndAlso node.Expression.IsKind(CS.SyntaxKind.SimpleAssignmentExpression) Then
                    ' Handle this case
                    'Dim x As New IO.FileInfo("C:\temp") With {.IsReadOnly = True}
                    IdentifierString = GetUniqueVariableNameInScope(node, "TempVar", mSemanticModel)
                    Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(IdentifierString)
                    Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(ModifiedIdentifier)
                    Dim exprNode As AssignmentStatementSyntax = DirectCast(node.Expression.Accept(mNodesVisitor), AssignmentStatementSyntax)
                    Dim AccessExpression As Syntax.MemberAccessExpressionSyntax = DirectCast(exprNode.Left, Syntax.MemberAccessExpressionSyntax)
                    Dim NewKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NewKeyword)
                    Dim ObjectCreateExpression As Syntax.ObjectCreationExpressionSyntax = DirectCast(AccessExpression.Expression, Syntax.ObjectCreationExpressionSyntax)
                    Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(exprNode.Right)
                    Dim FieldInitializer As FieldInitializerSyntax = SyntaxFactory.NamedFieldInitializer(name:=DirectCast(AccessExpression.Name, IdentifierNameSyntax), expression:=Initializer.Value)
                    Dim ObjectInitializer As ObjectMemberInitializerSyntax = SyntaxFactory.ObjectMemberInitializer(FieldInitializer)
                    Dim NewObject As Syntax.ObjectCreationExpressionSyntax = SyntaxFactory.ObjectCreationExpression(newKeyword:=NewKeyword, attributeLists:=ObjectCreateExpression.AttributeLists, type:=ObjectCreateExpression.Type, argumentList:=ObjectCreateExpression.ArgumentList, initializer:=ObjectInitializer)
                    Dim AsNewClause As AsNewClauseSyntax = SyntaxFactory.AsNewClause(newExpression:=NewObject.WithLeadingTrivia(SyntaxFactory.Space))
                    Dim Declarators As SeparatedSyntaxList(Of VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.VariableDeclarator(names:=Names, asClause:=AsNewClause, initializer:=Nothing))
                    statementSyntax1 = SyntaxFactory.LocalDeclarationStatement(Modifiers, Declarators)
                ElseIf node.ToString.StartsWith("(") AndAlso node.Expression.IsKind(CSharp.SyntaxKind.InvocationExpression) AndAlso DirectCast(node.Expression, CSS.InvocationExpressionSyntax).Expression.IsKind(CSharp.SyntaxKind.SimpleMemberAccessExpression) Then
                    ' Handle this case
                    ' (new Exception("RequestServiceAsync Timeout")).ReportServiceHubNFW("RequestServiceAsync Timeout");
                    Dim exprNode As Syntax.InvocationExpressionSyntax = DirectCast(node.Expression.Accept(mNodesVisitor), Syntax.InvocationExpressionSyntax)
                    Dim AccessExpression As Syntax.MemberAccessExpressionSyntax = DirectCast(exprNode.Expression, Syntax.MemberAccessExpressionSyntax)
                    Dim StmtList As New SyntaxList(Of StatementSyntax)
                    Dim Declarator As VariableDeclaratorSyntax
                    Select Case AccessExpression.Expression.Kind
                        Case SyntaxKind.CTypeExpression, SyntaxKind.TryCastExpression
                            IdentifierString = GetUniqueVariableNameInScope(node, "TempVar", mSemanticModel)
                            Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(IdentifierString)
                            Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(node:=ModifiedIdentifier)
                            Dim AsClause As AsClauseSyntax
                            Dim Initializer As EqualsValueSyntax
                            If AccessExpression.Expression.Kind = SyntaxKind.CTypeExpression Then
                                Dim CTypeExpression As Syntax.CTypeExpressionSyntax = CType(AccessExpression.Expression, Syntax.CTypeExpressionSyntax)
                                AsClause = SyntaxFactory.SimpleAsClause(CTypeExpression.Type.WithLeadingTrivia(SyntaxFactory.Space))
                                Initializer = SyntaxFactory.EqualsValue(CTypeExpression.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                            Else
                                Dim TryCastExpression As Syntax.TryCastExpressionSyntax = CType(AccessExpression.Expression, Syntax.TryCastExpressionSyntax)
                                AsClause = SyntaxFactory.SimpleAsClause(TryCastExpression.Type.WithLeadingTrivia(SyntaxFactory.Space))
                                Initializer = SyntaxFactory.EqualsValue(TryCastExpression.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                            End If
                            Declarator = SyntaxFactory.VariableDeclarator(Names, asClause:=AsClause, initializer:=Initializer)
                        Case SyntaxKind.ParenthesizedExpression
                            IdentifierString = GetUniqueVariableNameInScope(node, "TempVar", mSemanticModel)
                            Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(IdentifierString)
                            Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(node:=ModifiedIdentifier)
                            Dim ParenthesizedExpression As Syntax.ParenthesizedExpressionSyntax = CType(AccessExpression.Expression, Syntax.ParenthesizedExpressionSyntax)
                            Dim NewObject As Syntax.ObjectCreationExpressionSyntax = CType(ParenthesizedExpression.Expression, Syntax.ObjectCreationExpressionSyntax)
                            Dim AsNewClause As AsNewClauseSyntax = SyntaxFactory.AsNewClause(newExpression:=NewObject.WithLeadingTrivia(SyntaxFactory.Space))
                            Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(ParenthesizedExpression.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                            Declarator = SyntaxFactory.VariableDeclarator(Names, asClause:=AsNewClause, initializer:=Nothing)
                        Case SyntaxKind.IdentifierName
                            Declarator = Nothing
                            IdentifierString = AccessExpression.Expression.ToString
                        Case SyntaxKind.InvocationExpression
                            IdentifierString = GetUniqueVariableNameInScope(node, "TempVar", mSemanticModel)
                            Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(IdentifierString)
                            Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(node:=ModifiedIdentifier)
                            Dim MemberAccessExpression As Syntax.InvocationExpressionSyntax = CType(AccessExpression.Expression, Syntax.InvocationExpressionSyntax)
                            Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(MemberAccessExpression.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                            Declarator = SyntaxFactory.VariableDeclarator(Names, asClause:=Nothing, initializer:=Initializer)
                        Case SyntaxKind.SimpleMemberAccessExpression
                            IdentifierString = GetUniqueVariableNameInScope(node, "TempVar", mSemanticModel)
                            Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(IdentifierString)
                            Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(node:=ModifiedIdentifier)
                            Dim MemberAccessExpression As Syntax.MemberAccessExpressionSyntax = CType(AccessExpression.Expression, Syntax.MemberAccessExpressionSyntax)
                            Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(MemberAccessExpression.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                            Declarator = SyntaxFactory.VariableDeclarator(Names, asClause:=Nothing, initializer:=Initializer)
                        Case Else
                            Declarator = Nothing
                            IdentifierString = ""
                            Stop
                    End Select
                    If Declarator IsNot Nothing Then
                        Dim Declarators As SeparatedSyntaxList(Of VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(Declarator)
                        StmtList = StmtList.Add(SyntaxFactory.LocalDeclarationStatement(Modifiers, Declarators).WithConvertedLeadingTriviaFrom(node))
                    End If
                    With AccessExpression
                        Dim Identifier As ExpressionSyntax = SyntaxFactory.IdentifierName(IdentifierString)
                        Dim MemberAccessExpression As ExpressionSyntax = SyntaxFactory.MemberAccessExpression(kind:= .Kind, expression:=Identifier, operatorToken:= .OperatorToken, name:= .Name)
                        Dim AgrumentList As Syntax.ArgumentListSyntax = CType(DirectCast(node.Expression, CSS.InvocationExpressionSyntax).ArgumentList.Accept(mNodesVisitor), Syntax.ArgumentListSyntax)
                        Dim InvocationExpression As Syntax.InvocationExpressionSyntax = SyntaxFactory.InvocationExpression(MemberAccessExpression, AgrumentList).WithConvertedTrailingTriviaFrom(node)
                        StmtList = StmtList.Add(SyntaxFactory.ExpressionStatement(InvocationExpression))
                    End With
                    Return ReplaceStatementWithMarkedStatement(node, StmtList)
                Else
                    statementSyntax1 = ConvertSingleExpression(node.Expression)
                End If
                Dim FoundEOL As Boolean = False
                FoundEOL = RestructureTrailingTrivia(OriginalTrailingTrivia:=statementSyntax1.GetTrailingTrivia, FoundEOL:=FoundEOL, NewTrailingTrivia:=TrailingTrivia)
                FoundEOL = RestructureTrailingTrivia(ConvertTrivia(triviaToConvert:=node.GetTrailingTrivia), FoundEOL:=FoundEOL, NewTrailingTrivia:=TrailingTrivia)

                If FoundEOL Then
                    TrailingTrivia.Add(VB_EOLTrivia)
                End If
                Dim syntaxList1 As SyntaxList(Of StatementSyntax) = SyntaxFactory.SingletonList(statementSyntax1.WithConvertedLeadingTriviaFrom(node).WithTrailingTrivia(TrailingTrivia))
                Return ReplaceStatementWithMarkedStatement(node, syntaxList1)
            End Function

            Public Overrides Function VisitFixedStatement(node As FixedStatementSyntax) As SyntaxList(Of StatementSyntax)
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.EmptyStatement.WithLeadingTrivia(node.ConvertNodeToMultipleLineComment("TODO TASK: Fixed is not support by VB, original lines are below, no translation was attempted")))
            End Function

            Public Overrides Function VisitForEachStatement(ByVal node As CSS.ForEachStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim variable As VisualBasicSyntaxNode
#Disable Warning CC0014 ' Use Ternary operator.
                If node.Type.IsVar Then
#Enable Warning CC0014 ' Use Ternary operator.
                    variable = SyntaxFactory.IdentifierName(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False))
                Else
                    variable = SyntaxFactory.VariableDeclarator(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False))), asClause:=SyntaxFactory.SimpleAsClause(DirectCast(node.Type.Accept(mNodesVisitor), TypeSyntax)), initializer:=Nothing)
                End If

                Dim expression As ExpressionSyntax = DirectCast(node.Expression.Accept(mNodesVisitor), ExpressionSyntax)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim stmt As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                'If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                '    Stop
                'End If
                Dim NextStatement As NextStatementSyntax = SyntaxFactory.NextStatement().WithLeadingTrivia(ClosingBraceLeadingTrivia)
                Dim block As ForEachBlockSyntax = SyntaxFactory.ForEachBlock(SyntaxFactory.ForEachStatement(variable, expression), stmt, NextStatement)
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(block)
            End Function

            Public Overrides Function VisitForEachVariableStatement(node As ForEachVariableStatementSyntax) As SyntaxList(Of StatementSyntax)
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(NodesVisitor.CommentOutUnsupportedStatements(node, "For Each Variable Statement"))
            End Function

            Public Overrides Function VisitForStatement(ByVal node As CSS.ForStatementSyntax) As SyntaxList(Of StatementSyntax)
                '   ForStatement -> ForNextStatement when for-loop is simple

                ' only the following forms of the for-statement are allowed:
                ' for (TypeReference name = start; name < oneAfterEnd; name += step)
                ' for (name = start; name < oneAfterEnd; name += step)
                ' for (TypeReference name = start; name <= end; name += step)
                ' for (name = start; name <= end; name += step)
                ' for (TypeReference name = start; name > oneAfterEnd; name -= step)
                ' for (name = start; name > oneAfterEnd; name -= step)
                ' for (TypeReference name = start; name >= end; name -= step)
                ' for (name = start; name >= end; name -= step)
                Dim block As StatementSyntax = Nothing

                ' check if the form Is valid And collect TypeReference, name, start, end And step
                If Not ConvertForToSimpleForNext(node, block) Then
                    Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                    Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                    Dim stmts As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia).AddRange(node.Incrementors.[Select](AddressOf ConvertSingleExpression))
                    If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                        Stop
                    End If

                    Dim condition As ExpressionSyntax = If(node.Condition Is Nothing, Literal(True), DirectCast(node.Condition.Accept(mNodesVisitor), ExpressionSyntax))
                    block = SyntaxFactory.WhileBlock(SyntaxFactory.WhileStatement(condition), stmts)
                    Return SyntaxFactory.List(node.Initializers.[Select](AddressOf ConvertSingleExpression)).Add(block)
                End If

                Return SyntaxFactory.SingletonList(block)
            End Function

            Public Overrides Function VisitGotoStatement(ByVal node As CSS.GotoStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim label As LabelSyntax
                If node.IsKind(CS.SyntaxKind.GotoCaseStatement, CS.SyntaxKind.GotoDefaultStatement) Then
                    If mBlockInfo.Count = 0 Then Throw New InvalidOperationException("GoTo Case/GoTo Default outside switch Is illegal!")
                    Dim labelExpression As VisualBasicSyntaxNode = If(node.Expression?.Accept(mNodesVisitor), SyntaxFactory.ElseCaseClause())
                    mBlockInfo.Peek().GotoCaseExpressions.Add(labelExpression)
                    label = SyntaxFactory.Label(SyntaxKind.IdentifierLabel, MakeGotoSwitchLabel(labelExpression))
                Else
                    label = SyntaxFactory.Label(SyntaxKind.IdentifierLabel, GenerateSafeVBToken((DirectCast(node.Expression, CSS.IdentifierNameSyntax)).Identifier, IsQualifiedName:=False))
                End If

                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.GoToStatement(label).WithConvertedTriviaFrom(node))
            End Function

            Public Overrides Function VisitIfStatement(ByVal node As CSS.IfStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim name As IdentifierNameSyntax = Nothing
                Dim arguments As New List(Of ArgumentSyntax)()
                Dim stmt As StatementSyntax
                If node.[Else] Is Nothing AndAlso TryConvertRaiseEvent(node, name, arguments) Then
                    stmt = SyntaxFactory.RaiseEventStatement(name, SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList(arguments)))
                    Return SyntaxFactory.SingletonList(stmt)
                End If

                Dim ListOfElseIfBlocks As New List(Of ElseIfBlockSyntax)()
                Dim ElseBlock As ElseBlockSyntax = Nothing
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                CollectElseBlocks(node, ListOfElseIfBlocks, ElseBlock, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                Dim IFKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.IfKeyword).WithConvertedTriviaFrom(node.IfKeyword)
                Dim ThenKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ThenKeyword).WithConvertedTrailingTriviaFrom(node.CloseParenToken)
                Dim ConditionWithTrivia As ExpressionSyntax = DirectCast(node.Condition.Accept(mNodesVisitor), ExpressionSyntax).WithAppendedTrailingTrivia(ConvertTrivia(node.Condition.GetTrailingTrivia))
                If ConditionWithTrivia.HasTrailingTrivia Then
                    Dim ThenTrailingTrivia As New List(Of SyntaxTrivia)
                    Dim ConditionPreservedTrailingTrivia As New List(Of SyntaxTrivia)
                    RestructureConditionThenTrailingTrivia(ConvertTrivia(node.Condition.GetTrailingTrivia), ThenTrailingTrivia, ConditionPreservedTrailingTrivia)
                    ConditionWithTrivia = ConditionWithTrivia.WithTrailingTrivia(ConditionPreservedTrailingTrivia)
                    ThenKeyword = ThenKeyword.WithTrailingTrivia(ThenTrailingTrivia)
                End If
                Dim Braces As (SyntaxToken, SyntaxToken) = node.Statement.GetBraces
                Dim OpenBraces As SyntaxToken = Braces.Item1
                Dim CloseBraces As SyntaxToken = Braces.Item2
                Dim EndIfStatement As EndBlockStatementSyntax = SyntaxFactory.EndIfStatement(SyntaxFactory.Token(SyntaxKind.EndKeyword), SyntaxFactory.Token(SyntaxKind.IfKeyword)).WithConvertedTriviaFrom(CloseBraces)
                Dim ElseIfBlocks As SyntaxList(Of ElseIfBlockSyntax) = SyntaxFactory.List(ListOfElseIfBlocks)
                If ElseBlock IsNot Nothing AndAlso ElseBlock.Statements(0).IsKind(SyntaxKind.EmptyStatement) Then
                    EndIfStatement = EndIfStatement.WithLeadingTrivia(ElseBlock.GetTrailingTrivia)
                    ElseBlock = SyntaxFactory.ElseBlock(SyntaxFactory.ElseStatement(), statements:=Nothing)
                End If
                Dim Statements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                If ClosingBraceLeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                    EndIfStatement = EndIfStatement.WithLeadingTrivia(ClosingBraceLeadingTrivia)
                End If
                If TypeOf node.Statement Is CSS.BlockSyntax Then
                    stmt = SyntaxFactory.MultiLineIfBlock(SyntaxFactory.IfStatement(IFKeyword, ConditionWithTrivia, ThenKeyword),
                                                          Statements,
                                                          ElseIfBlocks,
                                                          ElseBlock,
                                                          EndIfStatement)
                Else
                    Dim IsInvocationExpression As Boolean = False
                    If node.Statement.IsKind(CSharp.SyntaxKind.ExpressionStatement) Then
                        Dim ExpressionStatement As CSS.ExpressionStatementSyntax = DirectCast(node.Statement, CSS.ExpressionStatementSyntax)
                        If ExpressionStatement.Expression.IsKind(CSharp.SyntaxKind.InvocationExpression) Then
                            IsInvocationExpression = ExpressionStatement.Expression.DescendantNodes().OfType(Of CSS.ConditionalExpressionSyntax).Any
                        End If
                    End If
                    If node.Statement.IsKind(CSharp.SyntaxKind.EmptyStatement) Then
                        ThenKeyword = ThenKeyword.WithAppendedTrailingTrivia(ConvertTrivia(node.Statement.GetTrailingTrivia))
                    End If
                    If ListOfElseIfBlocks.Any() OrElse
                        Not IsSimpleStatement(node.Statement) OrElse
                        IsInvocationExpression Then
                        stmt = SyntaxFactory.MultiLineIfBlock(SyntaxFactory.IfStatement(IFKeyword, ConditionWithTrivia, ThenKeyword),
                                                                  Statements,
                                                                  ElseIfBlocks,
                                                                  ElseBlock,
                                                                  EndIfStatement)
                    Else
                        Dim IfStatement As Syntax.IfStatementSyntax = SyntaxFactory.IfStatement(IFKeyword, ConditionWithTrivia, ThenKeyword)
                        If ThenKeyword.TrailingTrivia.ContainsEOLTrivia Then
                            Dim IFBlockStatements As SyntaxList(Of StatementSyntax) = Statements
                            stmt = SyntaxFactory.MultiLineIfBlock(ifStatement:=IfStatement,
                                                                    statements:=IFBlockStatements,
                                                                    elseIfBlocks:=ElseIfBlocks,
                                                                    elseBlock:=If(ElseBlock, Nothing),
                                                                    endIfStatement:=EndIfStatement)
                        Else
#Disable Warning CC0014 ' Use Ternary operator.
                            If ElseBlock IsNot Nothing Then
#Enable Warning CC0014 ' Use Ternary operator.
                                stmt = SyntaxFactory.MultiLineIfBlock(ifStatement:=IfStatement,
                                                                     statements:=Statements,
                                                                     elseIfBlocks:=Nothing,
                                                                     elseBlock:=ElseBlock,
                                                                     endIfStatement:=EndIfStatement)
                            Else
                                stmt = SyntaxFactory.SingleLineIfStatement(ifKeyword:=IFKeyword,
                                                                           condition:=ConditionWithTrivia, thenKeyword:=ThenKeyword,
                                                                           statements:=Statements,
                                                                           elseClause:=Nothing)
                            End If
                        End If
                    End If
                End If

                Return ReplaceStatementWithMarkedStatement(node, SyntaxFactory.SingletonList(stmt.WithConvertedLeadingTriviaFrom(node)))
            End Function

            Public Overrides Function VisitLabeledStatement(ByVal node As CSS.LabeledStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim Statements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                    Stop
                End If
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.LabelStatement(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False))).AddRange(Statements)
            End Function

            Public Overrides Function VisitLocalDeclarationStatement(ByVal node As CSS.LocalDeclarationStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, mNodesVisitor.IsModule, TokenContext.Local)
                If modifiers.Count = 0 Then
                    modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                End If
                Dim LeadingTrivia As New List(Of SyntaxTrivia)
                Dim declarators As SeparatedSyntaxList(Of VariableDeclaratorSyntax) = RemodelVariableDeclaration(node.Declaration, mNodesVisitor, LeadingTrivia)
                Dim localDeclarationStatement As Syntax.LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(
                    modifiers,
                    declarators).WithoutTrivia.WithLeadingTrivia(LeadingTrivia).WithAppendedTrailingTrivia(ConvertTrivia(node.GetTrailingTrivia)) ' this picks up end of line comments
                ' Don't repeat leading comments
                If Not localDeclarationStatement.GetLeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                    localDeclarationStatement = localDeclarationStatement.WithConvertedLeadingTriviaFrom(node)
                End If
                Dim syntaxList1 As SyntaxList(Of StatementSyntax) = SyntaxFactory.SingletonList(Of StatementSyntax)(
                    localDeclarationStatement
                    )

                Return ReplaceStatementWithMarkedStatement(node, syntaxList1)
            End Function

            Public Overrides Function VisitLocalFunctionStatement(node As LocalFunctionStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim syntaxList1 As SyntaxList(Of StatementSyntax) = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.EmptyStatement.WithLeadingTrivia(node.ConvertNodeToMultipleLineComment("TODO TASK: Local Functions are not support by VB, original lines are below, no translation was attempted")))
                Return syntaxList1
            End Function

            Public Overrides Function VisitLockStatement(ByVal node As CSS.LockStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim LockStatement As SyncLockStatementSyntax = SyntaxFactory.SyncLockStatement(DirectCast(node.Expression?.Accept(mNodesVisitor), ExpressionSyntax))
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim Statements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                    Stop
                End If
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.SyncLockBlock(LockStatement, Statements).WithConvertedTriviaFrom(node))
            End Function

            Public Overrides Function VisitReturnStatement(ByVal node As CSS.ReturnStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim stmt As StatementSyntax
                If node.Expression Is Nothing Then
                    stmt = SyntaxFactory.ReturnStatement().WithConvertedTriviaFrom(node)
                Else
                    Dim Expression As ExpressionSyntax = DirectCast(node.Expression.Accept(mNodesVisitor), ExpressionSyntax)
                    Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
                    Dim ExpressiontLeadingTrivia As SyntaxTriviaList = Expression.GetLeadingTrivia
                    If ExpressiontLeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                        NewLeadingTrivia.AddRange(ExpressiontLeadingTrivia)
                        Expression = Expression.WithLeadingTrivia(SyntaxFactory.ElasticSpace)
                    End If
                    NewLeadingTrivia.AddRange(ConvertTrivia(node.GetLeadingTrivia))
                    stmt = SyntaxFactory.ReturnStatement(Expression).With(NewLeadingTrivia, ConvertTrivia(node.GetTrailingTrivia))
                End If
                Return ReplaceStatementWithMarkedStatement(node, SyntaxFactory.SingletonList(stmt))
            End Function

            Public Overrides Function VisitSwitchStatement(ByVal node As CSS.SwitchStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim stmt As StatementSyntax
                mBlockInfo.Push(New BlockInfo())
                Try
                    Dim blocks As List(Of CaseBlockSyntax) = node.Sections.[Select](AddressOf ConvertSwitchSection).ToList
                    Dim OrderedBlocks As New List(Of CaseBlockSyntax)
                    Dim CaseElseIndex As Integer = -1
                    For i As Integer = 0 To blocks.Count - 1
                        If blocks(i).Kind = SyntaxKind.CaseElseBlock Then
                            CaseElseIndex = i
                        Else
                            OrderedBlocks.Add(blocks(i))
                        End If
                    Next
                    If CaseElseIndex >= 0 Then
                        OrderedBlocks.Add(blocks(CaseElseIndex))
                    End If
                    Dim SelectKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.SelectKeyword)
                    Dim CaseKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CaseKeyword)
                    Dim Expression As ExpressionSyntax = DirectCast(node.Expression.Accept(mNodesVisitor), ExpressionSyntax)
                    stmt = SyntaxFactory.SelectBlock(SyntaxFactory.SelectStatement(selectKeyword:=SelectKeyword,
                                                                                   caseKeyword:=CaseKeyword,
                                                                                   expression:=Expression),
                                                     SyntaxFactory.List(nodes:=AddLabels(blocks:=OrderedBlocks.ToArray,
                                                                                         gotoLabels:=mBlockInfo.Peek().GotoCaseExpressions)))
                    switchCount += 1
                Finally
                    mBlockInfo.Pop()
                End Try
                Return ReplaceStatementWithMarkedStatement(node, SyntaxFactory.SingletonList(stmt.WithConvertedLeadingTriviaFrom(node)))
            End Function

            Public Overrides Function VisitThrowStatement(ByVal node As CSS.ThrowStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim stmt As StatementSyntax
                stmt = If(node.Expression Is Nothing,
                            SyntaxFactory.ThrowStatement(),
                            SyntaxFactory.ThrowStatement(DirectCast(node.Expression.Accept(mNodesVisitor), ExpressionSyntax)))
                Return ReplaceStatementWithMarkedStatement(node, SyntaxFactory.SingletonList(stmt.WithConvertedTriviaFrom(node)))
            End Function

            Public Overrides Function VisitTryStatement(ByVal node As CSS.TryStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim TryStatement As Syntax.TryStatementSyntax = SyntaxFactory.TryStatement()
                Dim CatchBlocks As SyntaxList(Of CatchBlockSyntax) = SyntaxFactory.List(node.Catches.IndexedSelect(AddressOf ConvertCatchClause))
                Dim TriviaList As New List(Of SyntaxTrivia)
                Dim LastCatchBlockIndex As Integer = CatchBlocks.Count - 1
                For i As Integer = 0 To LastCatchBlockIndex
                    Dim CatchBlock As CatchBlockSyntax = CatchBlocks(i)
                    If CatchBlock.Statements(0).IsKind(SyntaxKind.EmptyStatement) Then
                        Dim TempTriviaList As New List(Of SyntaxTrivia)
                        TempTriviaList.AddRange(CatchBlock.Statements(0).GetTrailingTrivia)
                        TriviaList.AddRange(CatchBlock.GetBraces().Item2.LeadingTrivia)
                        CatchBlocks.Replace(CatchBlocks(i), SyntaxFactory.CatchBlock(CatchBlock.CatchStatement.WithLeadingTrivia(TriviaList)))
                        TriviaList = TempTriviaList
                    Else
                        CatchBlocks = CatchBlocks.Replace(CatchBlock, CatchBlock.WithLeadingTrivia(TriviaList))
                        TriviaList.Clear()
                    End If
                Next
                If LastCatchBlockIndex >= 0 Then
                    CatchBlocks = CatchBlocks.Replace(CatchBlocks(0), CatchBlocks(0).WithConvertedTriviaFrom(node.Block.CloseBraceToken))
                End If
                Dim FinallyBlock As FinallyBlockSyntax = Nothing
                If node.Finally IsNot Nothing Then
                    Dim FinallyStatements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.[Finally].Block, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                    If OpenBraceTrailingTrivia.Count > 0 OrElse ClosingBraceLeadingTrivia.Count > 0 Then
                        Stop
                    End If
                    FinallyBlock = SyntaxFactory.FinallyBlock(FinallyStatements).WithPrependedLeadingTrivia(TriviaList)
                    TriviaList.Clear()
                    If FinallyBlock.Statements(0).IsKind(SyntaxKind.EmptyStatement) Then
                        TriviaList.AddRange(FinallyBlock.Statements(0).GetTrailingTrivia)
                        FinallyBlock = FinallyBlock.WithTrailingTrivia(VB_EOLTrivia)
                    End If
                End If
                Dim EndTryStatement As EndBlockStatementSyntax = SyntaxFactory.EndTryStatement()
                If TriviaList.Any Then
                    EndTryStatement = EndTryStatement.WithLeadingTrivia(TriviaList)
                End If
                Dim TryBlockStatements As SyntaxList(Of StatementSyntax) = ConvertBlock(node.Block, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)
                Dim block As TryBlockSyntax = SyntaxFactory.TryBlock(TryStatement,
                                                                    TryBlockStatements,
                                                                    CatchBlocks,
                                                                    FinallyBlock,
                                                                    EndTryStatement
                                                                    )
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(block.WithConvertedTriviaFrom(node))
            End Function

            Public Overrides Function VisitUnsafeStatement(node As UnsafeStatementSyntax) As SyntaxList(Of StatementSyntax)
                Return SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.EmptyStatement.WithLeadingTrivia(node.ConvertNodeToMultipleLineComment("TODO TASK: Unsafe is not support by VB, original lines are below, no translation was attempted")))
            End Function

            Public Overrides Function VisitUsingStatement(ByVal node As CSS.UsingStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim UsingStmt As UsingStatementSyntax
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim Stmt As SyntaxList(Of StatementSyntax)
                Dim LeadingTrivia As New List(Of SyntaxTrivia)
                If node.Declaration Is Nothing Then
                    Dim UsingBlock As UsingBlockSyntax
                    LeadingTrivia.AddRange(ConvertTrivia(node.GetLeadingTrivia))
                    If node.Expression IsNot Nothing AndAlso node.Expression.IsKind(CS.SyntaxKind.ConditionalAccessExpression) Then
                        Dim NothingToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NothingKeyword)
                        Dim CS_ConditionalAccessExpression As CSS.ConditionalAccessExpressionSyntax = DirectCast(node.Expression, CSS.ConditionalAccessExpressionSyntax)
                        Dim VB_ConditionalAccessExpression As VisualBasicSyntaxNode = CS_ConditionalAccessExpression.Expression.Accept(mNodesVisitor)
                        Dim Condition As Syntax.BinaryExpressionSyntax = SyntaxFactory.IsNotExpression(left:=CType(VB_ConditionalAccessExpression, ExpressionSyntax), right:=SyntaxFactory.NothingLiteralExpression(token:=NothingToken))
                        Dim IfStatement As Syntax.IfStatementSyntax = SyntaxFactory.IfStatement(Condition)
                        UsingStmt = SyntaxFactory.UsingStatement(SyntaxFactory.ParseExpression($"{VB_ConditionalAccessExpression.ToString}{CS_ConditionalAccessExpression.WhenNotNull.Accept(mNodesVisitor)}"), SyntaxFactory.SeparatedList(Of VariableDeclaratorSyntax)())
                        UsingBlock = SyntaxFactory.UsingBlock(UsingStmt, ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)).WithLeadingTrivia(LeadingTrivia)
                        Dim IfStatementBlock As MultiLineIfBlockSyntax = SyntaxFactory.MultiLineIfBlock(IfStatement, SyntaxFactory.SingletonList(Of StatementSyntax)(UsingBlock), elseIfBlocks:=Nothing, elseBlock:=Nothing).WithLeadingTrivia(LeadingTrivia)
                        Stmt = SyntaxFactory.SingletonList(Of StatementSyntax)(IfStatementBlock)
                        Return ReplaceStatementWithMarkedStatement(node, Stmt)
                    Else
                        UsingStmt = SyntaxFactory.UsingStatement(DirectCast(node.Expression?.Accept(mNodesVisitor), ExpressionSyntax), SyntaxFactory.SeparatedList(Of VariableDeclaratorSyntax)())
                    End If
                Else
                    UsingStmt = SyntaxFactory.UsingStatement(expression:=Nothing, variables:=RemodelVariableDeclaration(node.Declaration, mNodesVisitor, LeadingTrivia))
                End If

                Stmt = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.UsingBlock(usingStatement:=UsingStmt, statements:=ConvertBlock(node.Statement, OpenBraceTrailingTrivia, ClosingBraceLeadingTrivia)).WithLeadingTrivia(LeadingTrivia))
                Return ReplaceStatementWithMarkedStatement(node:=node, Statements:=Stmt)
            End Function

            Public Overrides Function VisitWhileStatement(ByVal node As CSS.WhileStatementSyntax) As SyntaxList(Of StatementSyntax)
                Dim condition As ExpressionSyntax = DirectCast(node.Condition.Accept(mNodesVisitor), ExpressionSyntax)
                Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                Dim ClosingBraceLeadingTrivia As New List(Of SyntaxTrivia)
                Dim WhileStatements As SyntaxList(Of StatementSyntax) = ConvertBlock(node:=node.Statement, OpenBraceTrailiningTrivia:=OpenBraceTrailingTrivia, CloseBraceLeadingTrivia:=ClosingBraceLeadingTrivia)
                If OpenBraceTrailingTrivia.Count > 0 Then
                    Stop
                End If
                Dim EndWhileStatement As EndBlockStatementSyntax = SyntaxFactory.EndWhileStatement().WithLeadingTrivia(ClosingBraceLeadingTrivia)
                Dim block As WhileBlockSyntax = SyntaxFactory.WhileBlock(SyntaxFactory.WhileStatement(condition).WithConvertedLeadingTriviaFrom(node.WhileKeyword), WhileStatements, EndWhileStatement)
                Dim WhileStatementBlock As SyntaxList(Of StatementSyntax) = SyntaxFactory.SingletonList(Of StatementSyntax)(block)
                Return ReplaceStatementWithMarkedStatement(node:=node, Statements:=WhileStatementBlock)
            End Function

            Public Overrides Function VisitYieldStatement(ByVal node As CSS.YieldStatementSyntax) As SyntaxList(Of StatementSyntax)
                IsInterator = True
                Dim stmt As StatementSyntax
#Disable Warning CC0014 ' Use Ternary operator.
                If node.Expression Is Nothing Then
#Enable Warning CC0014 ' Use Ternary operator.
                    stmt = SyntaxFactory.ReturnStatement()
                Else
                    stmt = SyntaxFactory.YieldStatement(expression:=DirectCast(node.Expression.Accept(mNodesVisitor), ExpressionSyntax))
                End If
                Return SyntaxFactory.SingletonList(stmt.WithConvertedTriviaFrom(node))
            End Function

            Private Class BlockInfo
                Public ReadOnly GotoCaseExpressions As List(Of VisualBasicSyntaxNode) = New List(Of VisualBasicSyntaxNode)()
            End Class
        End Class
    End Class
End Namespace
