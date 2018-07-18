Option Explicit On
Option Infer Off
Option Strict On
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.Simplification
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports ArgumentListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArgumentListSyntax
Imports ArgumentSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArgumentSyntax
Imports ArrayRankSpecifierSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArrayRankSpecifierSyntax
Imports AttributeListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.AttributeListSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports IdentifierNameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.IdentifierNameSyntax
Imports InterpolatedStringContentSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.InterpolatedStringContentSyntax
Imports LambdaExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.LambdaExpressionSyntax
Imports ParameterListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ParameterListSyntax
Imports ParameterSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ParameterSyntax
Imports ReturnStatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ReturnStatementSyntax
Imports SimpleNameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.SimpleNameSyntax
Imports StatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.StatementSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax
Imports YieldStatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.YieldStatementSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)

            Private ReadOnly LiteralExpression_1 As ExpressionSyntax = SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(1))
            Private ReadOnly LiteralNothing As Syntax.LiteralExpressionSyntax = SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword))
            Private Shared Function CheckCorrectnessLeadingTrivia(StatementWithIssue As CS.CSharpSyntaxNode, MessageFragment As String) As SyntaxTriviaList
                Dim LeadingTrivia As New SyntaxTriviaList
                LeadingTrivia = LeadingTrivia.Add(SyntaxFactory.CommentTrivia($"' TODO TASK: {MessageFragment}:"))
                LeadingTrivia = LeadingTrivia.Add(SyntaxFactory.CommentTrivia($"' Original Statement:"))
                LeadingTrivia = LeadingTrivia.AddRange(StatementWithIssue.ConvertNodeToMultipleLineComment)
                LeadingTrivia = LeadingTrivia.Add(SyntaxFactory.CommentTrivia($"' An attempt was made to correctly port the code,check the code below for correctness"))
                Return LeadingTrivia
            End Function
            Private Shared Function Convert_ToObject(Type As String) As String
                Return If(Type = "_", "Object", Type)
            End Function

            Private Shared Function ConvertToInterpolatedStringTextToken(CSharpToken As SyntaxToken) As SyntaxToken
                Dim TokenString As String = ConvertCSharpEscapes(CSharpToken.ValueText)
                Return SyntaxFactory.InterpolatedStringTextToken(TokenString, TokenString)
            End Function
            Private Shared Sub RestructureCloseTokenTrivia(CS_Token As SyntaxToken, ByRef CloseToken As SyntaxToken)
                Dim TrailingTrivia As New List(Of SyntaxTrivia)
                If CS_Token.HasLeadingTrivia Then
                    For Each T As SyntaxTrivia In ConvertTrivia(CS_Token.LeadingTrivia)
                        Select Case T.Kind
                            Case SyntaxKind.CommentTrivia
                                TrailingTrivia.Add(T)
                            Case SyntaxKind.EndOfLineTrivia, SyntaxKind.WhitespaceTrivia, SyntaxKind.None
                                ' ignore
                            Case Else
                                Stop
                        End Select
                    Next
                End If
                For Each T As SyntaxTrivia In ConvertTrivia(CS_Token.TrailingTrivia)
                    Select Case T.Kind
                        Case SyntaxKind.CommentTrivia
                            TrailingTrivia.Add(T)
                        Case SyntaxKind.EndOfLineTrivia, SyntaxKind.None, SyntaxKind.WhitespaceTrivia
                            ' ignore
                        Case Else
                            Stop
                    End Select
                Next
                TrailingTrivia.AddRange(CloseToken.TrailingTrivia)
                CloseToken = CloseToken.WithTrailingTrivia(TrailingTrivia)
            End Sub

            Private Shared Function RestructureTrivia(TriviaList As SyntaxTriviaList, FoundEOL As Boolean, ByRef OperatorTrailingTrivia As List(Of SyntaxTrivia)) As Boolean
                For Each t As SyntaxTrivia In TriviaList
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            OperatorTrailingTrivia.Add(t)
                            FoundEOL = True
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            OperatorTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                        Case Else
                            Stop
                    End Select
                Next

                Return FoundEOL
            End Function

            Private Shared Function UnpackExpressionFromStatement(ByVal statementSyntax As StatementSyntax, <Out> ByRef expression As ExpressionSyntax) As Boolean
                If TypeOf statementSyntax Is ReturnStatementSyntax Then
                    expression = (DirectCast(statementSyntax, ReturnStatementSyntax)).Expression
                ElseIf TypeOf statementSyntax Is YieldStatementSyntax Then
                    expression = (DirectCast(statementSyntax, YieldStatementSyntax)).Expression
                Else
                    expression = Nothing
                End If

                Return expression IsNot Nothing
            End Function

            Private Function ConvertLambdaExpression(ByVal node As CSS.AnonymousFunctionExpressionSyntax, ByVal block As Object, ByVal parameters As SeparatedSyntaxList(Of CSS.ParameterSyntax), ByVal modifiers As SyntaxTokenList) As LambdaExpressionSyntax
                Dim symbol As IMethodSymbol = TryCast(ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol, IMethodSymbol)
                Dim parameterList As ParameterListSyntax = SyntaxFactory.ParameterList(SyntaxFactory.SeparatedList(parameters.Select(Function(p As CSS.ParameterSyntax) DirectCast(p.Accept(Me), ParameterSyntax))))
                Dim header As LambdaHeaderSyntax
                Dim IsFunction As Boolean = Not (symbol.ReturnsVoid OrElse TypeOf node.Body Is CSS.AssignmentExpressionSyntax)
#Disable Warning CC0014 ' Use Ternary operator.
                If IsFunction Then
#Enable Warning CC0014 ' Use Ternary operator.
                    header = SyntaxFactory.FunctionLambdaHeader(SyntaxFactory.List(Of AttributeListSyntax)(), ConvertModifiers(modifiers, IsModule, TokenContext.Local), parameterList, asClause:=Nothing)
                Else
                    header = SyntaxFactory.SubLambdaHeader(SyntaxFactory.List(Of AttributeListSyntax)(), ConvertModifiers(modifiers, IsModule, TokenContext.Local), parameterList, asClause:=Nothing)
                End If
                If TypeOf block Is CSS.BlockSyntax Then
                    block = DirectCast(block, CSS.BlockSyntax).Statements
                End If

                Dim Statements As New SyntaxList(Of StatementSyntax)
                Dim EndBlock As EndBlockStatementSyntax
                If TypeOf block Is CS.CSharpSyntaxNode Then
                    Dim body As VisualBasicSyntaxNode = DirectCast(block, CS.CSharpSyntaxNode).Accept(Me)
                    If body.IsKind(SyntaxKind.SimpleAssignmentStatement) Then
                        Dim SimpleAssignment As AssignmentStatementSyntax = DirectCast(body, AssignmentStatementSyntax)
                        If SimpleAssignment.Left.IsKind(SyntaxKind.SimpleMemberAccessExpression) Then
                            Dim MemberAccessExpression As Syntax.MemberAccessExpressionSyntax = DirectCast(SimpleAssignment.Left, Syntax.MemberAccessExpressionSyntax)
                            Select Case MemberAccessExpression.Expression.Kind
                                Case SyntaxKind.ObjectCreationExpression
                                    header = header.With({VB_EOLTrivia}, {VB_EOLTrivia})
                                    EndBlock = SyntaxFactory.EndBlockStatement(SyntaxKind.EndSubStatement, SyntaxFactory.Token(SyntaxKind.SubKeyword))
                                    Dim UniqueName As String = MethodBodyVisitor.GetUniqueVariableNameInScope(node, "tempVar", mSemanticModel)
                                    Dim UniqueIdentifier As IdentifierNameSyntax = SyntaxFactory.IdentifierName(SyntaxFactory.Identifier(UniqueName))
                                    Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                                    Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(UniqueName))
                                    Dim AsClause As AsClauseSyntax = SyntaxFactory.AsNewClause(DirectCast(MemberAccessExpression.Expression, NewExpressionSyntax))
                                    Dim VariableDeclaration As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.VariableDeclarator(Names, AsClause, initializer:=Nothing))
                                    Dim DimStatement As Syntax.LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(ModifiersTokenList, VariableDeclaration)
                                    Statements = Statements.Add(DimStatement)
                                    Statements = Statements.Add(SyntaxFactory.SimpleAssignmentStatement(SyntaxFactory.QualifiedName(UniqueIdentifier, MemberAccessExpression.Name), SimpleAssignment.Right))
                                    Return SyntaxFactory.MultiLineLambdaExpression(SyntaxKind.MultiLineSubLambdaExpression, header, Statements, EndBlock).WithConvertedTriviaFrom(node)
                                Case SyntaxKind.IdentifierName, SyntaxKind.InvocationExpression
                                    ' handled below
                                Case Else
                                    Stop
                            End Select
                        End If
                    End If
                    If IsFunction Then
                        Return SyntaxFactory.SingleLineLambdaExpression(SyntaxKind.SingleLineFunctionLambdaExpression, header, body).WithConvertedTriviaFrom(node)
                    End If
                    Return SyntaxFactory.SingleLineLambdaExpression(SyntaxKind.SingleLineSubLambdaExpression, header, body).WithConvertedTriviaFrom(node)
                End If
                If Not TypeOf block Is SyntaxList(Of CSS.StatementSyntax) Then
                    Throw New NotSupportedException()
                End If
                Statements = Statements.AddRange(SyntaxFactory.List(DirectCast(block, SyntaxList(Of CSS.StatementSyntax)).SelectMany(Function(s As CSS.StatementSyntax) s.Accept(New MethodBodyVisitor(mSemanticModel, Me)))))
                Dim expression As ExpressionSyntax = Nothing
                If Statements.Count = 1 AndAlso UnpackExpressionFromStatement(statementSyntax:=Statements(0), expression:=expression) Then
                    Dim lSyntaxKind As SyntaxKind = If(IsFunction, SyntaxKind.SingleLineFunctionLambdaExpression, SyntaxKind.SingleLineSubLambdaExpression)
                    Return SyntaxFactory.SingleLineLambdaExpression(lSyntaxKind, header, expression).WithConvertedTriviaFrom(node)
                End If

                Dim ExpressionKind As SyntaxKind
                If IsFunction Then
                    EndBlock = SyntaxFactory.EndBlockStatement(SyntaxKind.EndFunctionStatement, SyntaxFactory.Token(SyntaxKind.FunctionKeyword))
                    ExpressionKind = SyntaxKind.MultiLineFunctionLambdaExpression
                Else
                    EndBlock = SyntaxFactory.EndBlockStatement(SyntaxKind.EndSubStatement, SyntaxFactory.Token(SyntaxKind.SubKeyword))
                    ExpressionKind = SyntaxKind.MultiLineSubLambdaExpression
                End If
                Return SyntaxFactory.MultiLineLambdaExpression(kind:=ExpressionKind,
                                                               subOrFunctionHeader:=header,
                                                               statements:=Statements,
                                                               endSubOrFunctionStatement:=EndBlock).WithConvertedTriviaFrom(node)
            End Function
            Private Function MakeAssignmentStatement(ByVal node As CSS.AssignmentExpressionSyntax) As AssignmentStatementSyntax
                Dim kind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(node), IsModule, TokenContext.Local)

                Dim LeftNode As ExpressionSyntax = DirectCast(node.Left.Accept(Me), ExpressionSyntax)
                Dim OperatorToken As SyntaxToken = SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind))
                Dim RightNode As ExpressionSyntax = DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), ExpressionSyntax)
                If node.IsKind(CS.SyntaxKind.AndAssignmentExpression,
                               CS.SyntaxKind.OrAssignmentExpression,
                               CS.SyntaxKind.ExclusiveOrAssignmentExpression,
                               CS.SyntaxKind.ModuloAssignmentExpression) Then
                    Return SyntaxFactory.SimpleAssignmentStatement(LeftNode.WithConvertedTriviaFrom(node.Left), SyntaxFactory.BinaryExpression(kind:=kind, left:=LeftNode, operatorToken:=OperatorToken, right:=RightNode))
                End If
                Dim assignmentStatementSyntax1 As AssignmentStatementSyntax = SyntaxFactory.AssignmentStatement(kind:=kind,
                                                                                                                left:=LeftNode,
                                                                                                                operatorToken:=SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)),
                                                                                                                right:=RightNode)
                Return assignmentStatementSyntax1
            End Function
            Private Sub MarkPatchInlineAssignHelper(ByVal node As CS.CSharpSyntaxNode)
                Dim parentDefinition As CSS.BaseTypeDeclarationSyntax = node.AncestorsAndSelf().OfType(Of CSS.BaseTypeDeclarationSyntax)().FirstOrDefault()
                inlineAssignHelperMarkers.Add(parentDefinition)
            End Sub
            Private Function ReduceArrayUpperBoundExpression(ByVal expr As CSS.ExpressionSyntax) As ExpressionSyntax
                Dim constant As [Optional](Of Object) = mSemanticModel.GetConstantValue(expr)
                If constant.HasValue AndAlso TypeOf constant.Value Is Integer Then
                    Return SyntaxFactory.NumericLiteralExpression(SyntaxFactory.Literal(CInt(constant.Value) - 1))
                End If
                Return SyntaxFactory.BinaryExpression(kind:=SyntaxKind.SubtractExpression, left:=DirectCast(expr.Accept(Me), ExpressionSyntax), operatorToken:=SyntaxFactory.Token(SyntaxKind.MinusToken), right:=SyntaxFactory.NumericLiteralExpression(SyntaxFactory.Literal(1)))
            End Function

            Public Overrides Function VisitAnonymousMethodExpression(ByVal node As CSS.AnonymousMethodExpressionSyntax) As VisualBasicSyntaxNode
                Dim Parameters As New SeparatedSyntaxList(Of CSS.ParameterSyntax)
                If node.ParameterList IsNot Nothing Then
                    Parameters = CType((node.ParameterList?.Parameters), SeparatedSyntaxList(Of CSS.ParameterSyntax))
                End If
                Return ConvertLambdaExpression(node:=node, block:=node.Block.Statements, parameters:=Parameters, modifiers:=SyntaxFactory.TokenList(node.AsyncKeyword)).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitAnonymousObjectCreationExpression(ByVal node As CSS.AnonymousObjectCreationExpressionSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.AnonymousObjectCreationExpression(SyntaxFactory.ObjectMemberInitializer(SyntaxFactory.SeparatedList(node.Initializers.[Select](Function(i As CSS.AnonymousObjectMemberDeclaratorSyntax) DirectCast(i.Accept(Me), FieldInitializerSyntax))))) ' .WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitArrayCreationExpression(ByVal node As CSS.ArrayCreationExpressionSyntax) As VisualBasicSyntaxNode
                Dim upperBoundArguments As IEnumerable(Of ArgumentSyntax) = node.Type.RankSpecifiers.First()?.Sizes.Where(Function(s As CSS.ExpressionSyntax) Not (TypeOf s Is CSS.OmittedArraySizeExpressionSyntax)).[Select](Function(s As CSS.ExpressionSyntax) DirectCast(SyntaxFactory.SimpleArgument(ReduceArrayUpperBoundExpression(s)), ArgumentSyntax))
                Dim cleanUpperBounds As New List(Of ArgumentSyntax)
                For Each Argument As ArgumentSyntax In upperBoundArguments
                    If Argument.ToString <> "-1" Then
                        cleanUpperBounds.Add(Argument)
                    End If
                Next
                upperBoundArguments = cleanUpperBounds
                Dim rankSpecifiers As IEnumerable(Of ArrayRankSpecifierSyntax) = node.Type.RankSpecifiers.[Select](Function(rs As CSS.ArrayRankSpecifierSyntax) DirectCast(rs.Accept(Me), ArrayRankSpecifierSyntax))
                Dim NewKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NewKeyword)
                Dim AttributeLists As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(Of AttributeListSyntax)()
                Dim ArrayType As TypeSyntax = DirectCast(node.Type.ElementType.Accept(Me), TypeSyntax)
                Dim ArrayBounds As ArgumentListSyntax = If(upperBoundArguments.Any(), SyntaxFactory.ArgumentList(arguments:=SyntaxFactory.SeparatedList(upperBoundArguments)), Nothing)
                Dim RankSpecifierList As New List(Of ArrayRankSpecifierSyntax)
                Dim RankSpecifiers1 As SyntaxList(Of ArrayRankSpecifierSyntax) = If(upperBoundArguments.Any(), SyntaxFactory.List(rankSpecifiers.Skip(1)), SyntaxFactory.List(rankSpecifiers))
                RankSpecifierList.AddRange(RankSpecifiers1)
                Dim Initializer As CollectionInitializerSyntax = If(DirectCast(node.Initializer?.Accept(Me), CollectionInitializerSyntax), SyntaxFactory.CollectionInitializer())
                'If Initializer.Initializers.Count > 0 Then
                '    If RankSpecifierList.Count > 0 Then
                '        If Not rankSpecifiers.Last.GetTrailingTrivia.ContainsEOLTrivia Then
                '            RankSpecifierList(RankSpecifierList.Count - 1) = RankSpecifierList(RankSpecifierList.Count - 1).WithAppendedTrailingTrivia(VB_EOLTrivia)
                '        End If
                '    ElseIf ArrayBounds.Arguments.Count > 0 AndAlso ArrayBounds.HasTrailingTrivia AndAlso ArrayBounds.GetTrailingTrivia.ContainsEOLTrivia Then
                '    Else
                '        ArrayBounds = ArrayBounds.WithAppendedTrailingTrivia(VB_EOLTrivia)
                '    End If
                'End If
                Return SyntaxFactory.ArrayCreationExpression(
                                                    newKeyword:=NewKeyword,
                                                    attributeLists:=AttributeLists,
                                                    type:=ArrayType,
                                                    arrayBounds:=ArrayBounds,
                                                    rankSpecifiers:=RankSpecifiers1,
                                                    initializer:=Initializer
                                                            ).WithConvertedTriviaFrom(node)
            End Function



            Public Overrides Function VisitAssignmentExpression(ByVal node As CSS.AssignmentExpressionSyntax) As VisualBasicSyntaxNode
                'Dim errorHandler As EventHandler(Of AnalyzerLoadFailureEventArgs) = Sub(o, e) errors.Add(e)
                If TypeOf node.Parent Is CSS.ExpressionStatementSyntax Then
                    Dim RightTypeInfo As TypeInfo = ModelExtensions.GetTypeInfo(mSemanticModel, node.Right)
                    Dim IsDelegate As Boolean
                    If RightTypeInfo.ConvertedType IsNot Nothing Then
                        IsDelegate = RightTypeInfo.ConvertedType.IsDelegateType
                        If Not IsDelegate Then
                            IsDelegate = RightTypeInfo.ConvertedType.ToString.StartsWith("System.EventHandler")
                        End If
                    Else
                        If RightTypeInfo.Type IsNot Nothing Then
                            IsDelegate = RightTypeInfo.Type.IsDelegateType
                            If Not IsDelegate Then
                                IsDelegate = RightTypeInfo.Type.ToString.StartsWith("System.EventHandler")
                            End If
                        End If
                    End If
                    If IsDelegate OrElse node.Right.IsKind(CS.SyntaxKind.ParenthesizedLambdaExpression) Then
                        If SyntaxTokenExtensions.IsKind(node.OperatorToken, CS.SyntaxKind.PlusEqualsToken) Then
                            Return SyntaxFactory.AddHandlerStatement(DirectCast(node.Left.Accept(Me).WithLeadingTrivia(SyntaxFactory.ElasticSpace), ExpressionSyntax), DirectCast(node.Right.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
                        End If

                        If SyntaxTokenExtensions.IsKind(node.OperatorToken, CS.SyntaxKind.MinusEqualsToken) Then
                            ' TODO capture and leading comments from node.left
                            Return SyntaxFactory.RemoveHandlerStatement(DirectCast(node.Left.Accept(Me), ExpressionSyntax).WithLeadingTrivia(SyntaxFactory.ElasticSpace), DirectCast(node.Right.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
                        End If
                    End If
                    If node.Left.IsKind(CS.SyntaxKind.DeclarationExpression, CS.SyntaxKind.TupleExpression) Then
                        Dim TupleName As String
                        Dim kind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(node), IsModule, TokenContext.Local)
                        Dim OperatorToken As SyntaxToken = SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind))
                        Dim RightNode As ExpressionSyntax = DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), ExpressionSyntax)
                        Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(RightNode)
                        Dim DimToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.DimKeyword).WithConvertedLeadingTriviaFrom(node.Left.GetFirstToken())
                        Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(DimToken)
                        Dim StatementList As New SyntaxList(Of StatementSyntax)
                        Dim VariableNames As New List(Of String)
                        If node.Left.IsKind(CS.SyntaxKind.DeclarationExpression) Then
                            TupleName = MethodBodyVisitor.GetUniqueVariableNameInScope(node, "TupleTempVar", mSemanticModel)
                            Dim NodeLeft As CSS.DeclarationExpressionSyntax = DirectCast(node.Left, DeclarationExpressionSyntax)
                            Dim Designation As CSS.ParenthesizedVariableDesignationSyntax = DirectCast(NodeLeft.Designation, ParenthesizedVariableDesignationSyntax)
                            For i As Integer = 0 To Designation.Variables.Count - 1
                                VariableNames.Add(Designation.Variables(i).Accept(Me).ToString)
                            Next

                            Dim TempTupleIdentifier As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(TupleName))
                            Dim VariableDeclaration As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.VariableDeclarator(TempTupleIdentifier, asClause:=Nothing, Initializer))
                            Dim DimStatement As Syntax.LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(ModifiersTokenList, VariableDeclaration).WithPrependedLeadingTrivia(SyntaxFactory.CommentTrivia($" 'TODO: VB has no equivalent to C# deconstruction declarations, an attempt was made to convert."), VB_EOLTrivia)
                            StatementList = StatementList.Add(DimStatement)

                            For i As Integer = 0 To VariableNames.Count - 1
                                If VariableNames(i) = "_" Then
                                    Continue For
                                End If
                                Dim AsClause As AsClauseSyntax = Nothing
                                If NodeLeft.Type Is Nothing Then
                                    Stop
                                Else
                                    AsClause = SyntaxFactory.SimpleAsClause(CType(NodeLeft.Type.Accept(Me), TypeSyntax))
                                End If
                                Dim TupleExpression As ExpressionSyntax = SyntaxFactory.ParseExpression($"{TupleName}.Item{(i + 1)}")
                                Initializer = SyntaxFactory.EqualsValue(value:=SyntaxFactory.InvocationExpression(expression:=TupleExpression))
                                Dim Declarators As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(
                                    node:=SyntaxFactory.VariableDeclarator(
                                                                            names:=SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(VariableNames(i))),
                                                                            asClause:=AsClause,
                                                                            initializer:=Initializer
                                                                            )
                                                                                                                                                    )

                                Dim AssignmentStatement As Syntax.LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(modifiers:=ModifiersTokenList, declarators:=Declarators)
                                StatementList = StatementList.Add(AssignmentStatement)
                            Next
                        Else
                            ' Create declarations
                            TupleName = node.Left.Accept(Me).ToString
                        End If

                        ' Handle assignment to a Tuple of Variables that already exist
                        If node.Left.IsKind(CS.SyntaxKind.TupleExpression) Then
                            Dim LeftTupleNode As Syntax.TupleExpressionSyntax = DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), Syntax.TupleExpressionSyntax)

                            VariableNames = New List(Of String)
                            TupleName = MethodBodyVisitor.GetUniqueVariableNameInScope(node, "TupleTempVar", mSemanticModel)
                            For Each Argument As ArgumentSyntax In LeftTupleNode.Arguments
                                VariableNames.Add(Argument.ToString)
                            Next
                            Dim TupleList As New List(Of String)
                            If RightTypeInfo.Type.IsErrorType Then
                                For Each a As CSS.ArgumentSyntax In DirectCast(node.Left, CSS.TupleExpressionSyntax).Arguments
                                    If TypeOf a.Expression Is CSS.DeclarationExpressionSyntax Then
                                        Dim t As CSS.DeclarationExpressionSyntax = DirectCast(a.Expression, DeclarationExpressionSyntax)
                                        TupleList.Add(ConvertToType(t.[Type].ToString).ToString)
                                    Else
                                        ' We are going to ignore this
                                        TupleList.Add("_")
                                    End If
                                Next
                            Else
                                If TypeOf RightTypeInfo.Type Is INamedTypeSymbol Then
                                    For Each t As IFieldSymbol In DirectCast(RightTypeInfo.Type, INamedTypeSymbol).TupleElements
                                        ' Need to convert Types !!!!!!!
                                        TupleList.Add(ConvertToTypeString(t.Type))
                                    Next
                                ElseIf TypeOf RightTypeInfo.Type Is ITypeSymbol Then
                                    For i As Integer = 1 To LeftTupleNode.Arguments.Count
                                        TupleList.Add("Object")
                                    Next
                                Else
                                    Stop
                                End If
                            End If
                            Dim builder As New System.Text.StringBuilder()
                            builder.Append("(")
                            For i As Integer = 0 To TupleList.Count - 2
                                builder.Append(Convert_ToObject(TupleList(i)) & ", ")
                            Next
                            builder.Append(Convert_ToObject(TupleList.Last) & ")")
                            Dim TupleType As String = builder.ToString

                            Dim FieldSymbol As String = TupleList(TupleList.Count - 1)

                            Dim SimpleAs As SimpleAsClauseSyntax = SyntaxFactory.SimpleAsClause(asKeyword:=SyntaxFactory.Token(SyntaxKind.AsKeyword).WithTrailingTrivia(SyntaxFactory.Space), attributeLists:=Nothing, type:=SyntaxFactory.ParseTypeName(TupleType).WithLeadingTrivia(SyntaxFactory.Space)).WithLeadingTrivia(SyntaxFactory.Space)
                            Dim names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(TupleName))
                            Dim VariableDeclaration As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.VariableDeclarator(names:=names, asClause:=SimpleAs, initializer:=Initializer))
                            Dim DimStatement As Syntax.LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(modifiers:=ModifiersTokenList, declarators:=VariableDeclaration)
                            StatementList = StatementList.Add(DimStatement)

                            For i As Integer = 0 To TupleList.Count - 1
                                If TupleList(i) = "_" Then
                                    Continue For
                                End If
                                Dim NewLeftNode As ExpressionSyntax = SyntaxFactory.IdentifierName(VariableNames(i))
                                Dim NewRightNode As ExpressionSyntax = SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression($"{TupleName}.Item{(i + 1)}"))
                                Dim AssignmentStatement As AssignmentStatementSyntax = SyntaxFactory.AssignmentStatement(kind:=kind,
                                                                                                        left:=NewLeftNode,
                                                                                                        operatorToken:=OperatorToken,
                                                                                                        right:=NewRightNode)
                                StatementList = StatementList.Add(AssignmentStatement)
                            Next
                        End If
                        Dim FinallyBlock As FinallyBlockSyntax = SyntaxFactory.FinallyBlock()
                        Return SyntaxFactory.TryBlock(StatementList, catchBlocks:=Nothing, FinallyBlock)

                    End If

                    Return MakeAssignmentStatement(node).WithConvertedTriviaFrom(node)
                End If

                If TypeOf node.Parent Is CSS.ForStatementSyntax OrElse TypeOf node.Parent Is CSS.ParenthesizedLambdaExpressionSyntax Then
                    Return MakeAssignmentStatement(node).WithConvertedTriviaFrom(node)
                End If

                If TypeOf node.Parent Is CSS.InitializerExpressionSyntax Then
#Disable Warning CC0013 ' Use Ternary operator.
                    If TypeOf node.Left Is CSS.ImplicitElementAccessSyntax Then
                        Return SyntaxFactory.CollectionInitializer(SyntaxFactory.SeparatedList({DirectCast(node.Left.Accept(Me), ExpressionSyntax), DirectCast(node.Right.Accept(Me), ExpressionSyntax)})).WithConvertedTriviaFrom(node)
                    Else
                        Return SyntaxFactory.NamedFieldInitializer(DirectCast(node.Left.Accept(Me), IdentifierNameSyntax), DirectCast(node.Right.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
                    End If
                End If

                Dim LeftExpression As ExpressionSyntax = DirectCast(node.Left.Accept(Me), ExpressionSyntax)
                Dim RightExpression As ExpressionSyntax = DirectCast(node.Right.Accept(Me), ExpressionSyntax)
                If TypeOf node.Parent Is CSS.ArrowExpressionClauseSyntax Then
                    Return SyntaxFactory.SimpleAssignmentStatement(left:=LeftExpression, right:=RightExpression)
                End If
                MarkPatchInlineAssignHelper(node)
                Return SyntaxFactory.InvocationExpression(expression:=SyntaxFactory.IdentifierName("__InlineAssignHelper"), argumentList:=SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList((New ArgumentSyntax() {SyntaxFactory.SimpleArgument(LeftExpression), SyntaxFactory.SimpleArgument(RightExpression)})))).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitAwaitExpression(ByVal node As CSS.AwaitExpressionSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.AwaitExpression(expression:=DirectCast(node.Expression.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitBaseExpression(ByVal node As CSS.BaseExpressionSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.MyBaseExpression()
            End Function
            Public Overrides Function VisitBinaryExpression(ByVal node As CSS.BinaryExpressionSyntax) As VisualBasicSyntaxNode
                Try
                    Dim IfKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.IfKeyword)
                    Dim TryCastKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.TryCastKeyword)
                    Dim OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken)
                    Dim CommaToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CommaToken)
                    Dim CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken)
                    Dim FoundEOL As Boolean = False
                    Dim ElasticSpace As SyntaxTrivia = SyntaxFactory.ElasticSpace
                    If node.IsKind(CS.SyntaxKind.CoalesceExpression) Then
                        Dim FirstExpression As ExpressionSyntax = DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), ExpressionSyntax)
                        Dim SecondExpression As ExpressionSyntax = DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), ExpressionSyntax)
                        Dim SeparatorTrailingTrivia As New List(Of SyntaxTrivia)
                        If FirstExpression.ContainsEOLTrivia OrElse FirstExpression.ContainsCommentOrDirectiveTrivia Then
                            Dim IfLeadingTrivia As New List(Of SyntaxTrivia)
                            If FirstExpression.HasLeadingTrivia Then
                                IfKeyword = IfKeyword.WithLeadingTrivia(FirstExpression.GetLeadingTrivia)
                            End If
                            If FirstExpression.HasTrailingTrivia Then
                                For Each t As SyntaxTrivia In FirstExpression.GetTrailingTrivia
                                    Select Case t.Kind
                                        Case SyntaxKind.CommentTrivia
                                            SeparatorTrailingTrivia.Add(ElasticSpace)
                                            SeparatorTrailingTrivia.Add(t)
                                            SeparatorTrailingTrivia.Add(ElasticSpace)
                                            FoundEOL = True
                                        Case SyntaxKind.EndOfLineTrivia
                                            FoundEOL = True
                                        Case SyntaxKind.WhitespaceTrivia
                                            ' ignore
                                        Case Else
                                            Stop
                                    End Select
                                Next
                                If FoundEOL Then
                                    SeparatorTrailingTrivia.Add(VB_EOLTrivia)
                                    FoundEOL = False
                                End If
                                FirstExpression = FirstExpression.With({ElasticSpace}, {ElasticSpace})
                            End If
                            CommaToken = CommaToken.WithTrailingTrivia(SeparatorTrailingTrivia)
                            SecondExpression = SecondExpression.WithConvertedLeadingTriviaFrom(node.OperatorToken)
                        End If
                        Dim binaryConditionalExpressionSyntax1 As BinaryConditionalExpressionSyntax = SyntaxFactory.BinaryConditionalExpression(
                                                        ifKeyword:=IfKeyword,
                                                        openParenToken:=OpenParenToken,
                                                        firstExpression:=FirstExpression,
                                                        commaToken:=CommaToken,
                                                        secondExpression:=SecondExpression,
                                                        closeParenToken:=CloseParenToken)
                        Return binaryConditionalExpressionSyntax1
                    End If

                    If node.IsKind(CS.SyntaxKind.AsExpression) Then
                        Dim FirstExpression As ExpressionSyntax = DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), ExpressionSyntax)
                        If FirstExpression.ContainsEOLTrivia Then
                            FirstExpression = FirstExpression.WithRemovedTrailingEOLTrivia
                            CommaToken = CommaToken.WithTrailingTrivia(VB_EOLTrivia)
                        End If
                        Dim Type1 As TypeSyntax = DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), TypeSyntax)
                        Dim TryCastExpression As TryCastExpressionSyntax = SyntaxFactory.TryCastExpression(
                                                keyword:=TryCastKeyword,
                                                openParenToken:=OpenParenToken,
                                                expression:=FirstExpression,
                                                commaToken:=CommaToken,
                                                type:=Type1,
                                                closeParenToken:=CloseParenToken)
                        Return TryCastExpression
                    End If

                    If node.IsKind(CS.SyntaxKind.IsExpression) Then
                        Dim typeOfExpressionSyntax1 As Syntax.TypeOfExpressionSyntax = SyntaxFactory.TypeOfIsExpression(expression:=DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), ExpressionSyntax), type:=DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), TypeSyntax))
                        Return typeOfExpressionSyntax1
                    End If

                    If SyntaxTokenExtensions.IsKind(node.OperatorToken, CS.SyntaxKind.EqualsEqualsToken) Then
                        Dim otherArgument As ExpressionSyntax = Nothing
                        If node.Left.IsKind(CS.SyntaxKind.NullLiteralExpression) Then
                            otherArgument = DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), ExpressionSyntax)
                        End If

                        If node.Right.IsKind(CS.SyntaxKind.NullLiteralExpression) Then
                            otherArgument = DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), ExpressionSyntax)
                        End If

                        If otherArgument IsNot Nothing Then
                            Dim binaryExpressionSyntax2 As Syntax.BinaryExpressionSyntax = SyntaxFactory.IsExpression(otherArgument, LiteralNothing)
                            Return binaryExpressionSyntax2
                        End If
                    End If

                    If SyntaxTokenExtensions.IsKind(node.OperatorToken, CS.SyntaxKind.ExclamationEqualsToken) Then
                        Dim otherArgument As ExpressionSyntax = Nothing
                        If node.Left.IsKind(CS.SyntaxKind.NullLiteralExpression) Then
                            otherArgument = DirectCast(node.Right.Accept(Me).WithConvertedTriviaFrom(node.Right), ExpressionSyntax)
                        End If

                        If node.Right.IsKind(CS.SyntaxKind.NullLiteralExpression) Then
                            otherArgument = DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), ExpressionSyntax)
                        End If

                        If otherArgument IsNot Nothing Then
                            Dim binaryExpressionSyntax1 As Syntax.BinaryExpressionSyntax = SyntaxFactory.IsNotExpression(otherArgument, LiteralNothing.WithConvertedTriviaFrom(node.Right))
                            Return binaryExpressionSyntax1
                        End If
                    End If

                    Dim kind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(node), IsModule, TokenContext.Local)
                    If node.IsKind(CS.SyntaxKind.AddExpression) AndAlso (ModelExtensions.GetTypeInfo(mSemanticModel, node.Left).ConvertedType?.SpecialType = SpecialType.System_String OrElse ModelExtensions.GetTypeInfo(mSemanticModel, node.Right).ConvertedType?.SpecialType = SpecialType.System_String) Then
                        kind = SyntaxKind.ConcatenateExpression
                    End If

                    Dim LeftExpression As ExpressionSyntax = DirectCast(node.Left.Accept(Me).WithConvertedTriviaFrom(node.Left), ExpressionSyntax)
                    Dim LeftTrailingTrivia As SyntaxTriviaList = LeftExpression.GetTrailingTrivia
                    Dim RightExpression As ExpressionSyntax
                    If LeftTrailingTrivia.ToList.Count = 1 AndAlso LeftTrailingTrivia(0).ToString.Trim = "?" Then
                        RightExpression = DirectCast(node.Right.Accept(Me), ExpressionSyntax)
                        Dim OldIdentifier As Syntax.IdentifierNameSyntax = RightExpression.DescendantNodes.
                                                                           OfType(Of IdentifierNameSyntax).
                                                                           First(Function(b As IdentifierNameSyntax) b.Kind() = SyntaxKind.IdentifierName)
                        Dim NewIdentifierWithQuestionMark As IdentifierNameSyntax =
                                    SyntaxFactory.IdentifierName($"{DirectCast(node.Left.Accept(Me), ExpressionSyntax).ToString}?")
                        Return RightExpression.ReplaceNode(OldIdentifier, NewIdentifierWithQuestionMark)
                    End If

                    Dim MovedTrailingTrivia As New List(Of SyntaxTrivia)
                    If LeftExpression.HasLeadingTrivia AndAlso LeftExpression.GetLeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                        MovedTrailingTrivia.AddRange(LeftExpression.GetLeadingTrivia)
                        LeftExpression = LeftExpression.WithLeadingTrivia(ElasticSpace)
                    End If
                    If LeftExpression.HasTrailingTrivia Then
                        For Each t As SyntaxTrivia In LeftExpression.GetTrailingTrivia
                            Select Case t.Kind
                                Case SyntaxKind.CommentTrivia
                                    MovedTrailingTrivia.Add(ElasticSpace)
                                    MovedTrailingTrivia.Add(t)
                                    MovedTrailingTrivia.Add(ElasticSpace)
                                Case SyntaxKind.EndOfLineTrivia
                                    FoundEOL = True
                                Case SyntaxKind.WhitespaceTrivia
                                Case Else
                                    Stop
                            End Select
                        Next
                        LeftExpression = LeftExpression.WithTrailingTrivia(ElasticSpace)
                    End If
                    Dim operatorToken As SyntaxToken = SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)).WithConvertedTriviaFrom(node.OperatorToken)
                    If operatorToken.HasLeadingTrivia And operatorToken.LeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                        MovedTrailingTrivia.AddRange(operatorToken.LeadingTrivia)
                        operatorToken = operatorToken.WithLeadingTrivia(ElasticSpace)
                    End If

                    If operatorToken.HasTrailingTrivia Then
                        Dim NewOperatorTrailingTrivia As New List(Of SyntaxTrivia)
                        For Each t As SyntaxTrivia In operatorToken.TrailingTrivia
                            Select Case t.Kind
                                Case SyntaxKind.CommentTrivia
                                    If FoundEOL Then
                                        NewOperatorTrailingTrivia.AddRange(MovedTrailingTrivia)
                                        NewOperatorTrailingTrivia.Add(ElasticSpace)
                                        NewOperatorTrailingTrivia.Add(t)
                                        NewOperatorTrailingTrivia.Add(ElasticSpace)
                                        MovedTrailingTrivia.Clear()
                                    Else
                                        NewOperatorTrailingTrivia.Add(t)
                                    End If
                                Case SyntaxKind.EndOfLineTrivia
                                    FoundEOL = True
                                Case SyntaxKind.WhitespaceTrivia
                                    If FoundEOL Then
                                        MovedTrailingTrivia.Add(t)
                                    Else
                                        NewOperatorTrailingTrivia.Add(t)
                                    End If
                                Case Else
                                    Stop
                            End Select
                        Next
                        If FoundEOL Then
                            NewOperatorTrailingTrivia.Add(VB_EOLTrivia)
                        End If
                        operatorToken = operatorToken.WithTrailingTrivia(NewOperatorTrailingTrivia)
                    End If

                    Dim RightNode As ExpressionSyntax = DirectCast(node.Right.Accept(Me), ExpressionSyntax)

                    If node.Right.HasTrailingTrivia Then
                        MovedTrailingTrivia.Clear()
                        FoundEOL = False
                        For Each t As SyntaxTrivia In RightNode.GetTrailingTrivia
                            Select Case t.Kind
                                Case SyntaxKind.CommentTrivia
                                    MovedTrailingTrivia.Add(ElasticSpace)
                                    MovedTrailingTrivia.Add(t)
                                    MovedTrailingTrivia.Add(ElasticSpace)

                                Case SyntaxKind.EndOfLineTrivia
                                    FoundEOL = True
                                Case SyntaxKind.WhitespaceTrivia
                                    MovedTrailingTrivia.Add(t)
                                Case Else
                                    Stop
                            End Select
                        Next
                        If FoundEOL Then
                            MovedTrailingTrivia.Add(VB_EOLTrivia)
                        End If
                    End If
                    RightExpression = RightNode.WithLeadingTrivia(ElasticSpace).WithTrailingTrivia(MovedTrailingTrivia)
                    Dim binaryExpressionSyntax3 As Syntax.BinaryExpressionSyntax = SyntaxFactory.BinaryExpression(
                                                    kind:=kind,
                                                    left:=LeftExpression,
                                                    operatorToken:=operatorToken,
                                                    right:=RightExpression
                                                    )
                    Return binaryExpressionSyntax3
                Catch ex As Exception
                    Stop
                    Throw
                End Try
                Throw ExceptionUtilities.UnreachableException
            End Function
            Public Overrides Function VisitCastExpression(ByVal node As CSS.CastExpressionSyntax) As VisualBasicSyntaxNode
                Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                Dim CTypeExpressionSyntax As VisualBasicSyntaxNode = Nothing
                Try
                    Dim type As ITypeSymbol = ModelExtensions.GetTypeInfo(mSemanticModel, node.Type).Type
                    Dim expr As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                    NewTrailingTrivia.AddRange(expr.GetTrailingTrivia)
                    NewTrailingTrivia.AddRange(ConvertTrivia(node.GetTrailingTrivia))
                    expr = expr.WithoutTrivia

                    Dim ExpressionTypeStr As String = ModelExtensions.GetTypeInfo(mSemanticModel, node.Expression).Type?.ToString
                    Select Case type.SpecialType
                        Case SpecialType.System_Object
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CObjKeyword), expression:=expr)
                        Case SpecialType.System_Boolean
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CBoolKeyword), expression:=expr)
                        Case SpecialType.System_Char
                            Dim TestAgainst As String() = {"int", "ushort"}
                            If (node.Parent.IsKind(CS.SyntaxKind.AttributeArgument)) Then
                                CTypeExpressionSyntax = expr
                            ElseIf TestAgainst.Contains(ExpressionTypeStr, StringComparer.OrdinalIgnoreCase) Then
                                CTypeExpressionSyntax = SyntaxFactory.ParseExpression($"ChrW({expr.ToString})")
                            Else
                                CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CCharKeyword), expression:=expr)
                            End If
                        Case SpecialType.System_SByte
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CSByteKeyword), expression:=expr)
                        Case SpecialType.System_Byte
                            If expr.IsKind(SyntaxKind.CharacterLiteralExpression) Then
                                CTypeExpressionSyntax = SyntaxFactory.ParseExpression($"AscW({expr.ToString})")
                            Else
                                CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CByteKeyword), expression:=expr)
                            End If
                        Case SpecialType.System_Int16
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CShortKeyword), expression:=expr)
                        Case SpecialType.System_UInt16
                            If ExpressionTypeStr = "char" Then
                                CTypeExpressionSyntax = SyntaxFactory.ParseExpression(text:=$"ChrW({expr.ToString})")
                            Else
                                CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(keyword:=SyntaxFactory.Token(SyntaxKind.CUShortKeyword), expression:=expr)
                            End If
                        Case SpecialType.System_Int32
                            If ExpressionTypeStr = "char" Then
                                CTypeExpressionSyntax = SyntaxFactory.ParseExpression($"ChrW({expr.ToString})").WithTrailingTrivia(NewTrailingTrivia)
                            Else
                                CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CIntKeyword), expr)
                            End If
                        Case SpecialType.System_UInt32
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CUIntKeyword), expr)
                        Case SpecialType.System_Int64
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CLngKeyword), expr)
                        Case SpecialType.System_UInt64
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CULngKeyword), expr)
                        Case SpecialType.System_Decimal
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CDecKeyword), expr)
                        Case SpecialType.System_Single
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CSngKeyword), expr)
                        Case SpecialType.System_Double
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CDblKeyword), expr)
                        Case SpecialType.System_String
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CStrKeyword), expr)
                        Case SpecialType.System_DateTime
                            CTypeExpressionSyntax = SyntaxFactory.PredefinedCastExpression(SyntaxFactory.Token(SyntaxKind.CDateKeyword), expr)
                        Case Else
                            ' Added support to correctly handle AddressOf
                            Dim TypeOrAddressOf As VisualBasicSyntaxNode = node.Type.Accept(Me)
                            If TypeOrAddressOf.IsKind(SyntaxKind.AddressOfExpression) Then
                                Dim AddressOf1 As UnaryExpressionSyntax = CType(TypeOrAddressOf, UnaryExpressionSyntax)
                                If AddressOf1.Operand.ToString.StartsWith("&") Then
                                    node.FirstAncestorOrSelf(Of CSS.StatementSyntax).AddMarker(CommentOutUnsupportedStatements(node.FirstAncestorOrSelf(Of CSS.StatementSyntax), "pointers"), RemoveStatement:=True, AllowDuplicates:=True)
                                    CTypeExpressionSyntax = expr
                                Else
                                    CTypeExpressionSyntax = SyntaxFactory.CTypeExpression(expr, SyntaxFactory.ParseTypeName(AddressOf1.Operand.ToString.Replace("&", "")))
                                End If
                            Else
                                CTypeExpressionSyntax = SyntaxFactory.CTypeExpression(expr, DirectCast(TypeOrAddressOf, TypeSyntax))
                            End If
                    End Select
                Catch ex As Exception
                    Stop
                    Throw ex
                End Try
                Return CTypeExpressionSyntax.WithConvertedLeadingTriviaFrom(node).WithTrailingTrivia(NewTrailingTrivia)
            End Function
            Public Overrides Function VisitCheckedExpression(node As CheckedExpressionSyntax) As VisualBasicSyntaxNode
                Dim LeadingTrivia As New SyntaxTriviaList
                Dim IsField As Boolean = False
                Dim Unchecked As Boolean = node.Keyword.ValueText <> "checked"
                Dim msg As String = If(Unchecked, "VB has no direct equivalent to C# unchecked:", "VB default math is equivalent to C# checked:")
                Dim StatementWithIssue As CS.CSharpSyntaxNode = GetStatementwithIssues(node, IsField)
                Dim FieldWithIssue As CSS.FieldDeclarationSyntax = Nothing
                If IsField Then
                    FieldWithIssue = CType(StatementWithIssue, CSS.FieldDeclarationSyntax)
                    LeadingTrivia = CheckCorrectnessLeadingTrivia(FieldWithIssue, msg)
                    FieldWithIssue.AddMarker(LeadingTrivia, AllowDuplicates:=False)
                Else
                    LeadingTrivia = CheckCorrectnessLeadingTrivia(StatementWithIssue, msg)
                    ' Only notify once on one line TODO Merge the comments
                    StatementWithIssue.AddMarker(SyntaxFactory.EmptyStatement.WithLeadingTrivia(LeadingTrivia), RemoveStatement:=False, AllowDuplicates:=False)
                End If

                Dim Expression As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                If TypeOf Expression Is PredefinedCastExpressionSyntax Then
                    Dim CastExpression As PredefinedCastExpressionSyntax = DirectCast(Expression, PredefinedCastExpressionSyntax)
                    If Unchecked Then
                        Return SyntaxFactory.ParseExpression($"{CastExpression.Keyword.ToString}(Val(""&H"" & Hex({CastExpression.Expression.ToString})))")
                    Else
                        Return SyntaxFactory.ParseExpression($"{CastExpression.Keyword.ToString}({CastExpression.Expression.ToString})")
                    End If
                End If
                If TypeOf Expression Is Syntax.BinaryExpressionSyntax OrElse
                    TypeOf Expression Is Syntax.InvocationExpressionSyntax OrElse
                    TypeOf Expression Is Syntax.ObjectCreationExpressionSyntax Then
                    If Unchecked Then
                        Return SyntaxFactory.ParseExpression($"Unchecked({Expression.ToString})")
                    Else
                        Return Expression
                    End If
                End If
                If TypeOf Expression Is Syntax.CTypeExpressionSyntax Then
                    Return Expression
                End If

                Throw UnreachableException
            End Function
            Public Overrides Function VisitConditionalAccessExpression(ByVal node As CSS.ConditionalAccessExpressionSyntax) As VisualBasicSyntaxNode
                Dim expression As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                If expression.ContainsEOLTrivia Then
                    TrailingTriviaList.AddRange(expression.WithRemovedTrailingEOLTrivia.GetTrailingTrivia)
                    expression = expression.WithoutTrailingTrivia
                End If
                Return SyntaxFactory.ConditionalAccessExpression(expression, SyntaxFactory.Token(SyntaxKind.QuestionToken).WithTrailingTrivia(TrailingTriviaList), DirectCast(node.WhenNotNull.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitConditionalExpression(ByVal node As CSS.ConditionalExpressionSyntax) As VisualBasicSyntaxNode
                Dim Condition As ExpressionSyntax = DirectCast(node.Condition.Accept(Me), ExpressionSyntax)
                Dim WhenTrue As ExpressionSyntax = DirectCast(node.WhenTrue.Accept(Me).WithoutTrivia, ExpressionSyntax)
                Dim FoundEOL As Boolean
                Dim NewItemLeadingTrivia As New List(Of SyntaxTrivia)
                Dim NewSeparatorTrailingTrivia As New List(Of SyntaxTrivia)
                For Each t As SyntaxTrivia In ConvertTrivia(node.QuestionToken.LeadingTrivia)
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            NewSeparatorTrailingTrivia.Add(t)
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            NewItemLeadingTrivia.Add(t)
                        Case Else
                            Stop
                    End Select
                Next
                For Each t As SyntaxTrivia In ConvertTrivia(node.Condition.GetTrailingTrivia)
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            NewSeparatorTrailingTrivia.Add(t)
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            NewSeparatorTrailingTrivia.Add(SyntaxFactory.ElasticSpace)
                        Case Else
                            Stop
                    End Select
                Next
                Dim NewItemTrailingTrivia As New List(Of SyntaxTrivia)
                For Each t As SyntaxTrivia In ConvertTrivia(node.WhenTrue.GetTrailingTrivia)
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            NewSeparatorTrailingTrivia.Add(t)
                            FoundEOL = True
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            NewItemTrailingTrivia.Add(t)
                        Case Else
                            Stop
                    End Select
                Next
                WhenTrue = WhenTrue.With(NewItemLeadingTrivia, NewItemTrailingTrivia)
                NewItemLeadingTrivia.Clear()
                NewItemTrailingTrivia.Clear()

                If FoundEOL Then
                    NewSeparatorTrailingTrivia.Add(VB_EOLTrivia)
                End If

                Dim ElasticSpaceList As List(Of SyntaxTrivia) = New List(Of SyntaxTrivia) From {
                    SyntaxFactory.ElasticSpace
                }
                Dim FirstCommaToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CommaToken).With(ElasticSpaceList, NewSeparatorTrailingTrivia)
                NewSeparatorTrailingTrivia.Clear()
                Dim WhenFalse As ExpressionSyntax = DirectCast(node.WhenFalse.Accept(Me), ExpressionSyntax).WithConvertedLeadingTriviaFrom(node.ColonToken).WithoutTrailingTrivia
                For Each t As SyntaxTrivia In ConvertTrivia(node.WhenFalse.GetLeadingTrivia)
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            NewSeparatorTrailingTrivia.Add(t)
                            FoundEOL = True
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            NewItemTrailingTrivia.Add(t)
                        Case Else
                            Stop
                    End Select
                Next
                For Each t As SyntaxTrivia In ConvertTrivia(node.WhenFalse.GetTrailingTrivia)
                    Select Case t.Kind
                        Case SyntaxKind.CommentTrivia
                            NewSeparatorTrailingTrivia.Add(t)
                            FoundEOL = True
                        Case SyntaxKind.EndOfLineTrivia
                            FoundEOL = True
                        Case SyntaxKind.WhitespaceTrivia
                            NewItemTrailingTrivia.Add(t)
                        Case Else
                            Stop
                    End Select
                Next
                WhenFalse = WhenFalse.With(NewItemLeadingTrivia, NewItemTrailingTrivia)
                Dim ifKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.IfKeyword).WithConvertedLeadingTriviaFrom(node.Condition.GetFirstToken)
                Dim OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken)
                Dim SecondCommaToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CommaToken)
                Dim CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken).WithTrailingTrivia(NewSeparatorTrailingTrivia)

                Return SyntaxFactory.TernaryConditionalExpression(
                    ifKeyword,
                    OpenParenToken,
                    Condition.WithoutTrivia,
                    FirstCommaToken,
                    WhenTrue,
                    SecondCommaToken,
                    WhenFalse,
                    CloseParenToken)
            End Function
            Public Overrides Function VisitDeclarationExpression(ByVal Node As DeclarationExpressionSyntax) As VisualBasicSyntaxNode
                Dim Value As IdentifierNameSyntax
                If Node.Designation.IsKind(CS.SyntaxKind.SingleVariableDesignation) Then
                    Dim SingleVariableDesignation As SingleVariableDesignationSyntax = DirectCast(Node.Designation, CSS.SingleVariableDesignationSyntax)
                    Value = DirectCast(SingleVariableDesignation.Accept(Me), IdentifierNameSyntax).WithConvertedTriviaFrom(Node)
                    Return Value
                End If
                If Node.Designation.IsKind(CS.SyntaxKind.ParenthesizedVariableDesignation) Then
                    Dim ParenthesizedVariableDesignation As CSS.ParenthesizedVariableDesignationSyntax = DirectCast(Node.Designation, ParenthesizedVariableDesignationSyntax)
                    Dim VariableDeclaration As Syntax.VariableDeclaratorSyntax = DirectCast(ParenthesizedVariableDesignation.Accept(Me), Syntax.VariableDeclaratorSyntax)
                    Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(
                    SyntaxFactory.Token(SyntaxKind.DimKeyword))

                    Dim DeclarationToBeAdded As Syntax.LocalDeclarationStatementSyntax =
                    SyntaxFactory.LocalDeclarationStatement(
                                        ModifiersTokenList,
                                        SyntaxFactory.SingletonSeparatedList(VariableDeclaration)
                                        )

                    Dim StatementWithIssues As CS.CSharpSyntaxNode = GetStatementwithIssues(Node)
                    StatementWithIssues.AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=False)
                    Return SyntaxFactory.IdentifierName(Node.Designation.ToString.Replace(",", "").Replace(" ", "").Replace("(", "").Replace(")", ""))
                End If
                If Node.Designation.IsKind(CS.SyntaxKind.DiscardDesignation) Then
                    Dim DiscardDesignation As DiscardDesignationSyntax = DirectCast(Node.Designation, CSS.DiscardDesignationSyntax)
                    Value = DirectCast(DiscardDesignation.Accept(Me), IdentifierNameSyntax).WithConvertedTriviaFrom(Node)
                    Return Value

                End If
                Throw UnreachableException
            End Function
            Public Overrides Function VisitDefaultExpression(ByVal node As CSS.DefaultExpressionSyntax) As VisualBasicSyntaxNode
                Dim NothingKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NothingKeyword).WithConvertedTriviaFrom(node)
                Return SyntaxFactory.NothingLiteralExpression(NothingKeyword)
            End Function

            Public Overrides Function VisitDiscardDesignation(node As DiscardDesignationSyntax) As VisualBasicSyntaxNode
                Dim Identifier As SyntaxToken = GenerateSafeVBToken(node.UnderscoreToken, IsQualifiedName:=False)
                Dim IdentifierExpression As Syntax.IdentifierNameSyntax = SyntaxFactory.IdentifierName(Identifier)
                Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(Identifier)
                Dim SeparatedSyntaxListOfModifiedIdentifier As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) =
                    SyntaxFactory.SingletonSeparatedList(
                        ModifiedIdentifier
                        )
                Dim Parent As DeclarationExpressionSyntax
                Dim TypeName As VisualBasicSyntaxNode
                If TypeOf node.Parent Is DeclarationExpressionSyntax Then
                    Parent = DirectCast(node.Parent, DeclarationExpressionSyntax)
                    TypeName = Parent.Type.Accept(Me)
                    If TypeName.ToString = "var" Then
                        TypeName = SyntaxFactory.ParseTypeName("Object")
                    End If
                ElseIf node.ToString = "_" Then
                    TypeName = SyntaxFactory.ParseTypeName("Object")
                Else
                    Stop
                    TypeName = SyntaxFactory.ParseTypeName("Object")
                End If

                Dim SeparatedListOfvariableDeclarations As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) =
                    SyntaxFactory.SingletonSeparatedList(
                        SyntaxFactory.VariableDeclarator(
                            SeparatedSyntaxListOfModifiedIdentifier,
                            SyntaxFactory.SimpleAsClause(ConvertToType(TypeName.NormalizeWhitespace.ToString)),
                            SyntaxFactory.EqualsValue(NothingExpression)
                                )
                         )

                Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                Dim DeclarationToBeAdded As Syntax.LocalDeclarationStatementSyntax =
                    SyntaxFactory.LocalDeclarationStatement(
                        ModifiersTokenList,
                        SeparatedListOfvariableDeclarations).WithAdditionalAnnotations(Simplifier.Annotation)

                GetStatementwithIssues(node).AddMarker(Statement:=DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=True)
                Return IdentifierExpression
            End Function
            Public Overrides Function VisitElementAccessExpression(ByVal node As CSS.ElementAccessExpressionSyntax) As VisualBasicSyntaxNode
                Dim argumentList As ArgumentListSyntax = DirectCast(node.ArgumentList.Accept(Me), ArgumentListSyntax)
                Dim expression As ExpressionSyntax
                ' The next line is a double check, the second should be sufficient
                Dim r As New Regex("base *\[")
                If r.IsMatch(node.ToString) Then
                    If node.GetAncestor(Of CSS.IndexerDeclarationSyntax).IsKind(CS.SyntaxKind.IndexerDeclaration) Then
                        expression = SyntaxFactory.ParseExpression($"MyBase.Item")
                    Else
                        Return SyntaxFactory.ParseName($"MyBase.{argumentList.Arguments(0)}")
                    End If
                Else
                    expression = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                End If
                Return SyntaxFactory.InvocationExpression(expression, argumentList)
            End Function

            Public Overrides Function VisitElementBindingExpression(node As ElementBindingExpressionSyntax) As VisualBasicSyntaxNode
                Dim Arguments0 As VisualBasicSyntaxNode = node.ArgumentList.Arguments(0).Accept(Me)
                Dim expression As ExpressionSyntax = SyntaxFactory.ParseExpression(Arguments0.ToString)
                Dim ParenthesizedExpression As Syntax.ParenthesizedExpressionSyntax = SyntaxFactory.ParenthesizedExpression(expression)
                Return SyntaxFactory.InvocationExpression(ParenthesizedExpression)
            End Function

            Public Overrides Function VisitImplicitArrayCreationExpression(ByVal node As CSS.ImplicitArrayCreationExpressionSyntax) As VisualBasicSyntaxNode
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = node.Initializer.Expressions.GetSeparators
                Dim ExpressionItems As New List(Of ExpressionSyntax)
                Dim NamedFieldItems As New List(Of FieldInitializerSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = node.Initializer.Expressions.Count - 1
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.ExpressionSyntax = node.Initializer.Expressions(i)
                    Dim ItemWithTrivia As VisualBasicSyntaxNode
                    Try
                        ItemWithTrivia = e.Accept(Me).WithConvertedTrailingTriviaFrom(e)
                        If TypeOf ItemWithTrivia Is NamedFieldInitializerSyntax Then
                            NamedFieldItems.Add(DirectCast(ItemWithTrivia, NamedFieldInitializerSyntax))
                        Else
                            ExpressionItems.Add(DirectCast(ItemWithTrivia, ExpressionSyntax))
                        End If
                    Catch ex As Exception
                        Stop
                    End Try
                    If SeparatorCount > i Then
                        Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                    End If
                Next
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                Dim OpenBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken).WithConvertedTriviaFrom(node.Initializer.OpenBraceToken)
                Dim CloseBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseBraceToken).WithConvertedTrailingTriviaFrom(node.Initializer.CloseBraceToken)
                If ExpressionItems.Count > 0 Then
                    RestructureNodesAndSeparators(OpenBraceToken, ExpressionItems, Separators, CloseBraceToken, TrailingTriviaList)
                    Dim ExpressionInitializers As SeparatedSyntaxList(Of ExpressionSyntax) = SyntaxFactory.SeparatedList(ExpressionItems, Separators)
                    Dim CollectionInitializer As CollectionInitializerSyntax = SyntaxFactory.CollectionInitializer(OpenBraceToken, ExpressionInitializers, CloseBraceToken)
                    Return CollectionInitializer.WithConvertedLeadingTriviaFrom(node.NewKeyword)
                Else
                    RestructureNodesAndSeparators(OpenBraceToken, NamedFieldItems, Separators, CloseBraceToken, TrailingTriviaList)
                    Dim Initializers As SeparatedSyntaxList(Of FieldInitializerSyntax) = SyntaxFactory.SeparatedList(Of FieldInitializerSyntax)(NamedFieldItems)
                    Dim ObjectInitializer As ObjectMemberInitializerSyntax = SyntaxFactory.ObjectMemberInitializer(Initializers)
                    Return SyntaxFactory.AnonymousObjectCreationExpression(ObjectInitializer)
                End If
            End Function

            Public Overrides Function VisitInitializerExpression(ByVal node As CSS.InitializerExpressionSyntax) As VisualBasicSyntaxNode
                Dim CS_Separators As IEnumerable(Of SyntaxToken) = node.Expressions.GetSeparators
                Dim Expressions As New List(Of ExpressionSyntax)
                Dim Fields As New List(Of FieldInitializerSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim SeparatorCount As Integer = node.Expressions.Count - 1
                Dim FinalSeparator As Boolean = CS_Separators.Count > 0 And SeparatorCount <> CS_Separators.Count
                Dim FirstToken As SyntaxToken = If(node.Expressions.Count > 0, node.Expressions(0).GetFirstToken, Nothing)
                Dim OpenBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken).WithConvertedTriviaFrom(node.OpenBraceToken)
                'If FirstToken <> node.OpenBraceToken Then
                '    Dim OpenBraceTrailingTrivia As New List(Of SyntaxTrivia)
                '    OpenBraceTrailingTrivia.AddRange(ConvertTrivia(node.OpenBraceToken.TrailingTrivia))
                '    OpenBraceTrailingTrivia.AddRange(ConvertTrivia(FirstToken.LeadingTrivia))
                '    OpenBraceToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken).WithConvertedLeadingTriviaFrom(node.OpenBraceToken).WithTrailingTrivia(OpenBraceTrailingTrivia)
                'Else
                '    OpenBraceToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken).WithConvertedTriviaFrom(node.OpenBraceToken)
                'End If
                Dim TrailingTriviaList As New List(Of SyntaxTrivia)
                For i As Integer = 0 To SeparatorCount
                    Dim e As CSS.ExpressionSyntax = node.Expressions(i)
                    Dim Item As VisualBasicSyntaxNode = e.Accept(Me)
                    Dim ItemIsField As Boolean = False
                    If node.IsKind(CS.SyntaxKind.ObjectInitializerExpression) AndAlso TypeOf Item Is FieldInitializerSyntax Then
                        ItemIsField = True
                        Fields.Add(DirectCast(Item, FieldInitializerSyntax))
                    Else
                        Expressions.Add(DirectCast(Item, ExpressionSyntax))
                    End If


                    If SeparatorCount > i Then
                        Separators.Add(SyntaxFactory.Token(SyntaxKind.CommaToken).WithConvertedTrailingTriviaFrom(CS_Separators(i)))
                    Else
                        If FinalSeparator Then
                            If ItemIsField Then
                                Fields(i) = Fields(i).WithAppendedTrailingTrivia(ConvertTrivia(CS_Separators.Last.TrailingTrivia))
                            Else
                                Expressions(i) = Expressions(i).WithAppendedTrailingTrivia(ConvertTrivia(CS_Separators.Last.TrailingTrivia))
                            End If
                        End If
                    End If
                Next
                Dim CloseBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseBraceToken).WithConvertedTriviaFrom(node.CloseBraceToken)
                If node.IsKind(CS.SyntaxKind.ObjectInitializerExpression) Then
                    Dim WithKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.WithKeyword).WithTrailingTrivia(VB_EOLTrivia)
                    If Fields.Count > 0 Then
                        RestructureNodesAndSeparators(OpenBraceToken, Fields, Separators, CloseBraceToken, TrailingTriviaList)
                        If TrailingTriviaList.Count <> 0 Then
                            Throw UnexpectedValue($"TrailingTriviaList.Count <> 0 In {NameOf(VisitInitializerExpression)}")
                        End If
                        Return SyntaxFactory.ObjectMemberInitializer(WithKeyword, OpenBraceToken, SyntaxFactory.SeparatedList(Fields, Separators), CloseBraceToken).WithConvertedTriviaFrom(node)
                    End If
                    RestructureNodesAndSeparators(OpenBraceToken, Expressions, Separators, CloseBraceToken, TrailingTriviaList)
                    If TrailingTriviaList.Count <> 0 Then
                        Throw UnexpectedValue($"TrailingTriviaList.Count <> 0 In {NameOf(VisitInitializerExpression)}")
                    End If
                    If Expressions.Count > 0 Then
                        If Not Expressions(SeparatorCount).ContainsEOLTrivia Then
                            Expressions(SeparatorCount) = Expressions(SeparatorCount).WithAppendedTrailingTrivia(VB_EOLTrivia)
                            Return SyntaxFactory.ObjectCollectionInitializer(SyntaxFactory.CollectionInitializer(OpenBraceToken, SyntaxFactory.SeparatedList(Expressions.OfType(Of ExpressionSyntax), Separators), CloseBraceToken))
                        End If
                    Else
                        Return SyntaxFactory.ObjectCollectionInitializer(SyntaxFactory.CollectionInitializer(OpenBraceToken, SyntaxFactory.SeparatedList(Expressions.OfType(Of ExpressionSyntax), Separators), CloseBraceToken.WithoutTrivia)).WithTrailingTrivia(CloseBraceToken.LeadingTrivia)
                    End If
                End If

                RestructureNodesAndSeparators(OpenBraceToken, Expressions, Separators, CloseBraceToken, TrailingTriviaList)
                If node.IsKind(CS.SyntaxKind.ArrayInitializerExpression) OrElse node.IsKind(CS.SyntaxKind.CollectionInitializerExpression) Then
                    Dim initializers As SeparatedSyntaxList(Of ExpressionSyntax) = SyntaxFactory.SeparatedList(Expressions, Separators)
                    If FinalSeparator Then
                        RestructureCloseTokenTrivia(CS_Separators.Last, CloseBraceToken)
                    End If

                    Dim CollectionInitializer As CollectionInitializerSyntax = SyntaxFactory.CollectionInitializer(OpenBraceToken, initializers, CloseBraceToken)
                    Return CollectionInitializer
                End If
                Return SyntaxFactory.CollectionInitializer(OpenBraceToken, SyntaxFactory.SeparatedList(Expressions, Separators), CloseBraceToken)
            End Function

            Public Overrides Function VisitInterpolatedStringExpression(ByVal node As CSS.InterpolatedStringExpressionSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.InterpolatedStringExpression(node.Contents.[Select](Function(c As CSS.InterpolatedStringContentSyntax) DirectCast(c.Accept(Me), InterpolatedStringContentSyntax)).ToArray()).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitInterpolatedStringText(ByVal node As CSS.InterpolatedStringTextSyntax) As VisualBasicSyntaxNode
                Dim CSharpToken As SyntaxToken = node.TextToken
                Dim TextToken As SyntaxToken = ConvertToInterpolatedStringTextToken(CSharpToken)
                Return SyntaxFactory.InterpolatedStringText(TextToken).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitInterpolation(ByVal node As CSS.InterpolationSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.Interpolation(DirectCast(node.Expression.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitInterpolationFormatClause(ByVal node As CSS.InterpolationFormatClauseSyntax) As VisualBasicSyntaxNode
                Return MyBase.VisitInterpolationFormatClause(node).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitInvocationExpression(ByVal node As CSS.InvocationExpressionSyntax) As VisualBasicSyntaxNode
                If node.Expression.ToString().ToLower = "nameof" Then
                    Try
                        If node.ArgumentList.Arguments.Count <> 1 Then
                            Stop
                            Throw UnexpectedValue($"NameOf contains more than 1 Argument")
                        End If
                        Dim Name As Syntax.SimpleArgumentSyntax = CType(node.ArgumentList.Arguments(0).Accept(Me), Syntax.SimpleArgumentSyntax)
                        Return SyntaxFactory.NameOfExpression(argument:=Name.Expression).WithConvertedTriviaFrom(node)
                    Catch ex As Exception
                        Stop
                    End Try
                    Throw UnreachableException
                End If
                ' Remove Trailing trivia to force ( on same line
                Dim Expression As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax).WithoutTrailingTrivia
                Dim ArgumentList As ArgumentListSyntax = DirectCast(node.ArgumentList.Accept(Me), ArgumentListSyntax)
                Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                NewTrailingTrivia.AddRange(ArgumentList.GetTrailingTrivia)
                NewTrailingTrivia.AddRange(ConvertTrivia(node.GetTrailingTrivia))
                Dim invocationExpressionSyntax1 As Syntax.InvocationExpressionSyntax = SyntaxFactory.InvocationExpression(expression:=Expression, argumentList:=ArgumentList).WithConvertedLeadingTriviaFrom(node).WithTrailingTrivia(NewTrailingTrivia)
                Return invocationExpressionSyntax1
            End Function

            ''' <summary>
            '''
            ''' </summary>
            ''' <param name="node"></param>
            ''' <returns></returns>
            ''' <remarks>Added by PC</remarks>
            Public Overrides Function VisitIsPatternExpression(node As IsPatternExpressionSyntax) As VisualBasicSyntaxNode
                ' Correct
                'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
                'ORIGINAL LINE: if (ModelExtensions.GetSymbolInfo(semanticModel, node).Symbol is IMethodSymbol methodSymbol && methodSymbol.ReturnType.Equals(ModelExtensions.GetTypeInfo(semanticModel, node).ConvertedType))
                'If TypeOf ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol Is IMethodSymbol Then
                '    Dim methodSymbol As IMethodSymbol = DirectCast(ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol, IMethodSymbol)
                '    If methodSymbol.ReturnType.Equals(ModelExtensions.GetTypeInfo(mSemanticModel, node).ConvertedType) Then
                '        Return SyntaxFactory.InvocationExpression(MemberAccessExpressionSyntax, SyntaxFactory.ArgumentList())
                '    End If
                'End If

                ' Current
                ' Dim methodSymbol As IMethodSymbol = DirectCast(ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol, IMethodSymbol)

                Dim Pattern As PatternSyntax = node.Pattern
                If TypeOf Pattern Is DeclarationPatternSyntax Then
                    Dim StatementWithIssue As CS.CSharpSyntaxNode = GetStatementwithIssues(node)
                    Dim LeadingTrivia As New SyntaxTriviaList
                    LeadingTrivia = CheckCorrectnessLeadingTrivia(StatementWithIssue, "VB has no direct equivalent To C# pattern variables 'is' expressions")
                    Dim DeclarationPattern As DeclarationPatternSyntax = DirectCast(node.Pattern, DeclarationPatternSyntax)
                    Dim Designation As SingleVariableDesignationSyntax = DirectCast(DeclarationPattern.Designation, SingleVariableDesignationSyntax)

                    Dim value As ExpressionSyntax = SyntaxFactory.ParseExpression(text:=$"TryCast({node.Expression.Accept(Me).NormalizeWhitespace.ToFullString}, {DeclarationPattern.Type.Accept(Me).NormalizeWhitespace.ToFullString})")

                    Dim Identifier As SyntaxToken = GenerateSafeVBToken(id:=Designation.Identifier, IsQualifiedName:=False, IsTypeName:=False)
                    Dim VariableName As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(Identifier)
                    Dim SeparatedSyntaxListOfModifiedIdentifier As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) =
                        SyntaxFactory.SingletonSeparatedList(
                            VariableName)
                    Dim VariableType As TypeSyntax = DirectCast(DeclarationPattern.Type.Accept(Me), TypeSyntax)

                    Dim SeparatedListOfvariableDeclarations As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) =
                        SyntaxFactory.SingletonSeparatedList(
                            node:=SyntaxFactory.VariableDeclarator(
                                names:=SeparatedSyntaxListOfModifiedIdentifier,
                                asClause:=SyntaxFactory.SimpleAsClause(VariableType),
                                initializer:=SyntaxFactory.EqualsValue(value)
                                    )
                             )

                    Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(
                    token:=SyntaxFactory.Token(SyntaxKind.DimKeyword))

                    Dim DeclarationToBeAdded As Syntax.LocalDeclarationStatementSyntax =
                    SyntaxFactory.LocalDeclarationStatement(
                                        modifiers:=ModifiersTokenList,
                                        declarators:=SeparatedListOfvariableDeclarations
                                        ).WithAdditionalAnnotations(Simplifier.Annotation).WithLeadingTrivia(LeadingTrivia)

                    StatementWithIssue.AddMarker(Statement:=DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=True)
                    Return SyntaxFactory.IsNotExpression(SyntaxFactory.IdentifierName(Identifier.ToString), NothingExpression)
                ElseIf TypeOf Pattern Is ConstantPatternSyntax Then
                    Dim ConstantPattern As ConstantPatternSyntax = DirectCast(node.Pattern, ConstantPatternSyntax)
                    If ConstantPattern.ToString = "null" Then
                        Return SyntaxFactory.IsExpression(left:=DirectCast(node.Expression.Accept(Me), ExpressionSyntax), right:=SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword)))
                    End If
                    If ConstantPattern.Expression IsNot Nothing Then
                        Return SyntaxFactory.EqualsExpression(left:=DirectCast(node.Expression.Accept(Me), ExpressionSyntax), right:=CType(ConstantPattern.Expression.Accept(Me), ExpressionSyntax))
                    End If
                    Stop
                End If
                Throw UnreachableException
            End Function

            Public Overrides Function VisitLiteralExpression(ByVal node As CSS.LiteralExpressionSyntax) As VisualBasicSyntaxNode
                ' now this looks somehow like a hack... is there a better way?
                If node.IsKind(CS.SyntaxKind.StringLiteralExpression) Then
                    ' @"" have no escapes except quotes (ASCII and Unicode)
                    If node.Token.Text.StartsWith("@", StringComparison.Ordinal) Then
                        Return SyntaxFactory.StringLiteralExpression(
                                token:=SyntaxFactory.StringLiteralToken(
                                    text:=node.Token.Text.Substring(1).Replace(UnicodeOpenQuote, UnicodeDoubleOpenQuote).Replace(UnicodeCloseQuote, UnicodeDoubleCloseQuote).NormalizeLineEndings,
                                    value:=DirectCast(node.Token.Value, String).Replace("""", " """"").Replace(UnicodeOpenQuote, UnicodeDoubleOpenQuote).Replace(UnicodeCloseQuote, UnicodeDoubleCloseQuote).NormalizeLineEndings)).WithConvertedTriviaFrom(node.Token)
                    End If
                End If
                Dim ResultExpression As ExpressionSyntax = NothingExpression

                If node.IsKind(CS.SyntaxKind.CharacterLiteralExpression) Then
                    If node.Token.Text.Replace("'", "").Length <= 2 Then
                        Return GetLiteralExpression(value:=node.Token.Value, Token:=node.Token).WithConvertedTriviaFrom(node.Token)
                    End If
                End If

                If node.Token.ValueText.Contains("\") Then
                    ResultExpression = SyntaxFactory.InterpolatedStringExpression(
                                SyntaxFactory.InterpolatedStringText(
                                    ConvertToInterpolatedStringTextToken(node.Token))
                                )
                    Return ResultExpression
                End If
                Dim expressionSyntax1 As ExpressionSyntax = GetLiteralExpression(value:=node.Token.Value, Token:=node.Token).WithConvertedTriviaFrom(node.Token)
                Return expressionSyntax1
            End Function

            Public Overrides Function VisitMemberAccessExpression(ByVal node As CSS.MemberAccessExpressionSyntax) As VisualBasicSyntaxNode
                Dim expression As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax).WithConvertedTriviaFrom(node.Expression)
                Dim operatorToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.DotToken).WithConvertedTriviaFrom(node.OperatorToken)
                Dim name As SimpleNameSyntax = DirectCast(node.Name.Accept(Me).WithConvertedTriviaFrom(node.Name), SimpleNameSyntax)

                If expression.GetLastToken.ContainsEOLTrivia Then
                    Dim OperatorTrailingTrivia As New List(Of SyntaxTrivia)
                    Dim FoundEOL As Boolean = False
                    FoundEOL = RestructureTrivia(TriviaList:=expression.GetTrailingTrivia, FoundEOL:=FoundEOL, OperatorTrailingTrivia:=OperatorTrailingTrivia)
                    FoundEOL = RestructureTrivia(TriviaList:=operatorToken.LeadingTrivia, FoundEOL:=FoundEOL, OperatorTrailingTrivia:=OperatorTrailingTrivia)

                    If FoundEOL Then
                        OperatorTrailingTrivia.Add(VB_EOLTrivia)
                    End If
                    expression = expression.WithoutTrailingTrivia
                    operatorToken = operatorToken.WithoutTrivia.WithTrailingTrivia(OperatorTrailingTrivia)
                    name = name.WithLeadingTrivia(SyntaxFactory.ElasticSpace)
                End If
                Return WrapTypedNameIfNecessary(name:=SyntaxFactory.MemberAccessExpression(kind:=SyntaxKind.SimpleMemberAccessExpression, expression:=expression, operatorToken:=operatorToken, name:=name), originalName:=node).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitMemberBindingExpression(ByVal node As CSS.MemberBindingExpressionSyntax) As VisualBasicSyntaxNode
                Dim name As SimpleNameSyntax = DirectCast(node.Name.Accept(Me), SimpleNameSyntax)
                Return SyntaxFactory.SimpleMemberAccessExpression(name:=name)
            End Function

            Public Overrides Function VisitObjectCreationExpression(ByVal node As CSS.ObjectCreationExpressionSyntax) As VisualBasicSyntaxNode
                Dim NewKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NewKeyword)
                Dim type1 As TypeSyntax = DirectCast(node.Type.Accept(Me), TypeSyntax)
                Dim argumentList As ArgumentListSyntax = DirectCast(node.ArgumentList?.Accept(Me), ArgumentListSyntax)
                Dim PossibleInitializer As VisualBasicSyntaxNode = node.Initializer?.Accept(Me)
                Dim initializer As ObjectCollectionInitializerSyntax = Nothing
                If PossibleInitializer IsNot Nothing Then
                    Select Case PossibleInitializer.Kind
                        Case SyntaxKind.CollectionInitializer
                            initializer = SyntaxFactory.ObjectCollectionInitializer(initializer:=DirectCast(PossibleInitializer, CollectionInitializerSyntax))
                        Case SyntaxKind.ObjectCollectionInitializer
                            initializer = DirectCast(PossibleInitializer, ObjectCollectionInitializerSyntax)
                        Case SyntaxKind.ObjectMemberInitializer
                            ' Remove trailing trivia before with
                            If argumentList IsNot Nothing Then
                                argumentList = argumentList.WithCloseParenToken(SyntaxFactory.Token(SyntaxKind.CloseParenToken))
                            End If
                            Dim memberinitializer As ObjectMemberInitializerSyntax = DirectCast(PossibleInitializer, ObjectMemberInitializerSyntax)
                            Return SyntaxFactory.ObjectCreationExpression(newKeyword:=NewKeyword, attributeLists:=SyntaxFactory.List(Of AttributeListSyntax)(), type:=type1, argumentList:=argumentList, initializer:=memberinitializer)

                        Case Else
                            Throw UnexpectedValue(NameOf(PossibleInitializer))
                    End Select
                End If
                If argumentList IsNot Nothing AndAlso initializer?.GetFirstToken.IsKind(SyntaxKind.FromKeyword) Then
                    argumentList = argumentList.WithTrailingTrivia(SyntaxFactory.Space)
                End If
                Return SyntaxFactory.ObjectCreationExpression(SyntaxFactory.List(Of AttributeListSyntax)(), type1, argumentList, initializer)
            End Function

            Public Overrides Function VisitParenthesizedExpression(ByVal node As CSS.ParenthesizedExpressionSyntax) As VisualBasicSyntaxNode
                Dim Expression As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                If TypeOf Expression Is CTypeExpressionSyntax OrElse
                   TypeOf Expression Is IdentifierNameSyntax OrElse
                   TypeOf Expression Is Syntax.InvocationExpressionSyntax OrElse
                   TypeOf Expression Is TryCastExpressionSyntax Then
                    Return Expression
                End If
                Dim DeclarationToBeAdded As Syntax.LocalDeclarationStatementSyntax
                If TypeOf node.Parent Is CSS.MemberAccessExpressionSyntax OrElse TypeOf node.Parent Is CSS.ConditionalAccessExpressionSyntax Then
                    ' Statement with issues points to "Statement" Probably an Expression Statement. If this is part of a single Line If we need to go higher
                    Dim StatementWithIssue As CS.CSharpSyntaxNode = GetStatementwithIssues(node)
                    ' Statement with issues points to "Statement" Probably an Expression Statement. If this is part of an ElseIf we need to go higher
                    Dim UniqueName As String = MethodBodyVisitor.GetUniqueVariableNameInScope(node, "tempVar", mSemanticModel)
                    Dim UniqueIdentifier As IdentifierNameSyntax = SyntaxFactory.IdentifierName(SyntaxFactory.Identifier(UniqueName))
                    Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                    Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.ModifiedIdentifier(UniqueName))
                    Dim VariableDeclaration As Syntax.VariableDeclaratorSyntax
                    Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(Expression)
                    If TypeOf node.Parent Is CSS.MemberAccessExpressionSyntax Then
                        If TypeOf Expression Is BinaryConditionalExpressionSyntax OrElse TypeOf Expression Is TernaryConditionalExpressionSyntax Then
                            Dim EqualsToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.EqualsToken)
                            Dim IfKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.IfKeyword)
                            Dim ThenKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ThenKeyword)
                            If TypeOf Expression Is TernaryConditionalExpressionSyntax Then
                                Dim TExpression As TernaryConditionalExpressionSyntax = DirectCast(Expression, TernaryConditionalExpressionSyntax)
                                If TExpression.Condition.IsKind(SyntaxKind.IdentifierName) Then
                                    Dim NothingToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NothingKeyword)
                                    Dim IfStatement As Syntax.IfStatementSyntax =
                                       SyntaxFactory.IfStatement(ifKeyword:=IfKeyword, condition:=TExpression.Condition, thenKeyword:=ThenKeyword).WithConvertedLeadingTriviaFrom(node)
                                    Dim EndIfStatement As EndBlockStatementSyntax = SyntaxFactory.EndIfStatement(endKeyword:=SyntaxFactory.Token(SyntaxKind.EndKeyword), blockKeyword:=SyntaxFactory.Token(SyntaxKind.IfKeyword)).WithConvertedTrailingTriviaFrom(node)
                                    Dim IfBlockStatements As New SyntaxList(Of StatementSyntax)
                                    IfBlockStatements = IfBlockStatements.Add(SyntaxFactory.SimpleAssignmentStatement(left:=UniqueIdentifier, right:=TExpression.WhenTrue))
                                    Dim ElseBlockStatements As New SyntaxList(Of StatementSyntax)
                                    ElseBlockStatements = ElseBlockStatements.Add(SyntaxFactory.SimpleAssignmentStatement(left:=UniqueIdentifier, right:=TExpression.WhenFalse))
                                    Dim ElseBlock As ElseBlockSyntax = SyntaxFactory.ElseBlock(ElseBlockStatements)
                                    Dim IfBlockToBeAdded As StatementSyntax = SyntaxFactory.MultiLineIfBlock(
                                                                        IfStatement,
                                                                        IfBlockStatements,
                                                                        Nothing,
                                                                        ElseBlock,
                                                                        EndIfStatement)

                                    VariableDeclaration = SyntaxFactory.VariableDeclarator(Names, asClause:=Nothing, initializer:=Nothing)
                                    DeclarationToBeAdded =
                                       SyntaxFactory.LocalDeclarationStatement(
                                        ModifiersTokenList,
                                        SyntaxFactory.SingletonSeparatedList(VariableDeclaration)
                                        ).WithTrailingTrivia(SyntaxFactory.CommentTrivia($" ' TODO: Check, VB does not directly support MemberAccess off a Conditional If Expression"))

                                    StatementWithIssue.AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=False)
                                    StatementWithIssue.AddMarker(IfBlockToBeAdded, RemoveStatement:=False, AllowDuplicates:=True)
                                    Return UniqueIdentifier
                                Else
                                    ' This case is handled below
                                End If
                            ElseIf TypeOf Expression Is BinaryConditionalExpressionSyntax Then
                                Dim BExpression As BinaryConditionalExpressionSyntax = DirectCast(Expression, BinaryConditionalExpressionSyntax)
                                If BExpression.FirstExpression.IsKind(SyntaxKind.IdentifierName) Then
                                    UniqueIdentifier = DirectCast(BExpression.FirstExpression, IdentifierNameSyntax)
                                    Dim NothingToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NothingKeyword)
                                    Dim IfStatement As Syntax.IfStatementSyntax =
                                       SyntaxFactory.IfStatement(ifKeyword:=IfKeyword, SyntaxFactory.IsExpression(left:=UniqueIdentifier, right:=SyntaxFactory.NothingLiteralExpression(NothingToken)), thenKeyword:=ThenKeyword).WithConvertedLeadingTriviaFrom(node)
                                    Dim EndIfStatement As EndBlockStatementSyntax = SyntaxFactory.EndIfStatement(endKeyword:=SyntaxFactory.Token(SyntaxKind.EndKeyword), blockKeyword:=SyntaxFactory.Token(SyntaxKind.IfKeyword)).WithConvertedTrailingTriviaFrom(node)
                                    Dim Statements As New SyntaxList(Of StatementSyntax)
                                    Statements = Statements.Add(SyntaxFactory.SimpleAssignmentStatement(left:=UniqueIdentifier, right:=BExpression.SecondExpression))

                                    Dim IfBlockToBeAdded As StatementSyntax = SyntaxFactory.MultiLineIfBlock(
                                                                        ifStatement:=IfStatement,
                                                                        statements:=Statements,
                                                                        elseIfBlocks:=Nothing,
                                                                        elseBlock:=Nothing,
                                                                        endIfStatement:=EndIfStatement)

                                    VariableDeclaration = SyntaxFactory.VariableDeclarator(Names, asClause:=Nothing, initializer:=Nothing)
                                    DeclarationToBeAdded =
                                       SyntaxFactory.LocalDeclarationStatement(
                                        ModifiersTokenList,
                                        SyntaxFactory.SingletonSeparatedList(VariableDeclaration)
                                        ).WithTrailingTrivia(SyntaxFactory.CommentTrivia($" ' TODO: Check, VB does not directly support MemberAccess off a Conditional If Expression"))

                                    StatementWithIssue.AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=False)
                                    StatementWithIssue.AddMarker(IfBlockToBeAdded, RemoveStatement:=False, AllowDuplicates:=True)
                                    Return UniqueIdentifier
                                End If
                            Else
                                ' This case is handled below
                            End If
                            VariableDeclaration = SyntaxFactory.VariableDeclarator(Names, asClause:=Nothing, Initializer)
                            DeclarationToBeAdded =
                                       SyntaxFactory.LocalDeclarationStatement(
                                        ModifiersTokenList,
                                        SyntaxFactory.SingletonSeparatedList(VariableDeclaration)
                                        ).WithTrailingTrivia(SyntaxFactory.CommentTrivia($" ' TODO: Check, VB does not directly support MemberAccess off a Conditional If Expression"))

                            StatementWithIssue.AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=False)
                            Return UniqueIdentifier
                        End If
                    ElseIf TypeOf node.Parent Is CSS.ConditionalAccessExpressionSyntax Then
                        VariableDeclaration = SyntaxFactory.VariableDeclarator(Names, asClause:=Nothing, Initializer)
                        DeclarationToBeAdded = SyntaxFactory.LocalDeclarationStatement(
                                            ModifiersTokenList,
                                            SyntaxFactory.SingletonSeparatedList(VariableDeclaration)
                                            ).WithConvertedTriviaFrom(node)

                        ' Statement with issues points to "Statement" Probably an Expression Statement. If this is part of a single Line If we need to go higher
                        StatementWithIssue.AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=False)
                        Return UniqueIdentifier
                    End If
                End If

                If Expression.ContainsCommentOrDirectiveTrivia Then
                    Expression = Expression.RestructureCommentTrivia
                End If
                Dim TrailingTrivia As New List(Of SyntaxTrivia)
                TrailingTrivia.AddRange(Expression.GetTrailingTrivia)
                Dim parenthesizedExpressionSyntax1 As Syntax.ParenthesizedExpressionSyntax = SyntaxFactory.ParenthesizedExpression(Expression.WithoutTrailingTrivia).WithTrailingTrivia(TrailingTrivia)
                Return parenthesizedExpressionSyntax1
            End Function

            Public Overrides Function VisitParenthesizedLambdaExpression(ByVal node As CSS.ParenthesizedLambdaExpressionSyntax) As VisualBasicSyntaxNode
                Return ConvertLambdaExpression(node, node.Body, node.ParameterList.Parameters, SyntaxFactory.TokenList(node.AsyncKeyword))
            End Function

            Public Overrides Function VisitParenthesizedVariableDesignation(node As ParenthesizedVariableDesignationSyntax) As VisualBasicSyntaxNode
                Dim Variables As New List(Of Syntax.ModifiedIdentifierSyntax)
                For i As Integer = 0 To node.Variables.Count - 1
                    Dim VariableDeclarator As Syntax.ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(node.Variables(i).ToString)
                    Variables.Add(VariableDeclarator)
                Next

                Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SeparatedList(Of ModifiedIdentifierSyntax)(Variables)
                Dim ObjectType As Syntax.PredefinedTypeSyntax = SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ObjectKeyword))
                Return SyntaxFactory.VariableDeclarator(Names, SyntaxFactory.SimpleAsClause(ObjectType), initializer:=Nothing)
            End Function
            Public Overrides Function VisitPostfixUnaryExpression(ByVal node As CSS.PostfixUnaryExpressionSyntax) As VisualBasicSyntaxNode
                Dim kind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(node), IsModule, TokenContext.Local)
                Dim OperandExpression As ExpressionSyntax = DirectCast(node.Operand.Accept(Me), ExpressionSyntax)
                If TypeOf node.Parent Is CSS.ExpressionStatementSyntax OrElse TypeOf node.Parent Is CSS.ForStatementSyntax Then
                    Return SyntaxFactory.AssignmentStatement(
                        ConvertToken(CS.CSharpExtensions.Kind(node), IsModule, TokenContext.Local),
                        OperandExpression, SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)),
                        LiteralExpression_1)
                Else
                    Dim OperatorName As String
                    Dim minMax As String
                    Dim op As SyntaxKind
                    If kind = SyntaxKind.AddAssignmentStatement Then
                        OperatorName = "Increment"
                        minMax = "Min"
                        op = SyntaxKind.SubtractExpression
                    Else
                        OperatorName = "Decrement"
                        minMax = "Max"
                        op = SyntaxKind.AddExpression
                    End If
                    Dim MathExpression As Syntax.NameSyntax = SyntaxFactory.ParseName("Math." & minMax)
                    Dim InterlockedExpressionName As Syntax.NameSyntax = SyntaxFactory.ParseName("System.Threading.Interlocked." & OperatorName)

                    Dim OperandArgument As SimpleArgumentSyntax = SyntaxFactory.SimpleArgument(OperandExpression)

                    Dim OperandArgumentList As ArgumentListSyntax = SyntaxFactory.ArgumentList(SyntaxFactory.SingletonSeparatedList(Of ArgumentSyntax)(OperandArgument))
                    Dim ArgumentInvocationExpression As Syntax.InvocationExpressionSyntax = SyntaxFactory.InvocationExpression(InterlockedExpressionName, OperandArgumentList)
                    Dim SecondArgumentSyntax As SimpleArgumentSyntax = SyntaxFactory.SimpleArgument(SyntaxFactory.BinaryExpression(op, OperandExpression, SyntaxFactory.Token(GetExpressionOperatorTokenKind(op)), LiteralExpression_1))
                    Return SyntaxFactory.InvocationExpression(
                        MathExpression,
                        SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList((New ArgumentSyntax() {SyntaxFactory.SimpleArgument(ArgumentInvocationExpression),
                                                                                                      SecondArgumentSyntax})))).WithConvertedTriviaFrom(node)
                End If
            End Function
            Public Overrides Function VisitPrefixUnaryExpression(ByVal node As CSS.PrefixUnaryExpressionSyntax) As VisualBasicSyntaxNode
                Dim kind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(node), IsModule, TokenContext.Local)
                If kind = 1999 Then
                    node.FirstAncestorOrSelf(Of CSS.StatementSyntax).AddMarker(CommentOutUnsupportedStatements(node.FirstAncestorOrSelf(Of CSS.StatementSyntax), "IndirectPointer Expressions"), RemoveStatement:=True, AllowDuplicates:=False)
                    Return SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword))
                End If
                Dim OperandExpression As ExpressionSyntax = DirectCast(node.Operand.Accept(Me), ExpressionSyntax)
                If TypeOf node.Parent Is CSS.ExpressionStatementSyntax Then
                    Return SyntaxFactory.AssignmentStatement(kind, OperandExpression, SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)), LiteralExpression_1).WithConvertedTriviaFrom(node)
                End If
                If kind = SyntaxKind.AddAssignmentStatement OrElse kind = SyntaxKind.SubtractAssignmentStatement Then
                    If node.Parent.IsKind(CS.SyntaxKind.ForStatement) Then
                        If kind = SyntaxKind.AddAssignmentStatement Then
                            Return SyntaxFactory.AddAssignmentStatement(OperandExpression, SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)), LiteralExpression_1).WithConvertedTriviaFrom(node)
                        Else
                            Return SyntaxFactory.SubtractAssignmentStatement(OperandExpression, SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)), LiteralExpression_1).WithConvertedTriviaFrom(node)
                        End If
                    Else
                        Dim operatorName As String = If(kind = SyntaxKind.AddAssignmentStatement, "Increment", "Decrement")
                        Dim MathExpression As Syntax.NameSyntax = SyntaxFactory.ParseName("System.Threading.Interlocked." & operatorName)
                        Return SyntaxFactory.InvocationExpression(MathExpression, SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList((New ArgumentSyntax() {SyntaxFactory.SimpleArgument(OperandExpression)}))))
                    End If
                End If
#Enable Warning CC0013
                If kind = SyntaxKind.AddressOfExpression Then
                    Dim SpaceTriviaList As SyntaxTriviaList
                    SpaceTriviaList = SpaceTriviaList.Add(SyntaxFactory.Space)
                    Dim AddressOfToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.AddressOfKeyword).With(SpaceTriviaList, SpaceTriviaList)
                    Return SyntaxFactory.AddressOfExpression(AddressOfToken, OperandExpression).WithConvertedTriviaFrom(node)
                End If
                Return SyntaxFactory.UnaryExpression(kind, SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)), OperandExpression).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitSimpleLambdaExpression(ByVal node As CSS.SimpleLambdaExpressionSyntax) As VisualBasicSyntaxNode
                Return ConvertLambdaExpression(node, node.Body, SyntaxFactory.SingletonSeparatedList(node.Parameter), SyntaxFactory.TokenList(node.AsyncKeyword)).WithConvertedTriviaFrom(node)
            End Function

            ''' <summary>
            ''' Maps sizeof to Len(New {Type})
            ''' </summary>
            ''' <param name="node"></param>
            ''' <returns></returns>
            ''' <remarks>Added by PC</remarks>
            Public Overrides Function VisitSizeOfExpression(node As SizeOfExpressionSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.ParseExpression($"Len(New {node.Type.ToString}()) ")
            End Function
            Public Overrides Function VisitThisExpression(ByVal node As CSS.ThisExpressionSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.MeExpression().WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitThrowExpression(node As ThrowExpressionSyntax) As VisualBasicSyntaxNode
                Dim Expression As ExpressionSyntax = DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                Return SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression($"[Throw] ({Expression.ToString})"))
            End Function
            Public Overrides Function VisitTupleElement(node As CSS.TupleElementSyntax) As VisualBasicSyntaxNode
                Try
                    If node.Identifier.ValueText.IsEmptyNullOrWhitespace Then
                        Dim typedTupleElementSyntax1 As TypedTupleElementSyntax = SyntaxFactory.TypedTupleElement(DirectCast(node.Type.Accept(Me), TypeSyntax))
                        Return typedTupleElementSyntax1
                    End If
                    Dim namedTupleElementSyntax1 As NamedTupleElementSyntax = SyntaxFactory.NamedTupleElement(GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False).WithConvertedTriviaFrom(node.Type), SyntaxFactory.SimpleAsClause(DirectCast(node.Type.Accept(Me).WithConvertedTriviaFrom(node.Identifier), TypeSyntax)))
                    Return namedTupleElementSyntax1
                Catch ex As Exception
                    Stop
                End Try
                Throw UnreachableException
            End Function
            ''' <summary>
            ''' This only returns for Names VB
            ''' </summary>
            ''' <param name="node"></param>
            ''' <returns></returns>
            Public Overrides Function VisitTupleExpression(node As CSS.TupleExpressionSyntax) As VisualBasicSyntaxNode
                Dim lArgumentSyntax As New List(Of SimpleArgumentSyntax)
                If TypeOf node.Arguments(0).Expression IsNot DeclarationExpressionSyntax Then
                    For Each a As CSS.ArgumentSyntax In node.Arguments
                        lArgumentSyntax.Add(DirectCast(a.Accept(Me), SimpleArgumentSyntax))
                    Next
                    Return SyntaxFactory.TupleExpression(SyntaxFactory.SeparatedList(lArgumentSyntax))
                End If
                Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(
                    SyntaxFactory.Token(SyntaxKind.DimKeyword))

                For Each a As CSS.ArgumentSyntax In node.Arguments
                    Dim Identifier As Syntax.IdentifierNameSyntax
                    If a.Expression.IsKind(CS.SyntaxKind.IdentifierName) Then
                        Identifier = DirectCast(DirectCast(a.Expression, CSS.IdentifierNameSyntax).Accept(Me), IdentifierNameSyntax)
                    Else
                        Dim d As DeclarationExpressionSyntax = DirectCast(a.Expression, DeclarationExpressionSyntax)
                        Identifier = DirectCast(d.Designation.Accept(Me), Syntax.IdentifierNameSyntax)
                    End If
                    lArgumentSyntax.Add(SyntaxFactory.SimpleArgument(Identifier))
                Next
                Return SyntaxFactory.TupleExpression(SyntaxFactory.SeparatedList(lArgumentSyntax))
            End Function

            Public Overrides Function VisitTupleType(node As CSS.TupleTypeSyntax) As VisualBasicSyntaxNode
                Dim SSList As New List(Of Syntax.TupleElementSyntax)
                SSList.AddRange(node.Elements.Select(Function(a As CSS.TupleElementSyntax) DirectCast(a.Accept(Me), Syntax.TupleElementSyntax)))
                Return SyntaxFactory.TupleType(SSList.ToArray)
            End Function
            Public Overrides Function VisitTypeOfExpression(ByVal node As CSS.TypeOfExpressionSyntax) As VisualBasicSyntaxNode
                If TypeOf node.Type Is CSS.GenericNameSyntax Then
                    Dim NodeType As CSS.GenericNameSyntax = CType(node.Type, CSS.GenericNameSyntax)
                    Dim ArgumentList As CSS.TypeArgumentListSyntax = NodeType.TypeArgumentList
                    If ArgumentList.Arguments.Count = 1 AndAlso ArgumentList.Arguments(0).IsKind(CS.SyntaxKind.OmittedTypeArgument) Then
                        Return SyntaxFactory.GetTypeExpression(SyntaxFactory.ParseTypeName($"{NodeType.Identifier.ValueText}()"))
                    End If
                End If
                Return SyntaxFactory.GetTypeExpression(DirectCast(node.Type.Accept(Me), TypeSyntax))
            End Function
        End Class
    End Class
End Namespace