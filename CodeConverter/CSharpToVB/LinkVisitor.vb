Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports OrderingSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.OrderingSyntax
Imports QueryClauseSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.QueryClauseSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)
            Private Iterator Function ConvertQueryBody(ByVal body As CSS.QueryBodySyntax) As IEnumerable(Of QueryClauseSyntax)
                'If TypeOf body.SelectOrGroup Is CSS.GroupClauseSyntax AndAlso body.Continuation Is Nothing Then
                '    Throw New NotSupportedException("group by clause without into Not supported In VB")
                'End If
                If TypeOf body.SelectOrGroup Is CSS.SelectClauseSyntax Then
                    Yield CType(body.SelectOrGroup.Accept(Me), QueryClauseSyntax)
                Else
                    Dim group As CSS.GroupClauseSyntax = CType(body.SelectOrGroup, CSS.GroupClauseSyntax)
                    Dim Items As ExpressionRangeVariableSyntax = SyntaxFactory.ExpressionRangeVariable(CType(group.GroupExpression.Accept(Me), ExpressionSyntax))
                    Dim Identifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(GeneratePlaceholder("groupByKey"))
                    Dim NameEquals As VariableNameEqualsSyntax = SyntaxFactory.VariableNameEquals(Identifier)
                    Dim Expression As ExpressionSyntax = CType(group.ByExpression.Accept(Me), ExpressionSyntax)
                    Dim Keys As ExpressionRangeVariableSyntax
                    Dim FunctionName As SyntaxToken
                    If body.Continuation Is Nothing Then
                        Keys = SyntaxFactory.ExpressionRangeVariable(Expression)
                        FunctionName = SyntaxFactory.Identifier("Group")
                    Else
                        Keys = SyntaxFactory.ExpressionRangeVariable(NameEquals, Expression)
                        FunctionName = GenerateSafeVBToken(body.Continuation.Identifier, False)
                    End If
                    Dim Aggregation As FunctionAggregationSyntax = SyntaxFactory.FunctionAggregation(FunctionName)
                    Dim AggrationRange As AggregationRangeVariableSyntax = SyntaxFactory.AggregationRangeVariable(Aggregation)
                    Yield SyntaxFactory.GroupByClause(SyntaxFactory.SingletonSeparatedList(Items), SyntaxFactory.SingletonSeparatedList(Keys), SyntaxFactory.SingletonSeparatedList(AggrationRange))
                    If body.Continuation?.Body IsNot Nothing Then
                        For Each clause As QueryClauseSyntax In ConvertQueryBody(body.Continuation.Body)
                            Yield clause
                        Next
                    End If
                End If
            End Function
            Private Function GeneratePlaceholder(ByVal v As String) As String
                Return $"__{v}{Math.Min(Threading.Interlocked.Increment(placeholder), placeholder - 1)}__"
            End Function
            Public Overrides Function VisitFromClause(ByVal node As CSS.FromClauseSyntax) As VisualBasicSyntaxNode
                Dim expression As ExpressionSyntax = CType(node.Expression.Accept(Me), ExpressionSyntax).WithConvertedTriviaFrom(node.Expression)
                Dim identifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(node.Identifier, False))
                Dim CollectionRangevariable As CollectionRangeVariableSyntax = SyntaxFactory.CollectionRangeVariable(identifier, expression)
                Return SyntaxFactory.FromClause(CollectionRangevariable).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitJoinClause(ByVal node As CSS.JoinClauseSyntax) As VisualBasicSyntaxNode
#Disable Warning CC0013 ' Use Ternary operator.
                If node.Into IsNot Nothing Then
#Enable Warning CC0013 ' Use Ternary operator.
                    Return SyntaxFactory.GroupJoinClause(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.CollectionRangeVariable(SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(node.Identifier, False)), If(node.Type Is Nothing, Nothing, SyntaxFactory.SimpleAsClause(CType(node.Type.Accept(Me), TypeSyntax))), CType(node.InExpression.Accept(Me), ExpressionSyntax))), SyntaxFactory.SingletonSeparatedList(SyntaxFactory.JoinCondition(CType(node.LeftExpression.Accept(Me), ExpressionSyntax), CType(node.RightExpression.Accept(Me), ExpressionSyntax))), SyntaxFactory.SingletonSeparatedList(SyntaxFactory.AggregationRangeVariable(SyntaxFactory.VariableNameEquals(SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(node.Into.Identifier, False))), SyntaxFactory.GroupAggregation()))).WithConvertedTriviaFrom(node)
                Else
                    Return SyntaxFactory.SimpleJoinClause(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.CollectionRangeVariable(SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(node.Identifier, False)), If(node.Type Is Nothing, Nothing, SyntaxFactory.SimpleAsClause(CType(node.Type.Accept(Me), TypeSyntax))), CType(node.InExpression.Accept(Me), ExpressionSyntax))), SyntaxFactory.SingletonSeparatedList(SyntaxFactory.JoinCondition(CType(node.LeftExpression.Accept(Me), ExpressionSyntax), CType(node.RightExpression.Accept(Me), ExpressionSyntax)))).WithConvertedTriviaFrom(node)
                End If
            End Function

            Public Overrides Function VisitLetClause(ByVal node As CSS.LetClauseSyntax) As VisualBasicSyntaxNode
                Dim Identifier As SyntaxToken = GenerateSafeVBToken(node.Identifier, False)
                Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(Identifier)
                Dim NameEquals As VariableNameEqualsSyntax = SyntaxFactory.VariableNameEquals(ModifiedIdentifier)
                Dim Expression As ExpressionSyntax = CType(node.Expression.Accept(Me), ExpressionSyntax)
                Dim ExpressionRangeVariable As ExpressionRangeVariableSyntax = SyntaxFactory.ExpressionRangeVariable(NameEquals, Expression)
                Return SyntaxFactory.LetClause(SyntaxFactory.SingletonSeparatedList(ExpressionRangeVariable)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitOrderByClause(ByVal node As CSS.OrderByClauseSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.OrderByClause(SyntaxFactory.SeparatedList(node.Orderings.[Select](Function(o As CSS.OrderingSyntax) CType(o.Accept(Me), OrderingSyntax)))).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitOrdering(ByVal node As CSS.OrderingSyntax) As VisualBasicSyntaxNode
                If node.IsKind(CS.SyntaxKind.DescendingOrdering) Then
                    Return SyntaxFactory.Ordering(SyntaxKind.DescendingOrdering, CType(node.Expression.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
                Else
                    Return SyntaxFactory.Ordering(SyntaxKind.AscendingOrdering, CType(node.Expression.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
                End If
            End Function

            Public Overrides Function VisitQueryExpression(ByVal node As CSS.QueryExpressionSyntax) As VisualBasicSyntaxNode
                ' From trivia handled in VisitFromClause
                Dim FromClause As QueryClauseSyntax = CType(node.FromClause.Accept(Me), QueryClauseSyntax)
                Dim BodyClauses As IEnumerable(Of QueryClauseSyntax) = node.Body.Clauses.[Select](Function(c As CSS.QueryClauseSyntax) CType(c.Accept(Me).WithConvertedTriviaFrom(c), QueryClauseSyntax))
                Dim Body As IEnumerable(Of QueryClauseSyntax) = ConvertQueryBody(node.Body)
                Return SyntaxFactory.QueryExpression(SyntaxFactory.SingletonList(FromClause).AddRange(BodyClauses).AddRange(Body)).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitSelectClause(ByVal node As CSS.SelectClauseSyntax) As VisualBasicSyntaxNode
                Dim expression As ExpressionSyntax = CType(node.Expression.Accept(Me), ExpressionSyntax)
                Return SyntaxFactory.SelectClause(SyntaxFactory.ExpressionRangeVariable(expression)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitWhereClause(ByVal node As CSS.WhereClauseSyntax) As VisualBasicSyntaxNode
                Dim condition As ExpressionSyntax = CType(node.Condition.Accept(Me), ExpressionSyntax)
                Return SyntaxFactory.WhereClause(condition).WithConvertedTriviaFrom(node)
            End Function
        End Class
    End Class
End Namespace