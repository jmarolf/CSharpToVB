Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports System.Runtime.CompilerServices

Namespace IVisualBasicCode.CodeConverter.Util
    Friend Module VBUtil
        ''' <summary>
        '''Get VB Expression Operator Token Kind from VB Expression Kind
        ''' </summary>
        ''' <param name="op"></param>
        ''' <returns></returns>
        Public Function GetExpressionOperatorTokenKind(ByVal op As SyntaxKind) As SyntaxKind
            Select Case op
                Case SyntaxKind.EqualsExpression
                    Return SyntaxKind.EqualsToken
                Case SyntaxKind.NotEqualsExpression
                    Return SyntaxKind.LessThanGreaterThanToken
                Case SyntaxKind.GreaterThanExpression
                    Return SyntaxKind.GreaterThanToken
                Case SyntaxKind.GreaterThanOrEqualExpression
                    Return SyntaxKind.GreaterThanEqualsToken
                Case SyntaxKind.LessThanExpression
                    Return SyntaxKind.LessThanToken
                Case SyntaxKind.LessThanOrEqualExpression
                    Return SyntaxKind.LessThanEqualsToken
                Case SyntaxKind.OrExpression
                    Return SyntaxKind.OrKeyword
                Case SyntaxKind.OrElseExpression
                    Return SyntaxKind.OrElseKeyword
                Case SyntaxKind.AndExpression
                    Return SyntaxKind.AndKeyword
                Case SyntaxKind.AndAlsoExpression
                    Return SyntaxKind.AndAlsoKeyword
                Case SyntaxKind.AddExpression
                    Return SyntaxKind.PlusToken
                Case SyntaxKind.ConcatenateExpression
                    Return SyntaxKind.AmpersandToken
                Case SyntaxKind.SubtractExpression
                    Return SyntaxKind.MinusToken
                Case SyntaxKind.MultiplyExpression
                    Return SyntaxKind.AsteriskToken
                Case SyntaxKind.DivideExpression
                    Return SyntaxKind.SlashToken
                Case SyntaxKind.ModuloExpression
                    Return SyntaxKind.ModKeyword
                Case SyntaxKind.SimpleAssignmentStatement
                    Return SyntaxKind.EqualsToken
                Case SyntaxKind.LeftShiftAssignmentStatement
                    Return SyntaxKind.LessThanLessThanEqualsToken
                Case SyntaxKind.RightShiftAssignmentStatement
                    Return SyntaxKind.GreaterThanGreaterThanEqualsToken
                Case SyntaxKind.AddAssignmentStatement
                    Return SyntaxKind.PlusEqualsToken
                Case SyntaxKind.SubtractAssignmentStatement
                    Return SyntaxKind.MinusEqualsToken
                Case SyntaxKind.MultiplyAssignmentStatement
                    Return SyntaxKind.AsteriskEqualsToken
                Case SyntaxKind.DivideAssignmentStatement
                    Return SyntaxKind.SlashEqualsToken
                Case SyntaxKind.UnaryPlusExpression
                    Return SyntaxKind.PlusToken
                Case SyntaxKind.UnaryMinusExpression
                    Return SyntaxKind.MinusToken
                Case SyntaxKind.NotExpression
                    Return SyntaxKind.NotKeyword
                Case SyntaxKind.RightShiftExpression
                    Return SyntaxKind.GreaterThanGreaterThanToken
                Case SyntaxKind.LeftShiftExpression
                    Return SyntaxKind.LessThanLessThanToken
                Case SyntaxKind.AddressOfExpression
                    Return SyntaxKind.AddressOfKeyword
                Case SyntaxKind.ExclusiveOrExpression
                    Return SyntaxKind.XorKeyword
            End Select

            Throw New ArgumentOutOfRangeException($"op = {op.ToString}")
        End Function
        <Extension()>
        Public Function IsKind(ByVal node As SyntaxNode, ByVal kind1 As SyntaxKind, ByVal kind2 As SyntaxKind) As Boolean
            If node Is Nothing Then
                Return False
            End If

            Dim vbKind As SyntaxKind = node.Kind()
            Return vbKind = kind1 OrElse vbKind = kind2
        End Function
    End Module
End Namespace
