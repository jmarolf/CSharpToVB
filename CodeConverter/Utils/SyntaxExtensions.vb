Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.CSharp.Syntax

Namespace IVisualBasicCode.CodeConverter.Util
    Public Module SyntaxExtensions

        <Extension()>
        Public Function IsParentKind(ByVal node As SyntaxNode, ByVal kind As SyntaxKind) As Boolean
            Return node IsNot Nothing AndAlso node.Parent.IsKind(kind)
        End Function

        <Extension()>
        Public Function IsParentKind(ByVal node As SyntaxToken, ByVal kind As SyntaxKind) As Boolean
            Return node.Parent IsNot Nothing AndAlso node.Parent.IsKind(kind)
        End Function

    End Module
End Namespace