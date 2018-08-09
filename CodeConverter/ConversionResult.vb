Option Explicit On
Option Infer Off
Option Strict On
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports IVisualBasicCode.CodeConverter.Util
Imports System.Diagnostics.CodeAnalysis

Namespace IVisualBasicCode.CodeConverter
    Public Class ConversionResult
        Public Sub New(ByVal _ConvertedTree As SyntaxNode, InputLanguage As String, OutputLanguage As String)
            ConvertedCode = _ConvertedTree.NormalizeWhitespaceEx(useDefaultCasing:=True, PreserveCRLF:=True).ToFullString
            ConvertedTree = DirectCast(_ConvertedTree, VisualBasicSyntaxNode)
            Exceptions = New List(Of Exception)
            SourceLanguage = InputLanguage
            Success = True
            TargetLanguage = OutputLanguage
        End Sub

        <ExcludeFromCodeCoverage>
        Public Sub New(ParamArray ByVal exceptions() As Exception)
            Success = False
            Me.Exceptions = exceptions
        End Sub

        Public ReadOnly Property ConvertedCode As String
        Public Property ConvertedTree As VisualBasicSyntaxNode
        Public Property Exceptions As IReadOnlyList(Of Exception)
        Public Property FilteredListOfFailures As List(Of Diagnostic)
        Public Property SourceLanguage As String
        Public Property Success As Boolean
        Public Property TargetLanguage As String
    End Class
End Namespace
