Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace IVisualBasicCode.CodeConverter.Util

    Partial Public Module SyntaxExtensions
        ''' <summary>
        ''' Creates a new syntax node with all whitespace and end of line trivia replaced with
        ''' regularly formatted trivia.        ''' </summary>
        ''' <typeparam name="TNode">>The type of the node.</typeparam>
        ''' <param name="node">The node to format.</param>
        ''' <param name="useDefaultCasing"></param>
        ''' <param name="indentation">An optional sequence of whitespace characters that defines a single level of indentation.</param>
        ''' <param name="elasticTrivia"></param>
        ''' <param name="PreserveCRLF"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function NormalizeWhitespaceEx(Of TNode As SyntaxNode)(node As TNode, useDefaultCasing As Boolean, indentation As String, elasticTrivia As Boolean, PreserveCRLF As Boolean) As TNode
            Return DirectCast(SyntaxNormalizer.Normalize(node, indentation, DefaultEOL, elasticTrivia, useDefaultCasing, PreserveCRLF), TNode)
        End Function

        <Extension()>
        Public Function NormalizeWhitespaceEx(Of TNode As SyntaxNode)(node As TNode, useDefaultCasing As Boolean, Optional indentation As String = "    ", Optional eol As String = DefaultEOL, Optional elasticTrivia As Boolean = False, Optional PreserveCRLF As Boolean = False) As TNode
            Return DirectCast(SyntaxNormalizer.Normalize(node, indentation, eol, elasticTrivia, useDefaultCasing, PreserveCRLF), TNode)
        End Function

    End Module
End Namespace