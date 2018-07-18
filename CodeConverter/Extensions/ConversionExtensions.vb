Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports CS = Microsoft.CodeAnalysis.CSharp

Namespace IVisualBasicCode.CodeConverter.VB
    Friend Module ConversionExtensions
        Private Function MatchesNamespaceOrRoot(ByVal arg As SyntaxNode) As Boolean
            Return TypeOf arg Is CS.Syntax.NamespaceDeclarationSyntax OrElse TypeOf arg Is CS.Syntax.CompilationUnitSyntax
        End Function

        <Extension>
        Public Function HasUsingDirective(ByVal tree As CS.CSharpSyntaxTree, ByVal fullName As String) As Boolean
            If tree Is Nothing Then
                Throw New ArgumentNullException(NameOf(tree))
            End If
            If String.IsNullOrWhiteSpace(fullName) Then
                Throw New ArgumentException("given namespace cannot be null or empty.", NameOf(fullName))
            End If
            fullName = fullName.Trim()
            Return tree.GetRoot().DescendantNodes(AddressOf MatchesNamespaceOrRoot).OfType(Of CS.Syntax.UsingDirectiveSyntax)().Any(Function(u As CS.Syntax.UsingDirectiveSyntax) u.Name.ToString().Equals(fullName, StringComparison.OrdinalIgnoreCase))
        End Function
        <Extension>
        Public Iterator Function IndexedSelect(Of T, R)(ByVal source As IEnumerable(Of T), ByVal transform As Func(Of Integer, T, R)) As IEnumerable(Of R)
            Dim i As Integer = 0
            For Each item As T In source
                Yield transform(i, item)
                i += 1
            Next item
        End Function
    End Module
End Namespace
