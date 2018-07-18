Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

Public Module MethodKindExtensions
    <ExcludeFromCodeCoverage>
    <Extension>
    Public Function IsPropertyAccessor(ByVal kind As MethodKind) As Boolean
        Return kind = MethodKind.PropertyGet OrElse kind = MethodKind.PropertySet
    End Function
End Module

