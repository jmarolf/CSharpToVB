Imports System.Runtime.CompilerServices

Public Module DefaultExtension
    <MethodImpl(MethodImplOptions.AggressiveInlining)>
    <Extension>
    Public Function [Or](Of T)(test As T, [default] As T, Optional assert As Func(Of T, Boolean) = Nothing) As T
        Return If(Not assert Is Nothing, If(assert(test), test, [default]), If(test Is Nothing, [default], test))
    End Function
End Module
