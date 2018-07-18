Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices

Namespace IVisualBasicCode.CodeConverter.Util

    Partial Module EnumerableExtensions

        Private ReadOnly s_notNullTest As Func(Of Object, Boolean) = Function(x As Object) x IsNot Nothing

        <Extension>
        Public Function Contains(Of T)(ByVal sequence As IEnumerable(Of T), ByVal predicate As Func(Of T, Boolean)) As Boolean
            Return sequence.Any(predicate)
        End Function

        <Extension()>
        Public Function FirstOrNullable(Of T As Structure)(ByVal source As IEnumerable(Of T), ByVal predicate As Func(Of T, Boolean)) As T?
            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If

            Return source.Cast(Of T?)().FirstOrDefault(Function(v As T?) predicate(v.Value))
        End Function

    End Module
End Namespace