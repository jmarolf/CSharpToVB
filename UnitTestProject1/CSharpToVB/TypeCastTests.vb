﻿Imports Xunit

Namespace CodeConverter.Tests.CSharpToVB

    <TestClass()> Public Class TypeCastTests
        Inherits ConverterTestBase
        <Fact>
        Public Shared Sub CSharpToVB_CastObjectToInteger()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = 5;
    int i = (int) o;
}
", "Private Sub Test()
    Dim o As Object = 5
    Dim i As Integer = CInt(o)
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_CastObjectToString()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = ""Test"";
    string s = (string) o;
}
", "Private Sub Test()
    Dim o As Object = ""Test""
    Dim s As String = CStr(o)
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_CastObjectToGenericList()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = new System.Collections.Generic.List<int>();
    System.Collections.Generic.List<int> l = (System.Collections.Generic.List<int>) o;
}
", "Private Sub Test()
    Dim o As Object = New System.Collections.Generic.List(Of Integer)()
    Dim l As System.Collections.Generic.List(Of Integer) = CType(o, System.Collections.Generic.List(Of Integer))
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_TryCastObjectToInteger()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = 5;
    System.Nullable<int> i = o as int;
}
", "Private Sub Test()
    Dim o As Object = 5
    Dim i As System.Nullable(Of Integer) = TryCast(o, Integer)
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_TryCastObjectToGenericList()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = new System.Collections.Generic.List<int>();
    System.Collections.Generic.List<int> l = o as System.Collections.Generic.List<int>;
}
", "Private Sub Test()
    Dim o As Object = New System.Collections.Generic.List(Of Integer)()
    Dim l As System.Collections.Generic.List(Of Integer) = TryCast(o, System.Collections.Generic.List(Of Integer))
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_CastConstantNumberToLong()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = 5L;
}
", "Private Sub Test()
    Dim o As Object = 5L
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_CastConstantNumberToFloat()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = 5.0f;
}
", "Private Sub Test()
    Dim o As Object = 5F
End Sub
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_CastConstantNumberToDecimal()
            TestConversionCSharpToVisualBasic("void Test()
{
    object o = 5.0m;
}
", "Private Sub Test()
    Dim o As Object = 5.0D
End Sub
")
        End Sub
    End Class
End Namespace

