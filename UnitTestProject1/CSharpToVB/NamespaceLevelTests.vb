Imports Xunit

Namespace CodeConverter.Tests.CSharpToVB

    <TestClass()> Public Class NamespaceLevelTests
        Inherits ConverterTestBase


        <Fact>
        Public Shared Sub CSharpToVB_Namespace()
            TestConversionCSharpToVisualBasic("namespace Test
{
    
}", "Namespace Test

End Namespace")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_TopLevelAttribute()
            TestConversionCSharpToVisualBasic("[assembly: CLSCompliant(true)]", "<Assembly: CLSCompliant(True)>")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Imports()
            TestConversionCSharpToVisualBasic("using SomeNamespace;
using VB = Microsoft.VisualBasic;", "Imports SomeNamespace
Imports VB = Microsoft.VisualBasic")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Class()
            TestConversionCSharpToVisualBasic("namespace Test.@class
{
    class TestClass<T>
    {
    }
}", "Namespace Test.[class]

    Class TestClass(Of T)
    End Class
End Namespace")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_InternalStaticClass()
            TestConversionCSharpToVisualBasic("namespace Test.@class
{
    internal static class TestClass
    {
        public static void Test() {}
        static void Test2() {}
    }
}", "Namespace Test.[class]

    Friend Module TestClass

        Sub Test()
        End Sub
        Private Sub Test2()
        End Sub

    End Module
End Namespace")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_AbstractClass()
            TestConversionCSharpToVisualBasic("namespace Test.@class
{
    abstract class TestClass
    {
    }
}", "Namespace Test.[class]

    MustInherit Class TestClass
    End Class
End Namespace")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_SealedClass()
            TestConversionCSharpToVisualBasic("namespace Test.@class
{
    sealed class TestClass
    {
    }
}", "Namespace Test.[class]

    NotInheritable Class TestClass
    End Class
End Namespace")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Interface()
            TestConversionCSharpToVisualBasic("interface ITest : System.IDisposable
{
    void Test ();
}", "Interface ITest
    Inherits System.IDisposable

    Sub Test()

End Interface")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Enum()
            TestConversionCSharpToVisualBasic("internal enum ExceptionResource
{
    Argument_ImplementIComparable,
    ArgumentOutOfRange_NeedNonNegNum,
    ArgumentOutOfRange_NeedNonNegNumRequired,
    Arg_ArrayPlusOffTooSmall
}", "Friend Enum ExceptionResource
    Argument_ImplementIComparable
    ArgumentOutOfRange_NeedNonNegNum
    ArgumentOutOfRange_NeedNonNegNumRequired
    Arg_ArrayPlusOffTooSmall
End Enum")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ClassInheritanceList()
            TestConversionCSharpToVisualBasic("abstract class ClassA : System.IDisposable
{
    protected abstract void Test();
}", "MustInherit Class ClassA
    Implements System.IDisposable

    Protected MustOverride Sub Test()
End Class")
            TestConversionCSharpToVisualBasic("abstract class ClassA : System.EventArgs, System.IDisposable
{
    protected abstract void Test();
}", "MustInherit Class ClassA
    Inherits System.EventArgs
    Implements System.IDisposable

    Protected MustOverride Sub Test()
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Struct()
            TestConversionCSharpToVisualBasic("struct MyType : System.IComparable<MyType>
{
    void Test() {}
}", "Structure MyType
    Implements System.IComparable(Of MyType)

    Private Sub Test()
    End Sub

End Structure")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Delegate()
            TestConversionCSharpToVisualBasic("public delegate void Test();", "Public Delegate Sub Test()")
            TestConversionCSharpToVisualBasic("public delegate int Test();", "Public Delegate Function Test() As Integer")
            TestConversionCSharpToVisualBasic("public delegate void Test(int x);", "Public Delegate Sub Test(x As Integer)")
            TestConversionCSharpToVisualBasic("public delegate void Test(ref int x);", "Public Delegate Sub Test(ByRef x As Integer)")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_MoveImportsStatement()
            TestConversionCSharpToVisualBasic("namespace test { using SomeNamespace; }", "Imports SomeNamespace

Namespace test
End Namespace")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ClassImplementsInterface()
            TestConversionCSharpToVisualBasic("using System; class test : IComparable { }",
                                              "Imports System

Class test
    Implements IComparable

End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ClassImplementsInterface2()
            TestConversionCSharpToVisualBasic("class test : System.IComparable { }", "Class test
    Implements System.IComparable

End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ClassInheritsClass()
            TestConversionCSharpToVisualBasic("using System.IO; class test : InvalidDataException { }", "Imports System.IO

Class test
    Inherits InvalidDataException

End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ClassInheritsClass2()
            TestConversionCSharpToVisualBasic("class test : System.IO.InvalidDataException { }", "Class test
    Inherits System.IO.InvalidDataException

End Class")
        End Sub
    End Class
End Namespace
