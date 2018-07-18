Imports Xunit

Namespace CodeConverter.Tests.VBToCSharp

    <TestClass()> Public Class NamespaceLevelTests
        Inherits ConverterTestBase
        <Fact>
        Public Sub VBToCSharp_Namespace()
            TestConversionVisualBasicToCSharp("Namespace Test
End Namespace", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

namespace Test
{
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_TopLevelAttribute()
            TestConversionVisualBasicToCSharp("<Assembly: CLSCompliant(True)>", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

[assembly: CLSCompliant(true)]")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Imports()
            TestConversionVisualBasicToCSharp("Imports SomeNamespace
Imports VB = Microsoft.VisualBasic", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;
using SomeNamespace;
using VB = Microsoft.VisualBasic;")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Class()
            TestConversionVisualBasicToCSharp("Namespace Test.[class]
    Class TestClass(Of T)
    End Class
End Namespace", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

namespace Test.@class
{
    class TestClass<T>
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_InternalStaticClass()
            TestConversionVisualBasicToCSharp("Namespace Test.[class]
    Friend Module TestClass
        Sub Test()
        End Sub

        Private Sub Test2()
        End Sub
    End Module
End Namespace", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

namespace Test.@class
{
    internal static class TestClass
    {
        public static void Test()
        {
        }

        private static void Test2()
        {
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_AbstractClass()
            TestConversionVisualBasicToCSharp("Namespace Test.[class]
    MustInherit Class TestClass
    End Class
End Namespace", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

namespace Test.@class
{
    abstract class TestClass
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_SealedClass()
            TestConversionVisualBasicToCSharp("Namespace Test.[class]
    NotInheritable Class TestClass
    End Class
End Namespace", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

namespace Test.@class
{
    sealed class TestClass
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Interface()
            TestConversionVisualBasicToCSharp("Interface ITest
    Inherits System.IDisposable

    Sub Test()
End Interface", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

interface ITest : System.IDisposable
{
    void Test();
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Enum()
            TestConversionVisualBasicToCSharp("Friend Enum ExceptionResource
    Argument_ImplementIComparable
    ArgumentOutOfRange_NeedNonNegNum
    ArgumentOutOfRange_NeedNonNegNumRequired
    Arg_ArrayPlusOffTooSmall
End Enum", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

internal enum ExceptionResource
{
    Argument_ImplementIComparable,
    ArgumentOutOfRange_NeedNonNegNum,
    ArgumentOutOfRange_NeedNonNegNumRequired,
    Arg_ArrayPlusOffTooSmall
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ClassInheritanceList()
            TestConversionVisualBasicToCSharp("MustInherit Class ClassA
    Implements System.IDisposable

    Protected MustOverride Sub Test()
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

abstract class ClassA : System.IDisposable
{
    protected abstract void Test();
}")
            TestConversionVisualBasicToCSharp("MustInherit Class ClassA
    Inherits System.EventArgs
    Implements System.IDisposable

    Protected MustOverride Sub Test()
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

abstract class ClassA : System.EventArgs, System.IDisposable
{
    protected abstract void Test();
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Struct()
            TestConversionVisualBasicToCSharp("Structure MyType
    Implements System.IComparable(Of MyType)

    Private Sub Test()
    End Sub
End Structure", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

struct MyType : System.IComparable<MyType>
{
    private void Test()
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Delegate()
            Const usings As String = "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

"
            TestConversionVisualBasicToCSharp("Public Delegate Sub Test()", usings & "public delegate void Test();")
            TestConversionVisualBasicToCSharp("Public Delegate Function Test() As Integer", usings & "public delegate int Test();")
            TestConversionVisualBasicToCSharp("Public Delegate Sub Test(ByVal x As Integer)", usings & "public delegate void Test(int x);")
            TestConversionVisualBasicToCSharp("Public Delegate Sub Test(ByRef x As Integer)", usings & "public delegate void Test(ref int x);")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ClassImplementsInterface()
            TestConversionVisualBasicToCSharp("Class test
    Implements IComparable
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class test : IComparable
{
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ClassImplementsInterface2()
            TestConversionVisualBasicToCSharp("Class test
    Implements System.IComparable
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class test : System.IComparable
{
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ClassInheritsClass()
            TestConversionVisualBasicToCSharp("Imports System.IO

Class test
    Inherits InvalidDataException
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;
using System.IO;

class test : InvalidDataException
{
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ClassInheritsClass2()
            TestConversionVisualBasicToCSharp("Class test
    Inherits System.IO.InvalidDataException
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class test : System.IO.InvalidDataException
{
}")
        End Sub
    End Class
End Namespace
