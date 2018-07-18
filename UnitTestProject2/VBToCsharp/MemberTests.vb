Imports Xunit

Namespace CodeConverter.Tests.VBToCSharp

    <TestClass()> Public Class MemberTests
        Inherits ConverterTestBase
        <Fact>
        Public Sub VBToCSharp_Field()
            TestConversionVisualBasicToCSharp("Class TestClass
    Const answer As Integer = 42
    Private value As Integer = 10
    ReadOnly v As Integer = 15
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    const int answer = 42;
    private int value = 10;
    private readonly int v = 15;
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Method()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public Sub TestMethod(Of T As {Class, New}, T2 As Structure, T3)(<Out> ByRef argument As T, ByRef argument2 As T2, ByVal argument3 As T3)
        argument = Nothing
        argument2 = Nothing
        argument3 = Nothing
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public void TestMethod<T, T2, T3>(out T argument, ref T2 argument2, T3 argument3)
        where T : class, new()
        where T2 : struct
    {
        argument = null;
        argument2 = default(T2);
        argument3 = default(T3);
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_MethodWithReturnType()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public Function TestMethod(Of T As {Class, New}, T2 As Structure, T3)(<Out> ByRef argument As T, ByRef argument2 As T2, ByVal argument3 As T3) As Integer
        Return 0
    End Function
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public int TestMethod<T, T2, T3>(out T argument, ref T2 argument2, T3 argument3)
        where T : class, new()
        where T2 : struct
    {
        return 0;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_StaticMethod()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public Shared Sub TestMethod(Of T As {Class, New}, T2 As Structure, T3)(<Out> ByRef argument As T, ByRef argument2 As T2, ByVal argument3 As T3)
        argument = Nothing
        argument2 = Nothing
        argument3 = Nothing
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public static void TestMethod<T, T2, T3>(out T argument, ref T2 argument2, T3 argument3)
        where T : class, new()
        where T2 : struct
    {
        argument = null;
        argument2 = default(T2);
        argument3 = default(T3);
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_AbstractMethod()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public MustOverride Sub TestMethod()
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public abstract void TestMethod();
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_SealedMethod()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public NotOverridable Sub TestMethod(Of T As {Class, New}, T2 As Structure, T3)(<Out> ByRef argument As T, ByRef argument2 As T2, ByVal argument3 As T3)
        argument = Nothing
        argument2 = Nothing
        argument3 = Nothing
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public sealed void TestMethod<T, T2, T3>(out T argument, ref T2 argument2, T3 argument3)
        where T : class, new()
        where T2 : struct
    {
        argument = null;
        argument2 = default(T2);
        argument3 = default(T3);
    }
}")
        End Sub
        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_ExtensionMethod()
            TestConversionVisualBasicToCSharp("Imports System.Runtime.CompilerServices

Module TestClass
    <Extension()>
    Sub TestMethod(ByVal str As String)
    End Sub
End Module", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

static class TestClass
{
    public static void TestMethod(this String str)
    {
    }
}")
        End Sub
        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_ExtensionMethodWithExistingImport()
            TestConversionVisualBasicToCSharp("Imports System.Runtime.CompilerServices

Module TestClass
    <Extension()>
    Sub TestMethod(ByVal str As String)
    End Sub
End Module", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;
using System.Runtime.CompilerServices;

static class TestClass
{
    public static void TestMethod(this String str)
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Property()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public Property Test As Integer

    Public Property Test2 As Integer
        Get
            Return 0
        End Get
    End Property

    Private m_test3 As Integer

    Public Property Test3 As Integer
        Get
            Return Me.m_test3
        End Get
        Set(ByVal value As Integer)
            Me.m_test3 = value
        End Set
    End Property
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public int Test { get; set; }

    public int Test2
    {
        get
        {
            return 0;
        }
    }

    private int m_test3;

    public int Test3
    {
        get
        {
            return this.m_test3;
        }

        set
        {
            this.m_test3 = value;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Constructor()
            TestConversionVisualBasicToCSharp("Class TestClass(Of T As {Class, New}, T2 As Structure, T3)
    Public Sub New(<Out> ByRef argument As T, ByRef argument2 As T2, ByVal argument3 As T3)
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass<T, T2, T3>
    where T : class, new()
    where T2 : struct
{
    public TestClass(out T argument, ref T2 argument2, T3 argument3)
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Destructor()
            TestConversionVisualBasicToCSharp("Class TestClass
    Protected Overrides Sub Finalize()
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    ~TestClass()
    {
    }
}")
        End Sub
        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_Event()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public Event MyEvent As EventHandler
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public event EventHandler MyEvent;
}")
        End Sub
        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_CustomEvent()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private backingField As EventHandler

    Public Event MyEvent As EventHandler
        AddHandler(ByVal value As EventHandler)
            AddHandler Me.backingField, value
        End AddHandler
        RemoveHandler(ByVal value As EventHandler)
            RemoveHandler Me.backingField, value
        End RemoveHandler
    End Event
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    EventHandler backingField;

    public event EventHandler MyEvent {
        add {
            this.backingField += value;
        }
        remove {
            this.backingField -= value;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_SynthesizedBackingFieldAccess()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Shared Property First As Integer

    Private Second As Integer = _First
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private static int First { get; set; }

    private int Second = First;
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_PropertyInitializers()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private ReadOnly Property First As New List(Of String)
    Private Property Second As Integer = 0
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private List<string> First { get; } = new List<string>();
    private int Second { get; set; } = 0;
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_PartialFriendClassWithOverloads()
            TestConversionVisualBasicToCSharp("
Partial Friend MustInherit Class TestClass1
    Public Shared Sub CreateStatic()
    End Sub

    Public Sub CreateInstance()
    End Sub

    Public MustOverride Sub CreateAbstractInstance()

    Public Overridable Sub CreateVirtualInstance()
    End Sub
End Class

Friend Class TestClass2
    Inherits TestClass1
    Public Overloads Shared Sub CreateStatic()
    End Sub

    Public Overloads Sub CreateInstance()
    End Sub

    Public Overrides Sub CreateAbstractInstance()
    End Sub

    Public Overrides Sub CreateVirtualInstance()
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

internal abstract partial class TestClass1
{
    public static void CreateStatic()
    {
    }

    public void CreateInstance()
    {
    }

    public abstract void CreateAbstractInstance();

    public virtual void CreateVirtualInstance()
    {
    }
}

internal class TestClass2 : TestClass1
{
    public new static void CreateStatic()
    {
    }

    public new void CreateInstance()
    {
    }

    public override void CreateAbstractInstance()
    {
    }

    public override void CreateVirtualInstance()
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ClassWithGloballyQualifiedAttribute()
            TestConversionVisualBasicToCSharp("<Global.System.Diagnostics.DebuggerDisplay(""Hello World"")>
Class TestClass
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

[global::System.Diagnostics.DebuggerDisplay(""Hello World"")]
class TestClass
{
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_FieldWithAttribute()
            TestConversionVisualBasicToCSharp("Class TestClass
    <ThreadStatic>
    Private Shared First As Integer
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    [ThreadStatic]
    private static int First;
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ParamArray()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub SomeBools(ParamArray anyName As Boolean())
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void SomeBools(params bool[] anyName)
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ParamNamedBool()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub SomeBools(ParamArray bool As Boolean())
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void SomeBools(params bool[] @bool)
    {
    }
}")
        End Sub
        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_TestIndexer()
            TestConversionVisualBasicToCSharp("Class TestClass
    Default Public Property Item(ByVal index As Integer) As Integer

    Default Public Property Item(ByVal index As String) As Integer
        Get
            Return 0
        End Get
    End Property

    Private m_test3 As Integer

    Default Public Property Item(ByVal index As Double) As Integer
        Get
            Return Me.m_test3
        End Get
        Set(ByVal value As Integer)
            Me.m_test3 = value
        End Set
    End Property
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public int this[int index] { get; set; }
    public int this[string index] {
        get { return 0; }
    }
    int m_test3;
    public int this[double index] {
        get { return this.m_test3; }
        set { this.m_test3 = value; }
    }
}")
        End Sub
    End Class
End Namespace
