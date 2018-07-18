Imports Xunit

Namespace CodeConverter.Tests.VBToCSharp

    <TestClass()>
    Public Class StatementTests
        Inherits ConverterTestBase
        <Fact>
        Public Sub VBToCSharp_EmptyStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        If True Then
        End If

        While True
        End While

        Do
        Loop While True
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        if (true)
        {
        }

        while (true)
        {
        }

        do
        {
        }
        while (true);
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_AssignmentStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer
        b = 0
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int b;
        b = 0;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_AssignmentStatementInDeclaration()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer = 0
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int b = 0;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_AssignmentStatementInVarDeclaration()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b = 0
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        var b = 0;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ObjectInitializationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As String
        b = New String(""test"")
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        string b;
        b = new string(""test"");
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ObjectInitializationStatementInDeclaration()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As String = New String(""test"")
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        string b = new string(""test"");
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ObjectInitializationStatementInVarDeclaration()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b = New String(""test"")
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        var b = new string(""test"");
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ArrayDeclarationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer()
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[] b;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_EndStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        End
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        System.Environment.Exit(0);
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_StopStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Stop
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        System.Diagnostics.Debugger.Break();
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_WithBlock()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        With New System.Text.StringBuilder
            .Capacity = 20
            .Length = 0
        End With
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        {
            var withBlock = new System.Text.StringBuilder();
            withBlock.Capacity = 20;
            withBlock.Length = 0;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_NestedWithBlock()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        With New System.Text.StringBuilder
            Dim withBlock as Integer = 3
            With New System.Text.StringBuilder
                Dim withBlock1 as Integer = 4
                .Capacity = withBlock1
            End With

            .Length = withBlock
        End With
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        {
            var withBlock2 = new System.Text.StringBuilder();
            int withBlock = 3;
            {
                var withBlock3 = new System.Text.StringBuilder();
                int withBlock1 = 4;
                withBlock3.Capacity = withBlock1;
            }

            withBlock2.Length = withBlock;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ArrayInitializationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer() = {1, 2, 3}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[] b = new[] { 1, 2, 3 };
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ArrayInitializationStatementInVarDeclaration()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b = {1, 2, 3}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        var b = new[] { 1, 2, 3 };
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ArrayInitializationStatementWithType()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer() = New Integer() {1, 2, 3}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[] b = new int[] { 1, 2, 3 };
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_ArrayInitializationStatementWithLength()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer() = New Integer(2) {1, 2, 3}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[] b = new int[3] { 1, 2, 3 };
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_MultidimensionalArrayDeclarationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer(,)
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[,] b;
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_MultidimensionalArrayInitializationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer(,) = {{1, 2}, {3, 4}}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[,] b = { { 1, 2 }, { 3, 4 } };
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_MultidimensionalArrayInitializationStatementWithType()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer(,) = New Integer(,) {{1, 2}, {3, 4}}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[,] b = new int[,] { { 1, 2 }, { 3, 4 } };
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_MultidimensionalArrayInitializationStatementWithLengths()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer(,) = New Integer(1, 1) {{1, 2}, {3, 4}}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[,] b = new int[2, 2] { { 1, 2 }, { 3, 4 } };
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_JaggedArrayDeclarationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer()()
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[][] b;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_JaggedArrayInitializationStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer()() = {New Integer() {1, 2}, New Integer() {3, 4}}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[][] b = new[] { new int[] { 1, 2 }, new int[] { 3, 4 } };
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_JaggedArrayInitializationStatementWithType()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer()() = New Integer()() {New Integer() {1, 2}, New Integer() {3, 4}}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[][] b = new int[][] { new int[] { 1, 2 }, new int[] { 3, 4 } };
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_JaggedArrayInitializationStatementWithLength()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer()() = New Integer(1)() {New Integer() {1, 2}, New Integer() {3, 4}}
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[][] b = new int[2][] { new int[] { 1, 2 }, new int[] { 3, 4 } };
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_DeclarationStatements()
            TestConversionVisualBasicToCSharp("Class Test
    Private Sub TestMethod()
the_beginning:
        Dim value As Integer = 1
        Const myPIe As Double = System.Math.PI
        Dim text = ""This is my text!""
        GoTo the_beginning
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class Test {
    private void TestMethod()
    {
the_beginning:
        int value = 1;
        const double myPIe = System.Math.PI;
        var text = ""This is my text!"";
        goto the_beginning;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_IfStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod(ByVal a As Integer)
        Dim b As Integer

        If a = 0 Then
            b = 0
        ElseIf a = 1 Then
            b = 1
        ElseIf a = 2 OrElse a = 3 Then
            b = 2
        Else
            b = 3
        End If
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod(int a)
    {
        int b;
        if (a == 0)
            b = 0;
        else if (a == 1)
            b = 1;
        else if (a == 2 || a == 3)
            b = 2;
        else
            b = 3;
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_WhileStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer
        b = 0

        While b = 0
            If b = 2 Then Continue While
            If b = 3 Then Exit While
            b = 1
        End While
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int b;
        b = 0;
        while (b == 0)
        {
            if (b == 2)
                continue;
            if (b == 3)
                break;
            b = 1;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_DoWhileStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b As Integer
        b = 0

        Do
            If b = 2 Then Continue Do
            If b = 3 Then Exit Do
            b = 1
        Loop While b = 0
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int b;
        b = 0;
        do
        {
            if (b == 2)
                continue;
            if (b == 3)
                break;
            b = 1;
        }
        while (b == 0);
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ForEachStatementWithExplicitType()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod(ByVal values As Integer())
        For Each val As Integer In values
            If val = 2 Then Continue For
            If val = 3 Then Exit For
        Next
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod(int[] values)
    {
        foreach (int val in values)
        {
            if (val == 2)
                continue;
            if (val == 3)
                break;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ForEachStatementWithVar()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod(ByVal values As Integer())
        For Each val In values
            If val = 2 Then Continue For
            If val = 3 Then Exit For
        Next
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod(int[] values)
    {
        foreach (var val in values)
        {
            if (val == 2)
                continue;
            if (val == 3)
                break;
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_SyncLockStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod(ByVal nullObject As Object)
        If nullObject Is Nothing Then Throw New ArgumentNullException(NameOf(nullObject))

        SyncLock nullObject
            Console.WriteLine(nullObject)
        End SyncLock
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod(object nullObject)
    {
        if (nullObject == null)
            throw new ArgumentNullException(nameof(nullObject));
        lock (nullObject)
            Console.WriteLine(nullObject);
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ForWithSingleStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b, s As Integer()
        For i = 0 To [end]
            b(i) = s(i)
        Next
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[] b, s;
        for (var i = 0; i <= end; i++)
            b[i] = s[i];
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ForWithBlock()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod()
        Dim b, s As Integer()
        For i = 0 To [end] - 1
            b(i) = s(i)
        Next
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod()
    {
        int[] b, s;
        for (var i = 0; i <= end - 1; i++)
            b[i] = s[i];
    }
}")
        End Sub

        <Fact(Skip:="Not implemented!")>
        Public Sub VBToCSharp_LabeledAndForStatement()
            TestConversionVisualBasicToCSharp("Class GotoTest1
    Private Shared Sub Main()
        Dim x As Integer = 200, y As Integer = 4
        Dim count As Integer = 0
        Dim array As String(,) = New String(x - 1, y - 1) {}

        For i As Integer = 0 To x - 1

            For j As Integer = 0 To y - 1
                array(i, j) = (System.Threading.Interlocked.Increment(count)).ToString()
            Next
        Next

        Console.Write(""Enter the number to search for: "")
        Dim myNumber As String = Console.ReadLine()

        For i As Integer = 0 To x - 1

            For j As Integer = 0 To y - 1

                If array(i, j).Equals(myNumber) Then
                    GoTo Found
                End If
            Next
        Next

        Console.WriteLine(""The number {0} was not found."", myNumber)
        GoTo Finish
Found:
        Console.WriteLine(""The number {0} is found."", myNumber)
Finish:
        Console.WriteLine(""End of search."")
        Console.WriteLine(""Press any key to exit."")
        Console.ReadKey()
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class GotoTest1
{
    static void Main()
    {
        int x = 200, y = 4;
        int count = 0;
        string[,] array = new string[x, y];

        for (int i = 0; i < x; i++)

            for (int j = 0; j < y; j++)
                array[i, j] = (++count).ToString();

        Console.Write(""Enter the number to search for: "");

        string myNumber = Console.ReadLine();

                for (int i = 0; i < x; i++)
                {
                    for (int j = 0; j < y; j++)
                    {
                        if (array[i, j].Equals(myNumber))
                        {
                            goto Found;
                        }
                    }
                }

            Console.WriteLine(""The number {0} was not found."", myNumber);
            goto Finish;

        Found:
            Console.WriteLine(""The number {0} is found."", myNumber);

        Finish:
            Console.WriteLine(""End of search."");


            Console.WriteLine(""Press any key to exit."");
            Console.ReadKey();
        }
    }")
        End Sub
        <Fact>
        Public Sub VBToCSharp_ThrowStatement()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod(ByVal nullObject As Object)
        If nullObject Is Nothing Then Throw New ArgumentNullException(NameOf(nullObject))
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod(object nullObject)
    {
        if (nullObject == null)
            throw new ArgumentNullException(nameof(nullObject));
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_AddRemoveHandler()
            TestConversionVisualBasicToCSharp("Class TestClass
    Public Event MyEvent As EventHandler

    Private Sub TestMethod(ByVal e As EventHandler)
        AddHandler Me.MyEvent, e
        AddHandler Me.MyEvent, AddressOf MyHandler
    End Sub

    Private Sub TestMethod2(ByVal e As EventHandler)
        RemoveHandler Me.MyEvent, e
        RemoveHandler Me.MyEvent, AddressOf MyHandler
    End Sub

    Private Sub MyHandler(ByVal sender As Object, ByVal e As EventArgs)
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    public event EventHandler MyEvent;

    private void TestMethod(EventHandler e)
    {
        this.MyEvent += e;
        this.MyEvent += MyHandler;
    }

    private void TestMethod2(EventHandler e)
    {
        this.MyEvent -= e;
        this.MyEvent -= MyHandler;
    }

    private void MyHandler(object sender, EventArgs e)
    {
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_SelectCase1()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Sub TestMethod(ByVal number As Integer)
        Select Case number
            Case 0, 1, 2
                Console.Write(""number is 0, 1, 2"")
            Case 5
                Console.Write(""section 5"")
            Case Else
                Console.Write(""default section"")
        End Select
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private void TestMethod(int number)
    {
        switch (number)
        {
            case 0:
            case 1:
            case 2:
                {
                    Console.Write(""number is 0, 1, 2"");
                    break;
                }

            case 5:
                {
                    Console.Write(""section 5"");
                    break;
                }

            default:
                {
                    Console.Write(""default section"");
                    break;
                }
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_TryCatch()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Shared Function Log(ByVal message As String) As Boolean
        Console.WriteLine(message)
        Return False
    End Function

    Private Sub TestMethod(ByVal number As Integer)
        Try
            Console.WriteLine(""try"")
        Catch e As Exception
            Console.WriteLine(""catch1"")
        Catch
            Console.WriteLine(""catch all"")
        Finally
            Console.WriteLine(""finally"")
        End Try

        Try
            Console.WriteLine(""try"")
        Catch e2 As NotImplementedException
            Console.WriteLine(""catch1"")
        Catch e As Exception When Log(e.Message)
            Console.WriteLine(""catch2"")
        End Try

        Try
            Console.WriteLine(""try"")
        Finally
            Console.WriteLine(""finally"")
        End Try
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private static bool Log(string message)
    {
        Console.WriteLine(message);
        return false;
    }

    private void TestMethod(int number)
    {
        try
        {
            Console.WriteLine(""try"");
        }
        catch (Exception e)
        {
            Console.WriteLine(""catch1"");
        }
        catch
        {
            Console.WriteLine(""catch all"");
        }
        finally
        {
            Console.WriteLine(""finally"");
        }

        try
        {
            Console.WriteLine(""try"");
        }
        catch (NotImplementedException e2)
        {
            Console.WriteLine(""catch1"");
        }
        catch (Exception e) when (Log(e.Message))
        {
            Console.WriteLine(""catch2"");
        }

        try
        {
            Console.WriteLine(""try"");
        }
        finally
        {
            Console.WriteLine(""finally"");
        }
    }
}")
        End Sub
        <Fact>
        Public Sub VBToCSharp_Yield()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Iterator Function TestMethod(ByVal number As Integer) As IEnumerable(Of Integer)
        If number < 0 Then Return

        For i As Integer = 0 To number - 1
            Yield i
        Next
    End Function
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private IEnumerable<int> TestMethod(int number)
    {
        if (number < 0)
            yield break;
        for (int i = 0; i <= number - 1; i++)
            yield return i;
    }
}")
        End Sub
    End Class
End Namespace
