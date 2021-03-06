﻿Imports Xunit

Namespace CodeConverter.Tests.CSharpToVB
    <TestClass()> Public Class StatementTests
        Inherits ConverterTestBase

        <Fact>
        Public Shared Sub CSharpToVB_EmptyStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (true) ; // This is a Comment
        while (true) ;
        for (;;) ;
        do ; while (true);
    }
}", "Class TestClass

    Private Sub TestMethod()
        If True Then ' This is a Comment
        End If

        While True
        End While

        While True
        End While

        Do
        Loop While True
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_AssignmentStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int b; // This is a Comment
        b = 0;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer ' This is a Comment
        b = 0
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_AssignmentStatementInDeclaration()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int b = 0;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer = 0
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_AssignmentStatementInVarDeclaration()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        var b = 0;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b = 0
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ObjectInitializationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        string b;
        b = new string(""test"");
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As String
        b = New String(""test"")
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ObjectInitializationStatementInDeclaration()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        string b = new string(""test"");
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As String = New String(""test"")
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ObjectInitializationStatementInVarDeclaration()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        var b = new string(""test"");
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b = New String(""test"")
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ArrayDeclarationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[] b;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer()
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ArrayInitializationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[] b = { 1, 2, 3 };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer() = {1, 2, 3}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ArrayInitializationStatementInVarDeclaration()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        var b = { 1, 2, 3 };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b = {1, 2, 3}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ArrayInitializationStatementWithType()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[] b = new int[] { 1, 2, 3 };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer() = New Integer() {1, 2, 3}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ArrayInitializationStatementWithLength()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[] b = new int[3] { 1, 2, 3 };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer() = New Integer(2) {1, 2, 3}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_MultidimensionalArrayDeclarationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[,] b;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer(,)
    End Sub
End Class")
        End Sub

        <Fact>
        Public Shared Sub CSharpToVB_MultidimensionalArrayInitializationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[,] b = { { 1, 2 }, { 3, 4 } };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer(,) = {{1, 2}, {3, 4}}
    End Sub
End Class")
        End Sub

        <Fact>
        Public Shared Sub CSharpToVB_MultidimensionalArrayInitializationStatementWithType()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[,] b = new int[,] { { 1, 2 }, { 3, 4 } };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer(,) = New Integer(,) {{1, 2}, {3, 4}}
    End Sub
End Class")
        End Sub

        <Fact>
        Public Shared Sub CSharpToVB_MultidimensionalArrayInitializationStatementWithLengths()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[,] b = new int[2, 2] { { 1, 2 }, { 3, 4 } };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer(,) = New Integer(1, 1) {{1, 2}, {3, 4}}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_JaggedArrayDeclarationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[][] b;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer()()
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_JaggedArrayInitializationStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[][] b = { new int[] { 1, 2 }, new int[] { 3, 4 } };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer()() = {New Integer() {1, 2}, New Integer() {3, 4}}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_JaggedArrayInitializationStatementWithType()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[][] b = new int[][] { new int[] { 1, 2 }, new int[] { 3, 4 } };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer()() = New Integer()() {New Integer() {1, 2}, New Integer() {3, 4}}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_JaggedArrayInitializationStatementWithLength()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int[][] b = new int[2][] { new int[] { 1, 2 }, new int[] { 3, 4 } };
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer()() = New Integer(1)() {New Integer() {1, 2}, New Integer() {3, 4}}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_DeclarationStatements()
            TestConversionCSharpToVisualBasic("class Test {
    void TestMethod()
    {
the_beginning:
        int value = 1;
        const double myPIe = System.Math.PI;
        var text = ""This is my text!"";
        goto the_beginning;
    }
}", "Class Test

    Private Sub TestMethod()
the_beginning:
        Dim value As Integer = 1
        Const myPIe As Double = System.Math.PI
        Dim text = ""This is my text!""
        GoTo the_beginning
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_IfStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod (int a)
    {
        int b;
        if (a == 0) {
            b = 0;
        } else if (a == 1) {
            b = 1;
        } else if (a == 2 || a == 3) {
            b = 2;
        } else {
            b = 3;
        }
    }
}", "Class TestClass

    Private Sub TestMethod(a As Integer)
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
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_WhileStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
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
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer
        b = 0
        While b = 0
            If b = 2 Then
                Continue While
            End If

            If b = 3 Then
                Exit While
            End If

            b = 1
        End While
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_DoWhileStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
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
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer
        b = 0
        Do
            If b = 2 Then
                Continue Do
            End If

            If b = 3 Then
                Exit Do
            End If

            b = 1
        Loop While b = 0
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ForEachStatementWithExplicitType()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod(int[] values)
    {
        foreach (int val in values)
        {
            if (val == 2)
                continue;
            if (val == 3)
                break;
        }
    }
}", "Class TestClass

    Private Sub TestMethod(values As Integer())
        For Each val As Integer In values
            If val = 2 Then
                Continue For
            End If

            If val = 3 Then
                Exit For
            End If
        Next
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ForEachStatementWithVar()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod(int[] values)
    {
        foreach (var val in values)
        {
            if (val == 2)
                continue;
            if (val == 3)
                break;
        }
    }
}", "Class TestClass

    Private Sub TestMethod(values As Integer())
        For Each val In values
            If val = 2 Then
                Continue For
            End If

            If val = 3 Then
                Exit For
            End If
        Next
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_SyncLockStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod(object nullObject)
    {
        if (nullObject == null)
            throw new ArgumentNullException(nameof(nullObject));
        lock (nullObject) {
            Console.WriteLine(nullObject);
        }
    }
}", "Class TestClass

    Private Sub TestMethod(nullObject As Object)
        If nullObject Is Nothing Then Throw New ArgumentNullException(NameOf(nullObject))
        SyncLock nullObject
            Console.WriteLine(nullObject)
        End SyncLock
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ForWithUnknownConditionAndSingleStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        for (i = 0; unknownCondition; i++)
            b[i] = s[i];
    }
}", "Class TestClass

    Private Sub TestMethod()
        i = 0
        While unknownCondition
            b(i) = s(i)
            i += 1
        End While
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ForWithUnknownConditionAndBlock()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        for (i = 0; unknownCondition; i++) {
            b[i] = s[i];
        }
    }
}", "Class TestClass

    Private Sub TestMethod()
        i = 0
        While unknownCondition
            b(i) = s(i)
            i += 1
        End While
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ForWithSingleStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        for (i = 0; i < end; i++) b[i] = s[i];
    }
}", "Class TestClass

    Private Sub TestMethod()
        For i = 0 To [end] - 1
            b(i) = s(i)
        Next
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ForWithBlock()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        for (i = 0; i < end; i++) {
            b[i] = s[i];
        }
    }
}", "Class TestClass

    Private Sub TestMethod()
        For i = 0 To [end] - 1
            b(i) = s(i)
        Next
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_LabeledAndForStatement()
            TestConversionCSharpToVisualBasic("class GotoTest1
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
}", "Class GotoTest1

    Private Shared Sub Main()
        Dim x As Integer = 200, y As Integer = 4
        Dim count As Integer = 0
        Dim array As String(,) = New String(x - 1, y - 1) {}

        For i As Integer = 0 To x - 1

            For j As Integer = 0 To y - 1
                array(i, j) = System.Threading.Interlocked.Increment(count).ToString()
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
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_ThrowStatement()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod(object nullObject)
    {
        if (nullObject == null)
            throw new ArgumentNullException(nameof(nullObject));
    }
}", "Class TestClass

    Private Sub TestMethod(nullObject As Object)
        If nullObject Is Nothing Then Throw New ArgumentNullException(NameOf(nullObject))
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_AddRemoveHandler()
            TestConversionCSharpToVisualBasic("using System;

class TestClass
{
    public event EventHandler MyEvent;

    void TestMethod(EventHandler e)
    {
        this.MyEvent += e;
        this.MyEvent += MyHandler;
    }

    void TestMethod2(EventHandler e)
    {
        this.MyEvent -= e;
        this.MyEvent -= MyHandler;
    }

    void MyHandler(object sender, EventArgs e)
    {

    }
}", "Imports System

Class TestClass

    Public Event MyEvent As EventHandler

    Private Sub TestMethod(e As EventHandler)
        AddHandler Me.MyEvent, e
        AddHandler Me.MyEvent, AddressOf MyHandler
    End Sub

    Private Sub TestMethod2(e As EventHandler)
        RemoveHandler Me.MyEvent, e
        RemoveHandler Me.MyEvent, AddressOf MyHandler
    End Sub

    Private Sub MyHandler(sender As Object, e As EventArgs)

    End Sub
End Class")
        End Sub
        'Class RemoteAnalyzerFileReferenceTest
        '    Inherits MarshalByRefObject

        '    Private _analyzerLoadException As Exception

        '    Public Overrides Function InitializeLifetimeService() As Object
        '        Return Nothing
        '    End Function

        '    Public Function LoadAnalyzer(analyzerPath As String) As Exception
        '        _analyzerLoadException = Nothing
        '        Dim analyzerRef = New AnalyzerFileReference(analyzerPath, FromFileLoader.Instance)
        '        AddHandler analyzerRef.AnalyzerLoadFailed, Function(s, e) _analyzerLoadException = e.Exception
        '        Dim builder = ImmutableArray.CreateBuilder(Of DiagnosticAnalyzer)()
        '        analyzerRef.AddAnalyzers(builder, LanguageNames.CSharp)
        '        Return _analyzerLoadException
        '    End Function
        'End Class
        <Fact>
        Public Shared Sub CSharpToVB_AddHandler()
            TestConversionCSharpToVisualBasic("using System;

    class RemoteAnalyzerFileReferenceTest : MarshalByRefObject
    {
        private Exception _analyzerLoadException;

        public override object InitializeLifetimeService()
        {
            return null;
        }

        public Exception LoadAnalyzer(string analyzerPath)
        {
            _analyzerLoadException = null;
            var analyzerRef = new AnalyzerFileReference(analyzerPath, FromFileLoader.Instance);
            analyzerRef.AnalyzerLoadFailed += (s, e) => _analyzerLoadException = e.Exception;
            var builder = ImmutableArray.CreateBuilder<DiagnosticAnalyzer>();
            analyzerRef.AddAnalyzers(builder, LanguageNames.CSharp);
            return _analyzerLoadException;
        }
    }
", "Imports System

Class RemoteAnalyzerFileReferenceTest
    Inherits MarshalByRefObject

    Private _analyzerLoadException As Exception

    Public Overrides Function InitializeLifetimeService() As Object
        Return Nothing
    End Function

    Public Function LoadAnalyzer(analyzerPath As String) As Exception
        _analyzerLoadException = Nothing
        Dim analyzerRef = New AnalyzerFileReference(analyzerPath, FromFileLoader.Instance)
        AddHandler analyzerRef.AnalyzerLoadFailed, Sub(s, e) _analyzerLoadException = e.Exception
        Dim builder = ImmutableArray.CreateBuilder(Of DiagnosticAnalyzer)()
        analyzerRef.AddAnalyzers(builder, LanguageNames.CSharp)
        Return _analyzerLoadException
    End Function
End Class
")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_SelectCase1()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod(int number)
    {
        switch (number) {
            case 0:
            case 1:
            case 2:
                Console.Write(""number is 0, 1, 2"");
                break;
            case 3:
                Console.Write(""section 3"");
                goto case 5;
            case 4:
                Console.Write(""section 4"");
                goto default;
            case 5:
                Console.Write(""section 5"");
                break;
            default:
                Console.Write(""default section"");
                break;
        }
    }
}", "Class TestClass

    Private Sub TestMethod(number As Integer)
        Select Case number
            Case 0, 1, 2
                Console.Write(""number is 0, 1, 2"")
            Case 3
                Console.Write(""section 3"")
                GoTo _Select0_Case5
            Case 4
                Console.Write(""section 4"")
                GoTo _Select0_CaseDefault
            Case 5
_Select0_Case5:
                Console.Write(""section 5"")
            Case Else
_Select0_CaseDefault:
                Console.Write(""default section"")
        End Select
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_TryCatch()
            TestConversionCSharpToVisualBasic("class TestClass
{
    static bool Log(string message)
    {
        Console.WriteLine(message);
        return false;
    }

    void TestMethod(int number)
    {
        try {
            Console.WriteLine(""try"");
        } catch (Exception e) {
            Console.WriteLine(""catch1"");
        } catch {
            Console.WriteLine(""catch all"");
        } finally {
            Console.WriteLine(""finally"");
        }
        try {
            Console.WriteLine(""try"");
        } catch (System.IO.IOException) {
            Console.WriteLine(""catch1"");
        } catch (Exception e) when (Log(e.Message)) {
            Console.WriteLine(""catch2"");
        }
        try {
            Console.WriteLine(""try"");
        } finally {
            Console.WriteLine(""finally"");
        }
    }
}", "Class TestClass

    Private Shared Function Log(message As String) As Boolean
        Console.WriteLine(message)
        Return False
    End Function

    Private Sub TestMethod(number As Integer)
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
        Catch __unusedIOException1__ As System.IO.IOException
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
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_Yield()
            TestConversionCSharpToVisualBasic("class TestClass
{
    IEnumerable<int> TestMethod(int number)
    {
        if (number < 0)
            yield break;
        for (int i = 0; i < number; i++)
            yield return i;
    }
}", "Class TestClass

    Private Iterator Function TestMethod(number As Integer) As IEnumerable(Of Integer)
        If number < 0 Then
            Return
        End If

        For i As Integer = 0 To number - 1
            Yield i
        Next
    End Function
End Class")
        End Sub
    End Class
End Namespace
