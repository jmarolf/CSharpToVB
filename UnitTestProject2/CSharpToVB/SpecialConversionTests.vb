Imports Xunit

Namespace CodeConverter.Tests.CSharpToVB
    <TestClass()> Public Class SpecialConversionTests
        Inherits ConverterTestBase

        <Fact>
        Public Shared Sub CSharpToVB_SimpleInlineAssign()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int a, b;
        b = a = 5;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim a, b As Integer
        b = __InlineAssignHelper(a, 5)
    End Sub

    <Obsolete(""Please refactor code that uses this function, it is a simple work-around to simulate inline assignment in VB!"")>
    Private Shared Function __InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_SimplePostIncrementAssign()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        int a = 5, b;
        b = a++;
    }
}", "Class TestClass

    Private Sub TestMethod()
        Dim b As Integer, a As Integer = 5
        b = Math.Min(System.Threading.Interlocked.Increment(a), a - 1)
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_RaiseEvent()
            TestConversionCSharpToVisualBasic("class TestClass
{
    event EventHandler MyEvent;

    void TestMethod()
    {
        if (MyEvent != null) MyEvent(this, EventArgs.Empty);
    }
}", "Class TestClass

    Private Event MyEvent As EventHandler

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if ((MyEvent != null)) MyEvent(this, EventArgs.Empty);
    }
}", "Class TestClass

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (null != MyEvent) { MyEvent(this, EventArgs.Empty); }
    }
}", "Class TestClass

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (this.MyEvent != null) MyEvent(this, EventArgs.Empty);
    }
}", "Class TestClass

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (MyEvent != null) this.MyEvent(this, EventArgs.Empty);
    }
}", "Class TestClass

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if ((this.MyEvent != null)) { this.MyEvent(this, EventArgs.Empty); }
    }
}", "Class TestClass

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub CSharpToVB_IfStatementSimilarToRaiseEvent()
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (FullImage != null) DrawImage();
    }
}", "Class TestClass

    Private Sub TestMethod()
        If FullImage IsNot Nothing Then DrawImage()
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (FullImage != null) e.DrawImage();
    }
}", "Class TestClass

    Private Sub TestMethod()
        If FullImage IsNot Nothing Then e.DrawImage()
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (FullImage != null) { DrawImage(); }
    }
}", "Class TestClass

    Private Sub TestMethod()
        If FullImage IsNot Nothing Then
            DrawImage()
        End If
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (FullImage != null) { e.DrawImage(); }
    }
}", "Class TestClass

    Private Sub TestMethod()
        If FullImage IsNot Nothing Then
            e.DrawImage()
        End If
    End Sub
End Class")
            TestConversionCSharpToVisualBasic("class TestClass
{
    void TestMethod()
    {
        if (Tiles != null) foreach (Tile t in Tiles) this.TileTray.Controls.Remove(t);
    }
}", "Class TestClass

    Private Sub TestMethod()
        If Tiles IsNot Nothing Then
            For Each t As Tile In Tiles
                Me.TileTray.Controls.Remove(t)
            Next
        End If
    End Sub
End Class")
        End Sub
    End Class
End Namespace
