Imports Xunit

Namespace CodeConverter.Tests.VBToCSharp
    <TestClass()>
    Public Class SpecialConversionTests
        Inherits ConverterTestBase
        <Fact>
        Public Sub VBToCSharp_RaiseEvent()
            TestConversionVisualBasicToCSharp("Class TestClass
    Private Event MyEvent As EventHandler

    Private Sub TestMethod()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class", "using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;

class TestClass
{
    private event EventHandler MyEvent;

    private void TestMethod()
    {
        MyEvent?.Invoke(this, EventArgs.Empty);
    }
}")
        End Sub
    End Class
End Namespace
