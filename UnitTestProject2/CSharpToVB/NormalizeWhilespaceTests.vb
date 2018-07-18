Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Xunit
Imports IVisualBasicCode.CodeConverter.Util
Namespace CodeConverter.Tests.CSharpToVB

    <TestClass()> Public Class NormalizeWhilespaceTests
        Inherits ConverterTestBase

        <Fact>
        Public Shared Sub _NormalizeWhitespace()
            TestConversionCSharpToVisualBasic("public class CustomDebugInfoTests
{
    [Fact]
    public void TryGetCustomDebugInfoRecord1()
    {
        // incomplete record header
        var cdi = new byte[]
        {
            4, 1, 0, 0, // global header
            9, (byte)CustomDebugInfoKind.EditAndContinueLocalSlotMap,
        };
    }
}", "Public Class CustomDebugInfoTests

    <Fact>
    Public Sub TryGetCustomDebugInfoRecord1()
        ' incomplete record header
        Dim cdi = New Byte() {
            4, 1, 0, 0, ' global header
            9, CByte(CustomDebugInfoKind.EditAndContinueLocalSlotMap)}
    End Sub
End Class")
        End Sub
        <Fact>
        Public Shared Sub _OpenParenWithCRLF()
            Dim Separators As New List(Of SyntaxToken)
            Dim NodeList As New List(Of ArgumentSyntax) From {
                SyntaxFactory.SimpleArgument(SyntaxFactory.ParseExpression("A"))
            }
            Dim ArgumentList As ArgumentListSyntax =
                SyntaxFactory.ArgumentList(SyntaxFactory.Token(SyntaxKind.OpenParenToken).WithTrailingTrivia(SyntaxFactory.EndOfLineTrivia(vbCrLf)),
                                           SyntaxFactory.SeparatedList(NodeList, Separators),
                                           SyntaxFactory.Token(SyntaxKind.CloseParenToken).WithTrailingTrivia(SyntaxFactory.EndOfLineTrivia(vbCrLf))
                                            )
            Assert.Equal($"({vbCrLf}A){vbCrLf}", ArgumentList.ToFullString)
        End Sub
        <Fact>
        Public Shared Sub _BraceWithCRLF()
            Dim items As New List(Of ExpressionSyntax) From {
                SyntaxFactory.ParseExpression("A"),
                SyntaxFactory.ParseExpression("B")
            }
            Dim Separators As New List(Of SyntaxToken) From {
                SyntaxFactory.Token(SyntaxKind.CommaToken)
            }
            Dim initializers As SeparatedSyntaxList(Of ExpressionSyntax) = SyntaxFactory.SeparatedList(items, Separators)
            Dim OpenBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken).WithTrailingTrivia(SyntaxFactory.EndOfLineTrivia(vbCrLf))
            Dim CloseBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseBraceToken).WithTrailingTrivia(SyntaxFactory.EndOfLineTrivia(vbCrLf))
            Dim CollectionInitializer As CollectionInitializerSyntax = SyntaxFactory.CollectionInitializer(OpenBraceToken, initializers, CloseBraceToken)
            Assert.Equal($"{{{vbCrLf}A,B}}{vbCrLf}", CollectionInitializer.ToFullString)
            Assert.Equal($"{{A, B}}", CollectionInitializer.NormalizeWhitespace.ToFullString)
            Assert.Equal($"{{{vbCrLf}    A, B}}{vbCrLf}", CollectionInitializer.NormalizeWhitespaceEx(True, "    ", False, True).ToFullString)
        End Sub
    End Class
End Namespace
