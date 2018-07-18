Imports System.IO
Imports CSharpToVBApp
Imports IVisualBasicCode.CodeConverter
Imports Microsoft.CodeAnalysis
Imports Xunit
Namespace CSharpToVBAppCSharpToVB
    <TestClass()> Public Class TestCompile
        Dim LastFileProcessed As String
        Dim MaxPathLength As Integer = 0
        Dim ListOfFiles As New List(Of String)
        '<Fact(Skip:="Not implemented!")>
        <Fact>
        Public Sub VB_CompileAll()
            Const targetDirectory As String = "C:\Users\PaulM\Source\Repos\roslyn-master\src\Compilers\Core\CodeAnalysisTest\InternalUtilities\"
            Dim LanguageExtension As String = "cs"
            Dim FilesProcessed As Integer = 0
            Dim LastFileNameWithPath As String = ""
            Assert.True(ProcessDirectory(targetDirectory, Nothing, Nothing, Nothing, LastFileNameWithPath, LanguageExtension, FilesProcessed, AddressOf ProcessFile), $"Failing file {LastFileProcessed}")
        End Sub

        Private Function ProcessFile(ByVal PathWithFileName As String, LanguageExtension As String) As Boolean
            ' Do not delete the next line or the parameter it is needed by other versions of this routine
            LanguageExtension = ""
            LastFileProcessed = PathWithFileName
            Dim RequestToConvert As ConvertRequest = New ConvertRequest With {.RequestedConversion = "cs2vb"}
            Dim fs As FileStream = File.OpenRead(PathWithFileName)
            RequestToConvert.SourceCode = GetFileTextFromStream(fs)
            Dim ResultOfConversion As ConversionResult = ConvertInputRequest(RequestToConvert)
            Return ResultOfConversion.Success
        End Function

    End Class
End Namespace
