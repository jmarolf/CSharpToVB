Imports CSharpToVBApp
Imports Xunit

Namespace CodeConverter.Tests.VB
    <TestClass()> Public Class LoadSaveDictionaryTests
        Dim LastFileProcessed As String
        Dim MaxPathLength As Integer = 0
        Dim ListOfFiles As New List(Of String)
        <Fact>
        Public Shared Sub VB_DictionaryWriteTest()
            Dim filePath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "ColorDictionary.csv")
            ColorSelector.WriteColorDictionaryToFile(filePath)
        End Sub
        <Fact>
        Public Sub VB_MaxPathLength()
            Const targetDirectory As String = "C:\Users\PaulM\Source\Repos\roslyn-master\src"
            Dim LanguageExtension As String = "cs"
            Dim FilesProcessed As Integer = 0
            Dim LastFileNameWithPath As String = ""
            Dim Condition As Boolean = ProcessDirectory(targetDirectory, Nothing, Nothing, Nothing, LastFileNameWithPath, LanguageExtension, FilesProcessed, AddressOf GetMaxPathLength)
            Assert.True(MaxPathLength = 192, $"MaxPathLength = {MaxPathLength}")
        End Sub
        Private Function GetMaxPathLength(ByVal PathWithFileName As String, LanguageExtension As String) As Boolean
            ' Do not delete the next line of the parameter
            LanguageExtension = ""
            LastFileProcessed = PathWithFileName
            ListOfFiles.Add(LastFileProcessed)
            If PathWithFileName.Length > MaxPathLength Then
                MaxPathLength = PathWithFileName.Length
            End If
            Return True
        End Function
    End Class
End Namespace
