Imports CSharpToVBApp
Imports Microsoft.CodeAnalysis
Imports Xunit

Namespace CodeConverter.Tests.VB

    <TestClass()> Public Class LoadSaveDictionaryTests
        Dim LastFileProcessed As String
        Dim ListOfFiles As New List(Of String)
        Dim MaxPathLength As Integer = 0
        Private Function GetMaxPathLength(PathWithFileName As String, LanguageExtension As String, DontCare() As MetadataReference) As Boolean
            ' Do not delete the next line of the parameter
            LanguageExtension = ""
            LastFileProcessed = PathWithFileName
            ListOfFiles.Add(LastFileProcessed)
            If PathWithFileName.Length > MaxPathLength Then
                MaxPathLength = PathWithFileName.Length
            End If
            Return True
        End Function

        <Fact>
        Public Shared Sub VB_DictionaryWriteTest()
            Dim filePath As String = IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "ColorDictionary.csv")
            ColorSelector.WriteColorDictionaryToFile(filePath)
        End Sub

        <Fact>
        Public Sub VB_MaxPathLength()
            Const targetDirectory As String = "C:\Users\PaulM\Source\Repos\roslyn-PaulCohen\src"
            Dim LanguageExtension As String = "cs"
            Dim FilesProcessed As Integer = 0
            Dim LastFileNameWithPath As String = ""
            Dim Condition As Boolean = ProcessDirectory(targetDirectory:=targetDirectory, MeForm:=Nothing, StopButton:=Nothing, RichTextBoxFileList:=Nothing, LastFileNameWithPath:=LastFileNameWithPath, LanguageExtension:=LanguageExtension, FilesProcessed:=FilesProcessed, ProcessFile:=AddressOf GetMaxPathLength)
            Assert.True(MaxPathLength = 220, $"MaxPathLength = {MaxPathLength}")
        End Sub

        <Fact>
        Public Shared Sub VB_TestRenoveNewLine()
            Dim OriginalString As String = "This is a 2 Line
String"
            Dim ResultlString As String = "This is a 2 LineString"
            Assert.True(OriginalString.WithoutNewLines = ResultlString)
        End Sub
    End Class

End Namespace