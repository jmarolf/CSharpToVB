Option Explicit On
Option Infer Off
Option Strict On
Imports System.IO
Imports Microsoft.CodeAnalysis

Public Module ProcessDirectoriesModule
    Public Function GetFileTextFromStream(fileStream As Stream) As String
        Dim sw As New IO.StreamReader(fileStream)
        Dim SourceText As String = sw.ReadToEnd()
        sw.Close()
        Return SourceText
    End Function

    Public Sub LocalUseWaitCutsor(MeForm As Form1, Enable As Boolean)
        If MeForm Is Nothing Then
            Exit Sub
        End If
        If MeForm.UseWaitCursor <> Enable Then
            MeForm.UseWaitCursor = Enable
            Application.DoEvents()
        End If
    End Sub

    ''' <summary>
    ''' Needed for Unit Tests
    ''' </summary>
    ''' <param name="targetDirectory">Start location of where to process directories</param>
    ''' <param name="StopButton">Pass Nothing for Unit Tests</param>
    ''' <param name="LastFileNameWithPath">Pass Last File Name to Start Conversion where you left off</param>
    ''' <param name="LanguageExtension">vb or cs</param>
    ''' <param name="FilesProcessed">Count of the number of tiles processed</param>
    ''' <returns>
    ''' False if error and user wants to stop, True if success or user wants to ignore error
    ''' </returns>
    Public Function ProcessDirectory(targetDirectory As String, MeForm As Form1, StopButton As Button, RichTextBoxFileList As RichTextBox, ByRef LastFileNameWithPath As String, ByRef LanguageExtension As String, ByRef FilesProcessed As Integer, ProcessFile As Func(Of String, String, MetadataReference(), Boolean)) As Boolean
        ' Process the list of files found in the directory.
        Dim DirectoryList As String() = Directory.GetFiles(path:=targetDirectory, searchPattern:=$"*.{LanguageExtension}")
        For Each PathWithFileName As String In DirectoryList
            FilesProcessed += 1
            If LastFileNameWithPath.Length = 0 OrElse LastFileNameWithPath = PathWithFileName Then
                LastFileNameWithPath = ""
                If RichTextBoxFileList IsNot Nothing Then
                    RichTextBoxFileList.AppendText($"{FilesProcessed.ToString.PadLeft(5)} {PathWithFileName}{vbCrLf}")
                    RichTextBoxFileList.Select(RichTextBoxFileList.TextLength, 0)
                    RichTextBoxFileList.ScrollToCaret()
                    Application.DoEvents()
                End If
                If Not ProcessFile(PathWithFileName, LanguageExtension, References) Then
                    SetButtonStopAndCursor(MeForm:=MeForm, StopButton:=StopButton, VisibleValue:=False)
                    Return False
                End If
            End If
        Next PathWithFileName
        Dim subdirectoryEntries As String() = Directory.GetDirectories(path:=targetDirectory)
        ' Recurse into subdirectories of this directory.
        For Each subdirectory As String In subdirectoryEntries
            If (subdirectory.EndsWith("Test\Resources") OrElse subdirectory.EndsWith("Setup\Templates")) AndAlso (MeForm Is Nothing OrElse MeForm.SkipTestResourceFilesToolStripMenuItem.Checked) Then
                Continue For
            End If
            If Not ProcessDirectory(targetDirectory:=subdirectory, MeForm:=MeForm, StopButton:=StopButton, RichTextBoxFileList:=RichTextBoxFileList, LastFileNameWithPath:=LastFileNameWithPath, LanguageExtension:=LanguageExtension, FilesProcessed:=FilesProcessed, ProcessFile:=ProcessFile) Then
                SetButtonStopAndCursor(MeForm:=MeForm, StopButton:=StopButton, VisibleValue:=False)
                Return False
            End If
        Next subdirectory
        Return True
    End Function

    Public Sub SetButtonStopAndCursor(MeForm As Form1, StopButton As Button, VisibleValue As Boolean)
        If StopButton IsNot Nothing Then
            StopButton.Visible = VisibleValue
            MeForm.StopRequested = False
        End If
        LocalUseWaitCutsor(MeForm:=MeForm, Enable:=VisibleValue)
    End Sub
End Module
