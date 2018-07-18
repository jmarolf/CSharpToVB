Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable
Imports System.ComponentModel
Imports System.IO
Imports CSharpToVBApp
Imports IVisualBasicCode.CodeConverter
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic

Public Class Form1
    Private Const SPI_SETSCREENSAVETIMEOUT As Integer = 15

    Private Shared SnipitFileWithPath As String = Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "CSharpToVBLastSnipit.RTF")
    Private CurrentBuffer As RichTextBox = Nothing
    Private InputBufferCurrentIndex As Integer = 0

    Private InputBufferCurrentLength As Integer = 0

    Private OutputBufferCurrentIndex As Integer = 0

    Private OutputBufferCurrentLength As Integer = 0

    Private RequestToConvert As ConvertRequest

    Private ResultOfConversion As ConversionResult

    Private RTFLineStart As Integer
    Property StopRequested As Boolean = False

    Private Shared Sub RichTexBoxConversionOutput_HorizScrollBarRightClicked(ByVal sender As Object, ByVal loc As Point) Handles RichTextBoxConversionOutput.HorizScrollBarRightClicked
        Dim p As Integer = CType(sender, AdvancedRTB).HScrollPos
        MessageBox.Show($"Horizontal Right Clicked At: {loc.ToString}, HScrollPos = {p}")
    End Sub

    Private Shared Sub RichTexBoxConversionOutput_VertScrollBarRightClicked(ByVal sender As Object, ByVal loc As Point) Handles RichTextBoxConversionOutput.VertScrollBarRightClicked
        Dim p As Integer = CType(sender, AdvancedRTB).VScrollPos
        MessageBox.Show($"Vertical Right Clicked At: {loc.ToString}, VScrollPos = {p} ")
    End Sub

    Private Sub AutoCompileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoCompileToolStripMenuItem.Click
        My.Settings.AutoCompile = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub ButtonSearch_Click(sender As Object, e As EventArgs) Handles ButtonSearch.Click
        SearchBoxVisibility(False)
    End Sub

    Private Sub ButtonStop_Click(sender As Object, e As EventArgs) Handles ButtonStop.Click
        StopRequested = True
        ButtonStop.Visible = False
        ConvertSnipitToolStripMenuItem.Enabled = False
        Application.DoEvents()
    End Sub

    Private Sub ButtonStop_MouseEnter(sender As Object, e As EventArgs) Handles ButtonStop.MouseEnter
        LocalUseWaitCutsor(Me, False)
    End Sub

    Private Sub ButtonStop_MouseLeave(sender As Object, e As EventArgs) Handles ButtonStop.MouseLeave
        LocalUseWaitCutsor(Me, True And Not StopRequested)
    End Sub

    Private Sub Colorize(FragmentRange As IEnumerable(Of Range), ConversionBuffer As RichTextBox, Lines As Integer, Optional failures As IEnumerable(Of Diagnostic) = Nothing)
        Try ' Prevent crash when exiting
            If failures IsNot Nothing Then
                For Each dia As Diagnostic In failures
                    RichTextBoxErrorList.AppendText($"{dia.Id} Line = {dia.Location.GetLineSpan.StartLinePosition.Line + 1} {dia.GetMessage}")
                    RichTextBoxErrorList.AppendText(vbCrLf)
                Next
            End If
            With ConversionBuffer
                .Clear()
                .Select(.TextLength, 0)
                ProgressBar1.Value = 0
                ProgressBar1.Maximum = Lines
                For Each range As Range In FragmentRange
                    .Select(.TextLength, 0)
                    .SelectionColor = ColorSelector.GetColorFromName(range.ClassificationType)
                    .AppendText(range.Text)
                    If range.Text.Contains(vbLf) Then
                        ProgressBar1.Step = range.Text.Count(CType(vbLf, Char))
                        ProgressBar1.PerformStep()
                        ProgressBar1.Step = -1
                        ProgressBar1.PerformStep()
                        ProgressBar1.Step = 1
                        ProgressBar1.PerformStep()
                        ProgressBar1.Update()
                        'Dim percent As Integer = CInt(Math.Truncate((CDbl(ProgressBar1.Value - ProgressBar1.Minimum) / CDbl(ProgressBar1.Maximum - ProgressBar1.Minimum)) * 100))
                        Dim Text As String = $"{ProgressBar1.Value:N0} of {ProgressBar1.Maximum:N0} lines"
                        Using gr As Graphics = ProgressBar1.CreateGraphics()
                            gr.DrawString(Text, SystemFonts.DefaultFont, Brushes.Black, New PointF(ProgressBar1.Width \ 2 - (gr.MeasureString(Text, SystemFonts.DefaultFont).Width / 2.0F), ProgressBar1.Height \ 2 - (gr.MeasureString(Text, SystemFonts.DefaultFont).Height / 2.0F)))
                        End Using

                        Application.DoEvents()
                    End If
                    If StopRequested Then
                        Exit Sub
                    End If
                Next range
                Application.DoEvents()
                If failures?.Count > 0 Then
                    For Each dia As Diagnostic In failures
                        Dim ErrorLine As Integer = dia.Location.GetLineSpan.StartLinePosition.Line
                        Dim ErrorCharactorPosition As Integer = dia.Location.GetLineSpan.StartLinePosition.Character
                        Dim Length As Integer = dia.Location.GetLineSpan.EndLinePosition.Character - ErrorCharactorPosition
                        .Select(.GetFirstCharIndexFromLine(ErrorLine) + ErrorCharactorPosition, Length)
                        .SelectionColor = Color.Red
                        .Select(.TextLength, 0)
                    Next
                    .Select(.GetFirstCharIndexFromLine(failures(0).Location.GetLineSpan.StartLinePosition.Line), 0)
                    .ScrollToCaret()
                End If
            End With
            If failures?.Count > 0 Then
                LineNumbers_For_RichTextBoxInput.Visible = True
                LineNumbers_For_RichTextBoxOutput.Visible = True
            End If
        Catch ex As Exception
            Stop
        End Try
        ProgressBar1.Value = 0
    End Sub

    Private Sub ColorizeResultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorizeResultToolStripMenuItem.Click
        My.Settings.ColorizeOutput = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub ColorizeSourceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorizeSourceToolStripMenuItem.Click
        My.Settings.ColorizeInput = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub Compile_Colorize(TextToCompile As String)
        Dim CompileResult As EmitResult = CompileVisualBasicString(TextToCompile, ErrorsToBeIgnored, DiagnosticSeverity.Error, ResultOfConversion)
        LabelErrorCount.Text = $"Number of Errors: {ResultOfConversion.FilteredListOfFailures.Count}"
        Dim FragmentRange As IEnumerable(Of Range) = GetClassifiedRanges(TextToCompile, LanguageNames.VisualBasic)

        If Not CompileResult?.Success Then
            Dim Success As Boolean = ResultOfConversion.FilteredListOfFailures.Count = 0
            ResultOfConversion.Success = Success
            If Success Then
                If ColorizeResultToolStripMenuItem.Checked Then
                    Colorize(FragmentRange, RichTextBoxConversionOutput, TextToCompile.SplitLines.Length, ResultOfConversion.FilteredListOfFailures)
                Else
                    RichTextBoxConversionOutput.Text = ResultOfConversion.ConvertedCode
                End If
            Else
                Colorize(FragmentRange, RichTextBoxConversionOutput, TextToCompile.SplitLines.Length, ResultOfConversion.FilteredListOfFailures)
            End If
        Else
            If ColorizeResultToolStripMenuItem.Checked Then
                Colorize(FragmentRange, RichTextBoxConversionOutput, TextToCompile.SplitLines.Length)
            Else
                RichTextBoxConversionOutput.Text = ResultOfConversion.ConvertedCode
            End If
        End If
        Application.DoEvents()
    End Sub

    Private Sub CompileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompileToolStripMenuItem.Click
        LineNumbers_For_RichTextBoxInput.Visible = False
        LineNumbers_For_RichTextBoxOutput.Visible = False

        If RichTextBoxConversionOutput.Text.IsEmptyNullOrWhitespace Then
            Exit Sub
        End If
        RichTextBoxErrorList.Text = ""
        Compile_Colorize(RichTextBoxConversionOutput.Text)
    End Sub

    Private Function Convert_Compile_Colorize(RequestToConvert As ConvertRequest) As Boolean
        ResultOfConversion = ConvertInputRequest(RequestToConvert)
        SaveAsToolStripMenuItem.Enabled = ResultOfConversion.Success
        If ResultOfConversion.Success Then
            Compile_Colorize(ResultOfConversion.ConvertedCode)
            LabelErrorCount.Text = $"Number of Errors: {ResultOfConversion.FilteredListOfFailures.Count}"
        Else
            RichTextBoxConversionOutput.SelectionColor = Color.Red
            RichTextBoxConversionOutput.Text = GetExceptionsAsString(ResultOfConversion.Exceptions)
        End If
        Return ResultOfConversion.Success
    End Function

    Private Sub ConvertFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertFolderToolStripMenuItem.Click
        LineNumbers_For_RichTextBoxInput.Visible = False
        LineNumbers_For_RichTextBoxOutput.Visible = False
        InputFolderBrowserDialog1.Description = "Select Folder to convert..."
        InputFolderBrowserDialog1.ShowNewFolderButton = False
        InputFolderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop
        InputFolderBrowserDialog1.SelectedPath = My.Settings.DefaultProjectDirectory

        If InputFolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim FolderName As String = InputFolderBrowserDialog1.SelectedPath
            Dim LanguageExtension As String = RequestToConvert.RequestedConversion.Left(2)
            If Directory.Exists(FolderName) Then
                KeepMonitorActive()
                Dim LastFileNameWithPath As String = If(My.Settings.StartFolderConvertFromLastFile, My.Settings.LastFileNameWithPath, "")
                Dim FilesProcessed As Integer = 0
                If ProcessDirectory(FolderName, LastFileNameWithPath, LanguageExtension, FilesProcessed) Then
                    MsgBox($"Conversion completed, Successfully.")
                Else
                    MsgBox($"Conversion stopped.")
                End If
            Else
                MsgBox($"{FolderName} is not a directory.")
            End If
        End If
    End Sub

    Private Sub ConvertSnipitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertSnipitToolStripMenuItem.Click
        SetButtonStopAndCursor(Me, ButtonStop, True)
        RichTextBoxErrorList.Text = ""
        RichTextBoxFileList.Text = ""
        LineNumbers_For_RichTextBoxOutput.Visible = False
        ResizeRichTextBuffers()
        RequestToConvert.SourceCode = RichTextBoxConversionInput.Text
        Convert_Compile_Colorize(RequestToConvert)
        ConvertFolderToolStripMenuItem.Enabled = True
        SetButtonStopAndCursor(Me, ButtonStop, False)
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        CurrentBuffer.Copy()
    End Sub

    Private Sub CSharp2VB_CheckedChanged(sender As Object, e As EventArgs) Handles CSharp2VB.CheckedChanged
        RichTextBoxConversionInput.Text = ""
        RichTextBoxConversionOutput.Text = ""
        If CSharp2VB.Checked Then
            VB2CSharp.Checked = False
            RequestToConvert = New ConvertRequest With {.RequestedConversion = "cs2vb"}
            Text = "Convert C# To Visual Basic"
        End If
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        CurrentBuffer.Cut()
    End Sub

    Private Sub DelayBetweenConversionsToolStripMenuItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DelayBetweenConversionsToolStripMenuItem.SelectedIndexChanged
        Dim ComboBox As ToolStripComboBox = CType(sender, ToolStripComboBox)
        Select Case DelayBetweenConversionsToolStripMenuItem.Text.Substring("Delay Between Conversions = ".Length)
            Case "None"
                My.Settings.ConversionDelay = 0
            Case "5 Seconds"
                My.Settings.ConversionDelay = 5
            Case "10 Seconds"
                My.Settings.ConversionDelay = 10
            Case Else
        End Select
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    ''' <summary>
    ''' Look in SearchBuffer for text and highlight it
    '''  No error is displayed if not found
    ''' </summary>
    ''' <param name="SearchBuffer"></param>
    ''' <param name="StartLocation"></param>
    ''' <param name="SelectionTextLength"></param>
    ''' <returns>True if found, False is not found</returns>
    Private Function FindTextInBuffer(SearchBuffer As RichTextBox, ByRef StartLocation As Integer, ByRef SelectionTextLength As Integer) As Boolean
        If SelectionTextLength > 0 Then
            SearchBuffer.SelectionBackColor = Color.White
            ' Find the end index. End Index = number of characters in textbox
            ' remove highlight from the search string
            SearchBuffer.Select(StartLocation, SelectionTextLength)
            Application.DoEvents()
        End If
        Dim SearchForward As Boolean = SearchDirection.SelectedIndex = 0
        If StartLocation >= SearchBuffer.Text.Length Then
            StartLocation = If(SearchForward, 0, SearchBuffer.Text.Length - 1)
        End If
#Disable Warning CC0014 ' Use Ternary operator.
        If SearchForward Then
#Enable Warning CC0014 ' Use Ternary operator.
            ' Forward Search
            ' If string was found in the RichTextBox, highlight it
            StartLocation = SearchBuffer.Find(SearchInput.Text, StartLocation, RichTextBoxFinds.None)
        Else
            ' Back Search
            StartLocation = SearchBuffer.Find(SearchInput.Text, StartLocation, RichTextBoxFinds.Reverse)
        End If

        If StartLocation >= 0 Then
            SearchBuffer.ScrollToCaret()
            ' Set the highlight background color as Orange
            SearchBuffer.SelectionBackColor = Color.Orange
            ' Find the end index. End Index = number of characters in textbox
            SelectionTextLength = SearchInput.Text.Length
            ' Highlight the search string
            SearchBuffer.Select(StartLocation, SelectionTextLength)
            ' mark the start position after the position of
            ' last search string
            StartLocation = If(SearchForward, StartLocation + SelectionTextLength, StartLocation - 1)
            Return True
        End If
        StartLocation = If(SearchForward, 0, SearchInput.Text.Length - 1)
        SelectionTextLength = 0
        Return False
    End Function

    Private Sub FIndToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FIndToolStripMenuItem.Click
        SearchBoxVisibility(True)
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        RestoreMonitorSettings()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim items(ImageList1.Images.Count - 1) As String
        For i As Integer = 0 To ImageList1.Images.Count - 1
            items(i) = "Item " & i.ToString
        Next
        SearchDirection.Items.AddRange(items)
        SearchDirection.DropDownStyle = ComboBoxStyle.DropDownList
        SearchDirection.DrawMode = DrawMode.OwnerDrawVariable
        SearchDirection.ItemHeight = ImageList1.ImageSize.Height
        SearchDirection.Width = ImageList1.ImageSize.Width + 30
        SearchDirection.MaxDropDownItems = ImageList1.Images.Count
        SearchDirection.SelectedIndex = 0

        PictureBox1.Height = ImageList1.ImageSize.Height + 4
        PictureBox1.Width = ImageList1.ImageSize.Width + 4
        PictureBox1.Top = SearchDirection.Top + 2
        PictureBox1.Left = SearchDirection.Left + 2
        SearchWhere.SelectedIndex = 0
        SplitContainer1.SplitterDistance = SplitContainer1.Height - RichTextBoxErrorList.Height
        If My.Settings.DefaultProjectDirectory.IsEmptyNullOrWhitespace Then
            My.Settings.DefaultProjectDirectory = GetLatestVisualStudioProjectPath()
            My.Settings.Save()
            Application.DoEvents()
        End If
        AutoCompileToolStripMenuItem.Checked = My.Settings.AutoCompile
        ColorizeResultToolStripMenuItem.Checked = My.Settings.ColorizeOutput
        ColorizeSourceToolStripMenuItem.Checked = My.Settings.ColorizeInput
        PauseConvertOnSuccessToolStripMenuItem.Checked = My.Settings.PauseConvertOnSuccess
        RecentFile1ToolStripMenuItem.Text = My.Settings.LastFileNameWithPath
        SkipAutoGeneratedToolStripMenuItem.Checked = My.Settings.SkipAutoGenerated
        SkipTestResourceFilesToolStripMenuItem.Checked = My.Settings.SkipTestResourceFiles
        StartFolderConvertFromLastFileToolStripMenuItem.Checked = My.Settings.StartFolderConvertFromLastFile
        SnippetLoadLastToolStripMenuItem.Enabled = File.Exists(SnipitFileWithPath)
        LineNumbers_For_RichTextBoxInput.Visible = False
        LineNumbers_For_RichTextBoxOutput.Visible = False
        Width = Screen.PrimaryScreen.Bounds.Width
        Height = CInt(Screen.PrimaryScreen.Bounds.Height * 0.95)
        Select Case My.Settings.ConversionDelay
            Case 0
                DelayBetweenConversionsToolStripMenuItem.SelectedIndex = 0
            Case 5
                DelayBetweenConversionsToolStripMenuItem.SelectedIndex = 1
            Case 10
                DelayBetweenConversionsToolStripMenuItem.SelectedIndex = 2
            Case Else
                DelayBetweenConversionsToolStripMenuItem.SelectedIndex = 0
        End Select
        CenterToScreen()
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ResizeRichTextBuffers()
    End Sub

    Private Sub InputBufferLineNumbersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InputBufferLineNumbersToolStripMenuItem.Click
        LineNumbers_For_RichTextBoxInput.Visible = CType(sender, ToolStripMenuItem).Checked
        ResizeRichTextBuffers()
    End Sub

    Private Sub LineNumbers_For_RichTextBoxInput_Resize(sender As Object, e As EventArgs) Handles LineNumbers_For_RichTextBoxInput.Resize
        ResizeRichTextBuffers()
    End Sub

    Private Sub LineNumbers_For_RichTextBoxInput_VisibleChanged(sender As Object, e As EventArgs) Handles LineNumbers_For_RichTextBoxInput.VisibleChanged
        InputBufferLineNumbersToolStripMenuItem.Checked = CType(sender, LineNumbers_For_RichTextBox).Visible
        ResizeRichTextBuffers()
    End Sub

    Private Sub LineNumbers_For_RichTextBoxOutput_Resize(sender As Object, e As EventArgs) Handles LineNumbers_For_RichTextBoxOutput.Resize
        ResizeRichTextBuffers()
    End Sub

    Private Sub LineNumbers_For_RichTextBoxOutput_VisibleChanged(sender As Object, e As EventArgs) Handles LineNumbers_For_RichTextBoxOutput.VisibleChanged
        OuputBufferLineNumbersToolStripMenuItem.Checked = CType(sender, LineNumbers_For_RichTextBox).Visible
        ResizeRichTextBuffers()
    End Sub

    Private Function LoadInputBufferFromStream(LanguageExtension As String, fileStream As Stream, SkipAutoGenerated As Boolean) As Integer
        StopRequested = False
        LocalUseWaitCutsor(Me, True)
        Dim SourceText As String = GetFileTextFromStream(fileStream)
        Dim InputLines As Integer
        Dim ConversionInputLinesArray() As String = SourceText.SplitLines
        If SkipAutoGenerated Then
            If SkipAutoGenerated Then
                If SourceText.Contains({"< autogenerated", "<auto-generated"}, StringComparison.CurrentCultureIgnoreCase) Then
                    Return 0
                End If
            End If
        End If
        InputLines = ConversionInputLinesArray.Length
        If ColorizeSourceToolStripMenuItem.Checked Then
            Colorize(GetClassifiedRanges(ConversionInputLinesArray.Join(vbCrLf), If(LanguageExtension = "vb", LanguageNames.VisualBasic, LanguageNames.CSharp)), RichTextBoxConversionInput, InputLines)
        Else
            RichTextBoxConversionInput.Text = ConversionInputLinesArray.Join(vbCrLf)
        End If
        LocalUseWaitCutsor(Me, False)
        Return InputLines
    End Function

    Private Sub MiscellaneousOptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MiscellaneousOptionsToolStripMenuItem.Click
        OptionsDialog.Show(Me)
    End Sub
    Private Sub OpenFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileToolStripMenuItem.Click
        OpenFileDialog1.AddExtension = True
        Dim LanguageExtension As String = RequestToConvert.RequestedConversion.Left(2)
        OpenFileDialog1.DefaultExt = LanguageExtension
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = If(LanguageExtension = "vb", "VB Code Files (*.vb)|*.vb", "C# Code Files (*.cs)|*.cs")
        SaveFileDialog1.FilterIndex = 0
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.ReadOnlyChecked = True
        OpenFileDialog1.Title = $"Open {LanguageExtension.ToUpper} Source file"
        OpenFileDialog1.ValidateNames = True

        ' InputLines is used for future progress bar
        Dim FileOpenResult As DialogResult = OpenFileDialog1.ShowDialog
        If FileOpenResult = DialogResult.OK Then
            SetButtonStopAndCursor(Me, ButtonStop, True)
            Dim RecentFileNameWithPath As String = OpenFileDialog1.FileName
            Dim fileStream As Stream = File.OpenRead(RecentFileNameWithPath)
            Dim InputLines As Integer = LoadInputBufferFromStream(LanguageExtension, fileStream, False)
            ConvertFolderToolStripMenuItem.Enabled = False
            SetButtonStopAndCursor(Me, ButtonStop, False)
        Else
            ConvertFolderToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub OuputBufferLineNumbersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OuputBufferLineNumbersToolStripMenuItem.Click
        LineNumbers_For_RichTextBoxOutput.Visible = CType(sender, ToolStripMenuItem).Checked
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBoxConversionInput.SelectedText = Clipboard.GetText(TextDataFormat.Text)
    End Sub

    Private Sub PauseConvertOnSuccessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PauseConvertOnSuccessToolStripMenuItem.Click
        My.Settings.PauseConvertOnSuccess = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If SearchInput.Text.IsEmptyNullOrWhitespace Then
            Exit Sub
        End If
        Select Case SearchWhere.SelectedIndex
            Case 0
                If Not FindTextInBuffer(RichTextBoxConversionInput, InputBufferCurrentIndex, InputBufferCurrentLength) Then
                    MsgBox($"'{SearchInput.Text}' not found in Source Buffer!", MsgBoxStyle.OkOnly, "Not Found!")
                End If
            Case 1
                If Not FindTextInBuffer(RichTextBoxConversionOutput, OutputBufferCurrentIndex, OutputBufferCurrentLength) Then
                    MsgBox($"'{SearchInput.Text}' not found Conversion Buffer!", MsgBoxStyle.OkOnly, "Not Found!")
                End If
            Case 2
                If Not FindTextInBuffer(RichTextBoxConversionInput, InputBufferCurrentIndex, InputBufferCurrentLength) Then
                    MsgBox($"'{SearchInput.Text}' not found in Source Buffer!", MsgBoxStyle.OkOnly, "Not Found!")
                End If
                If Not FindTextInBuffer(RichTextBoxConversionOutput, OutputBufferCurrentIndex, OutputBufferCurrentLength) Then
                    MsgBox($"'{SearchInput.Text}' not found Conversion Buffer!", MsgBoxStyle.OkOnly, "Not Found!")
                End If
        End Select
    End Sub

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        PictureBox1.BorderStyle = BorderStyle.Fixed3D
    End Sub

    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        PictureBox1.BorderStyle = BorderStyle.FixedSingle
    End Sub

    ''' <summary>
    ''' Process all files in the directory passed in, recurse on any directories
    ''' that are found, and process the files they contain.
    ''' </summary>
    ''' <param name="targetDirectory">Start location of where to process directories</param>
    ''' <param name="LastFileNameWithPath">Pass Last File Name to Start Conversion where you left off</param>
    ''' <param name="LanguageExtension">vb or cs</param>
    ''' <param name="FilesProcessed">Count of the number of tiles processed</param>
    ''' <returns>
    ''' False if error and user wants to stop, True if success or user wants to ignore error
    ''' </returns>
    Private Function ProcessDirectory(targetDirectory As String, LastFileNameWithPath As String, LanguageExtension As String, ByRef FilesProcessed As Integer) As Boolean
        Try
            RichTextBoxErrorList.Text = ""
            RichTextBoxFileList.Text = ""
            SetButtonStopAndCursor(Me, ButtonStop, True)
            ' Process the list of files found in the directory.
            Return ProcessDirectoriesModule.ProcessDirectory(targetDirectory, Me, ButtonStop, RichTextBoxFileList, LastFileNameWithPath, LanguageExtension, FilesProcessed, AddressOf ProcessFile)
        Catch ex As Exception
            ' don't crash on exit
            Stop
        Finally
            RestoreMonitorSettings()
            SetButtonStopAndCursor(Me, ButtonStop, False)
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Convert one file
    ''' </summary>
    ''' <param name="PathWithFileName">Complete path including file name to file to be converted</param>
    ''' <param name="LanguageExtension">vb or cs</param>
    ''' <returns>False if error and user wants to stop, True if success or user wants to ignore error.</returns>
    Private Function ProcessFile(ByVal PathWithFileName As String, LanguageExtension As String) As Boolean
        ButtonStop.Visible = True
        RichTextBoxConversionOutput.Text = ""
        StopRequested = False
        RecentFile1ToolStripMenuItem.Text = PathWithFileName
        Dim fs As FileStream = File.OpenRead(PathWithFileName)
        Dim lines As Integer = LoadInputBufferFromStream(LanguageExtension, fs, SkipAutoGeneratedToolStripMenuItem.Checked)
        If lines > 0 Then
            RequestToConvert.SourceCode = RichTextBoxConversionInput.Text
            If Not Convert_Compile_Colorize(RequestToConvert) Then
                RestoreMonitorSettings()
                If MsgBox($"Conversion of {PathWithFileName} failed, stop conversion?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Return False
                End If
            Else
                If My.Settings.PauseConvertOnSuccess Then
                    RestoreMonitorSettings()
                    If MsgBox($"{PathWithFileName} successfully converted, Continue?", MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Return False
                    End If
                End If
            End If
            ' 5 second delay
            Const LoopSleep As Integer = 25
            Dim Delay As Integer = CInt((1000 * My.Settings.ConversionDelay) / LoopSleep)
            For i As Integer = 0 To Delay
                Application.DoEvents()
                Threading.Thread.Sleep(LoopSleep)
                If StopRequested Then
                    StopRequested = False
                    Return False

                End If
            Next
            Application.DoEvents()
        Else
            RichTextBoxConversionOutput.Clear()
        End If
        Return True
    End Function

    Private Sub RecentFile1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RecentFile1ToolStripMenuItem.Click
        SetButtonStopAndCursor(Me, ButtonStop, True)
        Dim RecentFileNameWithPath As String = CType(sender, ToolStripMenuItem).Text
        Dim fileStream As Stream = File.OpenRead(RecentFileNameWithPath)
        Dim LanguageExtension As String = RequestToConvert.RequestedConversion.Left(2)
        Dim InputLines As Integer = LoadInputBufferFromStream(LanguageExtension, fileStream, False)
        ConvertFolderToolStripMenuItem.Enabled = False
        LocalUseWaitCutsor(Me, False)
    End Sub

    Private Sub RecentFile1ToolStripMenuItem_MouseDown(sender As Object, e As MouseEventArgs) Handles RecentFile1ToolStripMenuItem.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Clipboard.SetText(RecentFile1ToolStripMenuItem.Text)
        End If
    End Sub

    Private Sub RecentFile1ToolStripMenuItem_TextChanged(sender As Object, e As EventArgs) Handles RecentFile1ToolStripMenuItem.TextChanged
        Dim RecentFileTSMI As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If RecentFileTSMI.Text = "RecentFile1" Then
            Exit Sub
        End If
        With RecentFileTSMI
            My.Settings.LastFileNameWithPath = .Text
            RecentFolder1ToolStripMenuItem.Text = Path.GetDirectoryName(.Text)
            My.Settings.Save()
            RecentFileTSMI.Visible = .Text.IsNotEmptyNullOrWhitespace
            Application.DoEvents()
        End With
    End Sub

    Private Sub RecentFolder1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RecentFolder1ToolStripMenuItem.Click
        Dim FolderName As String = CType(sender, ToolStripMenuItem).Text
        If Directory.Exists(FolderName) Then
            ' This path is a directory.
            Dim LanguageExtension As String = RequestToConvert.RequestedConversion.Left(2)
            Dim LastFileNameWithPath As String = If(My.Settings.StartFolderConvertFromLastFile, My.Settings.LastFileNameWithPath, "")
            Dim FilesProcessed As Integer = 0
            If ProcessDirectory(FolderName, LastFileNameWithPath, LanguageExtension, FilesProcessed) Then
                MsgBox($"Conversion completed.")
            End If
        Else
            MsgBox($"{FolderName} is not a directory.")
        End If
    End Sub

    Private Sub RecentFolder1ToolStripMenuItem_TextChanged(sender As Object, e As EventArgs) Handles RecentFolder1ToolStripMenuItem.TextChanged
        Dim RecentFolderTSMI As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If RecentFolderTSMI.Text = "RecentFolder1" Then
            Exit Sub
        End If
        My.Settings.LastPath = RecentFolderTSMI.Text
        My.Settings.Save()
        RecentFolderTSMI.Visible = RecentFolderTSMI.Text.IsNotEmptyNullOrWhitespace
        Application.DoEvents()
    End Sub

    Private Sub ResizeRichTextBuffers()
        Dim LineNumberInputWidth As Integer = If(LineNumbers_For_RichTextBoxInput.Visible, LineNumbers_For_RichTextBoxInput.Width, 0)
        Dim LineNumberOutputWidth As Integer = If(LineNumbers_For_RichTextBoxOutput.Visible, LineNumbers_For_RichTextBoxOutput.Width, 0)

        RichTextBoxConversionInput.Width = CInt((ClientSize.Width / 2 + 0.5)) - LineNumberInputWidth
        RichTextBoxFileList.Width = CInt(ClientSize.Width / 2 + 0.5)

        RichTextBoxConversionOutput.Width = ClientSize.Width - (RichTextBoxConversionInput.Width + LineNumberInputWidth + LineNumberOutputWidth)
        RichTextBoxConversionOutput.Left = RichTextBoxConversionInput.Width + LineNumberInputWidth + LineNumberOutputWidth

        RichTextBoxErrorList.Left = CInt(ClientSize.Width / 2)
        RichTextBoxErrorList.Width = CInt(ClientSize.Width / 2)
    End Sub
    Private Sub RichTexBoxErrorList_DoubleClick(sender As Object, e As EventArgs) Handles RichTextBoxErrorList.DoubleClick
        If RTFLineStart > 0 AndAlso RichTextBoxConversionOutput.SelectionStart <> RTFLineStart Then
            RichTextBoxConversionOutput.Select(RTFLineStart, 0)
            RichTextBoxConversionOutput.ScrollToCaret()
        End If
    End Sub

    Private Sub RichTexBoxErrorList_MouseDown(sender As Object, e As MouseEventArgs) Handles RichTextBoxErrorList.MouseDown
        Dim box As RichTextBox = DirectCast(sender, RichTextBox)
        If box.TextLength = 0 Then
            Exit Sub
        End If
        Dim index As Integer = box.GetCharIndexFromPosition(e.Location)
        Dim line As Integer = box.GetLineFromCharIndex(index)
        Dim lineStart As Integer = box.GetFirstCharIndexFromLine(line)
        Dim AfterEquals As Integer = box.GetFirstCharIndexFromLine(line) + 15
        Dim LineText As String = box.Text.Substring(box.GetFirstCharIndexFromLine(line))
        If Not LineText.StartsWith("BC") Then
            Exit Sub
        End If
        If Not LineText.Contains(" Line = ") Then
            Exit Sub
        End If
        Dim NumberCount As Integer = LineText.Substring(15).IndexOf(" ")
        Dim ErrorLine As Integer = CInt(Val(box.Text.Substring(AfterEquals, NumberCount)))
        If ErrorLine <= 0 Then
            Exit Sub
        End If
        RTFLineStart = RichTextBoxConversionOutput.GetFirstCharIndexFromLine(ErrorLine - 1)
    End Sub

    Private Sub RichTextBoxConversionInput_Enter(sender As Object, e As EventArgs) Handles RichTextBoxConversionInput.Enter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub RichTextBoxConversionInput_MouseEnter(sender As Object, e As EventArgs) Handles RichTextBoxConversionInput.MouseEnter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub RichTextBoxConversionInput_TextChanged(sender As Object, e As EventArgs) Handles RichTextBoxConversionInput.TextChanged
        Dim InputBufferInUse As Boolean = CType(sender, RichTextBox).TextLength > 0
        SnippetSaveToolStripMenuItem.Enabled = InputBufferInUse
        ConvertSnipitToolStripMenuItem.Enabled = InputBufferInUse
        ConvertFolderToolStripMenuItem.Enabled = InputBufferInUse
        'LineNumbers_For_RichTextBoxInput.Visible = LineNumbers_For_RichTextBoxInput.Visible And InputBufferInUse
    End Sub

    Private Sub RichTextBoxConversionOutput_Enter(sender As Object, e As EventArgs) Handles RichTextBoxConversionOutput.Enter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub RichTextBoxConversionOutput_MouseEnter(sender As Object, e As EventArgs) Handles RichTextBoxConversionOutput.MouseEnter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub RichTextBoxConversionOutput_TextChanged(sender As Object, e As EventArgs) Handles RichTextBoxConversionOutput.TextChanged
        Dim OutputBufferInUse As Boolean = CType(sender, RichTextBox).TextLength > 0
        CompileToolStripMenuItem.Enabled = OutputBufferInUse
    End Sub
    Private Sub RichTextBoxErrorList_Enter(sender As Object, e As EventArgs) Handles RichTextBoxErrorList.Enter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub RichTextBoxErrorList_MouseEnter(sender As Object, e As EventArgs) Handles RichTextBoxErrorList.MouseEnter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    'Private Sub RichTextBoxErrorList_TextChanged(sender As Object, e As EventArgs) Handles RichTextBoxErrorList.TextChanged
    '    Dim HasErrors As Boolean = RichTextBoxErrorList.TextLength > 0
    '    LineNumbers_For_RichTextBoxInput.Visible = HasErrors
    '    LineNumbers_For_RichTextBoxOutput.Visible = HasErrors AndAlso RichTextBoxConversionOutput.TextLength > 0
    'End Sub

    Private Sub RichTextBoxFileList_Enter(sender As Object, e As EventArgs) Handles RichTextBoxFileList.Enter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub RichTextBoxFileList_MouseEnter(sender As Object, e As EventArgs) Handles RichTextBoxFileList.MouseEnter
        CurrentBuffer = CType(sender, RichTextBox)
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim Extension As String = RequestToConvert.RequestedConversion.Right(2)

        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.CreatePrompt = False
        SaveFileDialog1.DefaultExt = Extension
        SaveFileDialog1.FileName = Path.ChangeExtension(OpenFileDialog1.SafeFileName, Extension)
        SaveFileDialog1.Filter = If(Extension = "vb", "VB Code Files (*.vb)|*.vb", "C# Code Files (*.cs)|*.cs")
        SaveFileDialog1.FilterIndex = 0
        SaveFileDialog1.OverwritePrompt = True
        SaveFileDialog1.SupportMultiDottedExtensions = False
        SaveFileDialog1.Title = $"Save {Extension.ToUpper} Output..."
        SaveFileDialog1.ValidateNames = True
        Dim FileSaveResult As DialogResult = SaveFileDialog1.ShowDialog
        If FileSaveResult = DialogResult.OK Then
            RichTextBoxConversionOutput.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.PlainText)
        End If
    End Sub

    Private Sub SearchBoxVisibility(Visible As Boolean)
        ButtonSearch.Visible = Visible
        SearchDirection.Visible = Visible
        PictureBox1.Visible = Visible
        SearchInput.Visible = Visible
    End Sub

    Private Sub SearchDirection_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs) Handles SearchDirection.DrawItem
        If e.Index <> -1 Then
            e.Graphics.DrawImage(ImageList1.Images(e.Index), e.Bounds.Left, e.Bounds.Top)
        End If
    End Sub

    Private Sub SearchDirection_MeasureItem(ByVal sender As Object, ByVal e As MeasureItemEventArgs) Handles SearchDirection.MeasureItem
        e.ItemHeight = ImageList1.ImageSize.Height
        e.ItemWidth = ImageList1.ImageSize.Width
    End Sub

    Private Sub SearchDirection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchDirection.SelectedIndexChanged
        PictureBox1.Image = ImageList1.Images(SearchDirection.SelectedIndex)
    End Sub

    Private Sub SearchInput_TextChanged(sender As Object, e As EventArgs) Handles SearchInput.TextChanged
        InputBufferCurrentIndex = 0
        OutputBufferCurrentIndex = 0
    End Sub

    Private Sub SkipAutoGeneratedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SkipAutoGeneratedToolStripMenuItem.Click
        My.Settings.SkipAutoGenerated = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub SkipTestResourceFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SkipTestResourceFilesToolStripMenuItem.Click
        My.Settings.SkipTestResourceFiles = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub SnippetLoadLastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SnippetLoadLastToolStripMenuItem.Click
        RichTextBoxConversionInput.LoadFile(SnipitFileWithPath)
    End Sub

    Private Sub SnippetSaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SnippetSaveToolStripMenuItem.Click
        If RichTextBoxConversionInput.TextLength = 0 Then
            Exit Sub
        End If
        RichTextBoxConversionInput.SaveFile(SnipitFileWithPath, RichTextBoxStreamType.RichText)
    End Sub

    Private Sub StartFolderConvertFromLastFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartFolderConvertFromLastFileToolStripMenuItem.Click
        My.Settings.StartFolderConvertFromLastFile = CType(sender, ToolStripMenuItem).Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub

    Private Sub ToolStripMenuItem_Copy_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_Copy.Click
        Dim RichTextContextMenu As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        CType(ContextMenuStrip1.SourceControl, RichTextBox).Copy()

    End Sub

    Private Sub ToolStripMenuItem_Cut_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_Cut.Click
        Dim RichTextContextMenu As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        CType(ContextMenuStrip1.SourceControl, RichTextBox).Cut()
    End Sub

    Private Sub ToolStripMenuItem_Paste_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_Paste.Click
        Dim RichTextContextMenu As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        CType(ContextMenuStrip1.SourceControl, RichTextBox).Paste()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        CurrentBuffer.Undo()
    End Sub

    Private Sub VB2CSharp_CheckedChanged(sender As Object, e As EventArgs) Handles VB2CSharp.CheckedChanged
        RichTextBoxConversionInput.Text = ""
        RichTextBoxConversionOutput.Text = ""
        If VB2CSharp.Checked Then
            CSharp2VB.Checked = False
            RequestToConvert = New ConvertRequest With {.RequestedConversion = "vb2cs"}
            Text = "Convert Visual Basic To C#"
        End If
    End Sub
End Class
