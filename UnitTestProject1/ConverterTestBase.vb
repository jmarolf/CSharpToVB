Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable

#If VB_TO_CSharp Then
Imports IVisualBasicCode.CodeConverter.CSharp
#End If
Imports IVisualBasicCode.CodeConverter.VB
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.CSharp.Formatting
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports IVisualBasicCode.CodeConverter.Util
Imports Xunit
Namespace CodeConverter.Tests
    Public Class ConverterTestBase
        ' Do not remove Version
        Public Property Version As Integer = 1
        Private Shared Sub CSharpWorkspaceSetup(ByRef workspace As TestWorkspace, ByRef doc As Document, Optional parseOptions As CSharpParseOptions = Nothing)
            workspace = New TestWorkspace()
            Dim projectId As ProjectId = ProjectId.CreateNewId()
            Dim documentId As DocumentId = DocumentId.CreateNewId(projectId)
            If parseOptions Is Nothing Then
                parseOptions = New CSharpParseOptions(
                   CSharp.LanguageVersion.CSharp6,
                    DocumentationMode.Diagnose Or DocumentationMode.Parse,
                    SourceCodeKind.Regular,
                    ImmutableArray.Create("DEBUG", "TEST"))
            End If
            workspace.Options.WithChangedOption(CSharpFormattingOptions.NewLinesForBracesInControlBlocks, False)
            workspace.Open(ProjectInfo.Create(
                projectId,
                VersionStamp.Create(),
                "TestProject",
                "TestProject",
                LanguageNames.CSharp,
                Nothing,
                Nothing,
                New CSharpCompilationOptions(
                    OutputKind.DynamicallyLinkedLibrary,
                    False,
                    "",
                    "",
                    "Script",
                    {NameOf(System), "System.Collections.Generic", "System.Linq"},
                    OptimizationLevel.Debug,
                    False,
                    True
                ),
                parseOptions,
                    {
                    DocumentInfo.Create(
                        documentId,
                        "a.cs",
                        Nothing,
                        SourceCodeKind.Regular
                    )
                },
                Nothing,
                DefaultMetadataReferences)
            )
            doc = workspace.CurrentSolution.GetProject(projectId).GetDocument(documentId)
        End Sub

        Private Shared Sub CSharpWorkspaceSetup(text As String, ByRef workspace As TestWorkspace, ByRef doc As Document, Optional parseOptions As CSharpParseOptions = Nothing)
            workspace = New TestWorkspace()
            Dim projectId As ProjectId = ProjectId.CreateNewId()
            Dim documentId As DocumentId = DocumentId.CreateNewId(projectId)
            If parseOptions Is Nothing Then
                parseOptions = New CSharpParseOptions(
                    languageVersion:=CSharp.LanguageVersion.CSharp6,
                    documentationMode:=DocumentationMode.Diagnose Or DocumentationMode.Parse,
                    kind:=SourceCodeKind.Regular,
                    preprocessorSymbols:=ImmutableArray.Create("DEBUG", "TEST")
                )
            End If
            workspace.Options.WithChangedOption([option]:=CSharpFormattingOptions.NewLinesForBracesInControlBlocks, value:=False)
            workspace.Open(ProjectInfo.Create(
                id:=projectId,
                version:=VersionStamp.Create(),
                name:="TestProject",
                assemblyName:="TestProject",
                language:=LanguageNames.CSharp,
                filePath:=Nothing,
                outputFilePath:=Nothing,
                compilationOptions:=New CSharpCompilationOptions(
                    OutputKind.DynamicallyLinkedLibrary,
                    False,
                    "",
                    "",
                    "Script",
                    {NameOf(System), "System.Collections.Generic", "System.Linq"},
                    OptimizationLevel.Debug,
                    False,
                    True
                    ),
                parseOptions:=parseOptions,
                documents:={
                    DocumentInfo.Create(
                        id:=documentId,
                        name:="a.cs",
                        folders:=Nothing,
                        sourceCodeKind:=SourceCodeKind.Regular,
                        loader:=TextLoader.From(textAndVersion:=TextAndVersion.Create(text:=SourceText.From(text), version:=VersionStamp.Create()))
                    )
                },
                projectReferences:=Nothing,
                metadataReferences:=DiagnosticTestBase.DefaultMetadataReferences)
                )
            doc = workspace.CurrentSolution.GetProject(projectId).GetDocument(documentId)
        End Sub

        Private Shared Sub VBWorkspaceSetup(ByRef workspace As TestWorkspace, ByRef doc As Document, Optional parseOptions As VisualBasicParseOptions = Nothing)
            workspace = New TestWorkspace()
            If parseOptions Is Nothing Then
                parseOptions = New VisualBasicParseOptions(
                    languageVersion:=VisualBasic.LanguageVersion.VisualBasic15_5,
                    documentationMode:=DocumentationMode.Diagnose Or DocumentationMode.Parse,
                    kind:=SourceCodeKind.Regular,
                    preprocessorSymbols:=New Dictionary(Of String, Object) From {{"NETSTANDARD2_0", Nothing}}
                )
            End If
            workspace.Options.WithChangedOption([option]:=CSharpFormattingOptions.NewLinesForBracesInControlBlocks, value:=False)
            Dim compilationOptions As VisualBasicCompilationOptions = New VisualBasicCompilationOptions(outputKind:=OutputKind.DynamicallyLinkedLibrary).
                WithRootNamespace(rootNamespace:="TestProject").
                WithGlobalImports(globalImports:=GlobalImport.Parse(NameOf(System), "System.Collections.Generic", "System.Linq", "Microsoft.VisualBasic"))

            Dim ProjectId As ProjectId = ProjectId.CreateNewId()
            Dim DocumentID As DocumentId = DocumentId.CreateNewId(projectId:=ProjectId)
            workspace.Open(ProjectInfo.Create(
                                            id:=ProjectId,
                                            version:=VersionStamp.Create(),
                                            name:="TestProject",
                                            assemblyName:="TestProject",
                                            language:=LanguageNames.VisualBasic,
                                            filePath:=Nothing,
                                            outputFilePath:=Nothing,
                                            compilationOptions:=compilationOptions,
                                            parseOptions:=parseOptions,
                                            documents:={
                                                DocumentInfo.Create(
                                                    DocumentID,
                                                    "a.vb",
                                                    Nothing,
                                                    SourceCodeKind.Regular
                                                )
                                            },
                                            projectReferences:=Nothing,
                                            metadataReferences:=DiagnosticTestBase.DefaultMetadataReferences))
            doc = workspace.CurrentSolution.GetProject(ProjectId).GetDocument(DocumentID)
        End Sub
        ' Converts C# to VB
        Public Shared Function Convert(input As CSharpSyntaxNode, semanticModel As SemanticModel, targetDocument As Document) As VisualBasicSyntaxNode
            Return CSharpConverter.Convert(input, semanticModel, targetDocument)
        End Function

        Public Shared Sub TestConversionCSharpToVisualBasic(csharpCode As String, DesiredResult As String, Optional csharpOptions As CSharpParseOptions = Nothing, Optional vbOptions As VisualBasicParseOptions = Nothing)
            Dim csharpWorkspace As TestWorkspace = Nothing
            Dim vbWorkspace As TestWorkspace = Nothing
            Dim inputDocument As Document = Nothing
            Dim outputDocument As Document = Nothing

            CSharpWorkspaceSetup(text:=csharpCode, workspace:=csharpWorkspace, doc:=inputDocument, parseOptions:=csharpOptions)
            VBWorkspaceSetup(workspace:=vbWorkspace, doc:=outputDocument, parseOptions:=vbOptions)

            Dim InputNode As SyntaxNode = inputDocument.GetSyntaxRootAsync().GetAwaiter().GetResult()
            Dim lSemanticModel As SemanticModel = inputDocument.GetSemanticModelAsync().GetAwaiter().GetResult()
            Dim outputNode As SyntaxNode = Convert(CType(InputNode, CSharpSyntaxNode), lSemanticModel, outputDocument)


            Dim ActualResult As String = outputDocument.WithSyntaxRoot(root:=outputNode.NormalizeWhitespaceEx(useDefaultCasing:=True, indentation:="    ", elasticTrivia:=False, PreserveCRLF:=True)).GetTextAsync().GetAwaiter().GetResult().ToString()
            ActualResult = HomogenizeEol(ActualResult).TrimEnd()
            DesiredResult = HomogenizeEol(DesiredResult).TrimEnd()
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(DesiredResult, ActualResult, FindFirstDifferenceLine(DesiredResult, ActualResult))

        End Sub

#If VB_TO_CSharp Then

        Private Shared Sub VBWorkspaceSetup(text As String, ByRef workspace As TestWorkspace, ByRef doc As Document, Optional parseOptions As VisualBasicParseOptions = Nothing)
            workspace = New TestWorkspace()
            Dim projectId As ProjectId = ProjectId.CreateNewId()
            Dim documentId As DocumentId = DocumentId.CreateNewId(projectId)
            If parseOptions Is Nothing Then
                parseOptions = New VisualBasicParseOptions(
                    VisualBasic.LanguageVersion.VisualBasic15_5,
                    DocumentationMode.Diagnose Or DocumentationMode.Parse,
                    SourceCodeKind.Regular)
            End If
            workspace.Options.WithChangedOption(CSharpFormattingOptions.NewLinesForBracesInControlBlocks, False)
            Dim compilationOptions As VisualBasicCompilationOptions = (New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)).WithRootNamespace("TestProject").WithGlobalImports(GlobalImport.Parse(NameOf(System), "System.Collections.Generic", "System.Linq", "Microsoft.VisualBasic"))
            workspace.Open(ProjectInfo.Create(
                                           projectId, VersionStamp.Create(),
                                           "TestProject",
                                           "TestProject",
                                           LanguageNames.VisualBasic,
                                           Nothing,
                                           Nothing,
                                           compilationOptions,
                                           parseOptions,
                                           {
                                                DocumentInfo.Create(
                                                    documentId,
                                                    "a.vb",
                                                    Nothing,
                                                    SourceCodeKind.Regular,
                                                    TextLoader.From(
                                                        TextAndVersion.Create(
                                                            SourceText.From(text),
                                                            VersionStamp.Create()
                                                            )
                                                        )
                                                    )
                                           },
                                           Nothing,
                                           DiagnosticTestBase.DefaultMetadataReferences)
                                   )
            doc = workspace.CurrentSolution.GetProject(projectId).GetDocument(documentId)
        End Sub

        Public Sub TestConversionVisualBasicToCSharp(visualBasicCode As String, expectedCsharpCode As String, Optional csharpOptions As CSharpParseOptions = Nothing, Optional vbOptions As VisualBasicParseOptions = Nothing)
            Dim csharpWorkspace As TestWorkspace = Nothing
            Dim vbWorkspace As TestWorkspace = Nothing
            Dim inputDocument As Document = Nothing
            Dim outputDocument As Document = Nothing
            VBWorkspaceSetup(visualBasicCode, vbWorkspace, inputDocument, vbOptions)
            CSharpWorkspaceSetup(csharpWorkspace, outputDocument, csharpOptions)
            Dim outputNode As CSharpSyntaxNode = Convert(CType(inputDocument.GetSyntaxRootAsync().Result, VisualBasicSyntaxNode), inputDocument.GetSemanticModelAsync().Result, outputDocument)

            Dim txt As String = outputDocument.WithSyntaxRoot(Formatter.Format(outputNode, vbWorkspace)).GetTextAsync().Result.ToString()
            txt = Utils.HomogenizeEol(txt).TrimEnd()
            expectedCsharpCode = Utils.HomogenizeEol(expectedCsharpCode).TrimEnd()
            If expectedCsharpCode <> txt Then
                Dim l As Integer = Math.Max(expectedCsharpCode.Length, txt.Length)
                Dim sb As New StringBuilder(l * 4)
                sb.AppendLine("expected:")
                sb.AppendLine(expectedCsharpCode)
                sb.AppendLine("got:")
                sb.AppendLine(txt)
                sb.AppendLine("diff:")
                For i As Integer = 0 To l - 1
                    If i >= expectedCsharpCode.Length OrElse i >= txt.Length OrElse expectedCsharpCode.Chars(i) <> txt.Chars(i) Then
                        sb.Append("x"c)
                    Else
                        sb.Append(expectedCsharpCode.Chars(i))
                    End If
                Next i
                Assert.True(False, sb.ToString())
            End If
        End Sub
        ' Converts VB to C#
        Private Function Convert(input As VisualBasicSyntaxNode, semanticModel As SemanticModel, targetDocument As Document) As CSharpSyntaxNode
            Return VisualBasicConverter.Convert(input, semanticModel, targetDocument)
        End Function
#End If
        Private Shared Function FindFirstDifferenceColumn(DesiredLine As String, ActualLine As String) As (ColumnIndex As Integer, Character As String)
            Dim minLength As Integer = Math.Min(DesiredLine.Length, ActualLine.Length) - 1
            For i As Integer = 0 To minLength
                If Not DesiredLine.Substring(i, 1).Equals(ActualLine.Substring(i, 1), StringComparison.CurrentCulture) Then
                    Return (i + 1, $"Desired Character ""{DesiredLine.Substring(i, 1)}"", Actual Character ""{ActualLine.Substring(i, 1)}""")
                End If
            Next
#Disable Warning CC0013 ' Use Ternary operator.
            If DesiredLine.Length > ActualLine.Length Then
#Enable Warning CC0013 ' Use Ternary operator.
                Return (minLength + 1, $"Desired Character ""{DesiredLine.Substring(minLength + 1, 1)}"", Actual Character Nothing")
            Else
                Return (minLength + 1, $"Desired Character Nothing, Actual Character ""{ActualLine.Substring(minLength + 1, 1)}""")
            End If
        End Function

        Private Shared Function FindFirstDifferenceLine(DesiredText As String, ActualText As String) As String
            Dim Desiredlines() As String = DesiredText.Replace(vbCr, "").Split(CType(vbLf, Char()))
            Dim ActuaLines() As String = ActualText.Replace(vbCr, "").Split(CType(vbLf, Char()))
            For i As Integer = 0 To Math.Min(Desiredlines.GetUpperBound(0), ActuaLines.GetUpperBound(0))
                Dim DesiredLine As String = Desiredlines(i)
                Dim ActualLine As String = ActuaLines(i)
                If Not DesiredLine.Equals(ActualLine, StringComparison.CurrentCulture) Then
                    Dim p As (ColumnIndex As Integer, Character As String) = FindFirstDifferenceColumn(DesiredLine, ActualLine)
                    Return $"{vbCrLf}Expected Line_{i + 1} {DesiredLine}{vbCrLf}Actual Line____{i + 1} {ActualLine}{vbCrLf}Column {p.ColumnIndex} {p.Character}"
                End If
            Next
            Return "Files identical"
        End Function
    End Class
End Namespace
