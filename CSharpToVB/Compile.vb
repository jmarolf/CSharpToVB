Imports System.Collections.Immutable
Imports System.IO
Imports IVisualBasicCode.CodeConverter
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic

Public Module Compile
    Public Function CompileVisualBasicString(StringToBeCompiler As String, ErrorsToBeIgnored() As String, SeverityToReport As DiagnosticSeverity, ByRef ResultOfConversion As ConversionResult) As EmitResult
        If StringToBeCompiler.IsEmptyNullOrWhitespace Then
            ResultOfConversion.FilteredListOfFailures = New List(Of Diagnostic)
            ResultOfConversion.Success = True
            Return Nothing
        End If
        Dim PreprocessorSymbols As New Dictionary(Of String, Object) From {
            {"NETSTANDARD2_0", True},
            {"NET46", True},
            {"NETCOREAPP2_0", True}
        }
        Dim ParseOptions As VisualBasicParseOptions = New VisualBasicParseOptions(
                languageVersion:=LanguageVersion.Latest,
                documentationMode:=DocumentationMode.Diagnose,
                kind:=SourceCodeKind.Regular,
                preprocessorSymbols:=PreprocessorSymbols)
        Dim syntaxTree As SyntaxTree = VisualBasicSyntaxTree.ParseText(text:=StringToBeCompiler, options:=ParseOptions)
        Dim assemblyName As String = Path.GetRandomFileName()



        Dim CompilationOptions As VisualBasicCompilationOptions = New VisualBasicCompilationOptions(
                                                                            outputKind:=OutputKind.DynamicallyLinkedLibrary,
                                                                            optionExplicit:=False,
                                                                            optionInfer:=True,
                                                                            optionStrict:=OptionStrict.Off,
                                                                            parseOptions:=ParseOptions
                                                                            )
        Dim compilation As VisualBasicCompilation = VisualBasicCompilation.Create(
                                                                    assemblyName:=assemblyName,
                                                                    syntaxTrees:={syntaxTree},
                                                                    references:=References,
                                                                    options:=CompilationOptions
                                                                                  )

        Dim CompileResult As EmitResult
        Using ms As MemoryStream = New MemoryStream()
            CompileResult = compilation.Emit(ms)
        End Using
        ResultOfConversion.FilteredListOfFailures = FilterDiagnostics(Diags:=compilation.GetParseDiagnostics(), Severity:=SeverityToReport, ErrorsToBeIgnored:=ErrorsToBeIgnored)
        Return CompileResult
    End Function
    Private Function FilterDiagnostics(Diags As ImmutableArray(Of Diagnostic), Severity As DiagnosticSeverity, ErrorsToBeIgnored As String()) As List(Of Diagnostic)
        Dim FilteredDiagnostics As New List(Of Diagnostic)
        For Each Diag As Diagnostic In Diags
            If Diag.Location.IsInSource = True AndAlso Diag.Severity >= Severity Then
                If ErrorsToBeIgnored.Contains(Diag.Id.ToString) Then
                    Continue For
                End If
                FilteredDiagnostics.Add(Diag)
            End If
        Next
        Return FilteredDiagnostics.OrderBy(Function(x As Diagnostic) x.Location.GetLineSpan.StartLinePosition.Line).ThenBy(Function(x As Diagnostic) x.Location.GetLineSpan.StartLinePosition.Character).ToList
    End Function

End Module
