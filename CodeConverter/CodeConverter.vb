Option Explicit On
Option Infer Off
Option Strict On

Imports IVisualBasicCode.CodeConverter.VB
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit

Namespace IVisualBasicCode.CodeConverter
    Public Module CodeConverter
        Private Function GetDefaultVersionForLanguage(ByVal language As String) As Integer
            If language Is Nothing Then
                Throw New ArgumentNullException(NameOf(language))
            End If
            If language.StartsWith("cs", StringComparison.OrdinalIgnoreCase) Then
                Return CSharp.LanguageVersion.Latest
            End If
            If language.StartsWith("vb", StringComparison.OrdinalIgnoreCase) Then
                Return VisualBasic.LanguageVersion.Latest
            End If
            Throw New ArgumentException($"{language} not supported!")
        End Function
        Private Function IsSupportedSource(ByVal fromLanguage As String, ByVal fromLanguageVersion As Integer) As Boolean
            Return (fromLanguage = LanguageNames.CSharp AndAlso fromLanguageVersion <= CSharp.LanguageVersion.Latest) OrElse
                (fromLanguage = LanguageNames.VisualBasic AndAlso fromLanguageVersion <= VisualBasic.LanguageVersion.Latest)
        End Function
        Private Function IsSupportedTarget(ByVal toLanguage As String, ByVal toLanguageVersion As Integer) As Boolean
            Return (toLanguage = LanguageNames.VisualBasic AndAlso toLanguageVersion <= VisualBasic.LanguageVersion.Latest) OrElse
                (toLanguage = LanguageNames.CSharp AndAlso toLanguageVersion <= CSharp.LanguageVersion.Latest)
        End Function
        Private Function ParseLanguage(ByVal language As String) As String
            If language Is Nothing Then
                Throw New ArgumentNullException(NameOf(language))
            End If
            If language.StartsWith("cs", StringComparison.OrdinalIgnoreCase) Then
                Return LanguageNames.CSharp
            End If
            If language.StartsWith("vb", StringComparison.OrdinalIgnoreCase) Then
                Return LanguageNames.VisualBasic
            End If
            Throw New ArgumentException($"{language} not supported!")
        End Function
        Public Function Convert(ByVal code As CodeWithOptions) As ConversionResult

            Select Case code.FromLanguage
                Case LanguageNames.CSharp
                    Select Case code.ToLanguage
                        Case LanguageNames.VisualBasic
                            Return CSharpConverter.ConvertText(code.Text, References)
                    End Select
                Case LanguageNames.VisualBasic
                    Select Case code.ToLanguage
                        Case LanguageNames.CSharp
#If VB_TO_CSharp = "True" Then
                            Return VisualBasicConverter.ConvertText(code.Text, code.References)
#Else
                            Throw New NotImplementedException("VB to C# is disabled")
#End If
                    End Select
            End Select
            If Not IsSupportedSource(code.FromLanguage, code.FromLanguageVersion) Then
                Return New ConversionResult(New NotSupportedException($"Source language {code.FromLanguage} {code.FromLanguageVersion} is not supported!"))
            End If
            If Not IsSupportedTarget(code.ToLanguage, code.ToLanguageVersion) Then
                Return New ConversionResult(New NotSupportedException($"Target language {code.ToLanguage} {code.ToLanguageVersion} is not supported!"))
            End If
            If code.FromLanguage = code.ToLanguage AndAlso code.FromLanguageVersion <> code.ToLanguageVersion Then
                Return New ConversionResult(New NotSupportedException($"Converting from {code.FromLanguage} {code.FromLanguageVersion} to {code.ToLanguage} {code.ToLanguageVersion} is not supported!"))
            End If
            Return New ConversionResult(New NotSupportedException($"Converting from {code.FromLanguage} {code.FromLanguageVersion} to {code.ToLanguage} {code.ToLanguageVersion} is not supported!"))
        End Function
        Public Function ConvertInputRequest(RequestToConvert As ConvertRequest) As ConversionResult
            Dim languages As String() = RequestToConvert.RequestedConversion.Split("2"c)
            Dim fromLanguage As String = LanguageNames.CSharp
            Dim toLanguage As String = LanguageNames.VisualBasic
            Dim fromVersion As Integer = CSharp.LanguageVersion.Latest
            Dim toVersion As Integer = VisualBasic.LanguageVersion.Latest
            If languages.Length = 2 Then
                fromLanguage = ParseLanguage(languages(0))
                fromVersion = GetDefaultVersionForLanguage(languages(0))
                toLanguage = ParseLanguage(languages(1))
                toVersion = GetDefaultVersionForLanguage(languages(1))
            End If
            Dim codeWithOptions As CodeWithOptions = (New CodeWithOptions(RequestToConvert.SourceCode)).SetFromLanguage(fromLanguage, fromVersion).SetToLanguage(toLanguage, toVersion)
            Dim ResultOfConversion As ConversionResult = CodeConverter.Convert(codeWithOptions)
            Return ResultOfConversion
        End Function
    End Module
End Namespace
