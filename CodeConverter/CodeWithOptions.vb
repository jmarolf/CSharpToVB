Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.CodeAnalysis

Namespace IVisualBasicCode.CodeConverter
    Public Class CodeWithOptions

        Public Sub New(ByVal text As String)
            Me.Text = text
            FromLanguage = LanguageNames.CSharp
            ToLanguage = LanguageNames.VisualBasic
            FromLanguageVersion = CSharp.LanguageVersion.Latest
            ToLanguageVersion = VisualBasic.LanguageVersion.Latest
        End Sub

        Public Property FromLanguage() As String
        Public Property FromLanguageVersion() As Integer
        Public Property Text() As String
        Public Property ToLanguage() As String
        Public Property ToLanguageVersion() As Integer
        Public Function SetFromLanguage(Optional ByVal name As String = LanguageNames.CSharp, Optional ByVal version As Integer = CSharp.LanguageVersion.Latest) As CodeWithOptions
            FromLanguage = name
            FromLanguageVersion = version
            Return Me
        End Function

        Public Function SetToLanguage(Optional ByVal name As String = LanguageNames.VisualBasic, Optional ByVal version As Integer = VisualBasic.LanguageVersion.Latest) As CodeWithOptions
            ToLanguage = name
            ToLanguageVersion = version
            Return Me
        End Function
    End Class
End Namespace
