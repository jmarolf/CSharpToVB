Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis
Imports Microsoft.VisualBasic.FileIO

Public Module SharedReferences
    Private ReadOnly CodeAnalysisReference As MetadataReference = MetadataReference.CreateFromFile(GetType(Compilation).Assembly.Location)
    Private ReadOnly ComponentModelEditorBrowsable As PortableExecutableReference = MetadataReference.CreateFromFile(GetType(System.ComponentModel.EditorBrowsableAttribute).GetAssemblyLocation())
    Private ReadOnly MSCorLibReference As MetadataReference = MetadataReference.CreateFromFile(GetType(Object).Assembly.Location)
    Private ReadOnly SystemAssembly As MetadataReference = MetadataReference.CreateFromFile(GetType(ComponentModel.BrowsableAttribute).Assembly.Location)
    Private ReadOnly SystemCore As MetadataReference = MetadataReference.CreateFromFile(GetType(Enumerable).Assembly.Location)
    Private ReadOnly SystemLinq As MetadataReference = MetadataReference.CreateFromFile(Path.Combine(SpecialDirectories.MyDocuments, "Visual Studio 2017\Projects\CSharpToVB\packages\System.Linq.4.3.0\lib\net463\System.Linq.dll"))
    Private ReadOnly SystemXmlLinq As MetadataReference = MetadataReference.CreateFromFile(GetType(XElement).Assembly.Location)
    Private ReadOnly VBPortable As PortableExecutableReference = MetadataReference.CreateFromFile("C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.1\Microsoft.VisualBasic.dll")
    Private ReadOnly VBRuntime As PortableExecutableReference = MetadataReference.CreateFromFile(GetType(CompilerServices.StandardModuleAttribute).Assembly.Location)
    Private ReadOnly XunitReferences As MetadataReference = MetadataReference.CreateFromFile(Path.Combine(SpecialDirectories.MyDocuments, "Visual Studio 2017\Projects\CSharpToVB\packages\xunit.extensibility.core.2.4.0\lib\netstandard2.0\xunit.core.dll"))

    Private _References As New List(Of MetadataReference)

    Private Sub BuildReferenceList()
        If _References?.Count > 0 Then
            Return
        End If
        Dim SystemReferences As New List(Of PortableExecutableReference)
        Const FrameworkDirectory As String = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\"
        For Each DLL As String In Directory.GetFiles(FrameworkDirectory, "*.dll")
            SystemReferences.Add(MetadataReference.CreateFromFile(DLL))
        Next
        _References = New List(Of MetadataReference)({
                                    CodeAnalysisReference,
                                    MSCorLibReference,
                                    SystemAssembly,
                                    SystemCore,
                                    SystemLinq,
                                    SystemXmlLinq,
                                    ComponentModelEditorBrowsable,
                                    VBPortable,
                                    VBRuntime,
                                    XunitReferences})
        _References.AddRange(SystemReferences)
        SystemReferences = Nothing
    End Sub

    <Extension>
    Public Function GetAssemblyLocation(ByVal type As Type) As String
        Dim asm As Assembly = type.GetTypeInfo().Assembly
        Dim locationProperty As PropertyInfo = asm.GetType().GetRuntimeProperties().Single(Function(p As PropertyInfo) p.Name = "Location")
        Return CStr(locationProperty.GetValue(asm))
    End Function
    Public Function References() As MetadataReference()
        If _References.Count = 0 Then
            BuildReferenceList()
        End If
        Return _References.ToArray()
    End Function

End Module