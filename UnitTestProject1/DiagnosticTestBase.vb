Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Linq
Imports Microsoft.CodeAnalysis

Namespace CodeConverter.Tests
    Public Module DiagnosticTestBase
        Private ReadOnly CodeAnalysisReference As MetadataReference = MetadataReference.CreateFromFile(GetType(Compilation).Assembly.Location)
        Private ReadOnly MSCorLibReference As MetadataReference = MetadataReference.CreateFromFile(GetType(Object).Assembly.Location)
        Private ReadOnly SystemAssembly As MetadataReference = MetadataReference.CreateFromFile(GetType(ComponentModel.BrowsableAttribute).Assembly.Location)
        Private ReadOnly SystemCore As MetadataReference = MetadataReference.CreateFromFile(GetType(Enumerable).Assembly.Location)
        Private ReadOnly SystemCoreReference As MetadataReference = MetadataReference.CreateFromFile(GetType(Enumerable).Assembly.Location)
        Private ReadOnly SystemLinq As MetadataReference = MetadataReference.CreateFromFile("C:\Users\PaulM\Documents\Visual Studio 2017\Projects\CSharpToVB\packages\System.Linq.4.3.0\lib\net463\System.Linq.dll")
        Private ReadOnly SystemXmlLinq As MetadataReference = MetadataReference.CreateFromFile(GetType(XElement).Assembly.Location)
        Private ReadOnly VBConstants As MetadataReference = MetadataReference.CreateFromFile(GetType(Microsoft.VisualBasic.Constants).Assembly.Location)

        Private ReadOnly ComponentModelEditorBrowsable As PortableExecutableReference = MetadataReference.CreateFromFile(GetType(System.ComponentModel.EditorBrowsableAttribute).GetAssemblyLocation())
        Private ReadOnly EnumerableReference As PortableExecutableReference = MetadataReference.CreateFromFile(GetType(Enumerable).GetAssemblyLocation())
        Private ReadOnly VBPortable As PortableExecutableReference = MetadataReference.CreateFromFile("C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.1\Microsoft.VisualBasic.dll")
        Private ReadOnly VBRuntime As PortableExecutableReference = MetadataReference.CreateFromFile(GetType(CompilerServices.StandardModuleAttribute).Assembly.Location)
        Private ReadOnly XunitReferences As MetadataReference = MetadataReference.CreateFromFile("C:\Users\PaulM\Documents\Visual Studio 2017\Projects\CSharpToVB\packages\xunit.extensibility.core.2.3.1\lib\netstandard1.1\xunit.core.dll")

        Friend Function DefaultMetadataReferences() As List(Of MetadataReference)
            Dim _References As New List(Of MetadataReference)({
                                    CodeAnalysisReference,
                                    MSCorLibReference,
                                    SystemAssembly,
                                    SystemCore,
                                    SystemCoreReference,
                                    SystemLinq,
                                    SystemXmlLinq,
                                    VBConstants,
                                    ComponentModelEditorBrowsable,
                                    EnumerableReference,
                                    VBPortable,
                                    VBRuntime,
                                    XunitReferences})
            _References.AddRange(AppDomain.CurrentDomain.GetAssemblies().Where(Function(x As Reflection.Assembly) x.GetName().Name.EndsWith(".dll")).Select(Function(a As Reflection.Assembly) MetadataReference.CreateFromFile(a.Location)))

            Return _References

        End Function
    End Module
End Namespace
