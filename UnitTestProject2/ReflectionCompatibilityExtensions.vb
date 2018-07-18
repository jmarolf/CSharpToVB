Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.CompilerServices

Namespace IVisualBasicCode.CodeConverter.Util
    <Flags>
    Friend Enum BindingFlags
        [Default] = 0
        Instance = 1
        [Static] = 2
        [Public] = 4
        NonPublic = 8
    End Enum

    Friend Module ReflectionCompatibilityExtensions
        <Extension>
        Public Function GetCustomAttributes(ByVal type As Type, ByVal inherit As Boolean) As Object()
            Return type.GetTypeInfo().GetCustomAttributes(inherit).ToArray()
        End Function

        <Extension>
        Public Function GetMethod(ByVal type As Type, ByVal name As String) As MethodInfo
            Return type.GetTypeInfo().DeclaredMethods.FirstOrDefault(Function(m As MethodInfo) m.Name = name)
        End Function

        <Extension>
        Public Function GetMethod(ByVal type As Type, ByVal name As String, ByVal bindingFlags As BindingFlags) As MethodInfo
            Return type.GetTypeInfo().DeclaredMethods.FirstOrDefault(Function(m As MethodInfo) (m.Name = name) AndAlso IsConformWithBindingFlags(m, bindingFlags))
        End Function

        '<Extension>
        'Public Function GetMethod(ByVal type As Type, ByVal name As String, ByVal types() As Type) As MethodInfo
        '    Return type.GetTypeInfo().DeclaredMethods.FirstOrDefault(Function(m As MethodInfo) (m.Name = name) AndAlso TypesAreEqual(m.GetParameters().Select(Function(p As ParameterInfo) p.ParameterType).ToArray(), types))
        'End Function

        '<Extension>
        'Public Function GetMethod(ByVal type As Type, ByVal name As String, ByVal bindingFlags As BindingFlags, ByVal binder As Object, ByVal types() As Type, ByVal modifiers As Object) As MethodInfo
        '    Return type.GetTypeInfo().DeclaredMethods.FirstOrDefault(Function(m As MethodInfo) (m.Name = name) AndAlso IsConformWithBindingFlags(m, bindingFlags) AndAlso TypesAreEqual(m.GetParameters().Select(Function(p As ParameterInfo) p.ParameterType).ToArray(), types))
        'End Function

        <Extension>
        Public Function GetMethods(ByVal type As Type) As IEnumerable(Of MethodInfo)
            Return type.GetTypeInfo().DeclaredMethods
        End Function

        <Extension>
        Public Function GetMethods(ByVal type As Type, ByVal name As String) As IEnumerable(Of MethodInfo)
            Return type.GetTypeInfo().DeclaredMethods.Where(Function(m As MethodInfo) m.Name = name)
        End Function

        <Extension>
        Public Function GetMethods(ByVal type As Type, ByVal bindingFlags As BindingFlags) As IEnumerable(Of MethodInfo)
            Return type.GetTypeInfo().DeclaredMethods.Where(Function(m As MethodInfo) IsConformWithBindingFlags(m, bindingFlags))
        End Function

        <Extension>
        Public Function GetField(ByVal type As Type, ByVal name As String) As FieldInfo
            Return type.GetTypeInfo().DeclaredFields.FirstOrDefault(Function(f As FieldInfo) f.Name = name)
        End Function

        <Extension>
        Public Function GetField(ByVal type As Type, ByVal name As String, ByVal bindingFlags As BindingFlags) As FieldInfo
            Return type.GetTypeInfo().DeclaredFields.FirstOrDefault(Function(f As FieldInfo) (f.Name = name) AndAlso IsConformWithBindingFlags(f, bindingFlags))
        End Function

        <Extension>
        Public Function GetProperty(ByVal type As Type, ByVal name As String) As PropertyInfo
            Return type.GetTypeInfo().DeclaredProperties.FirstOrDefault(Function(p As PropertyInfo) p.Name = name)
        End Function

        <Extension>
        Public Function GetProperties(ByVal type As Type) As IEnumerable(Of PropertyInfo)
            Return type.GetTypeInfo().DeclaredProperties
        End Function

        <Extension>
        Public Function GetAssemblyLocation(ByVal type As Type) As String
            Dim asm As Assembly = type.GetTypeInfo().Assembly
            Dim locationProperty As PropertyInfo = asm.GetType().GetRuntimeProperties().Single(Function(p As PropertyInfo) p.Name = "Location")
            Return CStr(locationProperty.GetValue(asm))
        End Function

        Private Function TypesAreEqual(ByVal memberTypes() As Type, ByVal searchedTypes() As Type) As Boolean
            If ((memberTypes Is Nothing) OrElse (searchedTypes Is Nothing)) AndAlso (memberTypes IsNot searchedTypes) Then
                Return False
            End If

            If memberTypes.Length <> searchedTypes.Length Then
                Return False
            End If

            For i As Integer = 0 To memberTypes.Length - 1
                If memberTypes(i) IsNot searchedTypes(i) Then
                    Return False
                End If
            Next i

            Return True
        End Function

        Private Function IsConformWithBindingFlags(ByVal method As MethodBase, ByVal bindingFlags As BindingFlags) As Boolean
            If method.IsPublic AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.Public) Then
                Return False
            End If
            If method.IsPrivate AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.NonPublic) Then
                Return False
            End If
            If method.IsStatic AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.Static) Then
                Return False
            End If
            If Not method.IsStatic AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.Instance) Then
                Return False
            End If

            Return True
        End Function

        Private Function IsConformWithBindingFlags(ByVal method As FieldInfo, ByVal bindingFlags As BindingFlags) As Boolean
            If method.IsPublic AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.Public) Then
                Return False
            End If
            If method.IsPrivate AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.NonPublic) Then
                Return False
            End If
            If method.IsStatic AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.Static) Then
                Return False
            End If
            If Not method.IsStatic AndAlso Not bindingFlags.HasFlag(IVisualBasicCode.CodeConverter.Util.BindingFlags.Instance) Then
                Return False
            End If

            Return True
        End Function
    End Module
End Namespace
