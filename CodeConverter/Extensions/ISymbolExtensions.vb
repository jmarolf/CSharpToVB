Option Explicit On
Option Infer Off
Option Strict On

Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices

Imports IVisualBasicCode.CodeConverter.Util

Imports Microsoft.CodeAnalysis

Namespace IVisualBasicCode.CodeConverter.Util

    ' All type argument must be accessible.

    Partial Public Module ISymbolExtensions

        Public Enum SymbolVisibility
            [Public]
            Internal
            [Private]
        End Enum

        <ExcludeFromCodeCoverage>
        <Extension>
        Private Function GetResultantVisibility(ByVal symbol As ISymbol) As SymbolVisibility
            ' Start by assuming it's visible.
            Dim visibility As SymbolVisibility = SymbolVisibility.Public

            Select Case symbol.Kind
                Case SymbolKind.Alias
                    ' Aliases are uber private.  They're only visible in the same file that they
                    ' were declared in.
                    Return SymbolVisibility.Private

                Case SymbolKind.Parameter
                    ' Parameters are only as visible as their containing symbol
                    Return GetResultantVisibility(symbol.ContainingSymbol)

                Case SymbolKind.TypeParameter
                    ' Type Parameters are private.
                    Return SymbolVisibility.Private
            End Select

            Do While symbol IsNot Nothing AndAlso symbol.Kind <> SymbolKind.Namespace
                Select Case symbol.DeclaredAccessibility
                    ' If we see anything private, then the symbol is private.
                    Case Accessibility.NotApplicable, Accessibility.Private
                        Return SymbolVisibility.Private

                    ' If we see anything internal, then knock it down from public to
                    ' internal.
                    Case Accessibility.Internal, Accessibility.ProtectedAndInternal
                        visibility = SymbolVisibility.Internal

                        ' For anything else (Public, Protected, ProtectedOrInternal), the
                        ' symbol stays at the level we've gotten so far.
                End Select

                symbol = symbol.ContainingSymbol
            Loop

            Return visibility
        End Function

        <ExcludeFromCodeCoverage>
        <Extension>
        Public Function GetReturnType(ByVal symbol As ISymbol) As ITypeSymbol
            If symbol Is Nothing Then
                Throw New ArgumentNullException(NameOf(symbol))
            End If
            Select Case symbol.Kind
                Case SymbolKind.Field
                    Dim field As IFieldSymbol = DirectCast(symbol, IFieldSymbol)
                    Return field.Type
                Case SymbolKind.Method
                    Dim method As IMethodSymbol = DirectCast(symbol, IMethodSymbol)
                    If method.MethodKind = MethodKind.Constructor Then
                        Return method.ContainingType
                    End If
                    Return method.ReturnType
                Case SymbolKind.Property
                    Dim [property] As IPropertySymbol = DirectCast(symbol, IPropertySymbol)
                    Return [property].Type
                Case SymbolKind.Event
                    Dim evt As IEventSymbol = DirectCast(symbol, IEventSymbol)
                    Return evt.Type
                Case SymbolKind.Parameter
                    Dim param As IParameterSymbol = DirectCast(symbol, IParameterSymbol)
                    Return param.Type
                Case SymbolKind.Local
                    Dim local As ILocalSymbol = DirectCast(symbol, ILocalSymbol)
                    Return local.Type
            End Select
            Return Nothing
        End Function

        <Extension()>
        Public Function IsInterfaceType(ByVal symbol As ISymbol) As Boolean
            If symbol Is Nothing OrElse TryCast(symbol, ITypeSymbol) Is Nothing Then
                Return False
            End If
            Return DirectCast(symbol, ITypeSymbol).IsInterfaceType() = True
        End Function

    End Module
End Namespace