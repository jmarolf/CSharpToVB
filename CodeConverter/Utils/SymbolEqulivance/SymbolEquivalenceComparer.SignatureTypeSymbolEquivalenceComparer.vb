Imports System.Diagnostics.CodeAnalysis

Imports Microsoft.CodeAnalysis

Partial Friend Class SymbolEquivalenceComparer

    <ExcludeFromCodeCoverage>
    Friend Class SignatureTypeSymbolEquivalenceComparer
        Implements IEqualityComparer(Of ITypeSymbol)

        Private ReadOnly _symbolEquivalenceComparer As SymbolEquivalenceComparer

        <ExcludeFromCodeCoverage>
        Public Sub New(ByVal symbolEquivalenceComparer As SymbolEquivalenceComparer)
            _symbolEquivalenceComparer = symbolEquivalenceComparer
        End Sub

        <ExcludeFromCodeCoverage>
        Public Shadows Function Equals(ByVal x As ITypeSymbol, ByVal y As ITypeSymbol) As Boolean Implements IEqualityComparer(Of ITypeSymbol).Equals
            Return Equals(x, y, Nothing)
        End Function

        <ExcludeFromCodeCoverage>
        Public Shadows Function Equals(ByVal x As ITypeSymbol, ByVal y As ITypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
            Return _symbolEquivalenceComparer.GetEquivalenceVisitor(compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=True).AreEquivalent(x, y, equivalentTypesWithDifferingAssemblies)
        End Function

        <ExcludeFromCodeCoverage>
        Public Shadows Function GetHashCode(ByVal x As ITypeSymbol) As Integer Implements IEqualityComparer(Of ITypeSymbol).GetHashCode
            Return _symbolEquivalenceComparer.GetGetHashCodeVisitor(compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=True).GetHashCode(x, currentHash:=0)
        End Function

    End Class

End Class