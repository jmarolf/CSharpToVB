﻿Imports System.Collections.Immutable
Imports System.Diagnostics.CodeAnalysis

Imports Microsoft.CodeAnalysis

''' <summary>
''' Provides a way to test two symbols for equivalence.  While there are ways to ask for
''' different sorts of equivalence, the following must hold for two symbols to be considered
''' equivalent.
'''
''' 1) The kinds of the two symbols must match.
'''
''' 2) The names of the two symbols must match.
'''
''' 3) The arity of the two symbols must match.
'''
''' 4) If the symbols are methods or parameterized properties, then the signatures of the two
''' symbols must match.
'''
''' 5) Both symbols must be definitions or must be instantiations.  If they are instantiations,
''' then they must be instantiated in the same manner.
'''
''' 6) The containing symbols of the two symbols must be equivalent.
'''
''' Note: equivalence does not concern itself with whole symbols.  Two types are considered
''' equivalent if the above hold, even if one type has different members than the other.  Note:
''' type parameters, and signature parameters are not considered 'children' when comparing
''' symbols.
'''
''' Options are provided to tweak the above slightly.  For example, by default, symbols are
''' equivalent only if they come from the same assembly or different assemblies of the same simple name.
''' However, one can ask if two symbols are equivalent even if their assemblies differ.
''' </summary>
Partial Friend Class SymbolEquivalenceComparer
    Implements IEqualityComparer(Of ISymbol)

    Private ReadOnly _assemblyComparerOpt As IEqualityComparer(Of IAssemblySymbol)
    Private ReadOnly _equivalenceVisitors As ImmutableArray(Of EquivalenceVisitor)
    Private ReadOnly _getHashCodeVisitors As ImmutableArray(Of GetHashCodeVisitor)

    Public Shared ReadOnly IgnoreAssembliesInstance As New SymbolEquivalenceComparer(assemblyComparerOpt:=Nothing, distinguishRefFromOut:=False)
    Public Shared ReadOnly Instance As New SymbolEquivalenceComparer(SimpleNameAssemblyComparer.Instance, distinguishRefFromOut:=False)
    <ExcludeFromCodeCoverage>
    Friend Sub New(ByVal assemblyComparerOpt As IEqualityComparer(Of IAssemblySymbol), ByVal distinguishRefFromOut As Boolean)
        _assemblyComparerOpt = assemblyComparerOpt

        ParameterEquivalenceComparer = New ParameterSymbolEqualityComparer(Me, distinguishRefFromOut)
        SignatureTypeEquivalenceComparer = New SignatureTypeSymbolEquivalenceComparer(Me)

        ' There are only so many EquivalenceVisitors and GetHashCodeVisitors we can have.
        ' Create them all up front.
        Dim equivalenceVisitorsBuilder As ImmutableArray(Of EquivalenceVisitor).Builder = ImmutableArray.CreateBuilder(Of EquivalenceVisitor)()
        equivalenceVisitorsBuilder.Add(New EquivalenceVisitor(Me, compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=True))
        equivalenceVisitorsBuilder.Add(New EquivalenceVisitor(Me, compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=False))
        equivalenceVisitorsBuilder.Add(New EquivalenceVisitor(Me, compareMethodTypeParametersByIndex:=False, objectAndDynamicCompareEqually:=True))
        equivalenceVisitorsBuilder.Add(New EquivalenceVisitor(Me, compareMethodTypeParametersByIndex:=False, objectAndDynamicCompareEqually:=False))
        _equivalenceVisitors = equivalenceVisitorsBuilder.ToImmutable()

        Dim getHashCodeVisitorsBuilder As ImmutableArray(Of GetHashCodeVisitor).Builder = ImmutableArray.CreateBuilder(Of GetHashCodeVisitor)()
        getHashCodeVisitorsBuilder.Add(New GetHashCodeVisitor(Me, compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=True))
        getHashCodeVisitorsBuilder.Add(New GetHashCodeVisitor(Me, compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=False))
        getHashCodeVisitorsBuilder.Add(New GetHashCodeVisitor(Me, compareMethodTypeParametersByIndex:=False, objectAndDynamicCompareEqually:=True))
        getHashCodeVisitorsBuilder.Add(New GetHashCodeVisitor(Me, compareMethodTypeParametersByIndex:=False, objectAndDynamicCompareEqually:=False))
        _getHashCodeVisitors = getHashCodeVisitorsBuilder.ToImmutable()
    End Sub

    Public ReadOnly Property ParameterEquivalenceComparer() As ParameterSymbolEqualityComparer
    Public ReadOnly Property SignatureTypeEquivalenceComparer() As SignatureTypeSymbolEquivalenceComparer
    Private Shared Function CheckContainingType(ByVal x As IMethodSymbol) As Boolean
        If x.MethodKind = MethodKind.DelegateInvoke AndAlso x.ContainingType IsNot Nothing AndAlso x.ContainingType.IsAnonymousType Then
            Return False
        End If

        Return True
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function GetKindAndUnwrapAlias(ByRef symbol As ISymbol) As SymbolKind
        Dim k As SymbolKind = symbol.Kind
        If k = SymbolKind.Alias Then
            symbol = DirectCast(symbol, IAliasSymbol).Target
            k = symbol.Kind
        End If

        Return k
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function GetTypeKind(ByVal x As INamedTypeSymbol) As TypeKind
        ' Treat static classes as modules.
        Dim k As TypeKind = x.TypeKind
        Return If(k = TypeKind.Module, TypeKind.Class, k)
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function GetVisitorIndex(ByVal compareMethodTypeParametersByIndex As Boolean, ByVal objectAndDynamicCompareEqually As Boolean) As Integer
        Return If(compareMethodTypeParametersByIndex, If(objectAndDynamicCompareEqually, 0, 1), If(objectAndDynamicCompareEqually, 2, 3))
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function IsConstructedFromSelf(ByVal symbol As INamedTypeSymbol) As Boolean
        Return DirectCast(symbol, Object).Equals(symbol.ConstructedFrom)
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function IsConstructedFromSelf(ByVal symbol As IMethodSymbol) As Boolean
        Return DirectCast(symbol, Object).Equals(symbol.ConstructedFrom)
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function IsObjectType(ByVal symbol As ISymbol) As Boolean
        Return symbol.Kind = SymbolKind.NamedType AndAlso DirectCast(symbol, ITypeSymbol).SpecialType = SpecialType.System_Object
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function IsPartialMethodDefinitionPart(ByVal symbol As IMethodSymbol) As Boolean
        Return symbol.PartialImplementationPart IsNot Nothing
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function IsPartialMethodImplementationPart(ByVal symbol As IMethodSymbol) As Boolean
        Return symbol.PartialDefinitionPart IsNot Nothing
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Iterator Function Unwrap(ByVal namedType As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
        Yield namedType

        Dim errorType As IErrorTypeSymbol = TryCast(namedType, IErrorTypeSymbol)

        If errorType IsNot Nothing Then
            For Each [type] As INamedTypeSymbol In errorType.CandidateSymbols.OfType(Of INamedTypeSymbol)()
                Yield [type]
            Next type
        End If
    End Function

    <ExcludeFromCodeCoverage>
    Private Shared Function UnwrapAlias(ByVal symbol As ISymbol) As ISymbol
        Return If(symbol.Kind = Global.Microsoft.CodeAnalysis.SymbolKind.Alias, DirectCast(symbol, IAliasSymbol).Target, symbol)
    End Function

    <ExcludeFromCodeCoverage>
    Private Function EqualsCore(ByVal x As ISymbol, ByVal y As ISymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
        Return GetEquivalenceVisitor().AreEquivalent(x, y, equivalentTypesWithDifferingAssemblies:=equivalentTypesWithDifferingAssemblies)
    End Function

    ' Very subtle logic here.  When checking if two parameters are the same, we can end up with
    ' a tricky infinite loop.  Specifically, consider the case if the parameter refers to a
    ' method type parameter.  i.e. "void Goo<T>(IList<T> arg)".  If we compare two method type
    ' parameters for equality, then we'll end up asking if their methods are the same.  And that
    ' will cause us to check if their parameters are the same.  And then we'll be right back
    ' here.  So, instead, when asking if parameters are equal, we pass an appropriate flag so
    ' that method type parameters are just compared by index and nothing else.
    <ExcludeFromCodeCoverage>
    Private Function GetEquivalenceVisitor(Optional ByVal compareMethodTypeParametersByIndex As Boolean = False, Optional ByVal objectAndDynamicCompareEqually As Boolean = False) As EquivalenceVisitor
        Dim visitorIndex As Integer = GetVisitorIndex(compareMethodTypeParametersByIndex, objectAndDynamicCompareEqually)
        Return _equivalenceVisitors(visitorIndex)
    End Function

    <ExcludeFromCodeCoverage>
    Private Function GetGetHashCodeVisitor(ByVal compareMethodTypeParametersByIndex As Boolean, ByVal objectAndDynamicCompareEqually As Boolean) As GetHashCodeVisitor
        Dim visitorIndex As Integer = GetVisitorIndex(compareMethodTypeParametersByIndex, objectAndDynamicCompareEqually)
        Return _getHashCodeVisitors(visitorIndex)
    End Function
    ''' <summary>
    ''' Compares given symbols <paramref name="x"/> and <paramref name="y"/> for equivalence.
    ''' </summary>
    <ExcludeFromCodeCoverage>
    Public Shadows Function Equals(ByVal x As ISymbol, ByVal y As ISymbol) As Boolean Implements IEqualityComparer(Of ISymbol).Equals
        Return EqualsCore(x, y, equivalentTypesWithDifferingAssemblies:=Nothing)
    End Function

    ''' <summary>
    ''' Compares given symbols <paramref name="x"/> and <paramref name="y"/> for equivalence and populates <paramref name="equivalentTypesWithDifferingAssemblies"/>
    ''' with equivalent non-nested named type key-value pairs that are contained in different assemblies.
    ''' These equivalent named type key-value pairs represent possibly equivalent forwarded types, but this API doesn't perform any type forwarding equivalence checks.
    ''' </summary>
    ''' <remarks>This API is only supported for <see cref="SymbolEquivalenceComparer.IgnoreAssembliesInstance"/>.</remarks>
    <ExcludeFromCodeCoverage>
    Public Shadows Function Equals(ByVal x As ISymbol, ByVal y As ISymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
        Debug.Assert(_assemblyComparerOpt Is Nothing)
        Return EqualsCore(x, y, equivalentTypesWithDifferingAssemblies:=equivalentTypesWithDifferingAssemblies)
    End Function

    <ExcludeFromCodeCoverage>
    Public Shadows Function GetHashCode(ByVal x As ISymbol) As Integer Implements IEqualityComparer(Of ISymbol).GetHashCode
        Return GetGetHashCodeVisitor(compareMethodTypeParametersByIndex:=False, objectAndDynamicCompareEqually:=False).GetHashCode(x, currentHash:=0)
    End Function

    <ExcludeFromCodeCoverage>
    Public Function ReturnTypeEquals(ByVal x As IMethodSymbol, ByVal y As IMethodSymbol, Optional ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol) = Nothing) As Boolean
        Return GetEquivalenceVisitor().ReturnTypesAreEquivalent(x, y, equivalentTypesWithDifferingAssemblies)
    End Function
End Class