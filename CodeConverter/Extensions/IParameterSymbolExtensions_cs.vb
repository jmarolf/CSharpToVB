' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'Namespace Microsoft.CodeAnalysis.Shared.Extensions
Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

Public Module IParameterSymbolExtensions
    <Extension>
    <ExcludeFromCodeCoverage>
    Public Function IsRefOrOut(ByVal symbol As IParameterSymbol) As Boolean
        Return symbol.RefKind <> RefKind.None
    End Function
End Module
'End Namespace
