﻿Option Explicit On
Option Infer Off
Option Strict On

Imports IVisualBasicCode.CodeConverter.VB

Namespace IVisualBasicCode.CodeConverter.Util
    Friend Module NameGenerator

        Private Sub EnsureUniquenessInPlace(ByVal names As IList(Of String), ByVal isFixed As IList(Of Boolean), ByVal canUse As Func(Of String, Boolean), Optional ByVal isCaseSensitive As Boolean = True)
            canUse = If(canUse, (Function(s As String) True))

            ' Don't enumerate as we will be modifying the collection in place.
            Dim i As Integer = 0
            Do While i < names.Count
                Dim name As String = names(i)
                Dim collisionIndices As List(Of Integer) = GetCollisionIndices(names, name, isCaseSensitive)

                If canUse(name) AndAlso collisionIndices.Count < 2 Then
                    ' no problems with this parameter name, move onto the next one.
                    i += 1
                    Continue Do
                End If

                HandleCollisions(isFixed, names, name, collisionIndices, canUse, isCaseSensitive)
                i += 1
            Loop
        End Sub

        Private Function GetCollisionIndices(ByVal names As IList(Of String), ByVal name As String, Optional ByVal isCaseSensitive As Boolean = True) As List(Of Integer)
            Dim comparer As StringComparer = If(isCaseSensitive, StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase)
            Dim collisionIndices As List(Of Integer) = names.Select(Function(currentName As String, index As Integer) New With {Key currentName, Key index}).Where(Function(t) comparer.Equals(t.currentName, name)).Select(Function(t) t.index).ToList()
            Return collisionIndices
        End Function

        Private Sub HandleCollisions(ByVal isFixed As IList(Of Boolean), ByVal names As IList(Of String), ByVal name As String, ByVal collisionIndices As List(Of Integer), ByVal canUse As Func(Of String, Boolean), Optional ByVal isCaseSensitive As Boolean = True)
            Dim suffix As Integer = 1
            Dim comparer As StringComparer = If(isCaseSensitive, StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase)
            For i As Integer = 0 To collisionIndices.Count - 1
                Dim collisionIndex As Integer = collisionIndices(i)
                If isFixed(collisionIndex) Then
                    ' can't do anything about this name.
                    Continue For
                End If

                Do
                    Dim newName As String = name & suffix
                    suffix += 1
                    If Not names.Contains(newName, comparer) AndAlso canUse(newName) Then
                        ' Found a name that doesn't conflict with anything else.
                        names(collisionIndex) = newName
                        Exit Do
                    End If
                Loop
            Next i
        End Sub

        ''' <summary>
        ''' Ensures that any 'names' is unique and does not collide with any other name.  Names that
        ''' are marked as IsFixed can not be touched.  This does mean that if there are two names
        ''' that are the same, and both are fixed that you will end up with non-unique names at the
        ''' end.
        ''' </summary>
        Public Function EnsureUniqueness(ByVal names As IList(Of String), ByVal isFixed As IList(Of Boolean), Optional ByVal canUse As Func(Of String, Boolean) = Nothing, Optional ByVal isCaseSensitive As Boolean = True) As IList(Of String)
            Dim copy As List(Of String) = names.ToList()
            EnsureUniquenessInPlace(copy, isFixed, canUse, isCaseSensitive)
            Return copy
        End Function

        ''' <summary>
        ''' Transforms baseName into a name that does not conflict with any name in 'reservedNames'
        ''' </summary>
        Public Function EnsureUniqueness(ByVal baseName As String, ByVal reservedNames As IEnumerable(Of String), Optional ByVal isCaseSensitive As Boolean = True) As String
            Dim names As List(Of String) = New List(Of String) From {baseName}
            Dim isFixed As List(Of Boolean) = New List(Of Boolean) From {False}
            For Each s As SymbolTableEntry In CSharpConverter.UsedIdentifiers.Values
                names.Add(s.Name)
            Next

            names.AddRange(reservedNames.Distinct())
            isFixed.AddRange(Enumerable.Repeat(True, names.Count - 1))

            Dim result As IList(Of String) = EnsureUniqueness(names, isFixed, isCaseSensitive:=isCaseSensitive)
            Return result.First()
        End Function

    End Module
End Namespace