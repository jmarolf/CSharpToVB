Public Class SymbolTableEntry
    Public Sub New(_Name As String, _IsType As Boolean)
        Name = _Name
        IsType = _IsType
    End Sub

    Public ReadOnly Property IsType As Boolean
    Public ReadOnly Property Name As String
End Class
