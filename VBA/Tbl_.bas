Attribute VB_Name = "Tbl_"
Option Explicit
Option Private Module

' =====================================================================
'   Table Constant
' =====================================================================

'   Build a 2D Array from a Table Constant String
'
Public Function Tbl_TableConst( _
    ByVal TableConst As String _
) As String()

    Dim Rows() As String
    Rows = Split(TableConst, vbLf)
    
    Dim Dense() As String
    ReDim Dense(0 To UBound(Rows))
    Dim DenseIndex As Long
    
    Dim RowsIndex As Long
    For RowsIndex = 0 To UBound(Rows)
        If Left(Rows(RowsIndex), 1) <> "'" Then
            Dense(DenseIndex) = Rows(RowsIndex)
            DenseIndex = DenseIndex + 1
        End If
    Next RowsIndex
    If DenseIndex < 1 Then DenseIndex = 1
    ReDim Preserve Dense(0 To DenseIndex - 1)
    
    Dim Cols() As String
    Cols = Split(Dense(0), "|")
    Dim TableArray() As String
    ReDim TableArray(0 To UBound(Dense), 0 To UBound(Cols))
    
    For DenseIndex = 0 To UBound(Dense)
        Cols = Split(Dense(DenseIndex), "|")
        Dim ColsIndex As Long
        For ColsIndex = 0 To UBound(Cols)
            TableArray(DenseIndex, ColsIndex) = Trim(Cols(ColsIndex))
        Next ColsIndex
    Next DenseIndex
    
    Tbl_TableConst = TableArray()

End Function

'   Build a 1D Array from a Table Constant String and a Column Index
'
Public Function Tbl_TableConstList( _
    ByVal TableConst As String, _
    Optional ByVal ColIndex As Long = 0, _
    Optional ByVal StartRow As Long = 1 _
) As String()

    Dim TableArray() As String
    TableArray = Tbl_TableConst(TableConst)
    
    Dim ListArray() As String
    ReDim ListArray(0 To UBound(TableArray, 1) - StartRow)
    
    Dim RowsIndex As Long
    For RowsIndex = 0 To UBound(ListArray)
        ListArray(RowsIndex) = TableArray(RowsIndex + StartRow, ColIndex)
    Next RowsIndex
    
    Tbl_TableConstList = ListArray()

End Function

'   Build a 1D Array of Column Values from a Table Constant String
'   for the first matching RowKey in KeyCol, starting at StartRow.
'
'   False   <-  If RowKey not found
'
Public Function Tbl_TableConstRow( _
    ByVal TableConst As String, _
    ByVal RowKey As String, _
    ByRef Cols() As String, _
    Optional ByVal KeyCol As Long = 0, _
    Optional ByVal StartRow As Long = 1 _
) As Boolean
Tbl_TableConstRow = False

    Dim TableArray() As String
    TableArray = Tbl_TableConst(TableConst)
    
    Dim Found As Boolean
    Dim RowIndex As Long
    For RowIndex = StartRow To UBound(TableArray, 1)
        If StrComp(TableArray(RowIndex, KeyCol), RowKey, vbTextCompare) = 0 Then
            Found = True
            Exit For
        End If
    Next RowIndex
    If Not Found Then Exit Function
    
    ReDim Cols(0 To UBound(TableArray, 2))
    Dim ColsIndex As Long
    For ColsIndex = 0 To UBound(Cols)
        Cols(ColsIndex) = TableArray(RowIndex, ColsIndex)
    Next ColsIndex

Tbl_TableConstRow = True
End Function

'   Get a Value from ColIndex in a Table Constant String
'   for the first matching RowKey in KeyCol, starting at StartRow.
'
'   False   <-  If RowKey not found
'
Public Function Tbl_TableConstFind( _
    ByVal TableConst As String, _
    ByVal RowKey As String, _
    ByVal ColIndex As Long, _
    ByRef value As String, _
    Optional ByVal KeyCol As Long = 0, _
    Optional ByVal StartRow As Long = 1 _
) As Boolean
Tbl_TableConstFind = False

    Dim TableArray() As String
    TableArray = Tbl_TableConst(TableConst)
    
    Dim RowIndex As Long
    For RowIndex = StartRow To UBound(TableArray, 1)
        If StrComp(TableArray(RowIndex, KeyCol), RowKey, vbTextCompare) = 0 Then
            value = TableArray(RowIndex, ColIndex)
            Tbl_TableConstFind = True
            Exit Function
        End If
    Next RowIndex
    value = ""

End Function

'   Does RowKey exist in a Table Constant String
'   In KeyCol, starting at StartRow?
'
'   False   <-  If RowKey not found
'
Public Function Tbl_TableConstExist( _
    ByVal TableConst As String, _
    ByVal RowKey As String, _
    Optional ByVal KeyCol As Long = 0, _
    Optional ByVal StartRow As Long = 1 _
) As Boolean

    Dim value As String
    Tbl_TableConstExist = Tbl_TableConstFind(TableConst, RowKey, 0, value, KeyCol, StartRow)

End Function

'   Get a ColIndex from a Table Constant Header Row (Row 0)
'   for the first matching ColKey in the Header Columns
'
'   False   <-  If ColKey not found
'
Public Function Tbl_TableConstHeaderCol( _
    ByVal TableConst As String, _
    ByVal ColKey As String, _
    ByRef ColIndex As Long _
) As Boolean

    Dim TableArray() As String
    TableArray = Tbl_TableConst(TableConst)
    
    For ColIndex = 0 To UBound(TableArray, 2)
    
        If StrComp(TableArray(0, ColIndex), ColKey, vbTextCompare) = 0 Then
            Tbl_TableConstHeaderCol = True
            Exit Function
        End If
    
    Next ColIndex

End Function


