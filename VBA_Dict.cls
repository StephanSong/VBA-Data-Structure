VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1
Option Explicit
'20200805 DONE

'all the supporting functions are in the module data_structure

Private key_list As vba_list
Private value_list As vba_list
Private p_length As Long

Private Sub Class_Initialize()
    p_length = 0
    Set key_list = New vba_list
    Set value_list = New vba_list
    
End Sub
Public Property Get length() As Integer
    length = p_length
End Property

Public Sub init(ByRef dat As Variant, Optional ByVal key_col As Integer = 1)
    ' the data has to include 2 columns
    ' for the duplicate entries, only the last one is kept.
    Dim num_col As Long
    Dim num_row As Long
    Dim val_col As Integer
    Dim icount As Long
    Dim ind As Integer
    
    num_row = getRows(dat)
    num_col = getCols(dat)
    
    If num_row * num_col = 0 Or num_col <> 2 Then
        Exit Sub
    End If
    
    If key_col = 1 Or key_col = 2 Then
        If key_col = 1 Then
            
            val_col = 2
        Else
            val_col = 1
        End If
    End If
    
    For icount = 1 To UBound(dat, 1)
        If key_list.isin(dat(icount, key_col)) = False Then
            
            key_list.Append dat(icount, key_col)
            value_list.Append (dat(icount, val_col))
        Else
            ind = key_list.index(dat(icount, key_col))
            value_list.setItem ind, dat(icount, val_col)
        End If
    Next icount
    p_length = key_list.length
    
End Sub
'add
Public Sub add(ByVal key As Variant, ByVal value As Variant)
    Dim ind As Integer
    'if the key already exists in the dictionary, update the value
    If key_list.isin(key) Then
        ind = key_list.index(key)
        value_list.setItem key, value
    Else
        key_list.Append (key)
        value_list.Append (value)
        p_length = key_list.length
    End If
End Sub

Public Function getItem(key As Variant) As Variant
    Dim ind As Integer
    
    If key_list.isin(key) Then
        ind = key_list.index(key)
        getItem = value_list.getItem(ind)
    Else
        getItem = "nan"
    End If
End Function

Public Function RemoveItem(key) As Variant
    If key_list.isin(key) Then
        ind = key_list.index(key)
        
        key_list.Remove (ind)
        value_list.Remove (ind)
        length = key_list.length
    End If
End Function

Public Function isin(key) As Boolean
    isin = key_list.isin(key)
End Function

Public Sub clear()
    key_list.clear
    value_list.clear
    length = 0
End Sub

Public Sub printDict()
    Dim icount As Integer
    For icount = 1 To p_length
        Debug.Print key_list.getItem(icount), value_list.getItem(icount)
    Next icount
End Sub

Public Function data() As Variant
    Dim result() As Variant
    ReDim result(1 To p_length, 2)
    Dim icount As Integer

    If p_length > 1 Then
        For icount = 1 To p_length
            result(icount, 1) = key_list.getItem(icount)
            result(icount, 2) = value_list.getItem(icount)
        Next icount
        data = result
    Else
        data = "nan"
    End If
End Function
Public Function keys() As Variant
    keys = key_list.values("h")
End Function

Public Function values() As Variant
    values = value_list.values("h")
End Function
