VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vba_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1
Option Explicit

'20200805 DONE


'all the supporting functions are in the module data_structure
'member variables
Private data() As Variant
Private p_length As Long
'dtype can be normal type or object. this determines how we assign values to the variable.
Private dtype As String 'not implemented yet

' initialization method
Private Sub Class_Initialize()
    p_length = 0
End Sub

Public Sub init(ByVal dat As Variant)
'the dtype of the elements in dat should be primitive type.
    
    Dim icount As Long
    Dim ndim As Integer
    Dim num_counts As Long
    Dim item As Variant
    
    'for different type the variant should be pre-processed differently
    If TypeName(dat) = "Range" Then
        dat = dat.value
    End If
    
    ndim = data_structure.getDimension(dat)
    If ndim = 0 Then
        ReDim data(1)
        data(1) = dat
        p_length = 1
    Else
        num_counts = data_structure.getElementsCounts(dat)
        ReDim data(1 To num_counts)
        icount = 0
        For Each item In dat
            icount = icount + 1
            data(icount) = item
        Next item
        p_length = icount
    End If
End Sub

'Property area
Public Property Get length() As Integer
    length = p_length
End Property

'method area
    'reverse the list
Public Sub Reverse()
    Dim tmp As Variant
    Dim icount As Integer
    For icount = 1 To CInt(p_length / 2)
        tmp = data(icount)
        data(icount) = data(p_length - icount + 1)
        data(p_length - icount + 1) = tmp
    Next icount
End Sub
    ' add one element to the list at the given position
Public Sub Append(ByVal el As Variant, Optional ByVal ind As Long = -1)

    If ind = -1 Or ind < 0 Or ind > length Then
        p_length = p_length + 1
        ReDim Preserve data(p_length) As Variant
        data(p_length) = el
    Else
        Dim tmp() As Variant
        Dim icount As Integer
        
        ReDim tmp(1 To length + 1)
        
        For icount = 1 To ind - 1
            tmp(icount) = data(icount)
        Next icount
        tmp(ind) = el
        For icount = ind + 1 To length + 1
            tmp(icount) = data(icount - 1)
        Next icount
        data = tmp
        p_length = p_length + 1
    End If
End Sub

    'get elements
Public Function getItem(Optional ByVal key As Variant = 1) As Variant
    If TypeName(key) = "Integer" Or TypeName(key) = "Double" Or TypeName(key) = "Long" Then
        getItem = data(key)
    End If
End Function
    'this function always returns a vba_list
Public Function subset(ByVal key As Variant) As vba_list
    Set subset = New vba_list
    Dim icount As Long
    
    If TypeName(key) = "Integer" Or TypeName(key) = "Double" Or TypeName(key) = "Long" Then
        
        If key <= 0 Or key > length Then
            Exit Function
        Else
            subset.init getOneItem(key)
            Exit Function
        End If
    End If
    
    Dim num_start As Long
    Dim num_end As Long

    If LCase(TypeName(key)) = "vba_slicer" Then
        num_start = key.start_with
        num_end = key.end_with
        If num_start < 0 Or num_start > length Or num_end < 0 Or num_end > length Or num_start > num_end Then
            Exit Function
        End If
        
        For icount = num_start To num_end
            subset.Append data(icount)
        Next icount
    End If
End Function
    'set the item value by position
Public Sub setItem(ByVal key As Variant, ByVal el As Variant)
    ' this has to be refined later since the key can also support slicer
    If TypeName(key) = "Integer" Or TypeName(key) = "Double" Or TypeName(key) = "Long" Then
        key = CInt(key)
        If key >= 0 And key <= length Then
            data(key) = el
            Exit Sub
        End If
    End If
    
    ' not implemented yet.
    If LCase(TypeName(key)) = "vba_list" Then
    
        
    End If
    
    If LCase(TypeName(key)) = "vba_dict" Then
    
        
    End If
    
    If LCase(TypeName(key)) = "vba_slicer" Then
    
        
    End If
End Sub

    'membership testing
Public Function isin(ByVal el As Variant) As Boolean
    
    isin = False
    Dim item As Variant
    If p_length = 0 Then
        Exit Function
    Else
        For Each item In data
            If item = el Then
                isin = True
                Exit Function
            End If
        Next item
    End If
End Function

'get item by indexing and slicing
Public Function getOneItem(Optional ByVal key As Variant = 1) As Variant
    If TypeName(key) = "Integer" Then
        
        If key <= 0 Or key > length Then
            getOneItem = False
            Exit Function
        Else
            getOneItem = data(key)
            Exit Function
        End If
    End If
End Function
    'count the elments - done
Public Function counts(ByVal el As Variant) As Long
    counts = 0
    Dim item As Variant
    For Each item In data
        If item = el Then
            counts = counts + 1
        End If
    Next item
End Function

'always return a number
Public Function index(ByVal val As Variant, Optional ByVal from_last As Boolean = False) As Long
    index = -1
    Dim icount As Long
    
    If from_last Then
        For icount = p_length To 1 Step -1
            If data(icount) = val Then
                index = icount
                Exit Function
            End If
        Next icount
    Else
        For icount = 1 To p_length
            If data(icount) = val Then
                index = icount
                Exit Function
            End If
        Next icount
    End If
End Function

    ' remove item by index
Public Sub RemoveItem(Optional ByVal ind As Long = -1)

    Dim tmp() As Variant
    Dim icount As Integer
    
    If ind = -1 Or ind < 0 Or ind > p_length Then
        ind = p_length
    End If
    
    ReDim tmp(1 To p_length - 1)
    
    For icount = 1 To ind - 1
        tmp(icount) = data(icount)
    Next icount
    
    For icount = ind + 1 To p_length
    
        tmp(icount - 1) = data(icount)
    Next icount
    data = tmp
    p_length = p_length - 1
End Sub

    'print the list from the beginning to the end
Public Sub PrintList()
    Dim icount  As Long
    For icount = 1 To p_length
        Debug.Print icount, data(icount)
    Next icount
End Sub
    'convert the list to an array
Public Function values(Optional ByVal transpose As String = "v") As Variant
    Dim tmp() As Variant
    ReDim tmp(1 To p_length, 1) As Variant
    Dim icount As Long
    
    If transpose = "v" Then
        For icount = 1 To p_length
            tmp(icount, 1) = data(icount)
        Next icount
        values = tmp
    Else
        values = data
    End If
End Function

Public Sub clear()
    Erase data
    p_length = 0
End Sub
' we need the dictionary
Public Function value_counts(Optional ByVal total_row As Boolean = True, Optional ByRef by_dict As VBA_Dict) As VBA_Dict
    Dim elements As Variant
    Dim tmp() As Variant
    Dim icount As Long
    Dim num_elements As Integer
    Dim elements_list As vba_list
    
    Set elements_list = New vba_list
    elements_list.init data_structure.getUniqueElements(data)
    num_elements = elements_list.length
    
    ReDim tmp(1 To num_elements, 2)
    
    If Not (by_dict Is Nothing) Then
        elements_list.sort by_dict
    End If
    
    
    For icount = 1 To num_elements
        tmp(icount, 1) = elements_list.getItem(icount)
        tmp(icount, 2) = counts(tmp(icount, 1))
    Next icount
    
    Set value_counts = New VBA_Dict
    value_counts.init tmp, 1
    value_counts.add "Total", data_structure.vba_sum(value_counts.values)
    
End Function

Public Function vba_sum() As Variant
    On Error GoTo errexit
    Dim icount As Long
    
    vba_sum = 0
    For icount = 1 To p_length

        vba_sum = vba_sum + CDbl(data(icount))
    Next icount
    Exit Function
errexit:
    err.clear
    vba_sum = "nan"
End Function

Public Function vba_mean()
    Dim mean_value As Variant
    If p_length > 0 Then
        mean_value = vba_sum()
        If CStr(mean_value) <> "nan" Then
            vba_mean = mean_value / p_length
        End If
    End If
End Function
'always use ascending order
Public Sub sort(Optional ByRef by_dict As VBA_Dict)
    
    Dim icount As Integer
    Dim jcount As Integer
    Dim tmp As Variant
    
    For icount = 1 To p_length - 1
        For jcount = 1 To p_length - icount
           If vba_compare(data(jcount), data(jcount + 1), by_dict) = 1 Then
           
                tmp = data(jcount)
                data(jcount) = data(jcount + 1)
                data(jcount + 1) = tmp
           End If
        Next jcount
    Next icount
End Sub
