VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Dot(ByVal m1 As Variant, ByVal m2 As Variant) As Variant
    If UBound(m1, 1) <> UBound(m2, 2) Then
        MsgBox "Invalid array sizes. Number of rows in m2 should be equal to number of columns in m2"
        Exit Function
    End If
    Dot = WorksheetFunction.MMult(m1, m2)
End Function
Public Function Transpose(ByVal m1 As Variant) As Variant
    Transpose = WorksheetFunction.Transpose(m1)
End Function
Public Function Append(ByVal x As Variant, val As Double) As Variant
    Dim i As Integer
    Dim columns As Integer
    Dim rows As Long
    
    columns = UBound(x, 2)
    rows = UBound(x, 1)
    
    ReDim Preserve x(1 To rows, 1 To columns + 1)
    For i = 1 To rows
        x(i, columns + 1) = val
    Next
    Append = x
End Function
Public Function substract(ByVal m1, ByVal m2) As Variant
    Dim m3 As Variant
    Dim rows As Long, columns As Long
    Dim i As Long, j As Long
    rows = UBound(m1, 1)
    columns = UBound(m1, 2)
    
    ReDim m3(1 To rows, 1 To columns)
    For i = LBound(m1, 1) To UBound(m1, 1)
        For j = LBound(m1, 2) To UBound(m1, 2)
            m3(i, j) = m1(i, j) - m2(i, j)
        Next
    Next
    substract = m3
End Function
Public Function Add(ByVal m1, ByVal m2) As Variant
    Dim m3 As Variant
    Dim rows As Long, columns As Long
    Dim i As Long, j As Long
    rows = UBound(m1, 1)
    columns = UBound(m1, 2)
    
    ReDim m3(1 To rows, 1 To columns)
    For i = LBound(m1, 1) To UBound(m1, 1)
        For j = LBound(m1, 2) To UBound(m1, 2)
            m3(i, j) = m1(i, j) + m2(i, j)
        Next
    Next
    Add = m3
End Function
Public Function Max(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    Dim i As Integer, j As Integer
    Dim result As Variant
    Dim maxval As Variant
    If Axis = 0 Then
        Max = WorksheetFunction.Max(m1)
    End If
    If Axis = 1 Then ' Columns
        
        ReDim result(1 To 1, 1 To UBound(m1, 2))
        
        For j = LBound(m1, 2) To UBound(m1, 2)
            maxval = m1(1, j)
            For i = LBound(m1, 1) To UBound(m1, 1)
                maxval = IIf(m1(i, j) > maxval, m1(i, j), maxval)
            Next
            result(1, j) = maxval
        Next
        
        Max = result
    End If
    If Axis = 2 Then ' Rows
        
        ReDim result(1 To UBound(m1, 1), 1 To 1)
        
        For i = LBound(m1, 1) To UBound(m1, 1)
            maxval = m1(i, 1)
            For j = LBound(m1, 2) To UBound(m1, 2)
                maxval = IIf(m1(i, j) > maxval, m1(i, j), maxval)
            Next
            result(i, 1) = maxval
        Next
        
        Max = result
    End If
End Function
Public Function Min(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    Dim i As Integer, j As Integer
    Dim result As Variant
    Dim minval As Variant
    
    If Axis = 0 Then 'All
        Min = WorksheetFunction.Min(m1)
    End If
    If Axis = 1 Then ' Columns
        
        ReDim result(1 To 1, 1 To UBound(m1, 2))
        
        For j = LBound(m1, 2) To UBound(m1, 2)
            minval = m1(1, j)
            For i = LBound(m1, 1) To UBound(m1, 1)
                minval = IIf(m1(i, j) < minval, m1(i, j), minval)
            Next
            result(1, j) = minval
        Next
        
        Min = result
    End If
    If Axis = 2 Then ' Rows
        
        ReDim result(1 To UBound(m1, 1), 1 To 1)
        
        For i = LBound(m1, 1) To UBound(m1, 1)
            minval = m1(i, 1)
            For j = LBound(m1, 2) To UBound(m1, 2)
                minval = IIf(m1(i, j) < minval, m1(i, j), minval)
            Next
            result(i, 1) = minval
        Next
        
        Min = result
    End If
End Function
Public Function CummSum(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    Dim i As Integer, j As Integer
    Dim result As Variant
    Dim sum As Double

    ReDim result(1 To UBound(m1, 1), 1 To UBound(m1, 2))
    If Axis = 0 Then 'All
        sum = 0
        For i = LBound(m1, 1) To UBound(m1, 1)
            For j = LBound(m1, 2) To UBound(m1, 2)
                sum = sum + m1(i, j)
                result(i, j) = sum
            Next
        Next
        CummSum = result
    End If
    If Axis = 1 Then ' Columns

        For j = LBound(m1, 2) To UBound(m1, 2)
            sum = 0
            For i = LBound(m1, 1) To UBound(m1, 1)
                sum = sum + m1(i, j)
                result(i, j) = sum
            Next
        Next

        CummSum = result
    End If
    If Axis = 2 Then ' Rows

        For i = LBound(m1, 1) To UBound(m1, 1)
            sum = 0
            For j = LBound(m1, 2) To UBound(m1, 2)
                sum = sum + m1(i, j)
                result(i, j) = sum
            Next
        Next

        CummSum = result
    End If
End Function

