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
    Dot = WorksheetFunction.MMult(m1, m2)
End Function
Public Function Transpose(ByVal m1 As Variant) As Variant
    Transpose = WorksheetFunction.Transpose(m1)
End Function
Public Function Append(ByVal X As Variant, val As Double) As Variant
    Dim i As Integer
    Dim columns As Integer
    Dim rows As Long
    
    columns = UBound(X, 2)
    rows = UBound(X, 1)
    
    ReDim Preserve X(1 To rows, 1 To columns + 1)
    For i = 1 To rows
        X(i, columns + 1) = val
    Next
    Append = X
End Function
Public Function Concatenate(ByVal X As Variant, ByVal y As Variant) As Variant
    Dim i As Integer
    Dim columns_x As Integer, columns_y As Integer
    Dim rows_x As Long, rows_y As Long, columns_z As Integer, rows_z As Long
    Dim z As Variant
    Dim val As Double
    
    columns_x = UBound(X, 2)
    rows_x = UBound(X, 1)
    
    columns_y = UBound(y, 2)
    rows_y = UBound(y, 1)

    
    If rows_x <> rows_y Then
        MsgBox "Number of Rows are not equal"
        Exit Function
    End If
    
    columns_z = columns_x + columns_y

    ReDim z(1 To rows_x, 1 To columns_z)
    
    ' X first
    For i = 1 To rows_x
        For j = 1 To columns_x
            z(i, j) = X(i, j)
        Next
    Next
    ' then y
    For i = 1 To rows_x
        For j = 1 To columns_y
            z(i, j + columns_x) = y(i, j)
        Next
    Next
    
    Concatenate = z
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
    Dim Sum As Double

    ReDim result(1 To UBound(m1, 1), 1 To UBound(m1, 2))
    If Axis = 0 Then 'All
        Sum = 0
        For i = LBound(m1, 1) To UBound(m1, 1)
            For j = LBound(m1, 2) To UBound(m1, 2)
                Sum = Sum + m1(i, j)
                result(i, j) = Sum
            Next
        Next
        CummSum = result
    End If
    If Axis = 1 Then ' Columns

        For j = LBound(m1, 2) To UBound(m1, 2)
            Sum = 0
            For i = LBound(m1, 1) To UBound(m1, 1)
                Sum = Sum + m1(i, j)
                result(i, j) = Sum
            Next
        Next

        CummSum = result
    End If
    If Axis = 2 Then ' Rows

        For i = LBound(m1, 1) To UBound(m1, 1)
            Sum = 0
            For j = LBound(m1, 2) To UBound(m1, 2)
                Sum = Sum + m1(i, j)
                result(i, j) = Sum
            Next
        Next

        CummSum = result
    End If
End Function
Public Function CummProduct(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    Dim i As Integer, j As Integer
    Dim result As Variant
    Dim product As Double

    ReDim result(1 To UBound(m1, 1), 1 To UBound(m1, 2))
    If Axis = 0 Then 'All
        product = 1
        For i = LBound(m1, 1) To UBound(m1, 1)
            For j = LBound(m1, 2) To UBound(m1, 2)
                product = product * m1(i, j)
                result(i, j) = product
            Next
        Next
        CummProduct = result
    End If
    If Axis = 1 Then ' Columns

        For j = LBound(m1, 2) To UBound(m1, 2)
            product = 1
            For i = LBound(m1, 1) To UBound(m1, 1)
                product = product * m1(i, j)
                result(i, j) = product
            Next
        Next

        CummProduct = result
    End If
    If Axis = 2 Then ' Rows

        For i = LBound(m1, 1) To UBound(m1, 1)
            product = 1
            For j = LBound(m1, 2) To UBound(m1, 2)
                product = product * m1(i, j)
                result(i, j) = product
            Next
        Next

        CummProduct = result
    End If
End Function
Public Function Sum(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    Sum = WorksheetFunction.Sum(m1)
End Function
Public Function mean(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    mean = WorksheetFunction.Average(m1)
End Function
Public Function std(ByVal m1 As Variant, Optional ByVal Axis As Integer = 0) As Variant
    std = WorksheetFunction.StDev(m1)
End Function
Public Function zeroes(rows As Integer, columns As Integer) As Variant
    Dim i As Integer, j As Integer
    Dim arr As Variant
    
    ReDim arr(1 To rows, 1 To columns)
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = 0
        Next
    Next
    zeroes = arr
End Function
Public Function ones(rows As Integer, columns As Integer) As Variant
    Dim i As Integer, j As Integer
    Dim arr As Variant
    
    ReDim arr(1 To rows, 1 To columns)
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = 1
        Next
    Next
    ones = arr
End Function
Public Function Sigmoid(m As Variant) As Variant
    Dim i As Integer, j As Integer
    Dim arr As Variant
    
    ReDim arr(1 To UBound(m, 1), 1 To UBound(m, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = 1 / (1 + Math.exp(-m(i, j)))
        Next
    Next
    Sigmoid = arr
End Function
Public Function divide(m As Variant, d As Double) As Variant
    Dim i As Integer, j As Integer
    Dim arr As Variant
    
    ReDim arr(1 To UBound(m, 1), 1 To UBound(m, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = m(i, j) / d
        Next
    Next
    divide = arr
End Function
Public Function multiply(m As Variant, d As Double) As Variant
    Dim i As Integer, j As Integer
    Dim arr As Variant
    
    ReDim arr(1 To UBound(m, 1), 1 To UBound(m, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = m(i, j) * d
        Next
    Next
    multiply = arr
End Function

