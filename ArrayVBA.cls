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

Public Function MMul(ByVal m1 As Variant, ByVal m2 As Variant) As Variant
    Dim i As Long, j As Long
    For i = LBound(m1, 1) To UBound(m1, 1)
        For j = LBound(m1, 2) To UBound(m1, 2)
            m1(i, j) = m1(i, j) * m2(i, j)
        Next j
    Next i
    MMul = m1
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
Public Function Concatenate(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim i As Integer
    Dim columns_x As Integer, columns_y As Integer
    Dim rows_x As Long, rows_y As Long, columns_z As Integer, rows_z As Long
    Dim z As Variant
    Dim val As Double
    
    columns_x = UBound(x, 2)
    rows_x = UBound(x, 1)
    
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
            z(i, j) = x(i, j)
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
Public Function MSubstract(ByVal m1, ByVal m2) As Variant
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
    MSubstract = m3
End Function
Public Function Add(ByVal m1, ByVal m2, broadcast As Boolean) As Variant
    Dim m3 As Variant
    Dim rows As Long, columns As Long
    Dim i As Long, j As Long
    
    If Not broadcast Then
        rows = UBound(m1, 1)
        columns = UBound(m1, 2)
        
        ReDim m3(1 To rows, 1 To columns)
        For i = LBound(m1, 1) To UBound(m1, 1)
            For j = LBound(m1, 2) To UBound(m1, 2)
                m3(i, j) = m1(i, j) + m2(i, j)
            Next
        Next
        Add = m3
        Exit Function
    End If
    
    For i = LBound(m1, 1) To UBound(m1, 1)
        For j = LBound(m1, 2) To UBound(m1, 2)
            'Debug.Print i, j
            m1(i, j) = m1(i, j) + m2(1, j)
        Next j
    Next i
    
    Add = m1
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
        Debug.Print 'Columns'
        ReDim result(1 To 1, 1 To UBound(m1, 2))
        
        For j = LBound(m1, 2) To UBound(m1, 2)
            Sum = 0
            For i = LBound(m1, 1) To UBound(m1, 1)
                Sum = Sum + m1(i, j)
            Next
            result(1, j) = Sum
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
Public Function randn(rows As Integer, columns As Integer) As Variant
    Dim x() As Variant
    Dim y() As Variant
    Dim i As Long, j As Long
    Dim temp As Double

    ' Initialize X with random numbers
    ReDim x(1 To rows, 1 To columns)
    For i = 1 To rows
        For j = 1 To columns
            x(i, j) = WorksheetFunction.Norm_Inv(Rnd(), 0, 1) ' random number from N(0,1)
        Next j
    Next i
    randn = x
End Function
Public Function PrintArray(arr As Variant, rows As Long) As Boolean
    Dim y As Long, x As Long
    
    If rows = -1 Then
        x = UBound(arr, 1)
    Else
        x = rows
    End If
    
    For i = LBound(arr, 1) To x
        row = ""
        For j = LBound(arr, 2) To UBound(arr, 2)
            row = row + CStr(arr(i, j)) + ","
        Next
        Debug.Print row
    Next
    PrintArray = True
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
            arr(i, j) = 1 / (1 + Math.Exp(-m(i, j)))
        Next
    Next
    Sigmoid = arr
End Function
Public Function ReLU(m As Variant) As Variant
    Dim i As Long, j As Long
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            If m(i, j) > 0 Then
                m(i, j) = m(i, j)
            Else
                m(i, j) = 0
            End If
        Next j
    Next i
    ReLU = m
End Function
Public Function ReLU_Derivative(m As Variant) As Variant
    Dim i As Long, j As Long
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            If m(i, j) > 0 Then
                m(i, j) = 1
            Else
                m(i, j) = 0
            End If
        Next j
    Next i
    ReLU_Derivative = m
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

Public Function subtract(m As Variant, d As Double) As Variant
    Dim i As Integer, j As Integer
    Dim arr As Variant
    
    ReDim arr(1 To UBound(m, 1), 1 To UBound(m, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = m(i, j) - d
        Next
    Next
    subtract = arr
End Function

Public Function square(m As Variant) As Variant
    Dim i As Long, j As Long
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            m(i, j) = m(i, j) ^ 2
        Next j
    Next i
    square = m
End Function


