# vba-array-operation

VBA has no native support for Array/Marix Operations

This is an attemp to provide an easy way to perform the most common array operations.

To see it in action, open up the attached excel file and look in the VBA editor on how the class is used.

Alternatively,

```VBA
Sub TestArrayOps()
    Dim avb As New ArrayVBA
    Dim X As Variant
    Dim B As Variant
    Dim i As Integer, j As Integer
    Dim row As String
    
    X = Sheet2.Range("B9:F58").Value
    B = Sheet2.Range("B2:B6").Value
    y = Sheet2.Range("G9:G58").Value
    
    
    Debug.Print "---------------------- Raw Data ------------------------------------"
    
    For i = LBound(X, 1) To UBound(X, 1)
        row = ""
        For j = LBound(X, 2) To UBound(X, 2)
            row = row + CStr(X(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Dim x_t As Variant
    
    x_t = avb.Transpose(X)

    Debug.Print "---------------------- Transpose ------------------------------------"
    For i = LBound(x_t, 1) To UBound(x_t, 1)
        row = ""
        For j = LBound(x_t, 2) To UBound(x_t, 2)
            row = row + CStr(x_t(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
        
    Dim X_dot As Variant
    
    X_dot = avb.Dot(X, x_t)

    Debug.Print "---------------------- Dot Product ------------------------------------"
    For i = LBound(X_dot, 1) To UBound(X_dot, 1)
        row = ""
        For j = LBound(X_dot, 2) To UBound(X_dot, 2)
            row = row + CStr(X_dot(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Dim x_a As Variant
    
    x_a = avb.Append(X, 99#)

    Debug.Print "---------------------- Append Value ------------------------------------"
    For i = LBound(x_a, 1) To UBound(x_a, 1)
        row = ""
        For j = LBound(x_a, 2) To UBound(x_a, 2)
            row = row + CStr(x_a(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next

    Dim x_s As Variant
    
    x_s = avb.substract(X, x_a)
    
    Debug.Print "---------------------- Subtract ------------------------------------"
    For i = LBound(x_s, 1) To UBound(x_s, 1)
        row = ""
        For j = LBound(x_s, 2) To UBound(x_s, 2)
            row = row + CStr(x_s(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Dim x_add As Variant
    
    x_add = avb.Add(X, x_a)

    Debug.Print "---------------------- Add ------------------------------------"
    For i = LBound(x_add, 1) To UBound(x_add, 1)
        row = ""
        For j = LBound(x_add, 2) To UBound(x_add, 2)
            row = row + CStr(x_add(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    
    Debug.Print "---------------------- Max - All ------------------------------------"
    Debug.Print avb.Max(X)

    Dim x_max As Variant
    
     x_max = avb.Max(X, 1)

    Debug.Print "---------------------- Max - Columns ------------------------------------"
    For i = LBound(X, 1) To UBound(x_max, 1)
        row = ""
        For j = LBound(x_max, 2) To UBound(x_max, 2)
            row = row + CStr(x_max(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next


     x_max = avb.Max(X, 2)

    Debug.Print "---------------------- Max - rows ------------------------------------"
    For i = LBound(x_max, 1) To UBound(x_max, 1)
        row = ""
        For j = LBound(x_max, 2) To UBound(x_max, 2)
            row = row + CStr(x_max(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Min - All ------------------------------------"
    Debug.Print avb.Min(X, 0)
    
    Dim x_min As Variant
    
    x_min = avb.Min(X, 1)
    
    Debug.Print "---------------------- Min - Columns ------------------------------------"
    For i = LBound(x_min, 1) To UBound(x_min, 1)
        row = ""
        For j = LBound(x_min, 2) To UBound(x_min, 2)
            row = row + CStr(x_min(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    x_min = avb.Min(X, 2)
    
    Debug.Print "---------------------- Min - Rows ------------------------------------"
    For i = LBound(x_min, 1) To UBound(x_min, 1)
        row = ""
        For j = LBound(x_min, 2) To UBound(x_min, 2)
            row = row + CStr(x_min(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Dim x_cummsum As Variant
    
    Debug.Print "---------------------- Cummulative Sum - All ------------------------------------"
    x_cummsum = avb.CummSum(X, 0)
    
    For i = LBound(x_cummsum, 1) To UBound(x_cummsum, 1)
        row = ""
        For j = LBound(x_cummsum, 2) To UBound(x_cummsum, 2)
            row = row + CStr(x_cummsum(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Cummulative Sum - Columns ------------------------------------"
    x_cummsum = avb.CummSum(X, 1)
    
    For i = LBound(x_cummsum, 1) To UBound(x_cummsum, 1)
        row = ""
        For j = LBound(x_cummsum, 2) To UBound(x_cummsum, 2)
            row = row + CStr(x_cummsum(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Cummulative Sum - Rows ------------------------------------"
    x_cummsum = avb.CummSum(X, 2)
    
    For i = LBound(x_cummsum, 1) To UBound(x_cummsum, 1)
        row = ""
        For j = LBound(x_cummsum, 2) To UBound(x_cummsum, 2)
            row = row + CStr(x_cummsum(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Array Sum ------------------------------------"
    Debug.Print avb.Sum(X)
    
    Debug.Print "---------------------- Array Mean ------------------------------------"
    Debug.Print avb.mean(X)

    Debug.Print "---------------------- Array Standard Deviation ------------------------------------"
    Debug.Print avb.std(X)
    
    Dim x_cummproduct As Variant
    
    Debug.Print "---------------------- Cummulative Product - All ------------------------------------"
    
    x_cummproduct = avb.CummProduct(X, 0)
    
    For i = LBound(x_cummproduct, 1) To UBound(x_cummproduct, 1)
        row = ""
        For j = LBound(x_cummproduct, 2) To UBound(x_cummproduct, 2)
            row = row + CStr(x_cummproduct(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Cummulative Product - Columns ------------------------------------"
    
    x_cummproduct = avb.CummProduct(X, 1)
    
    For i = LBound(x_cummproduct, 1) To UBound(x_cummproduct, 1)
        row = ""
        For j = LBound(x_cummproduct, 2) To UBound(x_cummproduct, 2)
            row = row + CStr(x_cummproduct(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Cummulative Product - Rows ------------------------------------"
    
    
    x_cummproduct = avb.CummProduct(X, 2)
    
    For i = LBound(x_cummproduct, 1) To UBound(x_cummproduct, 1)
        row = ""
        For j = LBound(x_cummproduct, 2) To UBound(x_cummproduct, 2)
            row = row + CStr(x_cummproduct(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
    
    Debug.Print "---------------------- Initialize - zeroes ------------------------------------"
    
    Dim zeroes As Variant
    
    zeroes = avb.zeroes(10, 2)
    
    For i = LBound(zeroes, 1) To UBound(zeroes, 1)
        row = ""
        For j = LBound(zeroes, 2) To UBound(zeroes, 2)
            row = row + CStr(zeroes(i, j)) + ","
        Next
        Debug.Print Left(row, Len(row) - 1)
    Next
End Sub

## Using the class in Logistic Regression Classification
```VBA
Sub logit()
    Application.ScreenUpdating = False
    
    Dim logit As New LogisticRegression
    Dim coefficients As Variant
    Dim lr As Double, threshold As Double
    Dim X As Variant, labels As Variant, x_pred As Variant, y_pred As Variant
    lr = 0.1
    threshold = 0.6
    
    Dim i As Integer, j As Integer, row As String, iter As Long
    Dim z As Variant, h As Variant, x_t As Variant, error As Variant, gradient As Variant
    Dim ones As Variant
    
    X = Sheets("NBA Data").Range("A2:G800").Value
    labels = Sheets("NBA Data").Range("H2:H800").Value
  
    x_pred = Sheets("TestData").Range("A2:G101").Value

    logit.InitializeProperties learning_rate:=lr, iterations:=10000
    
    Call logit.fit(X, labels)
    
    y_pred = logit.predict(x_pred, threshold)
    

    For i = LBound(y_pred, 1) To UBound(y_pred, 1)
        Sheets("TestData").Cells(i + 1, 9).Value = y_pred(i)
    Next
    
    Application.ScreenUpdating = True
    
End Sub
