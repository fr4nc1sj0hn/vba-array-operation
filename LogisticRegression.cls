VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogisticRegression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_learning_rate As Double
Dim m_iterations As Long
Dim m_fit_intecept As Boolean
Dim m_data As Variant
Dim m_labels As Variant
Dim m_thetas As Variant

Dim ARVBA As ArrayVBA

Public Function InitializeProperties(learning_rate As Double, iterations As Long, Optional fit_intecept As Boolean = True)
    m_learning_rate = learning_rate
    m_iterations = iterations
    m_fit_intecept = fit_intecept
    
    m_data = data
    m_labels = labels
    
    Set ARVBA = New ArrayVBA
    
End Function

Private Function AddIntercept(X As Variant) As Variant
    Dim intercept As Variant
    intercept = ARVBA.ones(UBound(m_data, 1), 1)
    
    AddIntercept = ARVBA.Concatenate(intercept, m_data)
End Function
Private Function Sigmoid(z As Variant) As Variant
    Sigmoid = ARVBA.Sigmoid(z)
End Function
Public Function fit(X As Variant, y As Variant)
    Dim X_data As Variant
    Dim thetas As Variant
    
    If fit_intecept Then
        X_data = AddIntercept(X)
    Else
        X_data = X
    End If
    
    thetas = ARVBA.zeroes(UBound(X_data, 2), 1)
    For iter = 1 To m_iterations
        z = ARVBA.Dot(X_data, thetas)
        h = ARVBA.Sigmoid(z)
        x_t = ARVBA.Transpose(X_data)
        
        gradient = ARVBA.divide(ARVBA.Dot(ARVBA.Transpose(X_data), ARVBA.substract(h, y)), UBound(y))
        
        thetas = ARVBA.substract(thetas, ARVBA.multiply(gradient, m_learning_rate))
    Next
    
    m_thetas = thetas
End Function
Private Function predict_proba(ByVal X As Variant) As Variant
    predict_proba = ARVBA.Sigmoid(ARVBA.Dot(X, m_thetas))
End Function
Public Function predict(ByVal X, threshold As Double) As Variant
    Dim i As Long, j As Integer
    Dim y_pred As Variant, final_result As Variant
    
    y_pred = predict_proba(X)

    ReDim final_result(1 To UBound(y_pred, 1))

    For i = 1 To UBound(y_pred)
        final_result(i) = IIf(y_pred(i, 1) > threshold, 1, 0)
    Next
    predict = final_result
End Function
