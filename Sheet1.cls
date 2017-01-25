VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim PromoCode As String
    
    PromoCode = ThisWorkbook.Worksheets("Global").Range("B5")
    
    Cart.Add item:="ult_small", Promo_Code:=PromoCode
End Sub

Private Sub CommandButton2_Click()
    Dim PromoCode As String
    
    PromoCode = ThisWorkbook.Worksheets("Global").Range("B5")
    
    Cart.Add item:="ult_medium", Promo_Code:=PromoCode
End Sub

Private Sub CommandButton3_Click()
    Dim PromoCode As String
    
    PromoCode = ThisWorkbook.Worksheets("Global").Range("B5")
    
    Cart.Add item:="ult_large", Promo_Code:=PromoCode
End Sub

Private Sub CommandButton4_Click()
        Dim PromoCode As String
    
    PromoCode = ThisWorkbook.Worksheets("Global").Range("B5")
    
    Cart.Add item:="1gb", Promo_Code:=PromoCode
End Sub

Private Sub CommandButton5_Click()
    Cart.NewCart
End Sub

Private Sub CommandButton6_Click()
    Cart.items
    Cart.Total
End Sub

Private Sub CommandButton7_Click()
    
    Dim correct_code As String
    
    ThisWorkbook.Worksheets("Global").Range("B5") = ThisWorkbook.Worksheets("Global").TextBox1
    
    correct_code = ThisWorkbook.Worksheets("Global").Range("B13")
    
    If correct_code = ThisWorkbook.Worksheets("Global").Range("B5") Then
        ThisWorkbook.Worksheets("Global").Range("B6") = "TRUE"
    Else
        ThisWorkbook.Worksheets("Global").Range("B6") = "FALSE"
    End If
End Sub

Private Sub TextBox1_Change()

End Sub
