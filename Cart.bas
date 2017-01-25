Attribute VB_Name = "Cart"
Sub NewCart()
    
    'This handles the memory space for the program. Do changes to prices, promo groupings, promo codes here
    'initialize
    ThisWorkbook.Worksheets("Global").Range("B1") = 0       'global_ult_small
    ThisWorkbook.Worksheets("Global").Range("B2") = 0       'global_ult_medium
    ThisWorkbook.Worksheets("Global").Range("B3") = 0       'global_ult_large
    ThisWorkbook.Worksheets("Global").Range("B4") = 0       'global_1gb
    ThisWorkbook.Worksheets("Global").Range("B5") = ""      'promo_code
    ThisWorkbook.Worksheets("Global").Range("B6") = "False" 'promo_code active status
    
    'Set correct promo code
    ThisWorkbook.Worksheets("Global").Range("B13") = "I<3AMAYSIM"
    
    'clear contents of purchases
    ThisWorkbook.Worksheets("Global").Range("M4:N8").ClearContents
    
    'Set the variables for grouping conditions
    ThisWorkbook.Worksheets("Global").Range("B9") = 3           'global_Rules_ult_small_grouping_promo
    ThisWorkbook.Worksheets("Global").Range("B10") = 1          'global_Rules_ult_med_grouping_promo
    ThisWorkbook.Worksheets("Global").Range("B11") = 3          'global_Rules_ult_large_grouping_promo
    
    
    'Set the variables for pricing conditions
    ThisWorkbook.Worksheets("Global").Range("B16") = 24.9       'global_Rules_ult_small_price
    ThisWorkbook.Worksheets("Global").Range("B17") = 49.8       'global_Rules_ult_small_price_group
    ThisWorkbook.Worksheets("Global").Range("B18") = 29.9       'global_Rules_ult_medium_price
    ThisWorkbook.Worksheets("Global").Range("B19") = 0          'global_Rules_ult_medium_price_group
    ThisWorkbook.Worksheets("Global").Range("B20") = 44.9       'global_Rules_ult_large_price
    ThisWorkbook.Worksheets("Global").Range("B21") = 39.9       'global_Rules_ult_large_price_group
    ThisWorkbook.Worksheets("Global").Range("B22") = 9.9        'global_Rules_1gb_price
    ThisWorkbook.Worksheets("Global").Range("B23") = 0          'global_Rules_1gb_price_group
    ThisWorkbook.Worksheets("Global").Range("B24") = 0.9        'global_promo_discount
    
    ThisWorkbook.Worksheets("Global").TextBox1 = ""

End Sub


Sub Add(ByVal item As String, Optional ByVal Promo_Code As String)

    Dim global_ult_small As Integer
    Dim global_ult_medium As Integer
    Dim global_ult_large As Integer
    Dim global_1gb As Integer
    Dim global_promo_code As String
    
    global_ult_small = ThisWorkbook.Worksheets("Global").Range("B1")
    global_ult_medium = ThisWorkbook.Worksheets("Global").Range("B2")
    global_ult_large = ThisWorkbook.Worksheets("Global").Range("B3")
    global_1gb = ThisWorkbook.Worksheets("Global").Range("B4")
    global_promo_code = ThisWorkbook.Worksheets("Global").Range("B5")

    If item = "ult_small" Then
        global_ult_small = global_ult_small + 1

    ElseIf item = "ult_medium" Then
        global_ult_medium = global_ult_medium + 1

    ElseIf item = "ult_large" Then
        global_ult_large = global_ult_large + 1
    
    ElseIf item = "1gb" Then
        global_1gb = global_1gb + 1

    End If

    If Promo_Code = "I<3AMAYSIM" Then
        global_promo_code = "TRUE"
    Else
        global_promo_code = "FALSE"
    End If
    
    'Pass the global values
    ThisWorkbook.Worksheets("Global").Range("B1") = global_ult_small
    ThisWorkbook.Worksheets("Global").Range("B2") = global_ult_medium
    ThisWorkbook.Worksheets("Global").Range("B3") = global_ult_large
    ThisWorkbook.Worksheets("Global").Range("B4") = global_1gb
    ThisWorkbook.Worksheets("Global").Range("B6") = global_promo_code
    
    
    
End Sub
       
Sub items()

    Dim global_ult_small As Integer
    Dim global_ult_medium As Integer
    Dim global_ult_large As Integer
    Dim global_1gb As Integer
    
    Dim gb_bonus As Integer
    Dim ult_medium_i As Integer
    
    'get global variable values
    global_ult_small = ThisWorkbook.Worksheets("Global").Range("B1")
    global_ult_medium = ThisWorkbook.Worksheets("Global").Range("B2")
    global_ult_large = ThisWorkbook.Worksheets("Global").Range("B3")
    global_1gb = ThisWorkbook.Worksheets("Global").Range("B4")
    
    'get rule variable values
    global_Rules_ult_med_grouping_promo = ThisWorkbook.Worksheets("Global").Range("B10")


    If global_ult_medium >= global_Rules_ult_med_grouping_promo Then
        
        ult_medium_i = global_ult_medium
        gb_bonus = 0

        Do While ult_medium_i >= global_Rules_ult_med_grouping_promo
        
            ult_medium_i = ult_medium_i - global_Rules_ult_med_grouping_promo
            gb_bonus = gb_bonus + 1
            
        Loop
        
        global_1gb = global_1gb + gb_bonus
    End If
    
    'save values
    ThisWorkbook.Worksheets("Global").Range("M4") = global_ult_small
    ThisWorkbook.Worksheets("Global").Range("M5") = global_ult_medium
    ThisWorkbook.Worksheets("Global").Range("M6") = global_ult_large
    ThisWorkbook.Worksheets("Global").Range("M7") = global_1gb
    
End Sub

Sub Total()

    Dim global_ult_small As Integer
    Dim global_ult_medium As Integer
    Dim global_ult_large As Integer
    Dim global_1gb As Integer
    
    Dim group_small_count As Integer
    Dim leftover_ult_small As Integer
    
    Dim global_Rules_ult_small_price As Single
    Dim global_Rules_ult_small_price_group As Single
    Dim global_Rules_ult_medium_price As Single
    Dim global_Rules_ult_medium_price_group As Single
    Dim global_Rules_ult_large_price As Single
    Dim global_Rules_ult_large_price_group As Single
    Dim global_Rules_1gb_price As Single
    Dim global_Rules_1gb_price_group As Single

    
    Dim price_ult_small_total As Single
    Dim price_ult_medium_total As Single
    Dim price_ult_large_total As Single
    Dim price_1gb_total As Single
    Dim Total_Price As Single
    
    Dim global_Rules_ult_small_grouping_promo As Integer
    Dim global_Rules_ult_med_grouping_promo As Integer
    Dim global_Rules_ult_large_grouping_promo As Integer
    
    Dim Promo_Code_Active As String
    Dim global_promo_discount As Single



    'get global variable values
    global_ult_small = ThisWorkbook.Worksheets("Global").Range("B1")
    global_ult_medium = ThisWorkbook.Worksheets("Global").Range("B2")
    global_ult_large = ThisWorkbook.Worksheets("Global").Range("B3")
    global_1gb = ThisWorkbook.Worksheets("Global").Range("B4")
    
    
    'START OF LOGIC FOR UNLIMITED 1GB
    '**************************************************************************************
    'get pricing
    global_Rules_ult_small_price = ThisWorkbook.Worksheets("Global").Range("B16")
    global_Rules_ult_small_price_group = ThisWorkbook.Worksheets("Global").Range("B17")
    
    'get promo item groupings
    global_Rules_ult_small_grouping_promo = ThisWorkbook.Worksheets("Global").Range("B9")

    
    
    '3 for 2 unlimited 1GB (ult_small) Sim rule
    'logic here is applicable for other price grouping combinationss (i.e. 4 for 2, 4 for 3, 5 for 2, etc)

    If global_ult_small >= global_Rules_ult_small_grouping_promo Then

        group_small_count = 0   'counts how many promo groups (3 per group in this casse)
        leftover_ult_small = 0

        Do While global_ult_small >= global_Rules_ult_small_grouping_promo
        
            global_ult_small = global_ult_small - global_Rules_ult_small_grouping_promo
            group_small_count = group_small_count + 1
        Loop

        leftover_ult_small = global_ult_small

        price_ult_small_total = group_small_count * global_Rules_ult_small_price_group + leftover_ult_small * global_Rules_ult_small_price
    Else
        
        price_ult_small_total = global_ult_small * global_Rules_ult_small_price
    End If
    
       
    
    'START OF LOGIC FOR UNLIMITED 2GB
    '***************************************************************************************************
    
    'get pricing
    global_Rules_ult_medium_price = ThisWorkbook.Worksheets("Global").Range("B18")
        
    'get grouping
    'NO PROMO GROUPING APPLICABLE
    
    price_ult_medium_total = global_ult_medium * global_Rules_ult_medium_price
    
      
    
    'START OF LOGIC FOR UNLIMITED 5GB
    '***************************************************************************************************
    
    'get pricing
    global_Rules_ult_large_price = ThisWorkbook.Worksheets("Global").Range("B20")
    global_Rules_ult_large_price_group = ThisWorkbook.Worksheets("Global").Range("B21")
        
    'get grouping
    global_Rules_ult_large_grouping_promo = ThisWorkbook.Worksheets("Global").Range("B11")

    
    If global_ult_large > global_Rules_ult_large_grouping_promo Then
    
        price_ult_large_total = global_ult_large * global_Rules_ult_large_price_group
    Else
    
        price_ult_large_total = global_ult_large * global_Rules_ult_large_price
    End If
    
    
    
    'START OF LOGIC FOR 1GB Data-Pack
    '***************************************************************************************************
    
    'get pricing
    global_Rules_1gb_price = ThisWorkbook.Worksheets("Global").Range("B22")
        
    'get grouping
    'NO PROMO GROUPING APPLICABLE
      
    price_1gb_total = global_1gb * global_Rules_1gb_price
    'End If
  
    
    
    'SAVING OF PRICES
    '***************************************************************************************************
    
    Total_Price = price_ult_small_total + price_ult_medium_total + price_ult_large_total + price_1gb_total
    
    'check promo code activation
    Promo_Code_Active = ThisWorkbook.Worksheets("Global").Range("B6")
    
    If Promo_Code_Active = "True" Then
        
        'get discount promo value
        global_promo_discount = ThisWorkbook.Worksheets("Global").Range("B24")
    
        Total_Price = Total_Price * global_promo_discount
        price_ult_small_total = price_ult_small_total * global_promo_discount
        price_ult_medium_total = price_ult_medium_total * global_promo_discount
        price_ult_large_total = price_ult_large_total * global_promo_discount
        price_1gb_total = price_1gb_total * global_promo_discount
    End If
    
    'Display for easy viewing
    ThisWorkbook.Worksheets("Global").Range("N4") = price_ult_small_total
    ThisWorkbook.Worksheets("Global").Range("N5") = price_ult_medium_total
    ThisWorkbook.Worksheets("Global").Range("N6") = price_ult_large_total
    ThisWorkbook.Worksheets("Global").Range("N7") = price_1gb_total
    ThisWorkbook.Worksheets("Global").Range("N8") = Total_Price

End Sub



