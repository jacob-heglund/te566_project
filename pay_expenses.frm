VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pay_expenses 
   Caption         =   "UserForm1"
   ClientHeight    =   10836
   ClientLeft      =   -348
   ClientTop       =   -1296
   ClientWidth     =   17808
   OleObjectBlob   =   "pay_expenses.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pay_expenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_Click()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub address_2_Change()

End Sub

Private Sub clear_customer_data_Click()

For Each ctrl In Me.Controls
Select Case TypeName(ctrl)
    Case "TextBox"
        ctrl.Text = ""
    
        End Select
    Next

End Sub

Private Sub company_name_Change()

End Sub

Private Sub display_Click()

End Sub

Private Sub exit_customer_data_Click()

Unload Me
main_menu.Show

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub state_Change()

End Sub

Private Sub submit_customer_data_Click()

'add customer data to the customer database
Workbooks("finances.xlsm").Activate
Worksheets("customers").Activate

'TODO check that all the fields are filled
'TODO double check submission (avoid mistakes, formatting of submissions)

Rows("2:2").Select
Selection.Insert shift:=xlDown
Worksheets("customers").Activate
Range("a2").Offset(0, 0) = company_name.Text
Range("a2").Offset(0, 1) = first_name.Text
Range("a2").Offset(0, 2) = last_name.Text
Range("a2").Offset(0, 3) = address_1.Text
Range("a2").Offset(0, 4) = address_2.Text
Range("a2").Offset(0, 5) = city.Text
Range("a2").Offset(0, 6) = state.Text
Range("a2").Offset(0, 7) = zip_code.Text
Range("a2").Offset(0, 8) = price.Text
Range("a2").Offset(0, 9) = Val(cash_percentage.Text) * Val(price.Text) / 100
Range("a2").Offset(0, 10) = Val(credit_percentage.Text) * Val(price.Text) / 100
Range("a2").Offset(0, 11) = date_of_sale.Text

price_cash = Val(cash_percentage.Text) * Val(price.Text) / 100
price_credit = Val(credit_percentage.Text) * Val(price.Text) / 100

display.ColumnCount = 10
display.RowSource = "A1:J100"

'update the balance sheet'

'subtract from inventory
Worksheets("balance_sheet").Activate

'add gross margin to cash and accounts reveivable
Range("B4") = Range("B4") + price_cash
Range("B5") = Range("B5") + price_credit



End Sub







Private Sub TextBox2_Change()

End Sub
