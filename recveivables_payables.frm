VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} recveivables_payables 
   Caption         =   "UserForm1"
   ClientHeight    =   10836
   ClientLeft      =   -348
   ClientTop       =   -1296
   ClientWidth     =   17808
   OleObjectBlob   =   "recveivables_payables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "recveivables_payables"
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

'update expense database
Workbooks("finances.xlsm").Activate
Worksheets("expenses").Activate

'TODO check that all the fields are filled
'TODO double check submission (avoid mistakes, formatting of submissions)

Rows("2:2").Select
Selection.Insert shift:=xlDown
Range("a2").Offset(0, 0) = amount.Text
Range("a2").Offset(0, 1) = date_of_payment.Text
Range("a2").Offset(0, 2) = notes.Text

display.ColumnCount = 10
display.RowSource = "A1:J100"

'update the balance sheet'
Worksheets("balance_sheet").Activate
'subtract from cash account
Range("B4") = Range("B4") - Val(amount.Text)

'update date so that time can pass in my simulation
Range("A2") = date_of_payment.Text


End Sub







Private Sub TextBox2_Change()

End Sub
