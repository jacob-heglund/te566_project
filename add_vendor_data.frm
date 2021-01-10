VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_vendor_data 
   Caption         =   "UserForm1"
   ClientHeight    =   10836
   ClientLeft      =   -348
   ClientTop       =   -1296
   ClientWidth     =   17808
   OleObjectBlob   =   "add_vendor_data.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "add_vendor_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO add a screen that shows the database

Private Sub Frame1_Click()

End Sub

Private Sub CommandButton1_Click()

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
Private Sub submit_customer_data_Click()
'TODO get this working
If Range("a2") <> "" Then
Rows("2:2").Select
Selection.Insert shift:=xlDown
End If
'TODO check that all the fields are filled
'TODO double check submission (avoid mistakes)
'TODO give feedback that it was submitted correctly

If Range("a2") = "" Then
Range("a2").Offset(0, 0) = company_name.Text
Range("a2").Offset(0, 1) = first_name.Text
Range("a2").Offset(0, 2) = last_name.Text
Range("a2").Offset(0, 3) = address_1.Text
Range("a2").Offset(0, 4) = address_2.Text
Range("a2").Offset(0, 5) = city.Text
Range("a2").Offset(0, 6) = state.Text
Range("a2").Offset(0, 7) = zip_code.Text
Range("a2").Offset(0, 8) = price.Text

display.ColumnCount = 10
display.RowSource = "B1:J100"
End If



End Sub

