VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} show_employees 
   Caption         =   "UserForm1"
   ClientHeight    =   10836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17808
   OleObjectBlob   =   "show_employees.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "show_employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub

Private Sub exit_window_Click()

Unload Me
main_menu.Show

End Sub

Private Sub show_data_Click()
Workbooks("finances.xlsm").Activate
Worksheets("customers").Activate

display.ColumnCount = 10
display.RowSource = "A1:J100"
End Sub

Private Sub display_Click()

End Sub


Private Sub UserForm_Click()

End Sub
