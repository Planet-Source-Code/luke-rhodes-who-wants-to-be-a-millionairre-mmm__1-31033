Attribute VB_Name = "Module3"
Option Explicit

Sub highlightmnu()
If a = 1 Then
Form1.Label1.BackColor = &H80FF&
Form2.Label2.Caption = "£100"
ElseIf a = 2 Then
Form1.Label13.BackColor = &H80FF&
Form1.Label1.BackColor = &H8000000D
Form2.Label2.Caption = "£200"

ElseIf a = 3 Then
Form1.Label12.BackColor = &H80FF&
Form1.Label13.BackColor = &H8000000D
Form2.Label2.Caption = "£300"

ElseIf a = 4 Then
Form1.Label11.BackColor = &H80FF&
Form1.Label12.BackColor = &H8000000D
Form2.Label2.Caption = "£500"

ElseIf a = 5 Then
Form1.Label10.BackColor = &H80FF&
Form1.Label11.BackColor = &H8000000D
Form1.Label10.ForeColor = &H0&
Form2.Label2.Caption = "£1,000"

ElseIf a = 6 Then
Form1.Label9.BackColor = &H80FF&
Form1.Label10.BackColor = &H8000000D
Form1.Label10.ForeColor = &H80FF&
Form2.Label2.Caption = "£2,000"

ElseIf a = 7 Then
Form1.Label8.BackColor = &H80FF&
Form1.Label9.BackColor = &H8000000D
Form2.Label2.Caption = "£4,000"


ElseIf a = 8 Then
Form1.Label7.BackColor = &H80FF&
Form1.Label8.BackColor = &H8000000D
Form2.Label2.Caption = "£8,000"


ElseIf a = 9 Then
Form1.Label6.BackColor = &H80FF&
Form1.Label7.BackColor = &H8000000D
Form2.Label2.Caption = "£16,000"


ElseIf a = 10 Then
Form1.Label5.BackColor = &H80FF&
Form1.Label6.BackColor = &H8000000D
Form1.Label5.ForeColor = &H0&
Form2.Label2.Caption = "£32,000"

ElseIf a = 11 Then
Form1.Label4.BackColor = &H80FF&
Form1.Label5.BackColor = &H8000000D
Form1.Label5.ForeColor = &H80FF&
Form2.Label2.Caption = "£64,000"

ElseIf a = 12 Then
Form1.Label3.BackColor = &H80FF&
Form1.Label4.BackColor = &H8000000D
Form2.Label2.Caption = "£125,000"

ElseIf a = 13 Then
Form1.Label2.BackColor = &H80FF&
Form1.Label3.BackColor = &H8000000D
Form2.Label2.Caption = "£250,000"

ElseIf a = 14 Then
Form1.Label15.BackColor = &H80FF&
Form1.Label2.BackColor = &H8000000D
Form2.Label2.Caption = "£500,000"

ElseIf a = 15 Then
Form1.Label14.BackColor = &H80FF&
Form1.Label15.BackColor = &H8000000D
Form1.Label14.ForeColor = &H0&
Form2.Label2.Caption = "£1,000,000"
Form1.Timer1.Enabled = True



End If

End Sub
