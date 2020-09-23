VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Money, Money, Money"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   1200
   End
   Begin VB.Label Label28 
      BackColor       =   &H80000002&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   27
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label27 
      BackColor       =   &H80000002&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label26 
      BackColor       =   &H80000002&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   25
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label25 
      BackColor       =   &H80000002&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000002&
      Caption         =   "Label21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   23
      Top             =   7200
      Width           =   3495
   End
   Begin VB.Label Label23 
      BackColor       =   &H80000002&
      Caption         =   "Label21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   22
      Top             =   7200
      Width           =   3495
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000002&
      Caption         =   "Label21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000002&
      Caption         =   "Label21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   20
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label Label18 
      BackColor       =   &H000080FF&
      Caption         =   "              Instructions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   840
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   360
      Shape           =   2  'Oval
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1080
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Caption         =   "    W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   4575
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14:         £500,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15:         £1,000,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2:           £200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3:           £300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4:           £500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5:           £1,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6:           £2,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7            £4,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8:           £8,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9:           £16,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10:         £32,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:         £64,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12:         £128,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13:         £250,000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:           £100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   9720
      Shape           =   2  'Oval
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H000080FF&
      Caption         =   " 50:50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   360
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   360
      Shape           =   2  'Oval
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   360
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Label20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   19
      Top             =   4440
      Width           =   8775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   9255
   End
   Begin VB.Label Label19 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3360
      TabIndex        =   18
      Top             =   840
      Width           =   3015
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   10830
      Left            =   0
      Top             =   0
      Width           =   11910
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Sub highlight()
Call Module3.highlightmnu
End Sub


Private Sub About_Click()

frmAbout.Show
End Sub

Private Sub Form_Load()

Image1.Picture = LoadPicture(App.Path & "\pics\phone.jpg")
Image3.Picture = LoadPicture(App.Path & "\pics\loading_screen.jpg")
Image2.Picture = LoadPicture(App.Path & "\pics\askaudience.jpg")
Label19.Caption = "A question will apear in the box and will have 4 answers to choose from below it. Slect a anser and a message will appear to confirm your actions. There will be 15 random questions that get harder as the money amount rises. There are 3 Safe points where if you reach them you will deffinatley get a check for that amount. You can choose out of 3 'Life Lines' that will help you proggress."
Form2.Label2.Caption = "£0"
ftyfty = True
Randomize
q = Int((2 * Rnd))
Call Module4.whichr
End Sub

Private Sub Image1_Click()
Image1.Enabled = False
Form4.Show
Form1.Hide
Form4.Label1.Caption = " Yeh I think I can help him. Im pretty sure it's " & UCase$(an)
End Sub

Private Sub Image2_Click()
Call Module5.sd
End Sub

Private Sub Label16_Click()
Dim Msg, Style, Title, Response
Msg = "Are You Sure That You Want To Quit?"
Style = vbYesNo
Title = "Are You Sure?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
Me.Hide
Form2.Show
End If

End Sub

Private Sub Label17_Click()
Call Module1.fiftyfifty
End Sub

Private Sub Label21_Click()
Dim Msg, Style, Title, Response
Msg = "Are You Sure?"
Style = vbYesNo
Title = "Are You Sure?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
If an = "a" Then
Form1.Label21.Enabled = True
Form1.Label22.Enabled = True
Form1.Label23.Enabled = True
Form1.Label24.Enabled = True
Call Module2.lbl21
Else
Form2.Show
Form1.Hide
If a < 5 Then
Form2.Label2.Caption = "£0"
ElseIf a < 10 And a >= 5 Then
Form2.Label2.Caption = "£1,000"
ElseIf a < 15 And a >= 10 Then
Form2.Label2.Caption = "£1,000"
End If
End If
End If


End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim Msg, Style, Title, Response
Msg = "Are You Sure That You Want To Quit?"
Style = vbYesNo
Title = "Are You Sure?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
Me.Hide
Form2.Show
End If
End Sub
Private Sub Label22_Click()
Dim Msg, Style, Title, Response
Msg = "Are You Sure?"
Style = vbYesNo
Title = "Are You Sure?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
If an = "b" Then
Form1.Label21.Enabled = True
Form1.Label22.Enabled = True
Form1.Label23.Enabled = True
Form1.Label24.Enabled = True
Call Module2.lbl22
Else
Form2.Show
Form1.Hide
If a < 5 Then
Form2.Label2.Caption = "£0"
ElseIf a < 10 And a >= 5 Then
Form2.Label2.Caption = "£1,000"
ElseIf a < 15 And a >= 10 Then
Form2.Label2.Caption = "£1,000"
End If
End If
End If

End Sub

Private Sub Label23_Click()
Dim Msg, Style, Title, Response
Msg = "Are You Sure?"
Style = vbYesNo
Title = "Are You Sure?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
If an = "c" Then
Form1.Label21.Enabled = True
Form1.Label22.Enabled = True
Form1.Label23.Enabled = True
Form1.Label24.Enabled = True
Call Module2.lbl23
Else
Form2.Show
Form1.Hide
If a < 5 Then
Form2.Label2.Caption = "£0"
ElseIf a < 10 And a >= 5 Then
Form2.Label2.Caption = "£1,000"
ElseIf a < 15 And a >= 10 Then
Form2.Label2.Caption = "£1,000"
End If
End If
End If
End Sub

Private Sub Label24_Click()
Dim Msg, Style, Title, Response
Msg = "Are You Sure?"
Style = vbYesNo
Title = "Are You Sure?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
If an = "d" Then
Form1.Label21.Enabled = True
Form1.Label22.Enabled = True
Form1.Label23.Enabled = True
Form1.Label24.Enabled = True
Call Module2.lbl24
Else
Form2.Show
Form1.Hide
If a < 5 Then
Form2.Label2.Caption = "£0"
ElseIf a < 10 And a >= 5 Then
Form2.Label2.Caption = "£1,000"
ElseIf a < 15 And a >= 10 Then
Form2.Label2.Caption = "£1,000"
End If
End If
End If
End Sub

Private Sub Label29_Click()

Form2.Show
Form1.Hide

End Sub

Private Sub Label30_Click()

End Sub



Private Sub Timer1_Timer()
Form1.Hide
Form2.Show
End Sub
