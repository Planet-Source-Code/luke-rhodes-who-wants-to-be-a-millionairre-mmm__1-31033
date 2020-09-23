VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phone A Friend"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   3600
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5685
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\pics\phone.jpg")
Image2.Picture = LoadPicture(App.Path & "\pics\loading_screen.jpg")
Label2.Caption = "Hi Dave It's Chris Tarrent here. He's ran into a spot of badluck and so mayb you could help him."
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

Private Sub Label3_Click()
Form1.Show
Form4.Hide
End Sub
