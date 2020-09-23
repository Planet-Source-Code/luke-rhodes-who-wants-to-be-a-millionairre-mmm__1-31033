VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Money, Money, Money (What Have You Won?)"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleMode       =   0  'User
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "TOTAL PRIZE MONEY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "We have located where you live through the goverment and your check will be posted to you within the week."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   0
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000080FF&
      Height          =   3015
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   10830
      Left            =   0
      Top             =   0
      Width           =   11910
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\pics\loading_screen.jpg")
Label2.Alignment = vbCenter

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label4_Click()
End
End Sub
