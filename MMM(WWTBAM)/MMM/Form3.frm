VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ask The Audience"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   1575
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4680
   Begin VB.Shape a 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   480
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape d 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   3720
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape c 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   2640
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape b 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   1560
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "D"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "C"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "B"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "A"
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
      TabIndex        =   0
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   3855
      Left            =   480
      Top             =   720
      Width           =   3855
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label5_Click()
Form1.Show
Form3.Hide
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
