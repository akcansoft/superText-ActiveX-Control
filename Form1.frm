VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin Project1.superText superText5 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Metin           =   "superText5"
      Giris           =   4
   End
   Begin Project1.superText superText4 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Metin           =   "superText4"
      Giris           =   3
   End
   Begin Project1.superText superText3 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Metin           =   "superText3"
      Giris           =   2
   End
   Begin Project1.superText superText2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Metin           =   "superText2"
      Giris           =   1
   End
   Begin Project1.superText superText1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Metin           =   "superText1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "https://akcansoft.blogspot.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   2700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    superText1.Metin = "Merhaba"
End Sub

Private Sub superText1_Change()
    List1.AddItem superText1.Metin
End Sub

Private Sub superText2_KeyPress(KeyAscii As Integer)
    superText1.Metin = superText1.Metin & Chr(KeyAscii)
End Sub
