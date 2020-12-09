VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1125
   ScaleHeight     =   630
   ScaleWidth      =   1125
   Begin VB.VScrollBar VScroll1 
      Height          =   405
      Left            =   720
      Max             =   40
      Min             =   -10
      TabIndex        =   1
      Top             =   120
      Value           =   15
      Width           =   265
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Mesut Akcan
'© 3 Kasým 2001
'Güncelleme: 12 Kasým 2020
'
'https://akcansoft.blogspot.com
'makcan@ gmail.com

Private Sub UserControl_Initialize()
EnB = VScroll1.Max
EnK = VScroll1.Min
VScroll1.Max = -EnK
VScroll1.Min = -EnB
VScroll1.Value = -VScroll1.Value
End Sub

Private Sub VScroll1_Change()
Text1.Text = -VScroll1.Value
End Sub
