VERSION 5.00
Begin VB.UserControl superText 
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   ScaleHeight     =   3210
   ScaleWidth      =   4170
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "superText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Mesut Akcan
'© 2001 - 2020
' 17 Kasým 2020
'https://akcansoft.blogspot.com

Private Sub UserControl_InitProperties()
Text1.Text = UserControl.Name
' ya da
' Text1.Text= "superText"
End Sub
Private Sub UserControl_Resize()
With Text1
    'Text1'in boyutunu eklenen kontrol'a eþitle
    .Height = UserControl.ScaleHeight
    .Width = UserControl.ScaleWidth
End With
End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = vbRed 'kýrmýzý
    Text1.ForeColor = vbWhite 'beyaz
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = vbWindowBackground '&H80000005 'Zemin rengi
    Text1.ForeColor = vbWindowText '&H80000008 'Metin rengi
End Sub

Public Property Get Metin() As String
    Metin = Text1.Text
End Property

Public Property Let Metin(ByVal yeniDeger As String)
    Text1.Text = yeniDeger
End Property
