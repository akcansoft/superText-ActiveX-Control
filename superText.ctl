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
'� 2001 - 2020
'https://akcansoft.blogspot.com

Dim mGiris As Byte 'giri� de�eri
Enum stGiris
    Hepsi
    Say�
    Harf
    [B�y�k Harf]
    [K���k Harf]
End Enum

Event Change()
Event KeyPress(KeyAscii As Integer)

Private Sub UserControl_InitProperties()
    'Text1.Text = UserControl.Name
    Text1.Text = Extender.Name
End Sub

Private Sub UserControl_Resize()
With Text1
    'Text1'in boyutunu eklenen kontrol'a e�itle
    .Height = UserControl.ScaleHeight
    .Width = UserControl.ScaleWidth
End With
End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = vbYellow 'sar�
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = vbWindowBackground '&H80000005 'Zemin rengi
End Sub

Public Property Get Metin() As String
    Metin = Text1.Text
End Property

Public Property Let Metin(ByVal yeniDeger As String)
    Text1.Text = yeniDeger
End Property

Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1.Text = PropBag.ReadProperty("Metin", UserControl.Name)
    mGiris = PropBag.ReadProperty("Giris", 0)
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Metin", Text1.Text, UserControl.Name)
    Call PropBag.WriteProperty("Giris", mGiris, 0)
End Sub

Public Property Get Giris() As stGiris
    Giris = mGiris
End Property

Public Property Let Giris(ByVal vNewValue As stGiris)
    mGiris = vNewValue
    PropertyChanged "Giris"
End Property

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Dim keyErr As Boolean 'hatal� giri�
    If KeyAscii = 8 Then Exit Sub 'TAB tu�u ise
    If KeyAscii = vbKeyReturn Then 'ENTER tu�u ise
        KeyAscii = 0
        SendKeys "{TAB}", True 'TAB tu�vuru�u g�nder
        Exit Sub
    End If
    Select Case mGiris
        Case 1: 'Say�
            '0-9 ve . harici ise hata
            If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 46 Then
                keyErr = True
            End If
            ' metinde . varsa . ise hata
            If KeyAscii = 46 And InStr(Text1, ".") Then keyErr = True
        Case 2: 'Harf
            '0-9 ise hata
            If IsNumeric(Chr(KeyAscii)) Then keyErr = True
        Case 3: 'BuyukHarf
            If KeyAscii = 253 Then
                KeyAscii = 73 '� => I
            ElseIf KeyAscii = 105 Then
                KeyAscii = 221 'i => �
            Else
                'k���k harfi b�y�k harfe d�n��t�r
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        Case 4: 'KucukHarf
            If KeyAscii = 73 Then
                KeyAscii = 253 'I => �
            ElseIf KeyAscii = 221 Then
                KeyAscii = 105 '� => i
            Else
                'b�y�k harfi k���k harfe d�n��t�r
                KeyAscii = Asc(LCase(Chr(KeyAscii)))
            End If
    End Select
    'hatal� giri� varsa
    If keyErr = True Then KeyAscii = 0: Beep
End Sub

Private Sub Text1_Change()
    RaiseEvent Change
End Sub

