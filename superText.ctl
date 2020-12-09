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
'©2020 Mesut Akcan
'01/12/2020
'https://akcansoft.blogspot.com
'makcan@gmail.com

Dim mGiris As Byte 'giriþ deðeri
Dim ozr As OLE_COLOR 'odak zemin rengi
Dim oyr As OLE_COLOR 'odak yazý rengi
Dim gZeminRengi As OLE_COLOR 'geçici zemin rengi
Dim gYaziRengi As OLE_COLOR 'geçici zemin rengi
Enum stGiris
    Hepsi
    Sayý
    Harf
    [Büyük Harf]
    [Küçük Harf]
End Enum

Event Change()
Event KeyPress(KeyAscii As Integer)

Private Sub UserControl_InitProperties()
    Text1.Text = Extender.Name
    ozr = vbWindowBackground 'odak zemin rengine varsayýlan deðeri ata
    oyr = vbWindowText 'odak yazý rengine varsayýlan deðeri ata
End Sub

Private Sub UserControl_Resize()
With Text1
    'Text1'in boyutunu eklenen kontrol'a eþitle
    .Height = UserControl.ScaleHeight
    .Width = UserControl.ScaleWidth
End With
End Sub

Private Sub Text1_GotFocus()
    gZeminRengi = Text1.BackColor
    gYaziRengi = Text1.ForeColor
    Text1.BackColor = ozr 'odak zemin rengi
    Text1.ForeColor = oyr 'odak yazý rengi
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = gZeminRengi 'geçici zemin rengi
    Text1.ForeColor = gYaziRengi 'geçici yazý rengi
End Sub

Public Property Get Metin() As String
Attribute Metin.VB_Description = "Kontrolün içerdiði metin"
    Metin = Text1.Text
End Property

Public Property Let Metin(ByVal yeniDeger As String)
    Text1.Text = yeniDeger
End Property

Public Property Get Giris() As stGiris
Attribute Giris.VB_Description = "Klavye giriþ tuþlarý kýsýtlamasý"
    Giris = mGiris
End Property

Public Property Let Giris(ByVal vNewValue As stGiris)
    mGiris = vNewValue
    PropertyChanged "Giris"
End Property

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Dim keyErr As Boolean 'hatalý giriþ
    If KeyAscii = 8 Then Exit Sub 'TAB tuþu ise
    If KeyAscii = vbKeyReturn Then 'ENTER tuþu ise
        KeyAscii = 0
        SendKeys "{TAB}", True 'TAB tuþvuruþu gönder
        Exit Sub
    End If
    Select Case mGiris
        Case 1: 'Sayý
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
                KeyAscii = 73 'ý => I
            ElseIf KeyAscii = 105 Then
                KeyAscii = 221 'i => Ý
            Else
                'küçük harfi büyük harfe dönüþtür
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        Case 4: 'KucukHarf
            If KeyAscii = 73 Then
                KeyAscii = 253 'I => ý
            ElseIf KeyAscii = 221 Then
                KeyAscii = 105 'Ý => i
            Else
                'büyük harfi küçük harfe dönüþtür
                KeyAscii = Asc(LCase(Chr(KeyAscii)))
            End If
    End Select
    'hatalý giriþ varsa
    If keyErr = True Then KeyAscii = 0: Beep
End Sub

Private Sub Text1_Change()
    RaiseEvent Change
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get OdakZeminRengi() As OLE_COLOR
Attribute OdakZeminRengi.VB_Description = "Odak, kontrolde iken zemin rengi"
  OdakZeminRengi = ozr
End Property

Public Property Let OdakZeminRengi(ByVal c As OLE_COLOR)
  ozr = c
  PropertyChanged "OdakZeminRengi"
End Property

Public Property Get OdakYaziRengi() As OLE_COLOR
  OdakYaziRengi = oyr
End Property

Public Property Let OdakYaziRengi(ByVal c As OLE_COLOR)
  oyr = c
  PropertyChanged "OdakYaziRengi"
End Property

Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1.Text = PropBag.ReadProperty("Metin", UserControl.Name)
    mGiris = PropBag.ReadProperty("Giris", 0)
    Text1.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    ozr = PropBag.ReadProperty("OdakZeminRengi", vbWindowBackground)
    oyr = PropBag.ReadProperty("OdakYaziRengi", vbWindowText)
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Metin", Text1.Text, UserControl.Name)
    Call PropBag.WriteProperty("Giris", mGiris, 0)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, vbWindowBackground)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, vbWindowText)
    Call PropBag.WriteProperty("OdakZeminRengi", ozr, vbWindowBackground)
    Call PropBag.WriteProperty("OdakYaziRengi", oyr, vbWindowText)
End Sub

'Sub Hakkinda()
'    MsgBox App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & _
'        vbCrLf & "©" & Year(Now) & " " & App.CompanyName & vbCrLf & vbCrLf & _
'        App.Comments, , UserControl.Name & " Hakkýnda"
'End Sub

Sub Hakkinda()
Attribute Hakkinda.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

