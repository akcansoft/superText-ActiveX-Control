VERSION 5.00
Begin VB.UserControl superText 
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   DataSourceBehavior=   1  'vbDataSource
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
'4/12/2020
'https://akcansoft.blogspot.com
'makcan@gmail.com

Dim gZeminRengi As OLE_COLOR 'geçici zemin rengi
Dim gYaziRengi As OLE_COLOR 'geçici zemin rengi
Enum e_Giris
    Hepsi
    Sayý
    Harf
    [Büyük Harf]
    [Küçük Harf]
End Enum
Enum e_Appearance
    Flat
    [3D]
End Enum
Enum e_Alignment
    [Left Justify]
    [Right Justify]
    Center
End Enum
Enum e_BorderStyle
    None
    [Fixed Single]
End Enum
Enum e_LinkMode
    None
    Automatic
    Manual
    Notify
End Enum
Enum e_MousePointer
    Default
    Arrow
    Cross
    [I-Beam]
    Icon
    Size
    [Size NE SW]
    [Size N S]
    [Size NW SE]
    [Size W E]
    [Up Arrow]
    Hourglass
    [No Drop]
    [Arrow and Hourglass]
    [Arrow and Question]
    [Size All]
    Custom = 99
End Enum
Enum e_OLEDragMode
    Manual
    Automatic
End Enum
Enum e_OLEDropMode
    None
    Manual
    Automatic
End Enum
'Enum e_ScrollBars
'    None
'    Horizontal
'    Vertical
'    Both
'End Enum

'Default Property Values:
Const m_def_Giris = 0
Const m_def_OdakZeminRengi = vbWindowBackground
Const m_def_OdakYaziRengi = vbWindowText

'Property Variables:
Dim m_Giris As e_Giris
Dim m_OdakZeminRengi As OLE_COLOR
Dim m_OdakYaziRengi As OLE_COLOR

'Event Declarations:
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=Text1,Text1,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=Text1,Text1,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=Text1,Text1,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=Text1,Text1,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=Text1,Text1,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event DblClick() 'MappingInfo=Text1,Text1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click() 'MappingInfo=Text1,Text1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600

Private Sub UserControl_InitProperties()
    Text1.Text = Extender.Name
    m_Giris = m_def_Giris
    m_OdakZeminRengi = m_def_OdakZeminRengi 'odak zemin rengine varsayýlan deðeri ata
    m_OdakYaziRengi = m_def_OdakYaziRengi 'odak yazý rengine varsayýlan deðeri ata
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
    Text1.BackColor = m_OdakZeminRengi 'odak zemin rengi
    Text1.ForeColor = m_OdakYaziRengi 'odak yazý rengi
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = gZeminRengi 'geçici zemin rengi
    Text1.ForeColor = gYaziRengi 'geçici yazý rengi
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Dim keyErr As Boolean 'hatalý giriþ
    If KeyAscii = 8 Then Exit Sub 'TAB tuþu ise
    If KeyAscii = vbKeyReturn Then 'ENTER tuþu ise
        KeyAscii = 0
        SendKeys "{TAB}", True 'TAB tuþvuruþu gönder
        Exit Sub
    End If
    Select Case m_Giris
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

Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1.Text = PropBag.ReadProperty("Metin", UserControl.Name)
    Text1.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Text1.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Text1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Text1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    Text1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    Text1.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
    Text1.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
    Text1.LinkMode = PropBag.ReadProperty("LinkMode", 0)
    Text1.LinkItem = PropBag.ReadProperty("LinkItem", "")
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Text1.DataMember = PropBag.ReadProperty("DataMember", "")
    Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Text1.Appearance = PropBag.ReadProperty("Appearance", 1)
    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Text1.Text = PropBag.ReadProperty("Metin", "Text1")
    m_Giris = PropBag.ReadProperty("Giris", m_def_Giris)
    m_OdakZeminRengi = PropBag.ReadProperty("OdakZeminRengi", m_def_OdakZeminRengi)
    m_OdakYaziRengi = PropBag.ReadProperty("OdakYaziRengi", m_def_OdakYaziRengi)
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Metin", Text1.Text, UserControl.Name)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, vbWindowBackground)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, vbWindowText)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("RightToLeft", Text1.RightToLeft, False)
    Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
    Call PropBag.WriteProperty("OLEDropMode", Text1.OLEDropMode, 0)
    Call PropBag.WriteProperty("OLEDragMode", Text1.OLEDragMode, 0)
    Call PropBag.WriteProperty("MousePointer", Text1.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("LinkTopic", Text1.LinkTopic, "")
    Call PropBag.WriteProperty("LinkTimeout", Text1.LinkTimeout, 50)
    Call PropBag.WriteProperty("LinkMode", Text1.LinkMode, 0)
    Call PropBag.WriteProperty("LinkItem", Text1.LinkItem, "")
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataMember", Text1.DataMember, "")
    Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
    Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
    Call PropBag.WriteProperty("Appearance", Text1.Appearance, 1)
    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("Giris", m_Giris, m_def_Giris)
    Call PropBag.WriteProperty("OdakZeminRengi", m_OdakZeminRengi, m_def_OdakZeminRengi)
    Call PropBag.WriteProperty("OdakYaziRengi", m_OdakYaziRengi, m_def_OdakYaziRengi)
End Sub

'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub Text1_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ScrollBars
'Public Property Get ScrollBars() As e_ScrollBars
'    ScrollBars = Text1.ScrollBars
'End Property
'
'Public Property Let ScrollBars(ByVal New_ScrollBars As e_ScrollBars)
'    Text1.ScrollBars() = New_ScrollBars
'    PropertyChanged "ScrollBars"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = Text1.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    Text1.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    Text1.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Private Sub Text1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Text1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,OLEDropMode
Public Property Get OLEDropMode() As e_OLEDropMode
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = Text1.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As e_OLEDropMode)
    Text1.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,OLEDragMode
Public Property Get OLEDragMode() As e_OLEDragMode
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = Text1.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As e_OLEDragMode)
    Text1.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Private Sub Text1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MultiLine
'Public Property Get MultiLine() As Boolean
'    MultiLine = Text1.MultiLine
'End Property

'Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
'    Text1.MultiLine() = New_MultiLine
'    PropertyChanged "MultiLine"
'End Property

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MousePointer
Public Property Get MousePointer() As e_MousePointer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Text1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As e_MousePointer)
    Text1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Text1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
    LinkTopic = Text1.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    Text1.LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
    LinkTimeout = Text1.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    Text1.LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,LinkMode
Public Property Get LinkMode() As e_LinkMode
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
    LinkMode = Text1.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As e_LinkMode)
    Text1.LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
    LinkItem = Text1.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    Text1.LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Specifies whether the selection in a Masked edit control is hidden when the control loses focus."
    HideSelection = Text1.HideSelection
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub Text1_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    DataMember = Text1.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    Text1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,BorderStyle
Public Property Get BorderStyle() As e_BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As e_BorderStyle)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Appearance
Public Property Get Appearance() As e_Appearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = Text1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As e_Appearance)
    Text1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As e_Alignment
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As e_Alignment)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Metin() As String
Attribute Metin.VB_Description = "Returns/sets the text contained in the control."
    Metin = Text1.Text
End Property

Public Property Let Metin(ByVal New_Metin As String)
    Text1.Text() = New_Metin
    PropertyChanged "Metin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,0
Public Property Get Giris() As e_Giris
Attribute Giris.VB_Description = "Klavye giriþ tuþlarý kýsýtlamasý"
    Giris = m_Giris
End Property

Public Property Let Giris(ByVal New_Giris As e_Giris)
    m_Giris = New_Giris
    PropertyChanged "Giris"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWindowBackground
Public Property Get OdakZeminRengi() As OLE_COLOR
Attribute OdakZeminRengi.VB_Description = "Odak, kontrolde iken zemin rengi"
    OdakZeminRengi = m_OdakZeminRengi
End Property

Public Property Let OdakZeminRengi(ByVal New_OdakZeminRengi As OLE_COLOR)
    m_OdakZeminRengi = New_OdakZeminRengi
    PropertyChanged "OdakZeminRengi"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWindowText
Public Property Get OdakYaziRengi() As OLE_COLOR
Attribute OdakYaziRengi.VB_Description = "Odak, kontrolde iken yazý rengi"
    OdakYaziRengi = m_OdakYaziRengi
End Property

Public Property Let OdakYaziRengi(ByVal New_OdakYaziRengi As OLE_COLOR)
    m_OdakYaziRengi = New_OdakYaziRengi
    PropertyChanged "OdakYaziRengi"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Sub Hakkinda()
Attribute Hakkinda.VB_Description = "superText ActiveX kontrol hakkýnda"
Attribute Hakkinda.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub
