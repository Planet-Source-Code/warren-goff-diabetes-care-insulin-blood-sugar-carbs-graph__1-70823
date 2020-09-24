VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.UserControl BBTime 
   BackColor       =   &H8000000A&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ScaleHeight     =   2880
   ScaleWidth      =   3540
   ToolboxBitmap   =   "BBTime.ctx":0000
   Begin VB.Frame frmFrame 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1635
      Begin VB.TextBox txtAMPM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "PM"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMinute 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   695
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "00"
         Top             =   0
         Width           =   315
      End
      Begin VB.TextBox txtHour 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   260
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "BBTime.ctx":0312
         Top             =   0
         Width           =   300
      End
      Begin ComCtl2.UpDown udMinute 
         Height          =   285
         Left            =   1005
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udHour 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   555
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "BBTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Authors: Bobby Orr & Brandon Lackey
'
'Notes: We slammed this thing together for a customer service app in about 4
'hours or so. It's really only missing a resize function. A slightly modified
'version is still in use today for a very large call center.
'
'brandonlackey@hotmail.com

Public Enum bbTimeType
    bbMilitaryTime = 0
    bbStandardTime = 1
End Enum
Private pbMinuteUp As Boolean
Private pbHourUp As Boolean
'Default Property Values:
Const m_def_ShowMeridiem = 0
Const m_def_TimeType = 0
Const m_def_Text = "12:00"
Const m_def_fontboldBold = False
Const m_def_Meridiem = "PM"
'Property Variables:
Dim m_ShowMeridiem As Boolean
Dim m_TimeType As bbTimeType
Dim m_Text As String
Dim m_Font As New StdFont

Private Sub txtAMPM_Change()
Me.Text = txtHour.Text & ":" & txtMinute.Text & " " & txtAMPM.Text
End Sub

Private Sub txtAMPM_Click()
    If txtAMPM.Text = "AM" Then
        txtAMPM.Text = "PM"
    Else
        txtAMPM.Text = "AM"
    End If
    txtAMPM.SelStart = 0
    txtAMPM.SelLength = Len(txtAMPM.Text)
End Sub

Private Sub txtAMPM_GotFocus()
txtAMPM.SelStart = 0
txtAMPM.SelLength = Len(txtAMPM.Text)
End Sub

Private Sub txtAMPM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If txtAMPM.Text = "AM" Then
            txtAMPM.Text = "PM"
        Else
            txtAMPM.Text = "AM"
        End If
        KeyCode = 0
        txtAMPM.SelStart = 0
        txtAMPM.SelLength = Len(txtAMPM.Text)
    Else
        KeyCode = 0
    End If
End Sub

Private Sub txtHour_Change()
    If Trim(txtHour.Text) <> "" Then
        If Me.TimeType = bbMilitaryTime Then
            If CInt(txtHour.Text) > 23 Then
                txtHour.Text = "0"
            End If
            If CInt(txtHour.Text) < 0 Then
                txtHour.Text = "23"
            End If
            If CInt(txtHour.Text) > 11 Then
                txtAMPM.Text = "PM"
            Else
                txtAMPM.Text = "AM"
            End If
        Else
            'If we wanted to flip the AM/PM box this is where we would do it!!!
            If CInt(txtHour.Text) > 12 Then
                txtHour.Text = "1"
            End If
            If CInt(txtHour.Text) < 1 Then
                txtHour.Text = "12"
            End If
        End If
        Me.Text = txtHour.Text & ":" & txtMinute.Text & " " & txtAMPM.Text
    End If
End Sub

Private Sub txtHour_GotFocus()
    txtHour.SelStart = 0
    txtHour.SelLength = Len(txtHour.Text)
End Sub

Private Sub txtHour_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        udHour_DownClick
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        udHour_UpClick
        KeyCode = 0
    End If
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtHour_LostFocus()
    If Trim(txtHour.Text) = "" Then
        txtHour.Text = "12"
    End If
End Sub

Private Sub txtMinute_Change()
    If Trim(txtMinute.Text) <> "" Then
        If CInt(txtMinute.Text) > 59 Then txtMinute.Text = "00"
        If CInt(txtMinute.Text) < 0 Then txtMinute.Text = "59"
        If Len(Trim(txtMinute.Text)) = 2 Then
            Me.Text = txtHour.Text & ":" & txtMinute.Text & " " & txtAMPM.Text
        End If
    End If
End Sub

Private Sub txtMinute_GotFocus()
    txtMinute.SelStart = 0
    txtMinute.SelLength = Len(txtMinute.Text)
End Sub

Private Sub txtMinute_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        udMinute_DownClick
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        udMinute_UpClick
        KeyCode = 0
    End If
End Sub

Private Sub txtMinute_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0

End Sub

Private Sub txtMinute_LostFocus()
    Select Case Len(txtMinute.Text)
        Case 1
            txtMinute.Text = "0" & txtMinute.Text
        Case 0
            txtMinute.Text = "00"
        Case Is > 2
            MsgBox "How did this happen!"
    End Select
End Sub

Private Sub udHour_DownClick()
pbHourUp = False
txtHour.Text = CStr(CInt(txtHour.Text) - 1)
txtHour.SetFocus
txtHour.SelStart = 0
txtHour.SelLength = Len(txtHour.Text)
End Sub

Private Sub udHour_UpClick()
    pbHourUp = True
    txtHour.Text = CStr(CInt(txtHour.Text) + 1)
    txtHour.SetFocus
    txtHour.SelStart = 0
    txtHour.SelLength = Len(txtHour.Text)
End Sub

Private Sub udMinute_DownClick()
    pbMinuteUp = False
    txtMinute.Text = CStr(CInt(txtMinute.Text) - 1)
    If CInt(txtMinute.Text) < 10 Then txtMinute.Text = "0" & txtMinute.Text
    txtMinute.SetFocus
    txtMinute.SelStart = 0
    txtMinute.SelLength = Len(txtMinute.Text)
End Sub

Private Sub udMinute_UpClick()
pbMinuteUp = True
txtMinute.Text = CStr(CInt(txtMinute.Text) + 1)
If CInt(txtMinute.Text) < 10 Then txtMinute.Text = "0" & txtMinute.Text
txtMinute.SetFocus
txtMinute.SelStart = 0
txtMinute.SelLength = Len(txtMinute.Text)
End Sub

Private Sub UserControl_Resize()

    UserControl.Width = frmFrame.Width
    UserControl.Height = frmFrame.Height


End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,12:00
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Dim iStart As Integer
    Dim iLen As Integer
    Dim sHour As String
    Dim sMinute As String
    Dim sMeridiem As String
    Dim iHourLen As Integer
    
    iLen = InStr(1, New_Text, ":", vbTextCompare)
    If iLen = 0 Then
        MsgBox "Invalid Time Format Entered"
        Exit Property
    End If
    If iLen = 2 Then
        iHourLen = 1
    Else
        iHourLen = 2
    End If
    sHour = Mid(New_Text, iLen - iHourLen, iHourLen)
    iStart = iLen + 1
    sMinute = Mid(New_Text, iStart, 2)
    sMeridiem = UCase(Right(New_Text, 2))
    
    'hour validation
    If TimeType = bbMilitaryTime Then
        If IsNumeric(sHour) Then
            sHour = CheckHour(sHour, sMeridiem)
            If CInt(sHour) < 0 Or CInt(sHour) > 23 Then
                MsgBox "Hour Out Of Range: " & sHour
                Exit Property
            Else
                txtHour.Text = sHour
            End If
        Else
            MsgBox "Invalid Hour Error: " & sHour
            Exit Property
        End If
    Else
        If IsNumeric(sHour) Then
            sHour = CheckHour(sHour, sMeridiem)
            If CInt(sHour) < 1 Or CInt(sHour) > 12 Then
                MsgBox "Hour Out Of Standard Time Range: " & sHour
                Exit Property
            Else
                txtHour.Text = sHour
            End If
        Else
            MsgBox "Invalid Hour Error: " & sHour
            Exit Property
        End If
    End If
    'minute validation
    If IsNumeric(sMinute) Then
        If CInt(sMinute) < 0 Or CInt(sMinute) > 59 Then
            MsgBox "Minutes Out Of Range: " & sMinute
            Exit Property
        Else
            txtMinute.Text = sMinute
        End If
    Else
        MsgBox "Invalid Minute Error: " & sMinute
        Exit Property
    End If
    'Meridiem validation
    If sMeridiem = "AM" Or sMeridiem = "PM" Then
        txtAMPM.Text = sMeridiem
    End If
    If Me.ShowMeridiem Then
        m_Text = sHour & ":" & sMinute & " " & sMeridiem
    Else
        m_Text = sHour & ":" & sMinute
    End If
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtHour,txtHour,-1,Text
Public Property Get Hour() As String
Attribute Hour.VB_Description = "Returns/sets the text contained in the control."
    Hour = txtHour.Text
End Property

Public Property Let Hour(ByVal New_Hour As String)
    txtHour.Text() = New_Hour
    PropertyChanged "Hour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMinute,txtMinute,-1,Text
Public Property Get Minute() As String
Attribute Minute.VB_Description = "Returns/sets the text contained in the control."
    Minute = txtMinute.Text
End Property

Public Property Let Minute(ByVal New_Minute As String)
    txtMinute.Text() = New_Minute
    PropertyChanged "Minute"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
    m_TimeType = m_def_TimeType
    m_ShowMeridiem = m_def_ShowMeridiem
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Font.Name = "Arial"
    m_Font.Size = 10
    m_Font.Bold = Me.FontBold
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    txtHour.Text = PropBag.ReadProperty("Hour", "12")
    txtMinute.Text = PropBag.ReadProperty("Minute", "00")
    m_TimeType = PropBag.ReadProperty("TimeType", m_def_TimeType)
    txtHour.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtMinute.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtAMPM.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtHour.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtMinute.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtAMPM.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtHour.FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    txtMinute.FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    txtAMPM.FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    txtAMPM.Text = PropBag.ReadProperty("Meridiem", m_def_Meridiem)
    m_ShowMeridiem = PropBag.ReadProperty("ShowMeridiem", m_def_ShowMeridiem)
    Set txtHour.Font = PropBag.ReadProperty("FontName", m_Font)
    Set txtMinute.Font = PropBag.ReadProperty("FontName", m_Font)
    Set txtAMPM.Font = PropBag.ReadProperty("FontName", m_Font)
End Sub

Private Sub UserControl_Show()
    If Me.ShowMeridiem Then
        txtAMPM.Visible = True
        frmFrame.Width = 1665
    Else
        txtAMPM.Visible = False
        frmFrame.Width = 1290
    End If
    UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Hour", txtHour.Text, "12")
    Call PropBag.WriteProperty("Minute", txtMinute.Text, "00")
    Call PropBag.WriteProperty("TimeType", m_TimeType, m_def_TimeType)
    Call PropBag.WriteProperty("BackColor", txtHour.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtHour.ForeColor, &H80000008)
    Call PropBag.WriteProperty("FontBold", txtHour.FontBold, m_def_FontBold)
    Call PropBag.WriteProperty("Meridiem", txtAMPM.Text, m_def_Meridiem)
    Call PropBag.WriteProperty("ShowMeridiem", m_ShowMeridiem, m_def_ShowMeridiem)
    Call PropBag.WriteProperty("FontName", txtHour.Font, m_Font)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=26,0,0,0
Public Property Get TimeType() As bbTimeType
Attribute TimeType.VB_Description = "Set time to either military or standard time. Biatch!"
    TimeType = m_TimeType
End Property

Public Property Let TimeType(ByVal New_TimeType As bbTimeType)
Dim boolConverttoMilitaryTime As Boolean
Dim boolConverttoStandardTime As Boolean
    boolConverttoMilitaryTime = False
    boolConverttoStandardTime = False
    If m_TimeType = bbStandardTime And New_TimeType = bbMilitaryTime Then
        boolConverttoMilitaryTime = True
    ElseIf m_TimeType = bbMilitaryTime And New_TimeType = bbStandardTime Then
        boolConverttoStandardTime = True
    End If
    m_TimeType = New_TimeType
    If boolConverttoMilitaryTime Then
        If Me.Meridiem = "PM" Then
            If CInt(txtHour.Text) < 12 Then
                txtHour.Text = CStr(CInt(txtHour.Text) + 12)
            End If
        ElseIf txtHour.Text = "12" Then
            txtHour.Text = "0"
        End If
    ElseIf boolConverttoStandardTime Then
        If CInt(txtHour.Text) > 12 Then
            txtHour.Text = CStr(CInt(txtHour.Text) - 12)
        ElseIf txtHour.Text = 0 Then
            txtHour.Text = "12"
        End If
    End If
    PropertyChanged "TimeType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtHour,txtHour,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtHour.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtHour.BackColor() = New_BackColor
    Label1.BackColor() = New_BackColor
    txtMinute.BackColor() = New_BackColor
    txtAMPM.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtHour,txtHour,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtHour.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtHour.ForeColor() = New_ForeColor
    txtMinute.ForeColor() = New_ForeColor
    Label1.ForeColor() = New_ForeColor
    txtAMPM.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtHour,txtHour,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = txtHour.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtHour.FontBold() = New_FontBold
    txtMinute.FontBold() = New_FontBold
    txtAMPM.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtAMPM,txtAMPM,-1,Text
Public Property Get Meridiem() As String
Attribute Meridiem.VB_Description = "Returns/sets the text contained in the control."
    Meridiem = txtAMPM.Text
End Property

Public Property Let Meridiem(ByVal New_Meridiem As String)
    txtAMPM.Text() = New_Meridiem
    PropertyChanged "Meridiem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowMeridiem() As Boolean
    ShowMeridiem = m_ShowMeridiem
End Property

Public Property Let ShowMeridiem(ByVal New_ShowMeridiem As Boolean)
    m_ShowMeridiem = New_ShowMeridiem
    If New_ShowMeridiem Then
        txtAMPM.Visible = True
        frmFrame.Width = 1665
    Else
        txtAMPM.Visible = False
        frmFrame.Width = 1290
    End If
    UserControl_Resize
    PropertyChanged "ShowMeridiem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtHour,txtHour,-1,Font
Public Property Get FontName() As Font
Attribute FontName.VB_Description = "Returns a Font object."
    Set FontName = txtHour.Font
End Property

Public Property Set FontName(ByVal New_FontName As Font)
    Set txtHour.Font = New_FontName
    Set txtMinute.Font = New_FontName
    Set txtAMPM.Font = New_FontName
    PropertyChanged "FontName"
End Property

Private Function CheckHour(ByVal sHour As String, sMeridiem As String)
    If TimeType = bbMilitaryTime Then
        If sMeridiem = "PM" Then
            If sHour < 12 Then
                sHour = CStr(CInt(sHour) + 12)
            End If
        ElseIf sHour = 12 And sMeridiem = "AM" Then
            sHour = "00"
        End If
    Else
        If sHour > 12 Then
            sHour = sHour - 12
            sMeridiem = "PM"
        Else
            If sMeridiem <> "AM" And sMeridiem <> "PM" Then
                sMeridiem = "AM"
            End If
        End If
    End If
    CheckHour = sHour
End Function
