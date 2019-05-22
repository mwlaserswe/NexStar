VERSION 5.00
Begin VB.UserControl zzSlider 
   BackStyle       =   0  'Transparent
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   4815
   Begin VB.PictureBox picBackGround 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Label lblSlider 
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "zzSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : zzSlider
' DateTime  : 11/11/2007 14:19
' Author    : Zeezee
' Purpose   : Use as a Slding Control
' Version   : Ver 1.02 Beta
' Disclaimer: You can use this code freely. Please give credit when used. Code/Control is supplied 'AS IS'
'               No Gurantee for Loss of work.
'---------------------------------------------------------------------------------------

'Custom Properties
'
'1. Back Colour - set the Back Colour for the slider- This maps to the Pciture Box Back Colour
'2. Border Style - Set the Border Style of the slider - This maps to Picture Box Border Style
'3. Enabled - Enable Or Disable the Control
'4. Font - Maps to Label's Font- Currently Not Used
'5. Fore Colour - Maps to Label's For Colour - Current;ly Not Used. Can be used to set Font Colour
'6. hWnd - Window Handle of the User Control - Currently not Used. Cant think of any use
'7. MaxValue - Maximum Value returned. Must be  Long integer
'8. MinValue - Minumum Value Returned.  Must be Long integer. If MinValue > MaxValue, nothing returend
'9. MousePointer - Canbe used to set custom Mouse Pointer
'10. Slider Colour - This wil set the colour for the Slider - This maps to Label's Back Colour
'11. Slider Style - Set the Slider Style- Horizontal or Vertical. Only at design time
'12. Value - The current Value of the Slider. Support Read/ Write properties
'13. ToolTipText - Set the Slider's Tool Tip Text. - Implemented but not working
'14. SmallChange - The amount of change in value when Arrow Keys are pressed
'15. LargeChange - The amount of change in value when PageUp / PageDown keys are pressed
'16. RoundValue - The Rounding of value to given decimal points. - minimum 0, maximum 5

'Feature upgrades
'01 - Supports Negative or Positive integers
'02 - Has Key Board Support
'03 - Support for Single Precision Floating Point Numbers with given rounding value
'04 - SmallChange and LargeChange now support Floating Point Values

'Known Bugs
'01 - ToolTip Doesnt Work
'02 - Smooth Scrolling not working. The test project, before the use control, changed values when the mouse was moving.
'       In the user control, the slide movement is visible but value change only occurs when the mouse is released.
'       Change even is raased in Value Change but still not working. needs to fix this.

'Credits
'YoungBuck - For supplying the code for FocusRect - http://www.vbforums.com/showthread.php?t=17483

Option Explicit
'Default Property Values:
Const m_def_ToolTipText = ""
Const m_def_SliderStyle = 0
Const m_def_MaxValue = 0
Const m_def_MinValue = 0
Const m_def_Value = 0
Const m_def_SmallChange = 1
Const m_def_LargeChange = 1
Const m_def_RoundValue = 0

'Property Variables:
Dim m_ToolTipText As String
Dim m_SliderStyle As enumSliderStyle
Dim m_MaxValue As Single
Dim m_MinValue As Single
Dim m_Value As Single
Dim m_SmallChange As Single
Dim m_LargeChange As Single
Dim m_RoundValue As Integer


Public Enum BorderStyle
    None = 0
    FixedSingle = 1
End Enum

Public Enum enumSliderStyle
    Horizontal = 0
    Verticle = 1
End Enum

Dim valSetting As Boolean

'Event Declarations:
Event Change() 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event SliderScroll()


Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lprect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim b_HasFocus As Boolean ' flag for paint event

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBackGround,picBackGround,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picBackGround.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picBackGround.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblSlider,lblSlider,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblSlider.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblSlider.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblSlider,lblSlider,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblSlider.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblSlider.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBackGround,picBackGround,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = picBackGround.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    picBackGround.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub


Private Sub lblSlider_Click()
    RaiseEvent Click
End Sub

Private Sub lblSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     setValue Button, X, lblSlider.Top + Y
End Sub

Private Sub lblSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     setValue Button, X, lblSlider.Top + Y
End Sub

Private Sub picBackGround_Click()
 RaiseEvent Click
End Sub


Private Sub picBackGround_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Or KeyCode = vbKeyUp Then
        Value = Value + SmallChange
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyDown Then
        Value = Value - SmallChange
    ElseIf KeyCode = vbKeyPageUp Then
        Value = Value + LargeChange
    ElseIf KeyCode = vbKeyPageDown Then
        Value = Value - LargeChange
    End If
End Sub

Private Sub picBackGround_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setValue Button, X, Y
End Sub

Private Sub picBackGround_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setValue Button, X, Y
End Sub


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    picBackGround.MousePointer() = New_MousePointer
    lblSlider.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property


Private Sub UserControl_Initialize()
     b_HasFocus = False ' initialize focus flag to false (if control gets focus first b_hasfocus will be set to True by EnterFocus event)

    picBackGround.ScaleMode = ScaleModeConstants.vbTwips
    
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_MaxValue = m_def_MaxValue
    m_MinValue = m_def_MinValue
    m_Value = m_def_Value
    m_SliderStyle = m_def_SliderStyle
    m_ToolTipText = m_def_ToolTipText
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     setValue Button, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     setValue Button, X, Y
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picBackGround.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    lblSlider.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblSlider.Font = PropBag.ReadProperty("Font", Ambient.Font)
    picBackGround.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    lblSlider.BackColor = PropBag.ReadProperty("SliderColor", &H8000000F)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_SliderStyle = PropBag.ReadProperty("SliderStyle", m_def_SliderStyle)
    m_SmallChange = PropBag.ReadProperty("SmallChange", m_def_SmallChange)
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_RoundValue = PropBag.ReadProperty("RoundValue", m_def_RoundValue)
    
    picBackGround.ToolTipText = m_ToolTipText


    
    If SliderStyle = Horizontal Then
        lblSlider.Width = 0
        lblSlider.Height = picBackGround.Height
    ElseIf SliderStyle = Verticle Then
        lblSlider.Height = 0
        lblSlider.Width = picBackGround.Width
    End If

    
End Sub

Private Sub UserControl_Resize()
    picBackGround.Width = UserControl.Width
    picBackGround.Height = UserControl.Height
        
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    

    Call PropBag.WriteProperty("BackColor", picBackGround.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", lblSlider.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblSlider.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", picBackGround.BorderStyle, 1)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("SliderColor", lblSlider.BackColor, &HFFC0C0)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("SliderStyle", m_SliderStyle, m_def_SliderStyle)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange, m_def_SmallChange)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("RoundValue", m_RoundValue, m_def_RoundValue)
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblSlider,lblSlider,-1,BackColor
Public Property Get SliderColor() As OLE_COLOR
Attribute SliderColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    SliderColor = lblSlider.BackColor
End Property

Public Property Let SliderColor(ByVal New_SliderColor As OLE_COLOR)
    lblSlider.BackColor() = New_SliderColor
    PropertyChanged "SliderColor"
End Property

Public Property Get MaxValue() As Single
Attribute MaxValue.VB_Description = "MaxValue of the Slider"
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Single)

On Error GoTo ERRHNDLE
'    If New_MaxValue < 0 Then
'        m_MaxValue = 0
'        PropertyChanged "MaxValue"
'        Err.Raise 380
'    Else
        m_MaxValue = New_MaxValue
        PropertyChanged "MaxValue"
'    End If
    
    Exit Property

ERRHNDLE:
    MsgBox Err.Number & " : " & Err.Description
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinValue() As Single
Attribute MinValue.VB_Description = "Minimum Value For The Slider"
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Single)

    On Error GoTo ERRHNDLE
'    If New_MinValue < 0 Then
'        m_MinValue = 0
'        PropertyChanged "MinValue"
'        Err.Raise 380
'    Else
        m_MinValue = New_MinValue
        PropertyChanged "MinValue"
'    End If
    
    Exit Property

ERRHNDLE:
    MsgBox "Invalid Property Value."


End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Single
Attribute Value.VB_Description = "Current Value of the slider"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)

    If New_Value < MinValue Then
        m_Value = MinValue
    ElseIf New_Value > MaxValue Then
        m_Value = MaxValue
    Else
        m_Value = New_Value
    End If
    m_Value = Round(m_Value, RoundValue)
    PropertyChanged "Value"

    setValuePos m_Value
    
    lblSlider.ToolTipText = Value

    RaiseEvent Change
    valSetting = False
    
End Property


Public Property Get SliderStyle() As enumSliderStyle
    SliderStyle = m_SliderStyle
End Property


Public Property Let SliderStyle(ByVal New_SliderStyle As enumSliderStyle)
     If Ambient.UserMode Then Err.Raise 382
    Dim changeStyle As Boolean
    If SliderStyle <> New_SliderStyle Then changeStyle = True

    m_SliderStyle = New_SliderStyle
    PropertyChanged "SliderStyle"

    If changeStyle = True Then setSliderStyle
End Property


Public Property Get SmallChange() As Single
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Single)

On Error GoTo ERRHNDLE
    If New_SmallChange < 0 Then
        m_SmallChange = m_def_SmallChange
        PropertyChanged "SmallChange"
        Err.Raise 380
    Else
        m_SmallChange = New_SmallChange
        PropertyChanged "SmallChange"
    End If
    
    Exit Property

ERRHNDLE:
    MsgBox Err.Number & " : " & Err.Description
End Property


Public Property Get LargeChange() As Single
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Single)

On Error GoTo ERRHNDLE
    If New_LargeChange < 0 Then
        m_LargeChange = m_def_LargeChange
        PropertyChanged "LargeChange"
        Err.Raise 380
    Else
        m_LargeChange = New_LargeChange
        PropertyChanged "LargeChange"
    End If
    
    Exit Property

ERRHNDLE:
    MsgBox Err.Number & " : " & Err.Description
End Property

Public Property Get RoundValue() As Integer
    RoundValue = m_RoundValue
End Property

Public Property Let RoundValue(ByVal New_RoundValue As Integer)

On Error GoTo ERRHNDLE
    If New_RoundValue < 0 Then
        m_RoundValue = m_def_RoundValue
        PropertyChanged "RoundValue"
        Err.Raise 380
    Else
        m_RoundValue = New_RoundValue
        PropertyChanged "RoundValue"
    End If
    
    Exit Property

ERRHNDLE:
    MsgBox Err.Number & " : " & Err.Description
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    picBackGround.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property



Private Sub setSliderStyle()
    Dim currentwidth As Single
    Dim currentheight As Single
    
    
    If SliderStyle = Horizontal Then
            currentwidth = UserControl.Width
            currentheight = UserControl.Height
            UserControl.Width = currentheight
            UserControl.Height = currentwidth
            lblSlider.Height = picBackGround.Height
            lblSlider.Width = 0
            
    ElseIf SliderStyle = Verticle Then
            currentwidth = UserControl.Width
            currentheight = UserControl.Height
            UserControl.Width = currentheight
            UserControl.Height = currentwidth
            
            lblSlider.Width = picBackGround.Width
            lblSlider.Height = 0
            lblSlider.Top = picBackGround.Top + picBackGround.Height
    End If
    
End Sub




Private Function setValue(ByVal Button As Integer, ByVal X As Single, Y As Single)
    
    If Button = vbLeftButton And valSetting = False Then
        If MinValue > MaxValue Then Exit Function
       
        
        If SliderStyle = Horizontal Then
             lblSlider.Height = picBackGround.Height
            If X > Screen.TwipsPerPixelX And X <= picBackGround.Width Then
'                    lblSlider.Width = X
                    valSetting = True
                    Value = ((X / picBackGround.Width) * (MaxValue - MinValue)) + MinValue


            ElseIf X <= Screen.TwipsPerPixelX Then
'                 lblSlider.Width = 0
                 valSetting = True
                 Value = CLng(IIf(X > MinValue, X, MinValue))
            ElseIf X > picBackGround.Width Then
'                 lblSlider.Width = picBackGround.Width
                valSetting = True
                Value = MaxValue
                 
            End If
            
        ElseIf SliderStyle = Verticle Then
            lblSlider.Width = picBackGround.Width
            If Y > Screen.TwipsPerPixelY And Y <= picBackGround.Height Then
                    valSetting = True
                    Value = (((picBackGround.Height - Y) / picBackGround.Height) * (MaxValue - MinValue)) + MinValue
                 
            ElseIf Y <= picBackGround.Top Then
                 
                valSetting = True
                Value = MaxValue
                 
            ElseIf Y > picBackGround.Height Then
            
                valSetting = True
                Value = MinValue
                
            End If
        End If
        RaiseEvent SliderScroll
            
    End If
End Function


Private Function setValuePos(ByVal CurrentValue As Single)
    If MinValue > MaxValue Then Exit Function
    If SliderStyle = Horizontal Then
        lblSlider.Height = picBackGround.Height
        lblSlider.Width = CSng(picBackGround.Width * ((CurrentValue - MinValue) / (MaxValue - MinValue)))

    ElseIf SliderStyle = Verticle Then
        lblSlider.Width = picBackGround.Width
        lblSlider.Top = CSng(picBackGround.Height - (picBackGround.Height * ((CurrentValue - MinValue) / (MaxValue - MinValue))))
        lblSlider.Height = picBackGround.Height - lblSlider.Top
        
    End If
    
End Function

Private Sub UserControl_EnterFocus()
    b_HasFocus = True ' flag paint event to display focus rect
    Call UserControl_Paint ' repaint object
End Sub


Private Sub ShowFocusRect()
Dim rctFocus As RECT
Dim ret As Long
    'get dimensions of usercontrol
      With rctFocus
        .Top = 3
        .Left = 3
        .Right = (UserControl.Width \ Screen.TwipsPerPixelX) - 3
        .Bottom = (UserControl.Height \ Screen.TwipsPerPixelY) - 3
    End With
      
    ret = DrawFocusRect(picBackGround.hdc, rctFocus) ' display focus rect
      
End Sub

Private Sub UserControl_ExitFocus()
    b_HasFocus = False 'flag paint even NOT to draw focus rect
    Call UserControl.Cls ' clear graphical contents of UserControl to remove focus rect
    Call picBackGround.Cls
    Call UserControl_Paint ' repaint object
End Sub


Private Sub UserControl_Paint()
    ' Anything & Everything that you draw on a usercontrol at runtime has to be in the paint event _
        or it will not show up on the screen
    If b_HasFocus = True Then ' if the control has focus
        Call ShowFocusRect ' paint it!
    End If
End Sub


