VERSION 5.00
Begin VB.UserControl Label3D 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "Label3D.ctx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   5145
End
Attribute VB_Name = "Label3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Enum Effects
    None
    [Carved Light]
    Carved
    [Carved Heavy]
    [Raised Light]
    Raised
    [Raised Heavy]
End Enum

Enum Align
    [Top Left]
    [Top Middle]
    [Top Right]
    [Center Left]
    [Center Middle]
    [Center Right]
    [Bottom Left]
    [Bottom Middle]
    [Bottom Right]
End Enum

Enum BackgroundStyle
    Transparent
    Opaque
End Enum
    
Private m_Caption As String
Private m_Effect As Effects
Private m_TextAlignment As Align

'Default Property Values:
Const m_def_Caption = "3D Label"
Const m_def_Effect = 2
Const m_def_TextAlignment = 4
'Property Variables:
'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackgroundStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackgroundStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
Dim ExtProp As String
    ExtProp = "I'm a custom control. My name is " _
              & UserControl.Extender.Name
    ExtProp = ExtProp & "I'm located at (" & UserControl.Extender.Left _
              & ", " & UserControl.Extender.Left & ")"
    ExtProp = ExtProp & vbCrLf & " My dimensions are " _
              & UserControl.Extender.Width & " by " _
              & UserControl.Extender.Height
    ExtProp = ExtProp & vbCrLf & "I'm tagged as " _
              & UserControl.Extender.Tag
    ExtProp = ExtProp & vbCrLf & "I'm sited on a control named " _
              & UserControl.Extender.Parent.Name
    ExtProp = ExtProp & vbCrLf & "whose dimensions are " _
              & UserControl.Extender.Parent.Width _
              & " by " & UserControl.Extender.Parent.Height
    MsgBox ExtProp
    RaiseEvent Click
End Sub

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
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    UserControl_Paint
End Property

Private Sub UserControl_Paint()
    DrawCaption
    OldFontSize = UserControl.Font.Size
    UserControl.Font.Size = 10
    If Not Ambient.UserMode Then
        UserControl.CurrentX = 0
        UserControl.CurrentY = 0
        UserControl.Print "Design Mode"
    End If
    UserControl.Font.Size = OldFontSize
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
    RaiseEvent Resize
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "CaptionProperties"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

Public Property Get Effect() As Effects
    Effect = m_Effect
End Property

Public Property Let Effect(ByVal New_Effect As Effects)
    m_Effect = New_Effect
    PropertyChanged "Effect"
    UserControl_Paint
End Property

Public Property Get TextAlignment() As Align
Attribute TextAlignment.VB_Description = "Determines how the caption will be aligned on the control"
    TextAlignment = m_TextAlignment
End Property

Public Property Let TextAlignment(ByVal New_TextAlignment As Align)
    m_TextAlignment = New_TextAlignment
    PropertyChanged "TextAlignment"
    UserControl_Paint
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    m_Caption = m_def_Caption
    m_Effect = m_def_Effect
    m_TextAlignment = m_def_TextAlignment
    UserControl.BorderStyle = 1
    UserControl.BackStyle = 1
    UserControl_Paint
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Effect = PropBag.ReadProperty("Effect", m_def_Effect)
    m_TextAlignment = PropBag.ReadProperty("TextAlignment", m_def_TextAlignment)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Show()
    UserControl_Paint
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Effect", m_Effect, m_def_Effect)
    Call PropBag.WriteProperty("TextAlignment", m_TextAlignment, m_def_TextAlignment)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

Private Sub DrawCaption()
Dim CaptionWidth As Long, CaptionHeight As Long
Dim CurrX As Long, CurrY As Long
Dim oldForeColor As OLE_COLOR

    CaptionHeight = TextHeight(m_Caption)
    CaptionWidth = TextWidth(m_Caption)
    Select Case m_TextAlignment
        Case 0:
            CurrX = 30
            CurrY = 0
        Case 1:
            CurrX = (UserControl.Width - CaptionWidth) / 2
            CurrY = 0
        Case 2:
            CurrX = UserControl.Width - CaptionWidth - 30
            CurrY = 0
        Case 3:
            CurrX = 30
            CurrY = (UserControl.Height - CaptionHeight) / 2
        Case 4:
            CurrX = (UserControl.Width - CaptionWidth) / 2
            CurrY = (UserControl.Height - CaptionHeight) / 2
        Case 5:
            CurrX = UserControl.Width - CaptionWidth - 30
            CurrY = (UserControl.Height - CaptionHeight) / 2
        Case 6:
            CurrX = 30
            CurrY = UserControl.Height - CaptionHeight - 45
        Case 7:
            CurrX = (UserControl.Width - CaptionWidth) / 2
            CurrY = UserControl.Height - CaptionHeight - 45
        Case 8:
            CurrX = UserControl.Width - CaptionWidth - 30
            CurrY = UserControl.Height - CaptionHeight - 45
    End Select
        
    oldForeColor = UserControl.ForeColor
    Select Case m_Effect
        Case 0:
            UserControl.Cls
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.Print m_Caption
        Case 1:
            UserControl.Cls
            UserControl.CurrentX = CurrX + 15
            UserControl.CurrentY = CurrY + 15
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 2:
            UserControl.Cls
            UserControl.CurrentX = CurrX + 30
            UserControl.CurrentY = CurrY + 30
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 3:
            UserControl.Cls
            UserControl.CurrentX = CurrX + 45
            UserControl.CurrentY = CurrY + 45
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX + 30
            UserControl.CurrentY = CurrY + 30
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX + 15
            UserControl.CurrentY = CurrY + 15
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 4:
            UserControl.Cls
            UserControl.CurrentX = CurrX - 15
            UserControl.CurrentY = CurrY - 15
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 5:
            UserControl.Cls
            UserControl.CurrentX = CurrX - 30
            UserControl.CurrentY = CurrY - 30
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 6:
            UserControl.Cls
            UserControl.CurrentX = CurrX - 45
            UserControl.CurrentY = CurrY - 45
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX - 30
            UserControl.CurrentY = CurrY - 30
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX - 15
            UserControl.CurrentY = CurrY - 15
            UserControl.ForeColor = RGB(255, 255, 255)
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        
        End Select
' UNCOMMENT THE FOLLOWING LINES TO DRAW TWO CROSS LINE
' AT THE CENTER OF THE USERCONTROL OBJECT
' I used these statements to verify the centering of the caption
'        UserControl.ForeColor = RGB(255, 255, 0)
'        UserControl.Line (0, UserControl.Height / 2)-(UserControl.Width, UserControl.Height / 2)
'        UserControl.Line (UserControl.Width / 2, 0)-(UserControl.Width / 2, UserControl.Height)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UserControl_Paint
End Property

