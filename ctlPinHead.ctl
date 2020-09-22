VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlPinHead 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   PropertyPages   =   "ctlPinHead.ctx":0000
   ScaleHeight     =   1965
   ScaleWidth      =   2850
   ToolboxBitmap   =   "ctlPinHead.ctx":001C
   Begin VB.PictureBox picToggle 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Picture         =   "ctlPinHead.ctx":032E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   120
      Width           =   240
   End
   Begin MSComctlLib.ImageList imglstState 
      Left            =   360
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlPinHead.ctx":08B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlPinHead.ctx":0E52
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlPinHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Const TOP_MOST_ICON = 2
Const NORMAL_ICON = 1

Dim TopMostWindow As Boolean
'Default Property Values:
Const m_def_Border_Style = 1
'Property Variables:
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14

'Dim m_TopMost As New clsTopMost
Dim m_TopMost As New clsTopMost

Public Function Topmost()

With m_TopMost
    .Topmost
End With

UpdateStateImage

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Normal()

With m_TopMost
    .Normal
End With

UpdateStateImage

End Function

Private Sub picToggle_Click()

With m_TopMost
    
    If .Current_State = fsTopmost Then
        .Normal
    Else
        .Topmost
    End If

End With

UpdateStateImage

End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
''< -     picToggle.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_Border_Style)
End Sub

Private Sub UserControl_Resize()

With picToggle
    .Top = 0
    .Left = 0
End With

With UserControl
    .Width = picToggle.Width
    .Height = picToggle.Height
End With

End Sub

Private Sub UserControl_Show()

Set m_TopMost.Target_Form = UserControl.Parent

UpdateStateImage

End Sub

Private Sub UserControl_Terminate()

With m_TopMost
   Set .Target_Form = Nothing
End With

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

'    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
' < -     Call PropBag.WriteProperty("BorderStyle", picToggle.BorderStyle, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
' MemberInfo=14,0,0,0
'Public Property Get FormState() As enmFormState
'    FormState = m_FormState
'End Property

'Public Property Let FormState(ByVal Updated_FormState As enmFormState)
'    m_FormState = Updated_FormState
'    PropertyChanged "FormState"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = UserControl.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    UserControl.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = UserControl.BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    UserControl.BackStyle() = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
   ' m_FormState = m_def_FormState

End Sub

Private Sub UpdateStateImage()

If m_TopMost.Current_State = fsNormal Then
    Set picToggle.Picture = imglstState.ListImages(NORMAL_ICON).Picture
End If

If m_TopMost.Current_State = fsTopmost Then
    Set picToggle.Picture = imglstState.ListImages(TOP_MOST_ICON).Picture
End If

End Sub

'Public Property Get BorderStyle() As enmBorderStyle
'
'    BorderStyle = picToggle.BorderStyle
'
'End Property

'Public Property Let BorderStyle(ByVal New_BorderStyle As enmBorderStyle)
'
'    picToggle.BorderStyle() = New_BorderStyle
'    UserControl_Resize
'    PropertyChanged "BorderStyle"
'
'End Property

