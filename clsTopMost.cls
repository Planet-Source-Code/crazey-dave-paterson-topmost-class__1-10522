VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTopMost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarTarget_Form As Form 'local copy
'local variable(s) to hold property value(s)
'Private mvarCurrent_State As enmFormState 'local copy

Public Enum enmFormState
    fsTopmost = 1
    fsNormal = 2
End Enum

Public Sub Toggle()

MsgBox "Not implemeted", vbCritical + vbOKOnly, "Public Sub Toggle()"

End Sub

Public Property Get Current_State() As enmFormState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Current_State
    
    If isTopMost(mvarTarget_Form) Then
        Current_State = fsTopmost
    Else
        Current_State = fsNormal
    End If

End Property

Public Property Set Target_Form(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Target = Form1
    Set mvarTarget_Form = vData
End Property

Public Property Get Target_Form() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Target
    Set Target_Form = mvarTarget_Form
End Property

Public Sub Topmost()
Dim lState As Long

If Not IsMissing(mvarTarget_Form) Then

    Call SetWindowPos(mvarTarget_Form.hWnd, HWND_TOPMOST, -1, -1, -1, -1, SWP_NOSIZE Or SWP_NOMOVE)

Else

    Err.Raise Err.Number, "mmInterfaceEnhancement::clsTopMost.Update_Target_Form", Err.Description + " Target Form should be set before use."

End If

End Sub

Public Sub Normal()

If Not IsMissing(mvarTarget_Form) Then

    Call SetWindowPos(mvarTarget_Form.hWnd, HWND_NOTOPMOST, -1, -1, -1, -1, SWP_NOSIZE Or SWP_NOMOVE)

Else

    Err.Raise Err.Number, "mmInterfaceEnhancement::clsTopMost.Update_Target_Form", Err.Description + " Target Form should be set before use."

End If

End Sub


Public Function isTopMost(Target_Form As Object) As Boolean
Dim lWindowStyle As Long


lWindowStyle = GetWindowLong(Target_Form.hWnd, GWL_EXSTYLE)

If lWindowStyle And WS_EX_TOPMOST Then
    isTopMost = True
Else
    isTopMost = False
End If

End Function

