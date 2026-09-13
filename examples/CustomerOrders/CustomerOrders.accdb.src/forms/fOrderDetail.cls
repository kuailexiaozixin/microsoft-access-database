Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' See "fOrderDetail.cls"

Option Compare Database
Option Explicit

Private Sub cmdSaveClose_Click()
    modNavigation.SaveAndCloseForm
End Sub

Private Sub cmdDiscardClose_Click()
    modNavigation.DiscardAndCloseForm
End Sub

Private Sub cmdPreview_Click()
    ' Save the current record first so the report can read it
    If Me.Dirty Then Me.Dirty = False
    If Not IsNull(Me.ID) Then
        modNavigation.PreviewOrderReport Me.ID
    Else
        MsgBox "Please save the order before previewing.", vbExclamation
    End If
End Sub
