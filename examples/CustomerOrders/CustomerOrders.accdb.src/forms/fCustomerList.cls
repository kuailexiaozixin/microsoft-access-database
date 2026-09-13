Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' See "fCustomerList.cls"

Option Compare Database
Option Explicit

Private Sub cmdOpen_Click()
    If Not IsNull(Me.txtID) Then
        DoCmd.OpenForm "fCustomerDetail", , , "ID = " & Me.txtID
    End If
End Sub

Private Sub cmdAdd_Click()
    DoCmd.OpenForm "fCustomerDetail", , , , acFormAdd
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub
