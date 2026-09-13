Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' See "fCustomerDetail.cls"

Option Compare Database
Option Explicit

Private Sub cmdSaveClose_Click()
    modNavigation.SaveAndCloseForm
End Sub

Private Sub cmdDiscardClose_Click()
    modNavigation.DiscardAndCloseForm
End Sub
