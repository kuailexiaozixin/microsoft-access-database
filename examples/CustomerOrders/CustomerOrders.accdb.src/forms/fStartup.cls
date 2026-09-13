Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' See "fStartup.cls"

Option Compare Database
Option Explicit

Private Sub cmdCustomers_Click()
    DoCmd.OpenForm "fCustomerList"
End Sub

Private Sub cmdOrders_Click()
    DoCmd.OpenForm "fOrderList"
End Sub

Private Sub cmdExit_Click()
    modNavigation.ExitApp
End Sub
