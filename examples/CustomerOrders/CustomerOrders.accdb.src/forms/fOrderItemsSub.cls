Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' See "fOrderItemsSub.cls"

Option Compare Database
Option Explicit

Private Sub txtUnitPriceNet_AfterUpdate()
    CalcSubTotal
End Sub

Private Sub txtQuantity_AfterUpdate()
    CalcSubTotal
End Sub

Private Sub CalcSubTotal()
    If Not IsNull(Me.txtUnitPriceNet) And Not IsNull(Me.txtQuantity) Then
        Me.txtSubTotalNet = Me.txtUnitPriceNet * Me.txtQuantity
    Else
        Me.txtSubTotalNet = Null
    End If
End Sub
