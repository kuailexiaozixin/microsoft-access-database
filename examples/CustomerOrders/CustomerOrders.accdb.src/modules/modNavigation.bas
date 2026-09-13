Attribute VB_Name = "modNavigation"
Option Compare Database
Option Explicit

' Open a form in the specified view
Public Sub OpenFormNav(strFormName As String, Optional lngView As AcFormView = acNormal, _
                       Optional strFilter As String = "", _
                       Optional strWhereCondition As String = "")
    If strWhereCondition <> "" Then
        DoCmd.OpenForm strFormName, lngView, , strWhereCondition
    Else
        DoCmd.OpenForm strFormName, lngView
    End If
End Sub

' Close the active form
Public Sub CloseActiveForm()
    DoCmd.Close acForm, Screen.ActiveForm.Name
End Sub

' Close a specific form by name
Public Sub CloseFormByName(strFormName As String)
    DoCmd.Close acForm, strFormName
End Sub

' Open the order report for a specific OrderID
Public Sub PreviewOrderReport(lngOrderID As Long)
    DoCmd.OpenReport "rOrderReport", acViewPreview, , "OrderID = " & lngOrderID
End Sub

' Save the current record and close the form
Public Sub SaveAndCloseForm()
    Dim frmName As String
    frmName = Screen.ActiveForm.Name
    With Screen.ActiveForm
        If .Dirty Then .Dirty = False
    End With
    ' Requery all other open forms so list forms show updated data
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> frmName Then
            frm.Requery
        End If
    Next frm
    DoCmd.Close acForm, frmName
End Sub

' Discard changes and close the form
Public Sub DiscardAndCloseForm()
    With Screen.ActiveForm
        If .Dirty Then .Undo
        DoCmd.Close acForm, .Name
    End With
End Sub

' Exit the application
Public Sub ExitApp()
    DoCmd.Quit
End Sub
