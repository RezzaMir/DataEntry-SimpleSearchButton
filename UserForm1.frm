VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "WECC Form"
   ClientHeight    =   11670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20415
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim updateRow As Integer


Private Sub cmdAddNew_Click()

Dim wks As Worksheet
Dim AddNew As Range

Set wks = Sheet1
Set AddNew = wks.Range("A65356").End(xlUp).Offset(1, 0)

AddNew.Offset(0, 0).Value = txtRef.Text
AddNew.Offset(0, 1).Value = txtFirstname.Text
AddNew.Offset(0, 2).Value = txtSurname.Text
AddNew.Offset(0, 3).Value = txtAddress.Text
AddNew.Offset(0, 4).Value = txtPostCode.Text
AddNew.Offset(0, 5).Value = txtTelephone.Text
AddNew.Offset(0, 6).Value = txtDataReg.Text
AddNew.Offset(0, 7).Value = txtProve.Text
AddNew.Offset(0, 8).Value = txtMemberType.Text
AddNew.Offset(0, 9).Value = txtMemberFees.Text

TheDisplay.ColumnCount = 11
TheDisplay.RowSource = "A1:J65356"

Call Refresh_Data

End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

For i = 1 To Range("A65356").End(xlUp).Row
    If TheDisplay.Selected(i) Then
        Rows(i + 2).Select
        Selection.Delete
    End If
Next i

Call Refresh_Data

End Sub

Private Sub cmdExit_Click()
Dim iExit As VbMsgBoxResult
iExit = MsgBox("Confirm if you want to exit", vbQuestion + vbYesNo, "Search System")
If iExit = vbYes Then
Unload Me
End If
End Sub

Private Sub cmdPrint_Click()
Application.Dialogs(xlDialogPrinterSetup).Show
ThisWorkbook.Sheets("Sheet1").PrintOut copies:=1
End Sub


Private Sub cmdReset_Click()
Dim txt
    For Each txt In Frame2.Controls
        If TypeOf txt Is MSForms.TextBox Then
            txt.Text = ""
        End If
    Next txt
    
    txtSearch.Text = ""
End Sub

Private Sub cmdSearch_Click()

Dim iSearch As Long, i As Long

iSearch = Worksheets("Sheet1").Range("A1").CurrentRegion.Rows.Count

For i = 1 To iSearch

If Trim(Sheet1.Cells(i, 1)) <> Trim(txtSearch.Text) And i = iSearch Then
MsgBox ("Data not found")
txtSearch.Text = ""
txtSearch.SetFocus
End If

If Trim(Sheet1.Cells(i, 1)) = Trim(txtSearch.Text) Or _
(Trim(Sheet1.Cells(i, 2)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 3)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 4)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 5)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 6)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 7)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 8)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 9)) = Trim(txtSearch.Text)) Or _
(Trim(Sheet1.Cells(i, 10)) = Trim(txtSearch.Text)) Then



txtRef.Text = Sheet1.Cells(i, 1)
txtFirstname.Text = Sheet1.Cells(i, 2)
txtSurname.Text = Sheet1.Cells(i, 3)
txtAddress.Text = Sheet1.Cells(i, 4)
txtPostCode.Text = Sheet1.Cells(i, 5)
txtTelephone.Text = Sheet1.Cells(i, 6)
txtDataReg.Text = Sheet1.Cells(i, 7)
txtProve.Text = Sheet1.Cells(i, 8)
txtMemberType.Text = Sheet1.Cells(i, 9)
txtMemberFees.Text = Sheet1.Cells(i, 10)

Exit For
End If
Next i

End Sub

Private Sub cmdUpdate_Click()
If Me.txtRef.Value = "" Then
MsgBox "Select the record to update"
Exit Sub
End If

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Sheet1")
Dim Selected_Row As Long
Selected_Row = Application.WorksheetFunction.Match(CLng(Me.txtRef.Value), sh.Range("A:A"), 0)


sh.Range("A" & Selected_Row).Value = Me.txtRef.Value
sh.Range("B" & Selected_Row).Value = Me.txtFirstname.Value
sh.Range("C" & Selected_Row).Value = Me.txtSurname.Value
sh.Range("D" & Selected_Row).Value = Me.txtAddress.Value
sh.Range("E" & Selected_Row).Value = Me.txtPostCode.Value
sh.Range("F" & Selected_Row).Value = Me.txtTelephone.Value
sh.Range("G" & Selected_Row).Value = Me.txtDataReg.Value
sh.Range("H" & Selected_Row).Value = Me.txtProve.Value
sh.Range("I" & Selected_Row).Value = Me.txtMemberType.Value
sh.Range("J" & Selected_Row).Value = Me.txtMemberFees.Value
sh.Range("K" & Selected_Row).Value = Now
'------------------------------------------------------------------------

Me.txtRef.Value = ""
Me.txtFirstname.Value = ""
Me.txtSurname.Value = ""
Me.txtAddress.Value = ""
Me.txtPostCode.Value = ""
Me.txtTelephone.Value = ""
Me.txtDataReg.Value = ""
Me.txtProve.Value = ""
Me.txtMemberType.Value = ""
Me.txtMemberFees.Value = ""

Call Refresh_Data


End Sub


Private Sub SpinButton1_Change()
Dim c1, c2

If SpinButton1.Value > 1 Then
c2 = SpinButton1.Value

c1 = "A" & c2
txtRef.ControlSource = c1
c1 = "B" & c2
txtFirstname.ControlSource = c1
c1 = "C" & c2
txtSurname.ControlSource = c1
c1 = "D" & c2
txtAddress.ControlSource = c1
c1 = "E" & c2
txtPostCode.ControlSource = c1
c1 = "F" & c2
txtTelephone.ControlSource = c1
c1 = "G" & c2
txtDataReg.ControlSource = c1
c1 = "H" & c2
txtProve.ControlSource = c1
c1 = "I" & c2
txtMemberType.ControlSource = c1
c1 = "J" & c2
txtMemberFees.ControlSource = c1

End If
End Sub

Private Sub TheDisplay_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'--------------By doble clock on a row, it will bring it up ----------------
Me.txtRef.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 0)
Me.txtFirstname.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 1)
Me.txtSurname.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 2)
Me.txtAddress.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 3)
Me.txtPostCode.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 4)
Me.txtTelephone.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 5)
Me.txtDataReg.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 6)
Me.txtProve.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 7)
Me.txtMemberType.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 8)
Me.txtMemberFees.Value = Me.TheDisplay.List(Me.TheDisplay.ListIndex, 9)

'---------------------------------------------------------------------------



End Sub

Private Sub UserForm_Activate()
Call Refresh_Data
End Sub


Sub Refresh_Data()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Sheet1")
Dim last_Row As Long
last_Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

With Me.TheDisplay
        .ColumnHeads = True
        .ColumnCount = 11
        .ColumnWidths = "60,60,60,80,60,60,60,60,60,60,80"
        
        If last_Row = 1 Then
        .RowSource = "Sheet1!A2:K2"
        Else
        .RowSource = "Sheet1!A2:K" & last_Row
        End If
        
End With
        
End Sub



