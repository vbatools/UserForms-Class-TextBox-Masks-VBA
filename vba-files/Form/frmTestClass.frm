VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestClass 
   Caption         =   "Test class:"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16830
   OleObjectBlob   =   "frmTestClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : UserForm1
'* Created    : 07-11-2025 10:06
'* Author     : VBATools
'* Copyright  : VBATools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Dim clsTB           As clsTextboxMask

Private Sub chbDec_Change()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtNumeric.Name)
    If item Is Nothing Then Exit Sub
    item.IsDecemal = chbDec.Value
End Sub

Private Sub chbMinus_Change()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtNumeric.Name)
    If item Is Nothing Then Exit Sub
    item.IsNegative = chbMinus.Value
End Sub

Private Sub chbVisible_Click()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtNumeric.Name)
    If item Is Nothing Then Exit Sub
    item.VisibleLabelPlaceholder = chbVisible.Value
End Sub

Private Sub btnRemove_Click()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtDate.Name)
    If item Is Nothing Then Exit Sub
    item.RemoveItem
End Sub

Private Sub btnClear_Click()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtDate.Name)
    If item Is Nothing Then Exit Sub
    item.Clear
End Sub

Private Sub btnSetFocus_Click()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtDate.Name)
    If item Is Nothing Then Exit Sub
    item.SetFocus
End Sub

Private Sub btnSetValue_Click()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtDate.Name)
    If item Is Nothing Then Exit Sub
    item.Value = VBA.Date
End Sub

Private Sub txtDate_Change()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtDate.Name)
    If item Is Nothing Then Exit Sub
    With item
        lbValid.Caption = "Is Valid: " & .IsValid
        lbValue.Caption = "Value: " & .Value
    End With
End Sub

Private Sub txtOther_Change()
    Dim item        As clsTextboxMask
    Set item = clsTB.GetItemByName(txtOther.Name)
    If item Is Nothing Then Exit Sub
    With item
        Label7.Caption = "Remaining Chars: " & .RemainingChars
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    Set clsTB = New clsTextboxMask
    Dim dt          As Date

    dt = VBA.Date - 360
    dtMin.Caption = "Date min: " & dt
    dtMax.Caption = "Date max: " & VBA.Date

    With clsTB
        Call .AddFieldDate(txtDate, "##.##.####", dt, VBA.Date, "dd.mm.yyyy")
        Call .AddFieldText(txtPhone, "+7(###) ### ## ##")
        Call .AddFieldText(txtOther, "+7(*##) @# A# #Á#", True, rgbGreen, rgbRed, rgbViolet)
        Call .AddFieldTime(txtTime, "##:##", 0, 1)

        Call .AddFieldNumeric(txtNumeric, 0, 100, False, False)

        Call .AddFieldRegex(txtRegex, "\w+@\w+\.\w+", "[\w\-\.@]")

        With clsTB.GetItemByName(txtRegex.Name)
            lbPatern.Caption = lbPatern.Caption & .RegexPattern
            lbFulterChrs.Caption = lbFulterChrs.Caption & .RegexFilter
        End With

        Call .AddFieldVariableLength(txtVariableLen, 10, "##")

        Call .AddFieldText(txtPhoneHolder, "+7(###) ### ## ##", True, , , , "Empty", "Partial", "Complete", "Invalid")
        Call .AddFieldVariableLength(txtVariableLenHolder, 10, "##", True, , , , "Empty", "Partial", "Complete", "Invalid")

        Call .AddFieldVariableLength(txtVariableLenHolder2, 10, "##", True, , , , "", "", "", "", "mask: {holder} rem: {percent}")
        
        lbCount.Caption = lbCount.Caption & .Count
        lbVersion.Caption = .Version
    End With
End Sub
