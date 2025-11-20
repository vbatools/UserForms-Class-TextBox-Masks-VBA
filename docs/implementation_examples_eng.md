# Implementation Examples for clsTextboxMask

## Introduction

This document presents various implementation examples of the `clsTextboxMask` class in real-world scenarios. These examples will help you understand how to use the class in your projects and how to adapt it to specific tasks.

## Example 1: Personal Data Entry Form

### Description
Form for entering personal data with validation of all fields.

### Implementation

```vba
' UserForm: frmPersonalData
' Controls: TextBoxName, TextBoxPhone, TextBoxEmail, TextBoxBirthDate, TextBoxPassport
' CommandButton: cmdSubmit, cmdClear

Dim personalMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set personalMasks = New clsTextboxMask
    
    ' Name field - letters and spaces only, up to 50 characters
    Call personalMasks.AddFieldVariableLength(TextBoxName, 50, "@@@@@@@@@@@@@@@@@@", _
                                             True, , , , "Full Name", RGB(128, 128, 128), "Partial", RGB(165, 102, 41), _
                                             "Entered", RGB(0, 128, 0), "Invalid", RGB(255, 0, 0))
    
    ' Phone field - format +7(XXX) XXX-XX-XX
    Call personalMasks.AddFieldText(TextBoxPhone, "+7(###) ###-##-##", _
                                   True, RGB(0, 128, 0), RGB(25, 0, 0), , "Phone", , "Partial", , "OK", , "Error")
    
    ' Email field with validation via regular expression
    Call personalMasks.AddFieldRegex(TextBoxEmail, _
                                    "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", _
                                    "[a-zA-Z0-9._%+-@]", _
                                    True, , , , "Email", , "Partial", , "OK", , "Invalid email")
    
    ' Birth date field
    Dim minBirthDate As Date
    minBirthDate = Date - 365 * 100 ' 100 years ago
    Call personalMasks.AddFieldDate(TextBoxBirthDate, "##.##.####", _
                                   minBirthDate, Date - 365 * 18, "dd.mm.yyyy", _
                                   True, , , "DOB", , "Partial", , "OK", , "dd.mm.yyyy")
    
    ' Passport field - format XXXX XXXXXX
    Call personalMasks.AddFieldText(TextBoxPassport, "#### ######", _
                                   True, , , "Passport", , "Partial", , "OK", , "XXXX XXXXXXX")
End Sub

Private Sub TextBoxName_Change()
    UpdateFieldStatus TextBoxName
End Sub

Private Sub TextBoxPhone_Change()
    UpdateFieldStatus TextBoxPhone
End Sub

Private Sub TextBoxEmail_Change()
    UpdateFieldStatus TextBoxEmail
End Sub

Private Sub TextBoxBirthDate_Change()
    UpdateFieldStatus TextBoxBirthDate
End Sub

Private Sub TextBoxPassport_Change()
    UpdateFieldStatus TextBoxPassport
End Sub

Private Sub UpdateFieldStatus(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = personalMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(240, 255, 240) ' Light green
        Else
            textBox.BackColor = RGB(255, 240, 240) ' Light red
        End If
    End If
End Sub

Private Sub cmdSubmit_Click()
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To personalMasks.Count
        Dim currentField As clsTextboxMask
        Set currentField = personalMasks.GetItemByIndex(i)
        
        If Not currentField.IsValid Then
            allValid = False
            MsgBox "Field '" & GetFieldName(currentField.TextBox.Name) & "' is filled incorrectly!"
            currentField.SetFocus
            Exit Sub
        End If
    Next i
    
    If allValid Then
        MsgBox "All data is entered correctly! Form can be submitted.", vbInformation
        ' Here you can add code for saving data
    End If
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To personalMasks.Count
        personalMasks.GetItemByIndex(i).Clear
    Next i
End Sub

Private Function GetFieldName(controlName As String) As String
    Select Case controlName
        Case "TextBoxName": GetFieldName = "Full Name"
        Case "TextBoxPhone": GetFieldName = "Phone"
        Case "TextBoxEmail": GetFieldName = "Email"
        Case "TextBoxBirthDate": GetFieldName = "Birth Date"
        Case "TextBoxPassport": GetFieldName = "Passport"
        Case Else: GetFieldName = controlName
    End Select
End Function
```

## Example 2: Financial Data Entry Form

### Description
Form for entering financial data with range constraints and formatting.

### Implementation

```vba
' UserForm: frmFinancialData
' Controls: TextBoxSalary, TextBoxBonus, TextBoxTax, TextBoxTotal
' CommandButton: cmdCalculate, cmdReset

Dim financialMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set financialMasks = New clsTextboxMask
    
    ' Salary field - from 1000 to 1000000, with decimal values
    Call financialMasks.AddFieldNumeric(TextBoxSalary, 1000, 1000000, True, _
                                       True, "#,##0.00", , , , "Salary", , "Partial", , "OK", , "1,000.00 - 1,000,000.00")
    
    ' Bonus field - from 0 to 5000, with decimal values
    Call financialMasks.AddFieldNumeric(TextBoxBonus, 0, 50000, True, _
                                       True, "#,##0.00", , , , "Bonus", , "Partial", , "OK", , "0.00 - 500,000.00")
    
    ' Tax field - from 0 to 100 (percentages), with decimal values
    Call financialMasks.AddFieldNumeric(TextBoxTax, 0, 100, True, _
                                       True, "0.00%", , , , "Tax", , "Partial", , "OK", , "0% - 100%")
End Sub

Private Sub TextBoxSalary_Change()
    ValidateAndFormatField TextBoxSalary
End Sub

Private Sub TextBoxBonus_Change()
    ValidateAndFormatField TextBoxBonus
End Sub

Private Sub TextBoxTax_Change()
    ValidateAndFormatField TextBoxTax
End Sub

Private Sub ValidateAndFormatField(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = financialMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(240, 255, 240)
        Else
            textBox.BackColor = RGB(255, 240, 240)
        End If
    End If
End Sub

Private Sub cmdCalculate_Click()
    ' Check validity of all fields
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To financialMasks.Count
        If Not financialMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Calculate total value
        Dim salary As Double, bonus As Double, tax As Double
        salary = CDbl(financialMasks.GetItemByName("TextBoxSalary").Value)
        bonus = CDbl(financialMasks.GetItemByName("TextBoxBonus").Value)
        tax = CDbl(financialMasks.GetItemByName("TextBoxTax").Value)
        
        Dim total As Double
        total = (salary + bonus) * (1 - tax / 100)
        
        ' Format total value
        TextBoxTotal.Value = Format(total, "#,##0.00")
        TextBoxTotal.BackColor = RGB(240, 255, 240)
        
        MsgBox "Calculation completed successfully!", vbInformation
    Else
        MsgBox "Please check the correctness of all fields.", vbExclamation
    End If
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer
    For i = 1 To financialMasks.Count
        financialMasks.GetItemByIndex(i).Clear
    Next i
    TextBoxTotal.Value = ""
    TextBoxTotal.BackColor = RGB(255, 255, 255)
End Sub
```

## Example 3: Technical Parameters Entry Form

### Description
Form for entering technical parameters of equipment with various mask types.

### Implementation

```vba
' UserForm: frmTechnicalParams
' Controls: TextBoxSerial, TextBoxMAC, TextBoxIP, TextBoxVersion, TextBoxDate
' CommandButton: cmdValidate, cmdExport

Dim techMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set techMasks = New clsTextboxMask
    
    ' Serial number - format XXXX-XXXX-XXXX
    Call techMasks.AddFieldText(TextBoxSerial, "????-????-????", _
                               True, , , , "Serial No", , "Partial", , "OK", , "XXXX-XXXX-XXXX")
    
    ' MAC address - format XX:XX:XX:XX
    Call techMasks.AddFieldText(TextBoxMAC, "AA:AA:AA:AA", _
                               True, , , , "MAC Address", , "Partial", , "OK", , "XX:XX:XX:XX:XX:XX")
    
    ' IP address - validation via regular expression
    Call techMasks.AddFieldRegex(TextBoxIP, _
                                "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$", _
                                "[0-9\.]", _
                                True, , , , "IP Address", , "Partial", , "OK", , "XXX.XXX.XXX.XXX")
    
    ' Software version - format X.X.X
    Call techMasks.AddFieldText(TextBoxVersion, "A.A", _
                               True, , , , "Version", , "Partial", , "OK", , "X.X.X")
    
    ' Installation date
    Call techMasks.AddFieldDate(TextBoxDate, "##.##.####", _
                               #1/1/2000#, Date, "dd.mm.yyyy", _
                               True, , "Date", , "Partial", , "OK", , "dd.mm.yyyy")
End Sub

Private Sub TextBoxSerial_Change()
    UpdateTechFieldStatus TextBoxSerial
End Sub

Private Sub TextBoxMAC_Change()
    UpdateTechFieldStatus TextBoxMAC
End Sub

Private Sub TextBoxIP_Change()
    UpdateTechFieldStatus TextBoxIP
End Sub

Private Sub TextBoxVersion_Change()
    UpdateTechFieldStatus TextBoxVersion
End Sub

Private Sub TextBoxDate_Change()
    UpdateTechFieldStatus TextBoxDate
End Sub

Private Sub UpdateTechFieldStatus(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = techMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(220, 255, 220)
            textBox.ForeColor = RGB(0, 100, 0)
        Else
            textBox.BackColor = RGB(255, 220, 220)
            textBox.ForeColor = RGB(150, 0, 0)
        End If
    End If
End Sub

Private Sub cmdValidate_Click()
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To techMasks.Count
        Dim currentField As clsTextboxMask
        Set currentField = techMasks.GetItemByIndex(i)
        
        If Not currentField.IsValid Then
            allValid = False
            MsgBox "Field '" & GetTechFieldName(currentField.TextBox.Name) & "' contains invalid data!"
            currentField.SetFocus
            Exit Sub
        End If
    Next i
    
    If allValid Then
        MsgBox "All parameters are entered correctly!", vbInformation
    End If
End Sub

Private Sub cmdExport_Click()
    ' Check validity before export
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To techMasks.Count
        If Not techMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Here you can add code for exporting data
        MsgBox "Data exported successfully!", vbInformation
    Else
        MsgBox "Cannot export data with invalid parameters.", vbExclamation
    End If
End Sub

Private Function GetTechFieldName(controlName As String) As String
    Select Case controlName
        Case "TextBoxSerial": GetTechFieldName = "Serial Number"
        Case "TextBoxMAC": GetTechFieldName = "MAC Address"
        Case "TextBoxIP": GetTechFieldName = "IP Address"
        Case "TextBoxVersion": GetTechFieldName = "Software Version"
        Case "TextBoxDate": GetTechFieldName = "Installation Date"
        Case Else: GetTechFieldName = controlName
    End Select
End Function
```

## Example 4: Report Data Entry Form

### Description
Form with combined fields for entering report data with validation and automatic filling.

### Implementation

```vba
' UserForm: frmReportData
' Controls: TextBoxReportID, TextBoxPeriod, TextBoxAmount, TextBoxCurrency, TextBoxDescription
' CommandButton: cmdGenerate, cmdSave, cmdLoad

Dim reportMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set reportMasks = New clsTextboxMask
    
    ' Report ID - format RPT-XXXXX
    Call reportMasks.AddFieldText(TextBoxReportID, "RPT-#####", _
                                 True, , , , "Report ID", , "Partial", , "OK", , "RPT-XXXXX")
    
    ' Report period - format MM.YYYY
    Call reportMasks.AddFieldText(TextBoxPeriod, "##.####", _
                                 True, , , "Period", , "Partial", , "OK", , "MM.YYYY")
    
    ' Amount - numeric field with constraints
    Call reportMasks.AddFieldNumeric(TextBoxAmount, 0, 99999.99, True, _
                                    True, "#,##0.00", , , , "Amount", , "Partial", , "OK", , "0.00 - 9,999,999.99")
    
    ' Currency code - 3 letters
    Call reportMasks.AddFieldText(TextBoxCurrency, "AAA", _
                                 True, , , "Currency", , "Partial", , "OK", , "USD/EUR/RUB")
    
    ' Description - variable length up to 200 characters
    Call reportMasks.AddFieldVariableLength(TextBoxDescription, 200, , _
                                           True, , , , "Description", , "Partial", , "OK", , "Max 200 characters")
End Sub

Private Sub TextBoxReportID_Change()
    UpdateReportFieldStatus TextBoxReportID
End Sub

Private Sub TextBoxPeriod_Change()
    UpdateReportFieldStatus TextBoxPeriod
End Sub

Private Sub TextBoxAmount_Change()
    UpdateReportFieldStatus TextBoxAmount
End Sub

Private Sub TextBoxCurrency_Change()
    UpdateReportFieldStatus TextBoxCurrency
End Sub

Private Sub TextBoxDescription_Change()
    ' Update remaining character counter for description
    Dim field As clsTextboxMask
    Set field = reportMasks.GetItemByName(TextBoxDescription.Name)
    
    If Not field Is Nothing Then
        Dim remaining As Integer
        remaining = field.RemainingChars
        
        ' Update hint for description field
        field.PlaceholderPartial = "Remaining: " & remaining & " characters"
        
        If remaining < 10 Then
            field.PlaceholderPartialColor = RGB(255, 0, 0) ' Red color when few characters remain
        Else
            field.PlaceholderPartialColor = RGB(128, 128, 128) ' Standard color
        End If
        
        field.UpdatePlaceholder
        
        UpdateReportFieldStatus TextBoxDescription
    End If
End Sub

Private Sub UpdateReportFieldStatus(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = reportMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(245, 255, 245)
        Else
            textBox.BackColor = RGB(255, 245, 245)
        End If
    End If
End Sub

Private Sub cmdGenerate_Click()
    ' Check validity of all fields
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To reportMasks.Count
        If Not reportMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Generate report
        Dim reportID As String, period As String, amount As String, currency As String, description As String
        reportID = reportMasks.GetItemByName("TextBoxReportID").Value
        period = reportMasks.GetItemByName("TextBoxPeriod").Value
        amount = reportMasks.GetItemByName("TextBoxAmount").Value
        currency = reportMasks.GetItemByName("TextBoxCurrency").Value
        description = reportMasks.GetItemByName("TextBoxDescription").Value
        
        MsgBox "Report " & reportID & " for period " & period & " with amount " & amount & " " & currency & " generated successfully!" & vbCrLf & vbCrLf & _
               "Description: " & description, vbInformation
    Else
        MsgBox "Please check all fields before generating the report.", vbExclamation
    End If
End Sub

Private Sub cmdSave_Click()
    ' Check validity before saving
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To reportMasks.Count
        If Not reportMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Here you can add code for saving data
        MsgBox "Report data saved successfully!", vbInformation
    Else
        MsgBox "Cannot save data with invalid values.", vbExclamation
    End If
End Sub

Private Sub cmdLoad_Click()
    ' Here you can add code for loading data
    MsgBox "Data loading functionality will be implemented later.", vbInformation
End Sub
```

## Conclusion

These examples demonstrate the flexibility and power of the `clsTextboxMask` class. With this class, you can create complex forms with validation, automatic formatting, and custom user interfaces. The class supports various mask types and allows customizing the appearance of fields based on their state.

Each example can be adapted to the specific requirements of your project. The key is to understand the principles of working with the class and correctly apply different mask types depending on the type of data that needs to be entered.