# User Guide for clsTextboxMask

## Introduction

The `clsTextboxMask` class provides a powerful tool for creating textboxes with input masks in VBA. This guide contains step-by-step instructions for installation, configuration, and usage of the class in your projects.

## Installation

### Requirements
- Microsoft Excel (2010 or newer)
- VBA support enabled
- Permission to use MSForms objects

### Installing the Class
1. Open your Excel file with VBA project (Alt+F11)
2. In the VBA editor window, select "File" → "Import File"
3. Select the `clsTextboxMask.cls` file
4. The class will be added to your project

## Quick Start

### Simple Usage Example

Create a UserForm and add a TextBox to it. Then use the following code:

```vba
Dim maskField As clsTextboxMask
Set maskField = New clsTextboxMask
Call maskField.AddFieldText(Me.TextBox1, "###-##-##")
```

This code will create a field for entering a phone number in the format "123-45-67".

### Numeric Field Example

```vba
Dim numField As clsTextboxMask
Set numField = New clsTextboxMask
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True)
```

## Detailed Usage Examples

### 1. Creating a Date Input Field

```vba
Private Sub UserForm_Initialize()
    Dim dateField As clsTextboxMask
    Set dateField = New clsTextboxMask
    Call dateField.AddFieldDate(inputTextBox:=Me.TextBoxDate, _
                               dateMask:="##.##.####", _
                               minDate:=#1/1/2020#, _
                               maxDate:=#12/31/2030#, _
                               dateFormat:="dd.mm.yyyy")
End Sub
```

This code creates a date input field in the format "dd.mm.yyyy" with a date range limitation from January 1, 2020 to December 31, 2030.

### 2. Creating a Time Input Field

```vba
Private Sub UserForm_Initialize()
    Dim timeField As clsTextboxMask
    Set timeField = New clsTextboxMask
    Call timeField.AddFieldTime(inputTextBox:=Me.TextBoxTime, _
                               timeMask:="##:##", _
                               minTime:=#0:00:00#, _
                               maxTime:=#23:59#, _
                               timeFormat:="hh:mm")
End Sub
```

This code creates a time input field in the format "hh:mm".

### 3. Creating a Phone Number Field

```vba
Private Sub UserForm_Initialize()
    Dim phoneField As clsTextboxMask
    Set phoneField = New clsTextboxMask
    Call phoneField.AddFieldText(inputTextBox:=Me.TextBoxPhone, _
                                textMask:="+7(###) ###-##-##")
End Sub
```

This code creates a field for entering a Russian phone number with automatic formatting.

### 4. Creating a Numeric Field with Constraints

```vba
Private Sub UserForm_Initialize()
    Dim numField As clsTextboxMask
    Set numField = New clsTextboxMask
    Call numField.AddFieldNumeric(inputTextBox:=Me.TextBoxNumber, _
                                 minValue:=-10, _
                                 maxValue:=100, _
                                 allowDecimal:=True)
End Sub
```

This code creates a field for entering numbers from -10 to 100, including decimal values.

### 5. Creating a Variable Length Field

```vba
Private Sub UserForm_Initialize()
    Dim varField As clsTextboxMask
    Set varField = New clsTextboxMask
    Call varField.AddFieldVariableLength(inputTextBox:=Me.TextBoxName, _
                                        maxLength:=50)
End Sub
```

This code creates a text input field with a maximum length of 50 characters.

### 6. Creating a Regular Expression Field

```vba
Private Sub UserForm_Initialize()
    Dim emailField As clsTextboxMask
    Set emailField = New clsTextboxMask
    Call emailField.AddFieldRegex(inputTextBox:=Me.TextBoxEmail, _
                                 RegexPattern:="^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", _
                                 RegexFilter:="[a-zA-Z0-9._%+-@]")
End Sub
```

This code creates an email input field with validation via regular expression.

## Appearance Customization

### Changing Border Colors

```vba
Private Sub UserForm_Initialize()
    Dim field As clsTextboxMask
    Set field = New clsTextboxMask
    Call field.AddFieldText(inputTextBox:=Me.TextBox1, _
                           textMask:="###-###", _
                           BorderColorValid:=RGB(0, 128, 0), _
                           BorderColorInvalid:=RGB(255, 0, 0))
End Sub
```

### Placeholder Configuration

```vba
Private Sub UserForm_Initialize()
    Dim field As clsTextboxMask
    Set field = New clsTextboxMask
    Call field.AddFieldText(inputTextBox:=Me.TextBox1, _
                           textMask:="###-###", _
                           PlaceholderEmpty:="Enter code", _
                           PlaceholderEmptyColor:=RGB(128, 128, 128), _
                           PlaceholderComplete:="Code entered", _
                           PlaceholderCompleteColor:=RGB(0, 128, 0))
End Sub
```

## Working with Field Collections

The class allows managing multiple fields simultaneously:

```vba
Private Sub UserForm_Initialize()
    Dim formMasks As clsTextboxMask
    Set formMasks = New clsTextboxMask
    
    ' Adding multiple fields
    Call formMasks.AddFieldText(Me.TextBox1, "###-##-##")
    Call formMasks.AddFieldDate(Me.TextBox2, "##.##.####", #1/1/2000#, #12/31/2030#)
    Call formMasks.AddFieldNumeric(Me.TextBox3, 0, 100, False)
    
    ' Checking validity of all fields
    Dim isValid As Boolean
    isValid = True
    
    Dim i As Integer
    For i = 1 To formMasks.Count
        If Not formMasks.GetItemByIndex(i).IsValid Then
            isValid = False
            Exit For
        End If
    Next i
    
    MsgBox "All fields are valid: " & isValid
End Sub
```

## Practical Tips

### 1. Form Event Handling

To respond to changes in masked fields, use textbox events:

```vba
Private Sub TextBox1_Change()
    Dim field As clsTextboxMask
    Set field = clsTB.GetItemByName(TextBox1.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            ' Field is filled correctly
            TextBox1.BackColor = RGB(240, 255, 240) ' Light green
        Else
            ' Field is filled incorrectly
            TextBox1.BackColor = RGB(255, 240, 240) ' Light red
        End If
    End If
End Sub
```

### 2. Focus Management

```vba
Private Sub CommandButton1_Click()
    ' Set focus to a specific field
    Dim field As clsTextboxMask
    Set field = clsTB.GetItemByName(TextBox1.Name)
    If Not field Is Nothing Then field.SetFocus
End Sub
```

### 3. Field Clearing

```vba
Private Sub CommandButton2_Click()
    ' Clear all fields
    Dim formMasks As clsTextboxMask
    Set formMasks = New clsTextboxMask
    
    Dim i As Integer
    For i = 1 To formMasks.Count
        formMasks.GetItemByIndex(i).Clear
    Next i
End Sub
```

### 4. Removing Mask Elements

```vba
Private Sub CommandButton3_Click()
    ' Remove a specific field
    Dim field As clsTextboxMask
    Set field = clsTB.GetItemByName(TextBox1.Name)
    If Not field Is Nothing Then field.RemoveItem
End Sub
```

## Common Errors and Solutions

### Error: "The item has already been created"

This error occurs when trying to add a mask to a textbox that already has a mask. Solution:

```vba
' Check if element already exists
Dim existingField As clsTextboxMask
Set existingField = clsTB.GetItemByName(TextBox1.Name)

If existingField Is Nothing Then
    ' Element doesn't exist, safe to add
    Call clsTB.AddFieldText(TextBox1, "###-###")
Else
    ' Element already exists, can update its properties
    existingField.Mask = "###-###"
End If
```

### Error: "TextBox cannot be Nothing"

Ensure the textbox exists and is not Nothing before adding a mask:

```vba
If Not Me.TextBox1 Is Nothing Then
    Call maskField.AddFieldText(Me.TextBox1, "###-###")
End If
```

## Real-World Scenarios

### User Registration Form

```vba
Private Sub UserForm_Initialize()
    Dim formMasks As clsTextboxMask
    Set formMasks = New clsTextboxMask
    
    ' Email field
    Call formMasks.AddFieldRegex(Me.TextBoxEmail, _
                                "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", _
                                "[a-zA-Z0-9._%+-@]", _
                                True, , , , "Email", , "Partial", , "OK", , "Invalid email")
    
    ' Phone field
    Call formMasks.AddFieldText(Me.TextBoxPhone, "+7(###) ###-##-##", _
                               True, , , , "Phone", , "Partial", , "OK", , "Invalid format")
    
    ' Age field
    Call formMasks.AddFieldNumeric(Me.TextBoxAge, 18, 100, False, _
                                  True, , , , "Age", , "Partial", , "OK", , "18-100 years")
    
    ' Birth date field
    Dim birthDate As Date
    birthDate = Date - 365 * 18 ' 18 years ago
    Call formMasks.AddFieldDate(Me.TextBoxBirth, "##.##.####", _
                               birthDate - 365 * 50, Date, _
                               "dd.mm.yyyy", True, , , , "DOB", , "Partial", , "OK", , "dd.mm.yyyy")
End Sub

Private Sub CommandButtonSubmit_Click()
    Dim formMasks As clsTextboxMask
    Set formMasks = New clsTextboxMask
    
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To formMasks.Count
        If Not formMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            MsgBox "Field " & formMasks.GetItemByIndex(i).TextBox.Name & " is invalid"
            formMasks.GetItemByIndex(i).SetFocus
            Exit Sub
        End If
    Next i
    
    If allValid Then
        MsgBox "All fields are valid! Form can be submitted."
    End If
End Sub
```

## Advanced Features

### Using Placeholder Templates

Placeholder templates allow dynamically displaying field status information:

```vba
Call maskField.AddFieldText(Me.TextBox1, "####-####-####", _
                           True, , , , , , , "Template: {holder} Remaining: {remaining}")
```

Available markers:
- `{mask}` - displays the mask
- `{filled}` - number of filled characters
- `{remaining}` - number of remaining characters
- `{holder}` - placeholder with mask
- `{RegexPattern}` - regular expression pattern
- `{RegexFilter}` - regular expression filter
- `{percent}` - fill percentage

### Custom Masks

You can create complex masks with combinations of various symbols:

```vba
' Car license plate mask: A123AA123
Call maskField.AddFieldText(Me.TextBoxCarNumber, "@###@@###", _
                           True, , , , "Number", , "Partial", , "OK", , "A123AA123")

' Cyrillic character mask: А123БВ456
Call maskField.AddFieldText(Me.TextBoxCyrillic, "Б###ББ###", _
                           True, , , , "Number", , "Partial", , "OK", , "А123БВ456")
```

## Conclusion

The `clsTextboxMask` class provides a powerful and flexible tool for creating validated textboxes in VBA. With it, you can improve the user interface of your applications, ensuring correct data entry and simplifying the validation process.

Use the provided examples as a starting point for creating your own solutions using this class.