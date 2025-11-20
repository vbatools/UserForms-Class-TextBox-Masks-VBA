# Developer Guide for clsTextboxMask

## Introduction

This guide is intended for developers who want to understand the internal architecture of the `clsTextboxMask` class, modify it, or create extensions. The document covers the internal mechanisms of the class, its structure, and recommendations for extending functionality.

## Class Architecture

### Class Structure

The `clsTextboxMask` class is built on the principle of managing a collection of mask elements. Key components:

- **Core properties**: Store settings for a specific textbox field
- **Items Collection**: Manages multiple mask elements
- **Event handling methods**: Process data input and update field states
- **Validation methods**: Check the correctness of entered data

### Core Internal Variables

```vba
Private mSimvolsMasks As String              ' Mask symbols
Private mBorderColorValid As Long            ' Border color for correct input
Private mBorderColorInvalid As Long          ' Border color for incorrect input
Private WithEvents mTextBox As MSForms.TextBox ' Textbox with event handling
Private mLabelPlaceholder As MSForms.Label    ' Placeholder label
Private mItems As Collection                  ' Collection of mask elements
Private mMask As String                       ' Input mask
Private mFormatValue As String                ' Value format
Private mRegexPattern As String               ' Regular expression pattern
Private mRegexFilter As String                ' Regular expression filter
Private mRegex As Object                      ' Regular expression object
```

## Internal Mechanisms

### Event Handling

The class uses `WithEvents` to monitor changes in the textbox:

```vba
Private WithEvents mTextBox As MSForms.TextBox
```

Handled events:
- `mTextBox_change` - updates placeholder and checks validity
- `mTextBox_KeyPress` - controls input characters

### Key Press Processing Algorithm

The `mTextBox_KeyPress` method implements logic for checking input characters:

```vba
Private Sub mTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Me.CurrentMaskType
        Case enumTypeMask.tDateFix, enumTypeMask.tTimeFix, enumTypeMask.tOtherFix:
            Call KeyAsciiFixLenText(KeyAscii)
        Case enumTypeMask.tNumeric
            Call NumericValue(KeyAscii)
        Case enumTypeMask.tVariableLen
            If Me.Mask <> vbNullString Then
                Call KeyAsciiFixLenText(KeyAscii)
            End If
        Case enumTypeMask.tRegex
            Call KeyAsciiRegex(KeyAscii)
    End Select
End Sub
```

### Data Validation

The `IsValidInput()` method checks the correctness of entered data depending on the mask type:

- For numeric fields, checks value range
- For dates, checks format and range
- For text masks, checks compliance with mask characters
- For regular expressions, uses the RegExp object

## Extending Functionality

### Adding New Mask Types

To add a new mask type, you need to:

1. Add a value to the `enumTypeMask` enumeration:

```vba
Public Enum enumTypeMask
    tOtherFix = 1
    tDateFix
    tTimeFix
    tNumeric
    tVariableLen
    tRegex
    tNewType  ' New mask type
    [_First] = tOtherFix
    [_Last] = tNewType
End Enum
```

2. Update handling in the `mTextBox_KeyPress` method:

```vba
Private Sub mTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Me.CurrentMaskType
        ' ... existing cases ...
        Case enumTypeMask.tNewType
            Call KeyAsciiNewType(KeyAscii)
    End Select
End Sub
```

3. Create a method to handle the new type:

```vba
Private Sub KeyAsciiNewType(ByRef KeyAscii As MSForms.ReturnInteger)
    ' Logic for handling new mask type
End Sub
```

4. Update the validation method:

```vba
Private Function IsValidInput() As Boolean
    ' ... existing logic ...
    Select Case Me.CurrentMaskType
        ' ... existing cases ...
        Case enumTypeMask.tNewType
            ' Validation for new type
    End Select
End Function
```

### Creating Custom Mask Symbols

To add new mask symbols:

1. Update the `mSimvolsMasks` constant in `class_initialize`:

```vba
Private Sub class_initialize()
    mSimvolsMasks = "#*@A" & VBA.ChrW$(1041) & VBA.ChrW$(1073) & "N"  ' Adding symbol "N"
    ' ... other settings ...
End Sub
```

2. Add handling for the new symbol in `KeyAsciiFixLenText`:

```vba
Private Sub KeyAsciiFixLenText(ByRef KeyAscii As MSForms.ReturnInteger)
    ' ... existing logic ...
    Select Case endLetter
        ' ... existing cases ...
        Case "N"  ' New mask symbol
            ' Processing for symbol N
    End Select
End Sub
```

## Performance Tuning

### Optimizing Event Handling

To improve performance when working with a large number of fields:

1. Use optimized checking methods:

```vba
' Instead of multiple function calls, cache values
Private Function IsValidInput() As Boolean
    Static cachedValue As String
    Static cachedResult As Boolean
    
    If Me.Value <> cachedValue Then
        cachedValue = Me.Value
        ' Perform validation and save result
        cachedResult = PerformValidation()
    End If
    
    IsValidInput = cachedResult
End Function
```

2. Limit placeholder update frequency:

```vba
Private lastUpdate As Double
Private Const UPDATE_INTERVAL As Double = 0.1 ' 100ms

Private Sub mTextBox_change()
    If Timer - lastUpdate > UPDATE_INTERVAL Then
        Call UpdatePlaceholder
        Call IsValidInput
        lastUpdate = Timer
    End If
End Sub
```

## Error Handling

### Internal Error Handlers

The class uses several approaches for error handling:

1. Parameter checking in the `AddField` method:

```vba
Private Sub AddField(ByRef inputTextBox As MSForms.TextBox, ...)
    If inputTextBox Is Nothing Then
        Call Err.Raise(vbObjectError + 101, "clsTextboxMask", "TextBox cannot be Nothing")
        Exit Sub
    End If
    
    If maskType < enumTypeMask.[_First] Or maskType > enumTypeMask.[_Last] Then
        Call Err.Raise(vbObjectError + 102, "clsTextboxMask", "Invalid mask type")
        Exit Sub
    End If
    
    If IsControlInCollection(inputTextBox.Name) Then
        Call Err.Raise(vbObjectError + 103, "clsTextboxMask", "The item has already been created")
        Exit Sub
    End If
End Sub
```

2. Error handling when working with regular expressions:

```vba
Private Function IsValidRegexInput() As Boolean
    If mRegex Is Nothing Then
        Call InitializeRegex
        If mRegex Is Nothing Then
            IsValidRegexInput = False
            Exit Function
        End If
    End If
    mRegex.pattern = Me.RegexPattern
    On Error GoTo ErrorHandler
    IsValidRegexInput = mRegex.Test(Me.Value)
    Exit Function

ErrorHandler:
    IsValidRegexInput = False
End Function
```

## Testing and Debugging

### Unit Testing

To test class functionality, it's recommended to create test scenarios:

```vba
' Testing module: modTextboxMaskTests
Sub TestNumericField()
    Dim mask As New clsTextboxMask
    Dim tb As MSForms.TextBox
    Set tb = CreateTestTextBox()
    
    Call mask.AddFieldNumeric(tb, 0, 100, True)
    
    ' Test valid value input
    tb.Value = "50.5"
    Debug.Assert mask.IsValid = True, "Valid value should be accepted"
    
    ' Test invalid value input
    tb.Value = "150"
    Debug.Assert mask.IsValid = False, "Invalid value should be rejected"
    
    Debug.Print "Numeric field test passed"
End Sub

Sub TestTextField()
    Dim mask As New clsTextboxMask
    Dim tb As MSForms.TextBox
    Set tb = CreateTestTextBox()
    
    Call mask.AddFieldText(tb, "###-###")
    
    ' Test valid value input
    tb.Value = "123-456"
    Debug.Assert mask.IsValid = True, "Valid value should be accepted"
    
    ' Test invalid value input
    tb.Value = "123-abc"
    Debug.Assert mask.IsValid = False, "Invalid value should be rejected"
    
    Debug.Print "Text field test passed"
End Sub

Private Function CreateTestTextBox() As MSForms.TextBox
    ' Creating test textbox
    Dim tb As MSForms.TextBox
    Set tb = New MSForms.TextBox
    Set CreateTestTextBox = tb
End Function
```

### Validation Debugging

To debug the validation process, you can add logging:

```vba
Private Function IsValidInput() As Boolean
    ' Logging for debugging
    #If DEBUG_MODE Then
        Debug.Print "Validation check: " & Me.TextBox.Name
        Debug.Print "Value: " & Me.Value
        Debug.Print "Mask type: " & Me.CurrentMaskType
    #End If
    
    ' Main validation logic
    Select Case Me.CurrentMaskType
        ' ... main logic ...
    End Select
End Function
```

## Usage Recommendations

### Best Practices

1. **Timely Initialization**:
   - Initialize masks in the `UserForm_Initialize` event, not in `Activate`
   - Ensure all textboxes exist before adding masks

2. **Memory Management**:
   - Use the `RemoveItem` method to remove mask elements
   - Release object references when closing the form:

```vba
Private Sub UserForm_Terminate()
    If Not mItems Is Nothing Then
        Dim i As Integer
        For i = mItems.Count To 1 Step -1
            Dim item As clsTextboxMask
            Set item = mItems(i)
            item.RemoveItem
        Next i
        Set mItems = Nothing
    End If
End Sub
```

3. **Exception Handling**:
   - Wrap method calls in error handling blocks
   - Check that elements exist before accessing them

### Performance Recommendations

1. **Optimization for Large Numbers of Fields**:
   - Use one collection for multiple fields
   - Avoid frequent placeholder updates

2. **Efficient Regular Expression Usage**:
   - Cache RegExp objects
   - Avoid complex regular expressions in real-time

## Extending the Class

### Creating Derived Classes

You can create specialized classes based on `clsTextboxMask`:

```vba
' clsPhoneMask.cls
Implements clsTextboxMask

Private baseMask As clsTextboxMask

Private Sub Class_Initialize()
    Set baseMask = New clsTextboxMask
End Sub

Public Sub AddPhoneField(ByRef inputTextBox As MSForms.TextBox)
    Call baseMask.AddFieldText(inputTextBox, "+7(###) ###-##-##", _
                              True, RGB(0, 128, 0), RGB(255, 0, 0))
End Sub

' Delegating properties
Public Property Get IsValid() As Boolean
    IsValid = baseMask.IsValid
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = baseMask.TextBox
End Property

' ... other properties and methods
```

### Adding Custom Validators

To add custom validation functions:

```vba
' Adding delegate for custom validation
Public Type ValidationDelegate
    ValidateProc As String  ' Validation procedure name
End Type

Private customValidator As ValidationDelegate

Public Sub SetCustomValidator(validatorName As String)
    customValidator.ValidateProc = validatorName
End Sub

Private Function IsValidInput() As Boolean
    ' ... standard validation ...
    
    ' Custom validation
    If customValidator.ValidateProc <> "" Then
        IsValidInput = Application.Run(customValidator.ValidateProc, Me.Value)
    End If
End Function
```

## Conclusion

The `clsTextboxMask` class provides a flexible and extensible architecture for creating validated textboxes in VBA. Its modular structure allows for easy addition of new mask types and validation functions.

When developing class extensions, it's recommended to follow these principles:
- Maintain compatibility with the existing API
- Ensure proper error handling
- Consider performance when working with large numbers of fields
- Write tests for new functions

These recommendations will help you effectively use and extend the class functionality in your projects.