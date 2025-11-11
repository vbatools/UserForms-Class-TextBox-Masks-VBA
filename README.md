# VBA TextBox Masks Class

**clsTextboxMask** is a powerful VBA class that allows creating text fields with input masks in Excel and other Office applications. It provides input validation, placeholder display, and visual indication of field fill status.

## Screenshots

![User Forms](User_Forms.png)

## Key Features

- Support for various input mask types (numbers, dates, time, text, regular expressions)
- Real-time input validation
- Display of placeholders with different statuses (empty, partially filled, completely filled, incorrect)
- Visual indication of input correctness through border color
- Support for numeric values with range, sign, and decimal constraints
- Support for variable-length text
- Support for validation via regular expressions
- Support for placeholder color customization based on field status
- Support for placeholder templates with markers: {mask}, {filled}, {remaining}, {holder}, {RegexPattern}, {RegexFilter}, {percent}

## Installation

1. Copy the `clsTextboxMask.cls` file to your VBA project
2. Use the class in your UserForms

## Main Properties

| Property | Type | Description |
|----------|------|-------------|
| `TextBox` | MSForms.TextBox | Reference to the text box to which the mask is applied |
| `LabelPlaceholder` | MSForms.Label | Reference to the placeholder label that displays hints |
| `Mask` | String | Input mask defining allowed characters |
| `Value` | String | Current value of the text field |
| `CurrentMaskType` | enumTypeMask | Type of current mask |
| `Min` | Single | Minimum value for numeric fields |
| `Max` | Single | Maximum value for numeric fields |
| `IsNegative` | Boolean | Whether negative values are allowed |
| `IsDecemal` | Boolean | Whether decimal values are allowed |
| `BorderColorValid` | Long | Border color when input is correct |
| `BorderColorInvalid` | Long | Border color when input is incorrect |
| `PlaceholderEmptyColor` | Long | Placeholder text color for empty field |
| `PlaceholderPartialColor` | Long | Placeholder text color for partially filled field |
| `PlaceholderCompleteColor` | Long | Placeholder text color for completely filled field |
| `PlaceholderInvalidColor` | Long | Placeholder text color for field with incorrect data |
| `PlaceholderEmpty` | String | Placeholder text for empty field |
| `PlaceholderPartial` | String | Placeholder text for partially filled field |
| `PlaceholderComplete` | String | Placeholder text for completely filled field |
| `PlaceholderInvalid` | String | Placeholder text for field with incorrect data |

## Mask Types

The class supports the following mask types:

| Mask Type | Value | Description |
|-----------|-------|-------------|
| `tOtherFix` | 1 | Fixed mask with various characters |
| `tDateFix` | 2 | Fixed mask for dates |
| `tTimeFix` | 3 | Fixed mask for time |
| `tNumeric` | 4 | Numeric mask with range limitation capability |
| `tVariableLen` | 5 | Variable-length mask |
| `tRegex` | 6 | Regular expression-based mask |

## Main Methods

### `AddFieldNumeric`
Adds a numeric field with specified validation parameters.

```vba
Dim numField As New clsTextboxMask
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True, _
                             allowNegative:=False)
```

**Parameters:**
- `inputTextBox` - text field to which the mask is applied
- `minValue` - minimum allowed value
- `maxValue` - maximum allowed value
- `allowDecimal` - permission for decimal values input
- `allowNegative` - permission for negative values input
- `showPlaceholder` - placeholder display (optional)
- `numberFormat` - number display format (optional)
- `BorderColorValid` - border color when input is correct (optional)
- `BorderColorInvalid` - border color when input is incorrect (optional)
- `PlaceholderColor` - placeholder color (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalid` - placeholder text for field with incorrect data (optional)
- `PlaceHolderTemplete` - placeholder template (optional)

### `AddFieldDate`
Adds a date input field with specified validation parameters.

```vba
Dim dateField As New clsTextboxMask
Call dateField.AddFieldDate(inputTextBox:=Me.TextBox2, _
                           dateMask:="##.##.####", _
                           minDate:=#1/1/2020#, _
                           maxDate:=#12/31/2030#, _
                           dateFormat:="dd.mm.yyyy")
```

**Parameters:**
- `inputTextBox` - text field to which the mask is applied
- `dateMask` - date input mask
- `minDate` - minimum allowed date
- `maxDate` - maximum allowed date
- `dateFormat` - date display format (optional)
- `showPlaceholder` - placeholder display (optional)
- `BorderColorValid` - border color when input is correct (optional)
- `BorderColorInvalid` - border color when input is incorrect (optional)
- `PlaceholderColor` - placeholder color (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalid` - placeholder text for field with incorrect data (optional)
- `PlaceHolderTemplete` - placeholder template (optional)

### `AddFieldTime`
Adds a time input field with specified validation parameters.

```vba
Dim timeField As New clsTextboxMask
Call timeField.AddFieldTime(inputTextBox:=Me.TextBox3, _
                           timeMask:="##:##", _
                           minTime:=#0:00:00#, _
                           maxTime:=#23:59#, _
                           timeFormat:="hh:mm")
```

**Parameters:**
- `inputTextBox` - text field to which the mask is applied
- `timeMask` - time input mask
- `minTime` - minimum allowed time
- `maxTime` - maximum allowed time
- `timeFormat` - time display format (optional)
- `showPlaceholder` - placeholder display (optional)
- `BorderColorValid` - border color when input is correct (optional)
- `BorderColorInvalid` - border color when input is incorrect (optional)
- `PlaceholderColor` - placeholder color (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalid` - placeholder text for field with incorrect data (optional)
- `PlaceHolderTemplete` - placeholder template (optional)

### `AddFieldText`
Adds a text field with a specified input mask.

```vba
Dim textField As New clsTextboxMask
Call textField.AddFieldText(inputTextBox:=Me.TextBox4, _
                           textMask:="+7(*##) @# A# #Б#")  ' Letters-digits
```

**Parameters:**
- `inputTextBox` - text field to which the mask is applied
- `textMask` - text input mask
- `showPlaceholder` - placeholder display (optional)
- `BorderColorValid` - border color when input is correct (optional)
- `BorderColorInvalid` - border color when input is incorrect (optional)
- `PlaceholderColor` - placeholder color (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalid` - placeholder text for field with incorrect data (optional)
- `PlaceHolderTemplete` - placeholder template (optional)

### `AddFieldVariableLength`
Adds a field with variable text length.

```vba
Dim varField As New clsTextboxMask
Call varField.AddFieldVariableLength(inputTextBox:=Me.TextBox5, _
                                    maxLength:=10, _
                                    textMask:="##")
```

**Parameters:**
- `inputTextBox` - text field to which the mask is applied
- `maxLength` - maximum text length
- `textMask` - text input mask (optional)
- `showPlaceholder` - placeholder display (optional)
- `BorderColorValid` - border color when input is correct (optional)
- `BorderColorInvalid` - border color when input is incorrect (optional)
- `PlaceholderColor` - placeholder color (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalid` - placeholder text for field with incorrect data (optional)
- `PlaceHolderTemplete` - placeholder template (optional)

### `AddFieldRegex`
Adds a field with validation via regular expression.

```vba
Dim regexField As New clsTextboxMask
Call regexField.AddFieldRegex(inputTextBox:=Me.TextBox6, _
                             RegexPattern:="^[A-Z]{2}\d{4}$", _
                             RegexFilter:="[A-Z0-9]")
```

**Parameters:**
- `inputTextBox` - text field to which the mask is applied
- `RegexPattern` - regular expression pattern for validation
- `RegexFilter` - regular expression filter
- `showPlaceholder` - placeholder display (optional)
- `BorderColorValid` - border color when input is correct (optional)
- `BorderColorInvalid` - border color when input is incorrect (optional)
- `PlaceholderColor` - placeholder color (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalid` - placeholder text for field with incorrect data (optional)
- `PlaceHolderTemplete` - placeholder template (optional)

### `IsValid`
Checks the correctness of data entered in the text field.

```vba
If myField.IsValid() Then
    MsgBox "Data is correct!"
Else
    MsgBox "Data is incorrect!"
End If
```

### `Clear`
Clears the text field.

```vba
myField.Clear()
```

### `SetFocus`
Sets focus to the text field.

```vba
myField.SetFocus()
```

### `RemoveItem`
Removes the text box mask item and associated components.

```vba
myField.RemoveItem()
```

## Mask Symbols

When creating text masks, the following symbols are used:

| Symbol | Description |
|--------|-------------|
| `#` | Digits (0-9) |
| `@` | Latin letters (A-Z, a-z) |
| `A` | Latin letters and digits (A-Z, a-z, 0-9) |
| `Б` | Cyrillic letters |
| `б` | Cyrillic letters and digits |
| `*` | Any characters |

## Usage Examples

### 1. Numeric field with constraints:
```vba
Dim numField As New clsTextboxMask
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True, _
                             allowNegative:=False)
```

### 2. Date field:
```vba
Dim dateField As New clsTextboxMask
Call dateField.AddFieldDate(inputTextBox:=Me.TextBox2, _
                           dateMask:="##.##.####", _
                           minDate:=#1/2020#, _
                           maxDate:=#12/31/2030#, _
                           dateFormat:="dd.mm.yyyy")
```

### 3. Text field with mask:
```vba
Dim textField As New clsTextboxMask
Call textField.AddFieldText(inputTextBox:=Me.TextBox3, _
                           textMask:="+7(*##) @# A# #Б#")  ' Letters-digits
```

## Dependencies

- MSForms.TextBox
- MSForms.Label
- VBScript.RegExp (for regular expression validation)

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.