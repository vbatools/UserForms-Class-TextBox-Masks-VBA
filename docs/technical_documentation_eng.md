# Technical Documentation for clsTextboxMask

## Table of Contents
1. [Class Overview](#class-overview)
2. [Class Architecture](#class-architecture)
3. [Properties](#properties)
4. [Methods](#methods)
5. [Events](#events)
6. [Constants and Enumerations](#constants-and-enumerations)
7. [Implementation Details](#implementation-details)
8. [Dependencies](#dependencies)

## Class Overview

The `clsTextboxMask` class is a powerful VBA tool that allows creating textboxes with input masks in Excel and other Office applications. It provides input validation, placeholder display, and visual indication of field fill status.

### Key Features
- Support for various input mask types (numbers, dates, time, text, regular expressions)
- Real-time input validation
- Display of placeholders with different statuses (empty, partially filled, completely filled, invalid)
- Visual indication of input correctness through border color
- Support for numeric values with range, sign and decimal restrictions
- Support for variable length text
- Support for validation via regular expressions
- Support for placeholder color customization based on field status
- Support for placeholder templates with markers

## Class Architecture

### enumTypeMask Enumeration
Defines the types of supported masks:
- `tOtherFix` (1) - Fixed mask with various characters
- `tDateFix` (2) - Fixed mask for dates
- `tTimeFix` (3) - Fixed mask for time
- `tNumeric` (4) - Numeric mask with range limitation capability
- `tVariableLen` (5) - Variable length mask
- `tRegex` (6) - Regular expression-based mask

### Core Class Properties

| Property | Type | Description |
|----------|------|-------------|
| `TextBox` | MSForms.TextBox | Reference to the textbox field to which the mask is applied |
| `LabelPlaceholder` | MSForms.Label | Reference to the placeholder label that displays hints |
| `Mask` | String | Input mask that defines allowed characters |
| `Value` | String | Current value of the textbox field |
| `CurrentMaskType` | enumTypeMask | Type of the current mask |
| `Min` | Single | Minimum value for numeric fields |
| `Max` | Single | Maximum value for numeric fields |
| `IsDecimal` | Boolean | Are decimal values allowed |
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

## Properties

### Core Properties
- `TextBox` - Gets or sets the textbox field to which the mask is applied
- `LabelPlaceholder` - Gets or sets the placeholder label that displays hints
- `Mask` - Gets or sets the input mask that defines allowed characters
- `Value` - Gets or sets the current value of the textbox field
- `CurrentMaskType` - Gets or sets the type of the current mask
- `Min` - Gets or sets the minimum value for numeric fields
- `Max` - Gets or sets the maximum value for numeric fields
- `IsDecimal` - Gets or sets whether decimal values are allowed
- `BorderColorValid` - Gets or sets the border color when input is correct
- `BorderColorInvalid` - Gets or sets the border color when input is incorrect
- `PlaceholderEmptyColor` - Gets or sets the placeholder text color for empty field
- `PlaceholderPartialColor` - Gets or sets the placeholder text color for partially filled field
- `PlaceholderCompleteColor` - Gets or sets the placeholder text color for completely filled field
- `PlaceholderInvalidColor` - Gets or sets the placeholder text color for field with incorrect data
- `PlaceholderEmpty` - Gets or sets the placeholder text for empty field
- `PlaceholderPartial` - Gets or sets the placeholder text for partially filled field
- `PlaceholderComplete` - Gets or sets the placeholder text for completely filled field
- `PlaceholderInvalid` - Gets or sets the placeholder text for field with incorrect data

### Additional Properties
- `PlaceholderMask` - Gets the current placeholder mask showing remaining characters to be filled
- `PlaceHolderTemplate` - Gets or sets the placeholder template with markers
- `VisibleLabelPlaceholder` - Gets or sets the visibility of the placeholder label
- `Items` - Gets the collection of all textbox mask items
- `Count` - Gets the number of items in the collection
- `RemainingChars` - Gets the number of remaining characters to fill
- `FormatValue` - Gets or sets the format for displaying the value
- `LenValue` - Gets the length of the current value
- `lenMask` - Gets the length of the mask
- `Version` - Gets version information about the class

## Methods

### AddFieldNumeric
Adds a numeric field with specified validation parameters.

**Syntax:**
```vba
Public Sub AddFieldNumeric(ByRef inputTextBox As MSForms.TextBox, _
        ByVal minValue As Single, _
        ByVal maxValue As Single, _
        ByVal allowDecimal As Boolean, _
        Optional showPlaceholder As Boolean = True, _
        Optional numberFormat As String = "#.0", _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderEmptyColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartialColor As XlRgbColor = 0, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderCompleteColor As XlRgbColor = 0, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalidColor As XlRgbColor = 0, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
```

**Parameters:**
- `inputTextBox` - textbox field to which the mask is applied
- `minValue` - minimum allowed value
- `maxValue` - maximum allowed value
- `allowDecimal` - allow input of decimal values
- `showPlaceholder` - show placeholder (optional)
- `numberFormat` - number display format (optional)
- `BorderColorValid` - border color for correct input (optional)
- `BorderColorInvalid` - border color for incorrect input (optional)
- `PlaceholderEmptyColor` - placeholder color for empty field (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartialColor` - placeholder color for partially filled field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderCompleteColor` - placeholder color for completely filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalidColor` - placeholder color for field with invalid data (optional)
- `PlaceholderInvalid` - placeholder text for field with invalid data (optional)
- `PlaceHolderTemplate` - placeholder template (optional)

### AddFieldDate
Adds a date input field with specified validation parameters.

**Syntax:**
```vba
Public Sub AddFieldDate(ByRef inputTextBox As MSForms.TextBox, ByVal dateMask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional dateFormat As String = "dd.mm.yyyy", _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderEmptyColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartialColor As XlRgbColor = 0, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderCompleteColor As XlRgbColor = 0, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalidColor As XlRgbColor = 0, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
```

**Parameters:**
- `inputTextBox` - textbox field to which the mask is applied
- `dateMask` - date input mask
- `minDate` - minimum allowed date
- `maxDate` - maximum allowed date
- `dateFormat` - date display format (optional)
- `showPlaceholder` - show placeholder (optional)
- `BorderColorValid` - border color for correct input (optional)
- `BorderColorInvalid` - border color for incorrect input (optional)
- `PlaceholderEmptyColor` - placeholder color for empty field (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartialColor` - placeholder color for partially filled field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderCompleteColor` - placeholder color for completely filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalidColor` - placeholder color for field with invalid data (optional)
- `PlaceholderInvalid` - placeholder text for field with invalid data (optional)
- `PlaceHolderTemplate` - placeholder template (optional)

### AddFieldTime
Adds a time input field with specified validation parameters.

**Syntax:**
```vba
Public Sub AddFieldTime(ByRef inputTextBox As MSForms.TextBox, ByVal timeMask As String, _
        ByVal minTime As Date, _
        ByVal maxTime As Date, _
        Optional timeFormat As String = "hh:mm", _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderEmptyColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartialColor As XlRgbColor = 0, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderCompleteColor As XlRgbColor = 0, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalidColor As XlRgbColor = 0, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
```

**Parameters:**
- `inputTextBox` - textbox field to which the mask is applied
- `timeMask` - time input mask
- `minTime` - minimum allowed time
- `maxTime` - maximum allowed time
- `timeFormat` - time display format (optional)
- `showPlaceholder` - show placeholder (optional)
- `BorderColorValid` - border color for correct input (optional)
- `BorderColorInvalid` - border color for incorrect input (optional)
- `PlaceholderEmptyColor` - placeholder color for empty field (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartialColor` - placeholder color for partially filled field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderCompleteColor` - placeholder color for completely filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalidColor` - placeholder color for field with invalid data (optional)
- `PlaceholderInvalid` - placeholder text for field with invalid data (optional)
- `PlaceHolderTemplate` - placeholder template (optional)

### AddFieldText
Adds a text field with specified input mask.

**Syntax:**
```vba
Public Sub AddFieldText(ByRef inputTextBox As MSForms.TextBox, _
        ByVal textMask As String, _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderEmptyColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartialColor As XlRgbColor = 0, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderCompleteColor As XlRgbColor = 0, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalidColor As XlRgbColor = 0, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
```

**Parameters:**
- `inputTextBox` - textbox field to which the mask is applied
- `textMask` - text input mask
- `showPlaceholder` - show placeholder (optional)
- `BorderColorValid` - border color for correct input (optional)
- `BorderColorInvalid` - border color for incorrect input (optional)
- `PlaceholderEmptyColor` - placeholder color for empty field (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartialColor` - placeholder color for partially filled field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderCompleteColor` - placeholder color for completely filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalidColor` - placeholder color for field with invalid data (optional)
- `PlaceholderInvalid` - placeholder text for field with invalid data (optional)
- `PlaceHolderTemplate` - placeholder template (optional)

### AddFieldVariableLength
Adds a field with variable text length.

**Syntax:**
```vba
Public Sub AddFieldVariableLength(ByRef inputTextBox As MSForms.TextBox, _
        ByVal maxLength As Integer, _
        Optional textMask As String = vbNullString, _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderEmptyColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartialColor As XlRgbColor = 0, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderCompleteColor As XlRgbColor = 0, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalidColor As XlRgbColor = 0, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
```

**Parameters:**
- `inputTextBox` - textbox field to which the mask is applied
- `maxLength` - maximum text length
- `textMask` - text input mask (optional)
- `showPlaceholder` - show placeholder (optional)
- `BorderColorValid` - border color for correct input (optional)
- `BorderColorInvalid` - border color for incorrect input (optional)
- `PlaceholderEmptyColor` - placeholder color for empty field (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartialColor` - placeholder color for partially filled field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderCompleteColor` - placeholder color for completely filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalidColor` - placeholder color for field with invalid data (optional)
- `PlaceholderInvalid` - placeholder text for field with invalid data (optional)
- `PlaceHolderTemplate` - placeholder template (optional)

### AddFieldRegex
Adds a field with validation via regular expression.

**Syntax:**
```vba
Public Sub AddFieldRegex(ByRef inputTextBox As MSForms.TextBox, _
        ByVal RegexPattern As String, _
        ByVal RegexFilter As String, _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderEmptyColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartialColor As XlRgbColor = 0, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderCompleteColor As XlRgbColor = 0, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalidColor As XlRgbColor = 0, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
```

**Parameters:**
- `inputTextBox` - textbox field to which the mask is applied
- `RegexPattern` - regular expression pattern for validation
- `RegexFilter` - regular expression filter
- `showPlaceholder` - show placeholder (optional)
- `BorderColorValid` - border color for correct input (optional)
- `BorderColorInvalid` - border color for incorrect input (optional)
- `PlaceholderEmptyColor` - placeholder color for empty field (optional)
- `PlaceholderEmpty` - placeholder text for empty field (optional)
- `PlaceholderPartialColor` - placeholder color for partially filled field (optional)
- `PlaceholderPartial` - placeholder text for partially filled field (optional)
- `PlaceholderCompleteColor` - placeholder color for completely filled field (optional)
- `PlaceholderComplete` - placeholder text for completely filled field (optional)
- `PlaceholderInvalidColor` - placeholder color for field with invalid data (optional)
- `PlaceholderInvalid` - placeholder text for field with invalid data (optional)
- `PlaceHolderTemplate` - placeholder template (optional)

### IsValid
Checks the correctness of the entered data in the textbox field.

**Syntax:**
```vba
Public Function IsValid() As Boolean
```

**Return Value:**
- `Boolean` - True if data is correct, otherwise False

### Clear
Clears the textbox field.

**Syntax:**
```vba
Public Sub Clear()
```

### SetFocus
Sets focus on the textbox field.

**Syntax:**
```vba
Public Sub SetFocus()
```

### RemoveItem
Removes the textbox mask element and related components.

**Syntax:**
```vba
Public Sub RemoveItem()
```

## Events

### mTextBox_Change
Event raised when the textbox value changes.

### mTextBox_KeyPress
Event raised when a key is pressed in the textbox.

## Constants and Enumerations

### enumTypeMask
```vba
Public Enum enumTypeMask
    tOtherFix = 1
    tDateFix
    tTimeFix
    tNumeric
    tVariableLen
    tRegex
    [_First] = tOtherFix
    [_Last] = tRegex
End Enum
```

## Implementation Details

### Mask Symbols
When creating text masks, the following symbols are used:

| Symbol | Description |
|--------|-------------|
| `#` | Digits (0-9) |
| `@` | Latin letters (A-Z, a-z) |
| `A` | Latin letters and digits (A-Z, a-z, 0-9) |
| `Б` | Cyrillic letters |
| `б` | Cyrillic letters and digits |
| `*` | Any characters |

### Placeholder Templates
The class supports the use of markers in placeholder templates:
- `{mask}` - displays the mask
- `{filled}` - number of filled characters
- `{remaining}` - number of remaining characters
- `{holder}` - placeholder with mask
- `{RegexPattern}` - regular expression pattern
- `{RegexFilter}` - regular expression filter
- `{percent}` - fill percentage

### Event Handling
The class automatically handles textbox events:
- `Change` - updates placeholder and checks validity
- `KeyPress` - controls input characters according to the mask

## Dependencies

- MSForms.TextBox
- MSForms.Label
- VBScript.RegExp (for regular expression validation)