# UserForms-Class-TextBox-Masks

## Description
This project contains the `clsTextboxMask` class for Microsoft VBA, which allows creating textboxes with input masks in UserForms. The class provides input validation, placeholder display, and visual indicators of the field filling status.

## Features
- Support for various input mask types:
  - Numeric masks (with range, sign, and decimal options)
  - Date masks (with date validation and range checking)
  - Time masks (with time validation)
  - Fixed-length text masks (with various character types)
- Visual validation indicators (border color changes based on input validity)
- Placeholder hint display with expected format
- Support for various character types in masks:
  - `#` - digits
  - `@` - Latin letters
  - `*` - any characters
  - `A` - Latin letters and digits
  - `А` - Cyrillic letters
  - `а` - Cyrillic letters and digits

## Installation
1. Copy the `clsTextboxMask.cls` file to your VBA project
2. Import it into the VBA editor (e.g., in Excel or Word)

## Usage
The class provides several methods for adding different mask types:

### Numeric Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox1, 0, 100, True, False)
```

### Date Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox1, "##.##.####", #1/1/2000#, #12/31/2030#, "dd.mm.yyyy")
```

### Time Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox1, "##:##", #0:00:00#, #23:59:59#, "hh:mm")
```

### Fixed-Length Text Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox1, "###@@@")
```

## Parameters
- `TextBox` - textbox object to apply the mask to
- `Mask` - input mask string
- `Min/Max` - minimum and maximum allowed values (for numeric and date masks)
- `IsDecimal` - allow decimal input
- `IsNegative` - allow negative input
- `formatValue` - value display format (for dates and numbers)
- `visibleLabelHolder` - visibility of placeholder hint
- `borderColorOn/borderColorOff` - border colors for valid and invalid input

## Properties
- `Value` - current textbox value
- `Mask` - input mask
- `LenValue` - length of current value
- `LenMask` - length of mask
- `RemainingChars` - number of remaining characters until full fill
- `IsValid` - input validation check

## Author
VBATools

## Version
1.0.3

## License
Apache License