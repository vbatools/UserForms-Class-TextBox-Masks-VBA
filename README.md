# VBA TextBox Masks Class

![Project Demo](User_Forms.png)

**clsTextboxMask** is a powerful VBA class that allows creating textboxes with input masks in Excel and other Office applications. It provides input validation, placeholder display, and visual indication of field fill status.

## Table of Contents
1. [Features](#features)
2. [Components](#components)
3. [Installation](#installation)
4. [Quick Start](#quick-start)
5. [Main Functions](#main-functions)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Troubleshooting](#troubleshooting)

## Features

- Support for various input mask types (numbers, dates, time, text, regular expressions)
- Real-time input validation
- Display of placeholders with different statuses (empty, partially filled, completely filled, invalid)
- Visual indication of input correctness through border color
- Support for numeric values with range, sign and decimal restrictions
- Support for variable length text
- Support for validation via regular expressions
- Support for placeholder color customization based on field status
- Support for placeholder templates with markers: {mask}, {filled}, {remaining}, {holder}, {RegexPattern}, {RegexFilter}, {percent}

## Components

- `clsTextboxMask.cls`: The main textbox mask class implementation
- `frmTestClass.frm`: Test form demonstrating usage
- `modShowForms.bas`: Module containing form display functions
- Documentation in the `docs/` folder:
 - [`docs/technical_documentation_eng.md`](docs/technical_documentation_eng.md) - Technical documentation in English
 - [`docs/technical_documentation_rus.md`](docs/technical_documentation_rus.md) - Technical documentation in Russian
 - [`docs/user_guide_eng.md`](docs/user_guide_eng.md) - User guide in English
 - [`docs/user_guide_rus.md`](docs/user_guide_rus.md) - User guide in Russian
  - [`docs/implementation_examples_eng.md`](docs/implementation_examples_eng.md) - Implementation examples in English
 - [`docs/implementation_examples_rus.md`](docs/implementation_examples_rus.md) - Implementation examples in Russian

## Installation

1. Copy the `clsTextboxMask.cls` file to your VBA project
2. Use the class in your UserForms

## Quick Start

### Simple Usage Example
```vba
' Create an instance of clsTextboxMask class
Dim numField As clsTextboxMask
Set numField = New clsTextboxMask

' Add a numeric field with restrictions
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True, _
                             allowNegative:=False)

' The class automatically applies input mask to the control
```

## Main Functions

- **Mask Initialization**: Methods `AddFieldNumeric`, `AddFieldDate`, `AddFieldTime`, `AddFieldText`, `AddFieldVariableLength`, `AddFieldRegex` allow applying appropriate masks to textboxes
- **Input Validation**: Automatic validation of input correctness in real-time
- **Placeholder Display**: Dynamic placeholder text changes depending on field state
- **Color Indication**: Visual indication of input correctness through border color changes
- **Various Format Support**: Support for numeric, text, date, time and regular expression formats

## Working with Controls

The `clsTextboxMask` class adds input mask functionality to TextBox controls with capabilities:
- Setting minimum and maximum values for numeric fields
- Defining date and time formats
- Applying text masks with various characters
- Using regular expressions for validation
- Configuring border and placeholder colors

## Style Configuration

The class allows customization of:
- Border colors for correct and incorrect input
- Placeholder text colors for various states
- Display formats for numbers, dates and time
- Placeholder templates using markers

Example of color configuration:
```vba
' Configure colors when adding numeric field
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True, _
                             allowNegative:=False, _
                             BorderColorValid:=RGB(0, 150, 0), _
                             BorderColorInvalid:=RGB(200, 0, 0), _
                             PlaceholderColor:=RGB(150, 150, 150))
```

## Troubleshooting

### Display Issues
- Ensure Microsoft Forms 2.0 Object Library is enabled in references
- Check that controls are added before calling class methods
- Ensure the MultiUse property is set to True for the class

### Interaction Issues
- Check that control events are not overloaded with other handlers
- Ensure control properties are not changed manually while the class is running
- Verify that the class is not initialized multiple times

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.