VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextboxMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================
' Class:         clsTextboxMask
' Author:        VBATools
' Version:       1.0.4
' Creation Date: 05.11.2025 10:30
' Update Date:   08.11.2025 12:00
' License:       Apache License
' Description:   Class for creating textboxes with input masks in VBA.
'              Supports various mask types: numeric, letters, dates,
'              phone numbers, etc. Provides input validation, placeholder display, and
'              visual indication of field fill status.
'
'              The class allows creating validated text fields with specified
'              constraints on input data. Supports the following mask types:
'              - Numeric (with range, sign, and decimal options)
'              - Dates (with date validity and range checks)
'              - Time (with time validity checks)
'              - Fixed-length text (with various character types)
'              - Variable-length text (with optional mask pattern)
'              - Regular expression-based masks (custom validation patterns)
'
'              Each masked field has visual indicators: border color
'              changes depending on the validity of input data, and
'              a placeholder hint with the expected format is displayed.
'
' Examples of usage:
'   Numeric field: obj.addNumericField(Me.TextBox1, 0, 100, True, False) ' 0 to 100, with decimal, without negative
'   Date field: obj.addDateField(Me.TextBox1, "##.##.####", #1/1/200#, #1/1/2030#, "dd.mm.yyyy") ' Date from 01.01.2000 to 01.01.2030
'   Time field: obj.addTimeField(Me.TextBox1, "##:##", #0:00#, #23:59#, "hh:mm") ' Time from 00:00 to 23:59
'   Text field: obj.addTextField(Me.TextBox1, "###@@@") ' 3 digits followed by 3 letters
' ======================================================================================
' Enum:          enumTypeMask
' Description:   Enumeration of mask types supported by the class
' Values:
'   tOtherFix - Fixed-length text mask (letters, numbers, special characters)
'   tDateFix  - Fixed-length date input mask
'   tTimeFix  - Fixed-length time input mask
'   tNumeric  - Numeric mask with range, sign and decimal options
'   tVariableLen - Variable-length text mask with optional pattern
'   tRegex    - Custom mask via regular expressions
'   [_First]  - Helper value to determine the start of the enumeration
'   [_Last]   - Helper value to determine the end of the enumeration
Public Enum enumTypeMask
    tOtherFix = 1
    tDateFix
    tTimeFix
    tNumeric
    tVariableLen  ' Variable length mask
    tRegex        ' Custom mask via regular expressions
    [_First] = tOtherFix
    [_Last] = tRegex
End Enum

Private Const SIMVOLS_MASKS As String = "#*@AБб"
Private Const COLOR_DEFAULT_ON As Long = &H80000006
Private Const COLOR_DEFAULT_OFF As Long = &HC0C0FF
Private Const COLOR_DEFAULT_HOLDER As Long = &H808080

Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private mLabelHolder As MSForms.Label

Private mItems      As Collection

Private mMask       As String
Private mFormatValue As String
Private mRegexPattern As String    ' Regular expression pattern for tRegex masks

Private mBorderColorValid As Long
Private mBorderColorNoValid As Long

Private mIsDecemal  As Boolean
Private mIsNegative As Boolean

Private mMin        As Single
Private mMax        As Single

Private mCurrentMaskType As enumTypeMask

' Property:      constSimvolsMasks
' Description:   Returns a string with mask characters used in the class
' Returns:       String "#*@A" containing mask characters:
'                # - digits, @ - Latin letters, * - any characters,
'                A - Latin letters and digits, Б - Cyrillic letters,
'                б - Cyrillic letters and digits
Public Property Get constSimvolsMasks() As String
    constSimvolsMasks = SIMVOLS_MASKS
End Property

' Property:      Min
' Description:   Property to get and set the minimum value
'                for numeric masks. Used during input validation.
' Returns/Parameters:
'   Value - Value of the minimum allowed number for numeric masks
Public Property Get Min() As Single
    Min = mMin
End Property

Public Property Let Min(ByVal Value As Single)
    mMin = Value
End Property

' Property:      Max
' Description:   Property to get and set the maximum value
'                for numeric masks. Used during input validation.
' Returns/Parameters:
'   Value - Value of the maximum allowed number for numeric masks
Public Property Get Max() As Single
    Max = mMax
End Property

Public Property Let Max(ByVal Value As Single)
    mMax = Value
End Property

' Property:      visibleLabelHolder
' Description:   Property to get and set the visibility of the placeholder hint
'                for the text field. The hint displays the input mask.
' Returns/Parameters:
'   Value - Boolean value (True/False) determining the visibility of the hint
Public Property Get visibleLabelHolder() As Boolean
    If Me.LabelHolder Is Nothing Then Exit Property
    visibleLabelHolder = Me.LabelHolder.Visible
End Property

Public Property Let visibleLabelHolder(ByVal Value As Boolean)
    If Me.LabelHolder Is Nothing Then Exit Property
    Me.LabelHolder.Visible = Value
End Property

' Property:      IsNegative
' Description:   Property to get and set permission for input of negative
'                values for numeric masks.
' Returns/Parameters:
'   Value - Boolean value (True/False) allowing input of negative numbers
Public Property Get IsNegative() As Boolean
    IsNegative = mIsNegative
End Property

Public Property Let IsNegative(ByVal Value As Boolean)
    mIsNegative = Value
End Property

' Property:      IsDecemal
' Description:   Property to get and set permission for input of decimal
'                (fractional) values for numeric masks.
' Returns/Parameters:
'   Value - Boolean value (True/False) allowing input of fractional numbers
Public Property Get IsDecemal() As Boolean
    IsDecemal = mIsDecemal
End Property

Public Property Let IsDecemal(ByVal Value As Boolean)
    mIsDecemal = Value
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = mTextBox
End Property

Public Property Set TextBox(ByRef TextBox As MSForms.TextBox)
    Set mTextBox = TextBox
End Property

Public Property Get LabelHolder() As MSForms.Label
    Set LabelHolder = mLabelHolder
End Property

Public Property Set LabelHolder(ByRef Label As MSForms.Label)
    Set mLabelHolder = Label
End Property

' Property:      InputControl
' Description:   Alternative property to get and set the input control (TextBox)
' Returns/Parameters:
'   Control - MSForms.TextBox object that serves as the input control
Public Property Get inputControl() As MSForms.TextBox
    Set inputControl = mTextBox
End Property

Public Property Set inputControl(ByRef Control As MSForms.TextBox)
    Set mTextBox = Control
End Property

Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Property Set Items(ByRef Items As Collection)
    Set mItems = Items
End Property

' Property:      PlaceHolder
' Description:   Property to get the placeholder string that is displayed
'                in the hint until the text field is completely filled
' Returns:       String containing the current field value plus the remaining part of the mask
Public Property Get PlaceHolder() As String
    If Me.TextBox Is Nothing Then Exit Property
    If Me.LenMask > Me.LenValue Then PlaceHolder = Me.Value & VBA.Mid$(Me.Mask, Me.LenValue + 1, Me.LenMask - Me.LenValue)
End Property

' Property:      CurrentMaskType
' Description:   Property to get and set the current input mask type
'                for this class instance
' Returns/Parameters:
'   maskType - Value from the enumTypeMask enumeration defining the mask type
Public Property Get CurrentMaskType() As enumTypeMask
    CurrentMaskType = mCurrentMaskType
End Property

Public Property Let CurrentMaskType(ByVal maskType As enumTypeMask)
    mCurrentMaskType = maskType
End Property

' Property:      LenValue
' Description:   Property to get the current length of the entered value
'                in the text field
' Returns:       Integer representing the number of characters in the current value
Public Property Get LenValue() As Integer
    If Me.TextBox Is Nothing Then Exit Property
    LenValue = VBA.Len(Me.Value)
End Property

' Property:      Value
' Description:   Property to get and set the value of the text field
' Returns/Parameters:
'   Value - String representing the value of the text field
Public Property Let Value(ByRef Value As String)
    If Me.TextBox Is Nothing Then Exit Property
    Me.TextBox.Value = Value
End Property

Public Property Get Value() As String
    If Me.TextBox Is Nothing Then Exit Property
    Value = Me.TextBox.Value
End Property

' Property:      Mask
' Description:   Property to get and set the input mask for the text field
' Returns/Parameters:
'   Mask - String defining the input mask using special characters:
'          # (digits), @ (Latin letters), * (any characters), A (Latin letters and digits),
'          Б (Cyrillic letters), б (Cyrillic letters and digits)
Public Property Let Mask(ByRef Mask As String)
    mMask = Mask
End Property

Public Property Get Mask() As String
    Mask = mMask
End Property

' Property:      InputMask
' Description:   Alternative property to get and set the input mask for the text field
' Returns/Parameters:
'   Mask - String defining the input mask using special characters:
'          # (digits), @ (Latin letters), * (any characters), A (Latin letters and digits),
'          Б (Cyrillic letters), б (Cyrillic letters and digits)
Public Property Let InputMask(ByRef Mask As String)
    mMask = Mask
End Property

Public Property Get InputMask() As String
    InputMask = mMask
End Property

' Property:      LenMask
' Description:   Property to get the length of the input mask
' Returns:       Integer representing the number of characters in the mask
Public Property Get LenMask() As Integer
    LenMask = VBA.Len(Me.Mask)
End Property

' Property:      formatValue
' Description:   Property to get and set the format for displaying the value
'                after input (used for dates and numbers)
' Returns/Parameters:
'   formatValue - Format string defining how the value will be displayed
Public Property Let formatValue(ByRef formatValue As String)
    mFormatValue = formatValue
End Property

Public Property Get formatValue() As String
    formatValue = mFormatValue
End Property

' Property:      RegexPattern
' Description:   Property to get and set the regular expression pattern
'                for regex-based masks
' Returns/Parameters:
'   RegexPattern - String defining the regular expression pattern
Public Property Let regexPattern(ByRef regexPattern As String)
    mRegexPattern = regexPattern
End Property

Public Property Get regexPattern() As String
    regexPattern = mRegexPattern
End Property

' Property:      borderColorValid
' Description:   Property to get and set the border color of the text field
'                when input is valid
' Returns/Parameters:
'   Color - Color value (Long) used for the border when input is valid
Public Property Let borderColorValid(ByRef Color As Long)
    mBorderColorValid = Color
End Property

Public Property Get borderColorValid() As Long
    borderColorValid = mBorderColorValid
End Property

' Property:      borderColorNoValid
' Description:   Property to get and set the border color of the text field
'                when input is invalid
' Returns/Parameters:
'   Color - Color value (Long) used for the border when input is invalid
Public Property Let borderColorNoValid(ByRef Color As Long)
    mBorderColorNoValid = Color
End Property

Public Property Get borderColorNoValid() As Long
    borderColorNoValid = mBorderColorNoValid
End Property

' Property:      ValidColor
' Description:   Alternative property to get and set the border color when input is valid
' Returns/Parameters:
'   Color - Color value (Long) used for the border when input is valid
Public Property Let validColor(ByRef Color As Long)
    mBorderColorValid = Color
End Property

Public Property Get validColor() As Long
    validColor = mBorderColorValid
End Property

' Property:      InvalidColor
' Description:   Alternative property to get and set the border color when input is invalid
' Returns/Parameters:
'   Color - Color value (Long) used for the border when input is invalid
Public Property Let invalidColor(ByRef Color As Long)
    mBorderColorNoValid = Color
End Property

Public Property Get invalidColor() As Long
    invalidColor = mBorderColorNoValid
End Property

' Property:      Count
' Description:   Property to get the number of elements in the collection
'                of input masks
' Returns:       Byte representing the number of elements in the Items collection
Public Property Get Count() As Byte
    Count = Me.Items.Count
End Property

' Property:      RemainingChars
' Description:   Property to get the number of remaining characters
'                until the text field is completely filled according to the mask
' Returns:       Integer representing the number of remaining characters
Public Property Get RemainingChars() As Integer
    If Me.TextBox Is Nothing Then
        RemainingChars = Me.LenMask
        Exit Property
    End If
    RemainingChars = Me.LenMask - Me.LenValue
End Property

' Property:      getItemByIndex
' Description:   Property to get an item from the input masks collection by index
' Parameters:
'   index - Integer representing the index of the item in the collection (starting from 1)
' Returns:       clsTextboxMask object corresponding to the specified index,
'                or Nothing if an item with that index does not exist
Public Property Get getItemByIndex(ByVal index As Integer) As clsTextboxMask
    On Error GoTo endGetItem
    If Not Me.Items Is Nothing Then Set getItemByIndex = Me.Items(index)
    Exit Property
endGetItem:
    ' Don't clear the error, but log it for diagnostics
    Select Case Err.Number
        Case 0, 5
        Case Else
            Debug.Print "Error in getItemByName: " & Err.Description & " (Number: " & Err.Number & ")"
            Set getItemByIndex = Nothing
    End Select
End Property

' Property:      getItemByName
' Description:   Property to get an item from the input masks collection by name
' Parameters:
'   Name - String representing the name of the item in the collection
' Returns:       clsTextboxMask object corresponding to the specified name,
'                or Nothing if an item with that name does not exist
Public Property Get getItemByName(ByVal Name As String) As clsTextboxMask
    On Error GoTo endGetItem
    If Not Me.Items Is Nothing Then Set getItemByName = Me.Items(Name)
    Exit Property
endGetItem:
    ' Don't clear the error, but log it for diagnostics
    Select Case Err.Number
        Case 0, 5
        Case Else
            Debug.Print "Error in getItemByName: " & Err.Description & " (Number: " & Err.Number & ")"
            Set getItemByName = Nothing
    End Select
End Property

' Property: Version
' Purpose: Gets version information about the class
Public Property Get Version() As String
    Version = "Class: clsTextboxMask" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.4" & vbNewLine & _
            "Date of creation: 05.11.2025 10:30" & vbNewLine & _
            "Date of update: 08.11.2025 12:00" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: Class for creating textboxes with input masks in VBA" & vbNewLine & _
            "Supports various mask types: digital, letters, dates, phone numbers, etc." & vbNewLine & _
            "Provides input validation, placeholder display, and visual indication" & vbNewLine & _
            "of field fill status."
End Property

' Method:        addNumericField
' Description:   Alternative method to add a new item with numeric mask type to the collection
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   minValue        - Minimum allowed value
'   maxValue        - Maximum allowed value
'   allowDecimal    - Flag allowing input of decimal numbers
'   allowNegative   - Flag allowing input of negative numbers
'   showPlaceholder - (Optional) Flag for placeholder hint visibility (default True)
'   numberFormat    - (Optional) Number display format (default "#.0")
'   validColor      - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   invalidColor    - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   placeholderColor - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
' Example:       obj.addNumericField(Me.TextBox1, 0, 100, True, False) ' Numeric field from 0 to 100, with decimal, without negative
Public Sub addNumericField(ByRef inputControl As MSForms.TextBox, _
        ByVal minValue As Single, _
        ByVal maxValue As Single, _
        ByVal allowDecimal As Boolean, _
        ByVal allowNegative As Boolean, _
        Optional showPlaceholder As Boolean = True, _
        Optional numberFormat As String = "#.0", _
        Optional validColor As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional invalidColor As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional placeholderColor As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addField(inputControl, String$(WorksheetFunction.Max(VBA.Len(minValue), VBA.Len(maxValue)), "#"), _
            enumTypeMask.tNumeric, numberFormat, allowDecimal, allowNegative, validColor, invalidColor, _
            minValue, maxValue, showPlaceholder, placeholderColor, vbNullString)
End Sub

' Method:        addDateField
' Description:   Alternative method to add a new item with fixed-length date mask to the collection
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   dateMask        - Mask string for date (e.g., "##.##.####")
'   minDate         - Minimum allowed date
'   maxDate         - Maximum allowed date
'   dateFormat      - (Optional) Date display format (default "dd.mm.yyyy")
'   showPlaceholder - (Optional) Flag for placeholder hint visibility (default True)
'   validColor      - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   invalidColor    - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   placeholderColor - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
' Example:       obj.addDateField(Me.TextBox1, "##.##.####", #1/1/2000#, #1/1/2030#, "dd.mm.yyyy") ' Date from 01.01.2000 to 01.01.2030
Public Sub addDateField(ByRef inputControl As MSForms.TextBox, ByVal dateMask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional dateFormat As String = "dd.mm.yyyy", _
        Optional showPlaceholder As Boolean = True, _
        Optional validColor As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional invalidColor As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional placeholderColor As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addField(inputControl, dateMask, enumTypeMask.tDateFix, dateFormat, False, False, validColor, invalidColor, _
            minDate, maxDate, showPlaceholder, placeholderColor, vbNullString)
End Sub

' Method:        addTimeField
' Description:   Alternative method to add a new item with fixed-length time mask to the collection
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   timeMask        - Mask string for time (e.g., "##:##")
'   minTime         - Minimum allowed time
'   maxTime         - Maximum allowed time
'   timeFormat      - (Optional) Time display format (default "hh:mm")
'   showPlaceholder - (Optional) Flag for placeholder hint visibility (default True)
'   validColor      - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   invalidColor    - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   placeholderColor - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
' Example:       obj.addTimeField(Me.TextBox1, "##:##", #0:00#, #23:59#, "hh:mm") ' Time from 00:00 to 23:59
Public Sub addTimeField(ByRef inputControl As MSForms.TextBox, ByVal timeMask As String, _
        ByVal minTime As Date, _
        ByVal maxTime As Date, _
        Optional timeFormat As String = "hh:mm", _
        Optional showPlaceholder As Boolean = True, _
        Optional validColor As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional invalidColor As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional placeholderColor As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addField(inputControl, timeMask, enumTypeMask.tTimeFix, timeFormat, False, False, validColor, invalidColor, _
            0, 1, showPlaceholder, placeholderColor, vbNullString)
End Sub

' Method:        addTextField
' Description:   Alternative method to add a new item with fixed-length text mask to the collection
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   textMask        - Mask string for text (e.g., "###@@@")
'   showPlaceholder - (Optional) Flag for placeholder hint visibility (default True)
'   validColor      - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   invalidColor    - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   placeholderColor - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
' Example:       obj.addTextField(Me.TextBox1, "###@@@") ' 3 digits followed by 3 letters
Public Sub addTextField(ByRef inputControl As MSForms.TextBox, ByVal textMask As String, _
        Optional showPlaceholder As Boolean = True, _
        Optional validColor As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional invalidColor As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional placeholderColor As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addField(inputControl, textMask, enumTypeMask.tOtherFix, vbNullString, False, False, validColor, invalidColor, _
            0, 0, showPlaceholder, placeholderColor, vbNullString)
End Sub

' Method:        addVariableLengthField
' Description:   Method to add a new item with variable-length text mask to the collection
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   maxLength       - Maximum allowed length for the input (0 = unlimited)
'   textMask        - Mask string for text (e.g., "###@@@") - optional pattern for input validation
'   showPlaceholder - (Optional) Flag for placeholder hint visibility (default True)
'   validColor      - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   invalidColor    - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   placeholderColor - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
' Example:       obj.addVariableLengthField(Me.TextBox1, 20, "###@@@") ' Up to 20 characters with mask pattern
Public Sub addVariableLengthField(ByRef inputControl As MSForms.TextBox, _
        ByVal maxLength As Integer, _
        Optional textMask As String = vbNullString, _
        Optional showPlaceholder As Boolean = True, _
        Optional validColor As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional invalidColor As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional placeholderColor As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Dim maskStr     As String
    If textMask <> vbNullString Then
        If VBA.Len(textMask) < maxLength Then
            maskStr = textMask & VBA.String$(maxLength - VBA.Len(textMask), "*")
        Else
            maskStr = textMask
        End If
    Else
        ' If no mask is specified, use arbitrary text up to maxLength
        maskStr = VBA.String$(maxLength, "*")
    End If

    Call addField(inputControl, maskStr, enumTypeMask.tVariableLen, vbNullString, False, False, validColor, invalidColor, _
            0, maxLength, showPlaceholder, placeholderColor, vbNullString)
End Sub

' Method:        addRegexField
' Description:   Method to add a new item with regex-based mask to the collection
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   regexPattern    - Regular expression pattern for validation
'   showPlaceholder - (Optional) Flag for placeholder hint visibility (default True)
'   validColor      - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   invalidColor    - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   placeholderColor - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
' Example:       obj.addRegexField(Me.TextBox1, "^[A-Z]{3}\d{3}$") ' 3 uppercase letters followed by 3 digits
Public Sub addRegexField(ByRef inputControl As MSForms.TextBox, _
        ByVal regexPattern As String, _
        Optional showPlaceholder As Boolean = True, _
        Optional validColor As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional invalidColor As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional placeholderColor As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addField(inputControl, vbNullString, enumTypeMask.tRegex, vbNullString, False, False, validColor, invalidColor, _
            0, 0, showPlaceholder, placeholderColor, regexPattern)
End Sub

' Method:        addField
' Description:   Internal method for adding a new input mask item to the collection
'                with full parameter configuration. Used by helper methods
'                addNumericField, addDateField, addTimeField and addTextField.
' Parameters:
'   inputControl    - Textbox object to apply the mask to
'   textMask        - Mask string for the text field
'   maskType        - Mask type from the enumTypeMask enumeration
'   formatValue     - Value display format (used for dates and numbers)
'   allowDecimal    - Flag allowing input of decimal numbers (for numeric masks)
'   allowNegative   - Flag allowing input of negative numbers (for numeric masks)
'   validColor      - Border color when input is valid
'   invalidColor    - Border color when input is invalid
'   minValue        - Minimum allowed value (for numeric and date masks)
'   maxValue        - Maximum allowed value (for numeric and date masks)
'   showPlaceholder - Flag for placeholder hint visibility
'   placeholderColor - Placeholder hint text color
Private Sub addField(ByRef inputControl As MSForms.TextBox, _
        ByVal textMask As String, _
        ByVal maskType As enumTypeMask, _
        ByVal formatValue As String, _
        ByVal allowDecimal As Boolean, _
        ByVal allowNegative As Boolean, _
        ByVal validColor As XlRgbColor, _
        ByVal invalidColor As XlRgbColor, _
        ByVal minValue As Single, _
        ByVal maxValue As Single, _
        ByVal showPlaceholder As Boolean, _
        ByVal placeholderColor As XlRgbColor, _
        ByVal regexPattern As String)

    If inputControl Is Nothing Then
        Err.Raise vbObjectError + 101, "clsTextboxMask", "TextBox cannot be Nothing"
        Exit Sub
    End If

    If maskType < enumTypeMask.[_First] Or maskType > enumTypeMask.[_Last] Then
        Err.Raise vbObjectError + 102, "clsTextboxMask", "Invalid mask type"
        Exit Sub
    End If

    If IsControlInCollection(inputControl.Name) Then
        Err.Raise vbObjectError + 103, "clsTextboxMask", "The item has already been created"
        Exit Sub
    End If

    '    If VBA.Len(Mask) <> VBA.Len(formatValue) And formatValue <> vbNullString Then
    '        Err.Raise vbObjectError + 104, "Length of the mask is not equal to the length of the format"
    '        Exit Sub
    '    End If

    Dim itemCls     As clsTextboxMask
    Set itemCls = New clsTextboxMask

    With itemCls
        Set .TextBox = inputControl
        .Mask = textMask
        If .LenMask > 0 Then .TextBox.maxLength = .LenMask

        .borderColorValid = validColor
        .borderColorNoValid = invalidColor

        .IsDecemal = allowDecimal
        .IsNegative = allowNegative

        .Max = maxValue
        .Min = minValue

        Dim labelName As String
        labelName = inputControl.Name & "_holder"

        ' Check if a label with this name already exists
        On Error Resume Next
        Dim existingLabel As MSForms.Control
        Set existingLabel = .TextBox.Parent.Controls(labelName)
        On Error GoTo 0

        If Not existingLabel Is Nothing Then
            .TextBox.Parent.Controls.Remove existingLabel.Name
        End If

        Set .LabelHolder = .TextBox.Parent.Controls.Add("Forms.Label.1", labelName, True)

        With .LabelHolder
            .Left = itemCls.TextBox.Left + 8
            .Width = itemCls.TextBox.Width
            .Height = itemCls.TextBox.Height * 0.6
            .Top = itemCls.TextBox.Top - .Height
            .Caption = itemCls.Mask
            .BackStyle = fmBackStyleTransparent
            '.BorderStyle = fmBorderStyleSingle
            .ForeColor = placeholderColor
            .Font.Name = itemCls.TextBox.Font.Name
            .Font.Size = itemCls.TextBox.Font.Size * 0.9
        End With

        .visibleLabelHolder = showPlaceholder
        .CurrentMaskType = maskType
        .formatValue = formatValue
        ' Set the regular expression for this item
        .regexPattern = regexPattern
        If formatValue = vbNullString Then .formatValue = maskType
        If mItems Is Nothing Then Set mItems = New Collection
        Call mItems.Add(itemCls, inputControl.Name)
        Call .IsValid
        Set .Items = mItems
    End With
End Sub

' Method:        IsValid
' Description:   Checks the validity of the entered value in the text field
'                according to the set mask
' Returns:       Boolean value (True/False) indicating input validity
Public Function IsValid() As Boolean
    IsValid = IsValidInput()
End Function

' Method:        Clear
' Description:   Clears the value of the text field
Public Sub Clear()
    If Me.TextBox Is Nothing Then Exit Sub
    Me.Value = vbNullString
End Sub

' Method:        SetFocus
' Description:   Sets focus to the text field
Public Sub SetFocus()
    If Me.TextBox Is Nothing Then Exit Sub
    Me.TextBox.SetFocus
End Sub

' Method:        RemoveItem
' Description:   Removes the input mask item from the collection and clears associated controls
'                (placeholder hint and text field)
Public Sub RemoveItem()
    On Error Resume Next
    If IsControlInCollection(Me.TextBox.Name) Then
        If Not Me.LabelHolder Is Nothing Then Call Me.TextBox.Parent.Controls.Remove(Me.LabelHolder.Name)
        Call Me.Items.Remove(Me.TextBox.Name)
    End If
    On Error GoTo 0

    Set Me.LabelHolder = Nothing
    Set Me.TextBox = Nothing
    Set Me.Items = Nothing
End Sub

Private Sub mTextBox_change()
    If Me.LabelHolder Is Nothing Then Exit Sub

    If Me.LenMask = VBA.Len(Me.Value) Then
        Me.LabelHolder.Caption = vbNullString
    Else
        Me.LabelHolder.Caption = Me.PlaceHolder
    End If
    Call IsValidInput
End Sub

' Method:        mTextBox_KeyPress
' Description:   Event handler for key press events in the masked textbox
'                Applies appropriate input validation based on the current mask type
' Parameters:
'   KeyAscii - ASCII code of the pressed key, modified in-place to allow/reject input
Private Sub mTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Me.CurrentMaskType
        Case enumTypeMask.tDateFix, enumTypeMask.tTimeFix, enumTypeMask.tOtherFix:
            Call KeyAsciiFixLenText(KeyAscii)
        Case enumTypeMask.tNumeric
            Call NumericAll(KeyAscii)
        Case enumTypeMask.tVariableLen
            ' For variable length masks, apply mask validation if mask is specified
            If Me.Mask <> vbNullString Then
                Call KeyAsciiFixLenText(KeyAscii)
            End If
        Case enumTypeMask.tRegex
            Call KeyAsciiRegex(KeyAscii)
        Case Else
            ' Default case - allow all characters
    End Select
End Sub

Private Sub KeyAsciiFixLenText(ByRef KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If Me.TextBox Is Nothing Or Me.LenMask = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    Dim currentLength As Integer
    currentLength = Me.LenValue
    Dim i           As Integer
    i = currentLength + 1
    If i > Me.LenMask Then
        KeyAscii = 0
        Exit Sub
    End If

    Dim endLetter   As String
    Do
        endLetter = VBA.Mid$(Me.Mask, i, 1)
        Select Case endLetter
            Case "#"
                Call NumericAll(KeyAscii)
            Case "@"
                Call LatinLetters(KeyAscii)
            Case "A"
                Call AlphaNumeric(KeyAscii)
            Case "Б"
                Call UniCodeLetters(KeyAscii)
            Case "б"
                Call UniCodeLettersNumeric(KeyAscii)
            Case "*"
                ' Allow any character
            Case Else:
                Me.Value = Me.Value & endLetter
        End Select
        i = i + 1
        If i > Me.LenMask Then Exit Do
    Loop While Not haveSimvol(endLetter)

    Exit Sub
ErrorHandler:
    KeyAscii = 0
End Sub

Private Function haveSimvol(ByVal sVal As String) As Boolean
    haveSimvol = (InStr(1, SIMVOLS_MASKS, sVal, vbBinaryCompare) > 0)
End Function

Private Sub NumericAll(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57:    ' ASCII codes for digits 0-9
        Case 45:
            If Me.IsNegative Then
                KeyAscii = IIf(Me.Value <> vbNullString, 0, 45)
            Else
                KeyAscii = 0
            End If
        Case 44:
            If Me.IsDecemal Then
                KeyAscii = IIf(InStr(1, Me.Value, ",") > 0, 0, 44)
            Else
                KeyAscii = 0
            End If
        Case 46:
            If Me.IsDecemal Then
                KeyAscii = IIf(InStr(1, Me.Value, ".") > 0, 0, 46)
            Else
                KeyAscii = 0
            End If
        Case 8:    ' ASCII code for Backspace key - allow
            ' Do nothing, character is allowed
        Case Else:
            KeyAscii = 0
    End Select
End Sub

Private Function IsValidLatinLetter(ByVal KeyAscii As Integer) As Boolean
    IsValidLatinLetter = (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)
End Function

Private Function IsValidCyrillicLetter(ByVal KeyAscii As Integer) As Boolean
    IsValidCyrillicLetter = (KeyAscii >= 192 And KeyAscii <= 255) Or (KeyAscii >= 1040 And KeyAscii <= 1103)
End Function

Private Function IsValidDigit(ByVal KeyAscii As Integer) As Boolean
    IsValidDigit = (KeyAscii >= 48 And KeyAscii <= 57)
End Function

Private Sub LatinLetters(ByRef KeyAscii As MSForms.ReturnInteger)
    If IsValidLatinLetter(KeyAscii) Then
        ' This is a Latin letter, leave as is
    ElseIf KeyAscii = 8 Then
        ' Backspace key, allow
    Else
        ' Does not meet requirements, reject
        KeyAscii = 0
    End If
End Sub

Private Sub UniCodeLetters(ByRef KeyAscii As MSForms.ReturnInteger)
    If IsValidCyrillicLetter(KeyAscii) Then
        ' This is a Cyrillic letter, leave as is
    ElseIf KeyAscii = 8 Then
        ' Backspace key, allow
    Else
        ' Does not meet requirements, reject
        KeyAscii = 0
    End If
End Sub

Private Sub AlphaNumeric(ByRef KeyAscii As MSForms.ReturnInteger)
    If IsValidDigit(KeyAscii) Or IsValidLatinLetter(KeyAscii) Then
        ' This is a digit or Latin letter, leave as is
    ElseIf KeyAscii = 8 Then
        ' Backspace key, allow
    Else
        ' Does not meet requirements, reject
        KeyAscii = 0
    End If
End Sub

Private Sub UniCodeLettersNumeric(ByRef KeyAscii As MSForms.ReturnInteger)
    If IsValidDigit(KeyAscii) Or IsValidCyrillicLetter(KeyAscii) Then
        ' This is a digit or Cyrillic letter, leave as is
    ElseIf KeyAscii = 8 Then
        ' Backspace key, allow
    Else
        ' Does not meet requirements, reject
        KeyAscii = 0
    End If
End Sub

' Function to validate input against mask characters
Private Function IsValidMaskInput() As Boolean
    Dim i           As Integer
    Dim maskChar    As String
    Dim inputChar   As String

    ' If mask is empty, consider it valid
    If Me.Mask = vbNullString Then
        IsValidMaskInput = True
        Exit Function
    End If

    ' Check each character of the entered value against the mask
    For i = 1 To Me.LenValue
        maskChar = VBA.Mid$(Me.Mask, i, 1)
        inputChar = VBA.Mid$(Me.Value, i, 1)

        Select Case maskChar
            Case "#":    ' Digits
                IsValidMaskInput = IsValidDigit(Asc(inputChar))
            Case "@":    ' Latin letters
                IsValidMaskInput = IsValidLatinLetter(Asc(inputChar))
            Case "A":    ' Latin letters and digits
                IsValidMaskInput = IsValidDigit(Asc(inputChar)) Or IsValidLatinLetter(Asc(inputChar))
            Case "Б":    ' Cyrillic letters
                IsValidMaskInput = IsValidCyrillicLetter(Asc(inputChar))
            Case "б":    ' Cyrillic letters and digits
                IsValidMaskInput = IsValidDigit(Asc(inputChar)) Or IsValidCyrillicLetter(Asc(inputChar))
            Case "*":    ' Any characters
                IsValidMaskInput = True
            Case Else:    ' Fixed characters
                IsValidMaskInput = (inputChar = maskChar)
        End Select

        If Not IsValidMaskInput Then Exit Function
    Next i

    ' If less characters have been entered than the mask length, check possible remaining characters
    If Me.LenValue < Me.LenMask Then
        For i = Me.LenValue + 1 To Me.LenMask
            maskChar = VBA.Mid$(Me.Mask, i, 1)
            ' If the remaining mask characters are not fixed, consider it valid
            ' (the user can still enter valid characters)
            If InStr(1, SIMVOLS_MASKS, maskChar, vbBinaryCompare) > 0 Then
                IsValidMaskInput = True
                Exit Function
            End If
        Next i
    End If
End Function

' Function to validate input against regex pattern
Private Function IsValidRegexInput() As Boolean
    ' Check if the regular expression library is available
    On Error GoTo ErrorHandler
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .pattern = Me.regexPattern
        .IgnoreCase = False
        .Global = False
    End With

    IsValidRegexInput = regex.Test(Me.Value)
    Exit Function

ErrorHandler:
    ' If failed to create the regular expression object, return False
    IsValidRegexInput = False
End Function

Private Function IsValidInput() As Boolean
    If mTextBox Is Nothing Then
        IsValidInput = False
        Exit Function
    End If

    Select Case Me.CurrentMaskType
        Case enumTypeMask.tDateFix, enumTypeMask.tTimeFix
            If Me.LenMask = Me.LenValue Then
                If IsDate(Me.Value) Then
                    Me.Value = VBA.Format(VBA.CDate(Me.Value), Me.formatValue)
                    IsValidInput = Not (Me.Max < VBA.CDate(Me.Value) Or Me.Min > VBA.CDate(Me.Value))
                Else
                    IsValidInput = False
                End If
            Else
                IsValidInput = False
            End If
        Case enumTypeMask.tNumeric
            If IsNumeric(Me.Value) Then
                IsValidInput = Not (Me.Max < Me.Value Or Me.Min > Me.Value)
                'Me.Value = VBA.Format(Me.Value, Me.formatValue)
            Else
                IsValidInput = False
            End If
        Case enumTypeMask.tVariableLen
            ' For variable-length masks, check that the length does not exceed the maximum
            If Me.Max = 0 Then    ' If maximum length is not set, consider it valid
                IsValidInput = True
            Else
                IsValidInput = (Me.LenValue <= Me.Max)
            End If
            ' Also check mask compliance if it's specified
            If IsValidInput And Me.Mask <> vbNullString Then
                IsValidInput = IsValidMaskInput()
            End If
        Case enumTypeMask.tRegex
            ' For regular expressions, use the appropriate validation
            IsValidInput = IsValidRegexInput()
        Case Else
            IsValidInput = (Me.LenValue = Me.LenMask)
    End Select
    Call colorContr(IsValidInput)
End Function

Private Sub colorContr(ByRef Value As Boolean)
    If Me.TextBox Is Nothing Then Exit Sub
    With Me.TextBox
        If Value Then
            .BorderColor = Me.borderColorValid
        Else
            .BorderColor = Me.borderColorNoValid
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Function: IsControlInCollection
' Purpose:  Checks if a control with the specified name exists in the class
'           collection. Used to prevent duplication of items when adding new input masks.
' Parameters:
'   controlName - String containing the name of the control to check
' Returns: True if a control with that name already exists in the collection,
'          False otherwise
'---------------------------------------------------------------------------------------
Private Function IsControlInCollection(ByVal controlName As String) As Boolean
    Dim tempItem    As clsTextboxMask
    On Error Resume Next
    Set tempItem = getItemByName(controlName)
    On Error GoTo 0
    If Not tempItem Is Nothing Then IsControlInCollection = True
End Function

Private Sub Class_Terminate()
    ' Clear all objects when destroying the class
    Set mTextBox = Nothing
    Set mLabelHolder = Nothing
    Set mItems = Nothing
End Sub

' Procedure:     KeyAsciiRegex
' Description:   Handles key press events for regex-based masks, allowing only characters that match the regex pattern
' Parameters:    KeyAscii - ASCII code of the pressed key
Private Sub KeyAsciiRegex(ByRef KeyAscii As MSForms.ReturnInteger)
    ' Allow backspace key
    If KeyAscii = 8 Then
        Exit Sub
    End If

    ' Check if the test value matches the regex pattern (or at least the beginning of it)
    On Error GoTo ErrorHandler
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    On Error GoTo 0

    With regex
        .pattern = Me.regexPattern
        .IgnoreCase = False
        .Global = False
        ' Check if the new character is valid according to the pattern
        If Not .Test(VBA.ChrW$(KeyAscii)) Then KeyAscii = 0    ' Reject the character
    End With
    Exit Sub

ErrorHandler:
    ' If there's an error with the regex, allow the character to avoid blocking input
    ' This could happen if the regex pattern is invalid
End Sub
