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
' Version:       1.0.3
' Creation Date: 05.11.2025 10:30
' Update Date:   07.11.2025 16:50
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
'
'              Each masked field has visual indicators: border color
'              changes depending on the validity of input data, and
'              a placeholder hint with the expected format is displayed.
' ======================================================================================
' Enum:          enumTypeMask
' Description:   Enumeration of mask types supported by the class
' Values:
'   tOtherFix - Fixed-length text mask (letters, numbers, special characters)
'   tDateFix  - Fixed-length date input mask
'   tTimeFix  - Fixed-length time input mask
'   tNumeric  - Numeric mask with range, sign and decimal options
'   [_First]  - Helper value to determine the start of the enumeration
'   [_Last]   - Helper value to determine the end of the enumeration
Public Enum enumTypeMask
    tOtherFix = 1
    tDateFix
    tTimeFix
    tNumeric
    [_First] = tOtherFix
    [_Last] = tNumeric
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
    Set getItemByIndex = Me.Items(index)
    Exit Property
endGetItem:
    Err.Clear
End Property

' Property:      getItemByName
' Description:   Property to get an item from the input masks collection by name
' Parameters:
'   Name - String representing the name of the item in the collection
' Returns:       clsTextboxMask object corresponding to the specified name,
'                or Nothing if an item with that name does not exist
Public Property Get getItemByName(ByVal Name As String) As clsTextboxMask
    On Error GoTo endGetItem
    Set getItemByName = Me.Items(Name)
    Exit Property
endGetItem:
    Err.Clear
End Property

' Property: Version
' Purpose: Gets version information about the class
Public Property Get Version() As String
    Version = "Class: clsTextboxMask" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.3" & vbNewLine & _
            "Date of creation: 05.11.2025 10:30" & vbNewLine & _
            "Date of update: 07.11.2025 16:50" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: Class for creating textboxes with input masks in VBA" & vbNewLine & _
            "Supports various mask types: digital, letters, dates, phone numbers, etc." & vbNewLine & _
            "Provides input validation, placeholder display, and visual indication" & vbNewLine & _
            "of field fill status."
End Property

' Method:        addItemNumeric
' Description:   Adds a new item with numeric mask type to the collection
' Parameters:
'   TextBox         - Textbox object to apply the mask to
'   snMin           - Minimum allowed value
'   snMax           - Maximum allowed value
'   IsDecemal       - Flag allowing input of decimal numbers
'   IsNegative      - Flag allowing input of negative numbers
'   visibleLabelHolder - (Optional) Flag for placeholder hint visibility (default True)
'   formatNumeric   - (Optional) Number display format (default "#.0")
'   borderColorValid   - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   borderColorNoValid  - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   foreColorHolder - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
Public Sub addItemNumeric(ByRef TextBox As MSForms.TextBox, _
        ByVal snMin As Single, _
        ByVal snMax As Single, _
        ByVal IsDecemal As Boolean, _
        ByVal IsNegative As Boolean, _
        Optional visibleLabelHolder As Boolean = True, _
        Optional formatNumeric As String = "#.0", _
        Optional borderColorValid As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorNoValid As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, String$(WorksheetFunction.Max(VBA.Len(snMin), VBA.Len(snMax)), "#"), _
            enumTypeMask.tNumeric, formatNumeric, IsDecemal, IsNegative, borderColorValid, borderColorNoValid, _
            snMin, snMax, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItemFixLenDate
' Description:   Adds a new item with fixed-length date mask to the collection
' Parameters:
'   TextBox         - Textbox object to apply the mask to
'   Mask            - Mask string for date (e.g., "##.##.####")
'   minDate         - Minimum allowed date
'   maxDate         - Maximum allowed date
'   formatDate      - (Optional) Date display format (default "dd.mm.yyyy")
'   visibleLabelHolder - (Optional) Flag for placeholder hint visibility (default True)
'   borderColorValid   - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   borderColorNoValid  - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   foreColorHolder - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
Public Sub addItemFixLenDate(ByRef TextBox As MSForms.TextBox, ByVal Mask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional formatDate As String = "dd.mm.yyyy", _
        Optional visibleLabelHolder As Boolean = True, _
        Optional borderColorValid As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorNoValid As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, Mask, enumTypeMask.tDateFix, formatDate, False, False, borderColorValid, borderColorNoValid, _
            minDate, maxDate, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItemFixLenTime
' Description:   Adds a new item with fixed-length time mask to the collection
' Parameters:
'   TextBox         - Textbox object to apply the mask to
'   Mask            - Mask string for time (e.g., "##:##")
'   minDate         - Minimum allowed time
'   maxDate         - Maximum allowed time
'   formatDate      - (Optional) Time display format (default "hh:mm")
'   visibleLabelHolder - (Optional) Flag for placeholder hint visibility (default True)
'   borderColorValid   - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   borderColorNoValid  - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   foreColorHolder - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
Public Sub addItemFixLenTime(ByRef TextBox As MSForms.TextBox, ByVal Mask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional formatDate As String = "hh:mm", _
        Optional visibleLabelHolder As Boolean = True, _
        Optional borderColorValid As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorNoValid As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, Mask, enumTypeMask.tTimeFix, formatDate, False, False, borderColorValid, borderColorNoValid, _
            0, 1, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItemFixLen
' Description:   Adds a new item with fixed-length text mask to the collection
' Parameters:
'   TextBox         - Textbox object to apply the mask to
'   Mask            - Mask string for text (e.g., "###@@@")
'   visibleLabelHolder - (Optional) Flag for placeholder hint visibility (default True)
'   borderColorValid   - (Optional) Border color when input is valid (default COLOR_DEFAULT_ON)
'   borderColorNoValid  - (Optional) Border color when input is invalid (default COLOR_DEFAULT_OFF)
'   foreColorHolder - (Optional) Hint text color (default COLOR_DEFAULT_HOLDER)
Public Sub addItemFixLen(ByRef TextBox As MSForms.TextBox, ByVal Mask As String, _
        Optional visibleLabelHolder As Boolean = True, _
        Optional borderColorValid As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorNoValid As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, Mask, enumTypeMask.tOtherFix, vbNullString, False, False, borderColorValid, borderColorNoValid, _
            0, 0, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItem
' Description:   Internal method for adding a new input mask item to the collection
'                with full parameter configuration. Used by helper methods
'                addItemNumeric, addItemFixLenDate, addItemFixLenTime and addItemFixLen.
' Parameters:
'   TextBox         - Textbox object to apply the mask to
'   Mask            - Mask string for the text field
'   maskType        - Mask type from the enumTypeMask enumeration
'   formatValue     - Value display format (used for dates and numbers)
'   IsDecemal       - Flag allowing input of decimal numbers (for numeric masks)
'   IsNegative      - Flag allowing input of negative numbers (for numeric masks)
'   borderColorValid   - Border color when input is valid
'   borderColorNoValid  - Border color when input is invalid
'   snMin           - Minimum allowed value (for numeric and date masks)
'   snMax           - Maximum allowed value (for numeric and date masks)
'   visibleLabelHolder - Flag for placeholder hint visibility
'   foreColorHolder - Placeholder hint text color
Private Sub addItem(ByRef TextBox As MSForms.TextBox, _
        ByVal Mask As String, _
        ByVal maskType As enumTypeMask, _
        ByVal formatValue As String, _
        ByVal IsDecemal As Boolean, _
        ByVal IsNegative As Boolean, _
        ByVal borderColorValid As XlRgbColor, _
        ByVal borderColorNoValid As XlRgbColor, _
        ByVal snMin As Single, _
        ByVal snMax As Single, _
        ByVal visibleLabelHolder As Boolean, _
        ByVal foreColorHolder As XlRgbColor)

    If TextBox Is Nothing Then
        Err.Raise vbObjectError + 101, "clsTextboxMask", "TextBox cannot be Nothing"
        Exit Sub
    End If

    If maskType < enumTypeMask.[_First] Or maskType > enumTypeMask.[_Last] Then
        Err.Raise vbObjectError + 102, "clsTextboxMask", "Invalid mask type"
        Exit Sub
    End If

    If IsControlInCollection(TextBox.Name) Then
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
        Set .TextBox = TextBox
        .Mask = Mask
        If .LenMask > 0 Then .TextBox.MaxLength = .LenMask

        .borderColorValid = borderColorValid
        .borderColorNoValid = borderColorNoValid

        .IsDecemal = IsDecemal
        .IsNegative = IsNegative

        .Max = snMax
        .Min = snMin

        Dim labelName As String
        labelName = TextBox.Name & "_holder"

        On Error Resume Next
        Set Me.LabelHolder = .TextBox.Parent.Controls(labelName)
        On Error GoTo 0

        If Not Me.LabelHolder Is Nothing Then
            If Not .TextBox.Parent.Controls(labelName) Is Nothing Then
                .TextBox.Parent.Controls.Remove labelName
            End If
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
            .ForeColor = foreColorHolder
            .Font.Name = itemCls.TextBox.Font.Name
            .Font.Size = itemCls.TextBox.Font.Size * 0.9
        End With

        .visibleLabelHolder = visibleLabelHolder
        .CurrentMaskType = maskType
        .formatValue = formatValue
        If formatValue = vbNullString Then .formatValue = maskType
        If mItems Is Nothing Then Set mItems = New Collection
        Call mItems.Add(itemCls, TextBox.Name)
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
    If Not Me.LabelHolder Is Nothing Then
        On Error Resume Next
        If IsControlInCollection(Me.TextBox.Name) Then
            If Not Me.LabelHolder Is Nothing Then Call Me.TextBox.Parent.Controls.Remove(Me.LabelHolder.Name)
            Call Me.Items.Remove(Me.TextBox.Name)
        End If
        Set Me.LabelHolder = Nothing
        On Error GoTo 0
    End If
    Set Me.TextBox = Nothing
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

Private Sub mTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Me.CurrentMaskType
        Case enumTypeMask.tDateFix, enumTypeMask.tTimeFix, enumTypeMask.tOtherFix:
            Call KeyAsciiFixLenText(KeyAscii)
        Case enumTypeMask.tNumeric
            Call KeyAsciiNumeric(KeyAscii)
    End Select
End Sub

Private Sub KeyAsciiNumeric(ByRef KeyAscii As MSForms.ReturnInteger)
    Call NumericAll(KeyAscii)
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
    Dim i           As Byte

    For i = 1 To VBA.Len(SIMVOLS_MASKS)
        If sVal = VBA.Mid$(SIMVOLS_MASKS, i, 1) Then
            haveSimvol = True
            Exit Function
        End If
    Next i
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
                KeyAscii = IIf(InStr(1, Me.Value, ",") > 0, 0, 44)
            Else
                KeyAscii = 0
            End If
        Case 8:
            ' Do nothing, character is allowed
        Case Else:
            KeyAscii = 0
    End Select
End Sub

Private Sub LatinLetters(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 65 To 90, 97 To 122:    ' ASCII codes for uppercase and lowercase Latin letters
        Case 8:    ' ASCII code for Backspace key - allow
            ' Do nothing, character is allowed
        Case Else:
            KeyAscii = 0
    End Select
End Sub

Private Sub UniCodeLetters(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 192 To 255, 1040 To 1103:    ' ASCII codes for uppercase and lowercase UniCode letters
        Case 8:    ' ASCII code for Backspace key - allow
            ' Do nothing, character is allowed
        Case Else:
            KeyAscii = 0
    End Select
End Sub

Private Sub AlphaNumeric(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 65 To 90, 97 To 122:    ' ASCII codes for digits and Latin letters
        Case 8:    ' ASCII code for Backspace key - allow
            ' Do nothing, character is allowed
        Case Else:
            KeyAscii = 0
    End Select
End Sub

Private Sub UniCodeLettersNumeric(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 192 To 255, 1040 To 1103:    ' ASCII codes for uppercase and lowercase UniCode letters
        Case 8:    ' ASCII code for Backspace key - allow
            ' Do nothing, character is allowed
        Case Else:
            KeyAscii = 0
    End Select
End Sub

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
        Case Else
            IsValidInput = (Me.LenValue = Me.LenMask)
    End Select
    Call colorContr(IsValidInput)
End Function

Private Sub colorContr(ByRef Value As Boolean)
    If Me.TextBox Is Nothing Then Exit Sub
    With Me.TextBox
        If Value Then
            .BorderColor = Me.BorderColorValid
        Else
            .BorderColor = Me.BorderColorNoValid
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
    Set tempItem = getItemByName(controlName)
    If Not tempItem Is Nothing Then IsControlInCollection = True
End Function

Private Sub Class_Terminate()
    Set Me.TextBox = Nothing
    Set Me.LabelHolder = Nothing
    Set Me.Items = Nothing
End Sub
