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

' #####################################################################################################################
' Class: clsTextboxMask
' Author: VBATools
' Version: 1.0.6
' Creation Date: 05.11.2025 10:30
' Last Update Date: 10.11.2025 19:57
' License: Apache License
' #####################################################################################################################
' Purpose:
'   Class for creating textbox fields with input masks in VBA. Provides input validation, placeholder display
'   and visual indication of field fill status. Supports various mask types: numeric, letters, dates,
'   phone numbers and others.
'
' Main Features:
'   - Support for various input mask types (numbers, dates, time, text, regular expressions)
'   - Real-time input validation
'   - Display of placeholders with different statuses (empty, partially filled, completely filled, invalid)
'   - Visual indication of input correctness through border color
'   - Support for numeric values with range, sign and decimal restrictions
'   - Support for variable length text
'   - Support for validation via regular expressions
'
' Usage Examples:
'   1. Creating a numeric field with constraints:
'      Dim numField As New clsTextboxMask
'      Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
'                                   minValue:=0, _
'                                   maxValue:=100, _
'                                   allowDecimal:=True, _
'                                   allowNegative:=False)
'
'   2. Creating a date field:
'      Dim dateField As New clsTextboxMask
'      Call dateField.AddFieldDate(inputTextBox:=Me.TextBox2, _
'                                 dateMask:="##.##.####", _
'                                 minDate:=#1/1/2020#, _
'                                 maxDate:=#12/31/2030#, _
'                                 dateFormat:="dd.mm.yyyy")
'
'   3. Creating a text field with mask:
'      Dim textField As New clsTextboxMask
'      Call textField.AddFieldText(inputTextBox:=Me.TextBox3, _
'                                 textMask:="+7(*##) @# A# #Б#")  ' Letters-digits
'
' Dependencies:
'   - MSForms.TextBox
'   - MSForms.Label
'   - VBScript.RegExp (for validation via regular expressions)
'
' #####################################################################################################################

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

Private mSimvolsMasks As String
Private mBorderColorValid As Long
Private mBorderColorInvalid As Long
Private mPlaceholderColor As Long

Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private mLabelPlaceholder As MSForms.Label

Private mItems      As Collection

Private mMask       As String
Private mFormatValue As String

Private mRegexPattern As String             ' Regular expression pattern for tRegex masks
Private mRegexFilter As String              ' Regular expression filter
Private mRegex      As Object               ' Cached regex object for tRegex masks

Private mPlaceholderEmpty As String         ' Placeholder text when field is empty
Private mPlaceholderPartial As String       ' Placeholder text when field is partially filled
Private mPlaceholderComplete As String      ' Placeholder text when field is completely filled
Private mPlaceholderInvalid As String       ' Placeholder text when field contains invalid data
Private mPlaceholderTemplate As String      ' Example text for placeholders

Private mIsDecimal  As Boolean
Private mIsNegative As Boolean

Private mMin        As Single
Private mMax        As Single

Private mCurrentMaskType As enumTypeMask

Public Property Get ConstSimvolsMasks() As String
    ConstSimvolsMasks = mSimvolsMasks
End Property

' #####################################################################################################################
' Property: PlaceholderMask
' #####################################################################################################################
' Purpose:
'   Returns the current placeholder mask, which shows the remaining characters that need to be filled in the textbox
'   If the textbox value is shorter than the mask length, the placeholder will show the remaining mask characters
'   If the textbox value is equal to or longer than the mask length, the placeholder will be empty
'
' Return Value:
'   String - The placeholder mask showing remaining characters to be filled
' #####################################################################################################################
Public Property Get PlaceholderMask() As String
    If Me.TextBox Is Nothing Then Exit Property
    If Me.LenMask >= Me.LenValue Then
        PlaceholderMask = Me.Value & VBA.Mid$(Me.Mask, Me.LenValue + 1, Me.LenMask - Me.LenValue)
    Else
        PlaceholderMask = Me.Mask
    End If
End Property

Public Property Get PlaceholderEmpty() As String
    PlaceholderEmpty = mPlaceholderEmpty
End Property

Public Property Let PlaceholderEmpty(ByVal Value As String)
    mPlaceholderEmpty = Value
    'Call UpdatePlaceholder
End Property

Public Property Get PlaceholderPartial() As String
    PlaceholderPartial = mPlaceholderPartial
End Property

Public Property Let PlaceholderPartial(ByVal Value As String)
    mPlaceholderPartial = Value
    'Call UpdatePlaceholder
End Property

Public Property Get PlaceholderComplete() As String
    PlaceholderComplete = mPlaceholderComplete
End Property

Public Property Let PlaceholderComplete(ByVal Value As String)
    mPlaceholderComplete = Value
    'Call UpdatePlaceholder
End Property

Public Property Get PlaceholderInvalid() As String
    PlaceholderInvalid = mPlaceholderInvalid
End Property

Public Property Let PlaceholderInvalid(ByVal Value As String)
    mPlaceholderInvalid = Value
    'Call UpdatePlaceholder
End Property

Public Property Get PlaceHolderTemplate() As String
    PlaceHolderTemplate = mPlaceholderTemplate
End Property

Public Property Let PlaceHolderTemplate(ByVal Value As String)
    mPlaceholderTemplate = Value
    'Call UpdatePlaceholder
End Property

Public Property Get VisibleLabelPlaceholder() As Boolean
    If Me.LabelPlaceholder Is Nothing Then Exit Property
    VisibleLabelPlaceholder = Me.LabelPlaceholder.Visible
End Property

Public Property Let VisibleLabelPlaceholder(ByVal Value As Boolean)
    If Me.LabelPlaceholder Is Nothing Then Exit Property
    Me.LabelPlaceholder.Visible = Value
End Property

Public Property Get PlaceholderColor() As Long
    PlaceholderColor = mPlaceholderColor
End Property

Public Property Let PlaceholderColor(ByVal Value As Long)
    mPlaceholderColor = Value
End Property

Public Property Let BorderColorValid(ByRef Color As Long)
    mBorderColorValid = Color
End Property

Public Property Get BorderColorValid() As Long
    BorderColorValid = mBorderColorValid
End Property

Public Property Let BorderColorInvalid(ByRef Color As Long)
    mBorderColorInvalid = Color
End Property

Public Property Get BorderColorInvalid() As Long
    BorderColorInvalid = mBorderColorInvalid
End Property

Public Property Get Min() As Single
    Min = mMin
End Property

Public Property Let Min(ByVal Value As Single)
    mMin = Value
End Property

Public Property Get Max() As Single
    Max = mMax
End Property

Public Property Let Max(ByVal Value As Single)
    mMax = Value
End Property

Public Property Get IsNegative() As Boolean
    IsNegative = mIsNegative
End Property

Public Property Let IsNegative(ByVal Value As Boolean)
    mIsNegative = Value
End Property

Public Property Get IsDecimal() As Boolean
    IsDecimal = mIsDecimal
End Property

Public Property Let IsDecimal(ByVal Value As Boolean)
    mIsDecimal = Value
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = mTextBox
End Property

Public Property Set TextBox(ByRef TextBox As MSForms.TextBox)
    Set mTextBox = TextBox
End Property

Public Property Get LabelPlaceholder() As MSForms.Label
    Set LabelPlaceholder = mLabelPlaceholder
End Property

Public Property Set LabelPlaceholder(ByRef Label As MSForms.Label)
    Set mLabelPlaceholder = Label
End Property

Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Property Set Items(ByRef Items As Collection)
    Set mItems = Items
End Property

Public Property Get CurrentMaskType() As enumTypeMask
    CurrentMaskType = mCurrentMaskType
End Property

Public Property Let CurrentMaskType(ByVal maskType As enumTypeMask)
    mCurrentMaskType = maskType
End Property

Public Property Get LenValue() As Integer
    If Me.TextBox Is Nothing Then Exit Property
    LenValue = VBA.Len(Me.Value)
End Property

Public Property Let Value(ByRef Value As String)
    If Me.TextBox Is Nothing Then Exit Property
    Me.TextBox.Value = Value
End Property

Public Property Get Value() As String
    If Me.TextBox Is Nothing Then Exit Property
    Value = Me.TextBox.Value
End Property

Public Property Let Mask(ByRef Mask As String)
    mMask = Mask
End Property

Public Property Get Mask() As String
    Mask = mMask
End Property

Public Property Get LenMask() As Integer
    LenMask = VBA.Len(Me.Mask)
End Property

Public Property Let FormatValue(ByRef FormatValue As String)
    mFormatValue = FormatValue
End Property

Public Property Get FormatValue() As String
    FormatValue = mFormatValue
End Property

Public Property Let RegexPattern(ByRef RegexPattern As String)
    mRegexPattern = RegexPattern
    If RegexPattern <> vbNullString Then Call InitializeRegex
End Property

Public Property Get RegexPattern() As String
    RegexPattern = mRegexPattern
End Property

Public Property Let RegexFilter(ByRef RegexFilter As String)
    mRegexFilter = RegexFilter
    If RegexFilter <> vbNullString Then Call InitializeRegex
End Property

Public Property Get RegexFilter() As String
    RegexFilter = mRegexFilter
End Property

Public Property Get Count() As Byte
    Count = Me.Items.Count
End Property

Public Property Get RemainingChars() As Integer
    If Me.TextBox Is Nothing Then
        RemainingChars = Me.LenMask
        Exit Property
    End If
    RemainingChars = Me.LenMask - Me.LenValue
End Property

Public Property Get GetItemByIndex(ByVal index As Integer) As clsTextboxMask
    On Error GoTo endGetItem
    If Not Me.Items Is Nothing Then Set GetItemByIndex = Me.Items(index)
    Exit Property
endGetItem:
    ' Don't clear the error, but log it for diagnostics
    Select Case Err.Number
        Case 0, 5
        Case Else
            Debug.Print "Error in getItemByName: " & Err.Description & " (Number: " & Err.Number & ")"
            Set GetItemByIndex = Nothing
    End Select
End Property

Public Property Get GetItemByName(ByVal Name As String) As clsTextboxMask
    On Error GoTo endGetItem
    If Not Me.Items Is Nothing Then Set GetItemByName = Me.Items(Name)
    Exit Property
endGetItem:
    ' Don't clear the error, but log it for diagnostics
    Select Case Err.Number
        Case 0, 5
        Case Else
            Debug.Print "Error in getItemByName: " & Err.Description & " (Number: " & Err.Number & ")"
            Set GetItemByName = Nothing
    End Select
End Property

Public Property Get Version() As String
    Version = "Class: clsTextboxMask" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.6" & vbNewLine & _
            "Date of creation: 05.11.2025 10:30" & vbNewLine & _
            "Date of update: 10.11.2025 19:57" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: Class for creating textboxes with input masks in VBA" & vbNewLine & _
            "Supports various mask types: digital, letters, dates, phone numbers, etc." & vbNewLine & _
            "Provides input validation, placeholder display, and visual indication" & vbNewLine & _
            "of field fill status."
End Property
' #####################################################################################################################
' Procedure: AddFieldNumeric
' #####################################################################################################################
' Purpose:
'   Adds a numeric field with specified validation parameters
'
' Parameters:
'   inputTextBox - textbox field to which the mask is applied
'   minValue - minimum allowed value
'   maxValue - maximum allowed value
'   allowDecimal - allow input of decimal values
'   allowNegative - allow input of negative values
'   showPlaceholder - show placeholder
'   numberFormat - number display format
'   BorderColorValid - border color for correct input
'   BorderColorInvalid - border color for incorrect input
'   PlaceholderColor - placeholder color
'   PlaceholderEmpty - placeholder text for empty field
'   PlaceholderPartial - placeholder text for partially filled field
'   PlaceholderComplete - placeholder text for completely filled field
'   PlaceholderInvalid - placeholder text for field with invalid data
'   PlaceHolderTemplete - placeholder template
' #####################################################################################################################

Public Sub AddFieldNumeric(ByRef inputTextBox As MSForms.TextBox, _
        ByVal minValue As Single, _
        ByVal maxValue As Single, _
        ByVal allowDecimal As Boolean, _
        ByVal allowNegative As Boolean, _
        Optional showPlaceholder As Boolean = True, _
        Optional numberFormat As String = "#.0", _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
    Call AddField(inputTextBox:=inputTextBox, textMask:=VBA.String$(WorksheetFunction.Max(VBA.Len(minValue), VBA.Len(maxValue)), "#"), _
            maskType:=enumTypeMask.tNumeric, FormatValue:=numberFormat, allowDecimal:=allowDecimal, allowNegative:=allowNegative, _
            BorderColorValid:=BorderColorValid, BorderColorInvalid:=BorderColorInvalid, minValue:=minValue, maxValue:=maxValue, _
            showPlaceholder:=showPlaceholder, PlaceholderColor:=PlaceholderColor, RegexPattern:=vbNullString, _
            PlaceholderEmpty:=PlaceholderEmpty, _
            PlaceholderPartial:=PlaceholderPartial, _
            PlaceholderComplete:=PlaceholderComplete, _
            PlaceholderInvalid:=PlaceholderInvalid, _
            PlaceHolderTemplate:=PlaceHolderTemplate, _
            RegexFilter:=vbNullString)
End Sub
' #####################################################################################################################
' Procedure: AddFieldDate
' #####################################################################################################################
' Purpose:
'   Adds a date input field with specified validation parameters
'
' Parameters:
'   inputTextBox - textbox field to which the mask is applied
'   dateMask - date input mask
'   minDate - minimum allowed date
'   maxDate - maximum allowed date
'   dateFormat - date display format
'   showPlaceholder - show placeholder
'   BorderColorValid - border color for correct input
'   BorderColorInvalid - border color for incorrect input
'   PlaceholderColor - placeholder color
'   PlaceholderEmpty - placeholder text for empty field
'   PlaceholderPartial - placeholder text for partially filled field
'   PlaceholderComplete - placeholder text for completely filled field
'   PlaceholderInvalid - placeholder text for field with invalid data
'   PlaceHolderTemplete - placeholder template
' #####################################################################################################################

Public Sub AddFieldDate(ByRef inputTextBox As MSForms.TextBox, ByVal dateMask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional dateFormat As String = "dd.mm.yyyy", _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
    Call AddField(inputTextBox:=inputTextBox, textMask:=dateMask, maskType:=enumTypeMask.tDateFix, _
            FormatValue:=dateFormat, allowDecimal:=False, allowNegative:=False, BorderColorValid:=BorderColorValid, _
            BorderColorInvalid:=BorderColorInvalid, minValue:=minDate, maxValue:=maxDate, showPlaceholder:=showPlaceholder, _
            PlaceholderColor:=PlaceholderColor, RegexPattern:=vbNullString, _
            PlaceholderEmpty:=PlaceholderEmpty, _
            PlaceholderPartial:=PlaceholderPartial, _
            PlaceholderComplete:=PlaceholderComplete, _
            PlaceholderInvalid:=PlaceholderInvalid, _
            PlaceHolderTemplate:=PlaceHolderTemplate, _
            RegexFilter:=vbNullString)
End Sub
' #####################################################################################################################
' Procedure: AddFieldTime
' #####################################################################################################################
' Purpose:
'   Adds a time input field with specified validation parameters
'
' Parameters:
'   inputTextBox - textbox field to which the mask is applied
'   timeMask - time input mask
'   minTime - minimum allowed time
'   maxTime - maximum allowed time
'   timeFormat - time display format
'   showPlaceholder - show placeholder
'   BorderColorValid - border color for correct input
'   BorderColorInvalid - border color for incorrect input
'   PlaceholderColor - placeholder color
'   PlaceholderEmpty - placeholder text for empty field
'   PlaceholderPartial - placeholder text for partially filled field
'   PlaceholderComplete - placeholder text for completely filled field
'   PlaceholderInvalid - placeholder text for field with invalid data
'   PlaceHolderTemplete - placeholder template
' #####################################################################################################################

Public Sub AddFieldTime(ByRef inputTextBox As MSForms.TextBox, ByVal timeMask As String, _
        ByVal minTime As Date, _
        ByVal maxTime As Date, _
        Optional timeFormat As String = "hh:mm", _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
    Call AddField(inputTextBox:=inputTextBox, textMask:=timeMask, maskType:=enumTypeMask.tTimeFix, _
            FormatValue:=timeFormat, allowDecimal:=False, allowNegative:=False, BorderColorValid:=BorderColorValid, _
            BorderColorInvalid:=BorderColorInvalid, minValue:=0, maxValue:=1, showPlaceholder:=showPlaceholder, _
            PlaceholderColor:=PlaceholderColor, RegexPattern:=vbNullString, _
            PlaceholderEmpty:=PlaceholderEmpty, _
            PlaceholderPartial:=PlaceholderPartial, _
            PlaceholderComplete:=PlaceholderComplete, _
            PlaceholderInvalid:=PlaceholderInvalid, _
            PlaceHolderTemplate:=PlaceHolderTemplate, _
            RegexFilter:=vbNullString)
End Sub
' #####################################################################################################################
' Procedure: AddFieldText
' #####################################################################################################################
' Purpose:
'   Adds a text field with specified input mask
'
' Parameters:
'   inputTextBox - textbox field to which the mask is applied
'   textMask - text input mask
'   showPlaceholder - show placeholder
'   BorderColorValid - border color for correct input
'   BorderColorInvalid - border color for incorrect input
'   PlaceholderColor - placeholder color
'   PlaceholderEmpty - placeholder text for empty field
'   PlaceholderPartial - placeholder text for partially filled field
'   PlaceholderComplete - placeholder text for completely filled field
'   PlaceholderInvalid - placeholder text for field with invalid data
'   PlaceHolderTemplete - placeholder template
' #####################################################################################################################

Public Sub AddFieldText(ByRef inputTextBox As MSForms.TextBox, ByVal textMask As String, _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
    Call AddField(inputTextBox:=inputTextBox, textMask:=textMask, maskType:=enumTypeMask.tOtherFix, _
            FormatValue:=vbNullString, allowDecimal:=False, allowNegative:=False, BorderColorValid:=BorderColorValid, _
            BorderColorInvalid:=BorderColorInvalid, minValue:=0, maxValue:=0, showPlaceholder:=showPlaceholder, _
            PlaceholderColor:=PlaceholderColor, RegexPattern:=vbNullString, _
            PlaceholderEmpty:=PlaceholderEmpty, _
            PlaceholderPartial:=PlaceholderPartial, _
            PlaceholderComplete:=PlaceholderComplete, _
            PlaceholderInvalid:=PlaceholderInvalid, _
            PlaceHolderTemplate:=PlaceHolderTemplate, _
            RegexFilter:=vbNullString)
End Sub
' #####################################################################################################################
' Procedure: AddFieldVariableLength
' #####################################################################################################################
' Purpose:
'   Adds a field with variable text length
'
' Parameters:
'   inputTextBox - textbox field to which the mask is applied
'   maxLength - maximum text length
'   textMask - text input mask (optional)
'   showPlaceholder - show placeholder
'   BorderColorValid - border color for correct input
'   BorderColorInvalid - border color for incorrect input
'   PlaceholderColor - placeholder color
'   PlaceholderEmpty - placeholder text for empty field
'   PlaceholderPartial - placeholder text for partially filled field
'   PlaceholderComplete - placeholder text for completely filled field
'   PlaceholderInvalid - placeholder text for field with invalid data
'   PlaceHolderTemplete - placeholder template
' #####################################################################################################################

Public Sub AddFieldVariableLength(ByRef inputTextBox As MSForms.TextBox, _
        ByVal maxLength As Integer, _
        Optional textMask As String = vbNullString, _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
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

    Call AddField(inputTextBox:=inputTextBox, textMask:=maskStr, maskType:=enumTypeMask.tVariableLen, _
            FormatValue:=vbNullString, allowDecimal:=False, allowNegative:=False, BorderColorValid:=BorderColorValid, _
            BorderColorInvalid:=BorderColorInvalid, minValue:=0, maxValue:=maxLength, showPlaceholder:=showPlaceholder, _
            PlaceholderColor:=PlaceholderColor, RegexPattern:=vbNullString, _
            PlaceholderEmpty:=PlaceholderEmpty, _
            PlaceholderPartial:=PlaceholderPartial, _
            PlaceholderComplete:=PlaceholderComplete, _
            PlaceholderInvalid:=PlaceholderInvalid, _
            PlaceHolderTemplate:=PlaceHolderTemplate, _
            RegexFilter:=vbNullString)
End Sub
' #####################################################################################################################
' Procedure: AddFieldRegex
' #####################################################################################################################
' Purpose:
'   Adds a field with validation via regular expression
'
' Parameters:
'   inputTextBox - textbox field to which the mask is applied
'   RegexPattern - regular expression pattern for validation
'   RegexFilter - regular expression filter
'   showPlaceholder - show placeholder
'   BorderColorValid - border color for correct input
'   BorderColorInvalid - border color for incorrect input
'   PlaceholderColor - placeholder color
'   PlaceholderEmpty - placeholder text for empty field
'   PlaceholderPartial - placeholder text for partially filled field
'   PlaceholderComplete - placeholder text for completely filled field
'   PlaceholderInvalid - placeholder text for field with invalid data
'   PlaceHolderTemplete - placeholder template
' #####################################################################################################################

Public Sub AddFieldRegex(ByRef inputTextBox As MSForms.TextBox, _
        ByVal RegexPattern As String, _
        ByVal RegexFilter As String, _
        Optional showPlaceholder As Boolean = True, _
        Optional BorderColorValid As XlRgbColor = 0, _
        Optional BorderColorInvalid As XlRgbColor = 0, _
        Optional PlaceholderColor As XlRgbColor = 0, _
        Optional PlaceholderEmpty As String = vbNullString, _
        Optional PlaceholderPartial As String = vbNullString, _
        Optional PlaceholderComplete As String = vbNullString, _
        Optional PlaceholderInvalid As String = vbNullString, _
        Optional PlaceHolderTemplate As String = "{holder}")
    Call AddField(inputTextBox:=inputTextBox, textMask:=vbNullString, maskType:=enumTypeMask.tRegex, _
            FormatValue:=vbNullString, allowDecimal:=False, allowNegative:=False, BorderColorValid:=BorderColorValid, _
            BorderColorInvalid:=BorderColorInvalid, minValue:=0, maxValue:=0, showPlaceholder:=showPlaceholder, _
            PlaceholderColor:=PlaceholderColor, RegexPattern:=RegexPattern, _
            PlaceholderEmpty:=PlaceholderEmpty, _
            PlaceholderPartial:=PlaceholderPartial, _
            PlaceholderComplete:=PlaceholderComplete, _
            PlaceholderInvalid:=PlaceholderInvalid, _
            PlaceHolderTemplate:=PlaceHolderTemplate, _
            RegexFilter:=RegexFilter)
End Sub

Private Sub AddField(ByRef inputTextBox As MSForms.TextBox, _
        ByVal textMask As String, _
        ByVal maskType As enumTypeMask, _
        ByVal FormatValue As String, _
        ByVal allowDecimal As Boolean, _
        ByVal allowNegative As Boolean, _
        ByVal BorderColorValid As XlRgbColor, _
        ByVal BorderColorInvalid As XlRgbColor, _
        ByVal minValue As Single, _
        ByVal maxValue As Single, _
        ByVal showPlaceholder As Boolean, _
        ByVal PlaceholderColor As XlRgbColor, _
        ByVal PlaceholderEmpty As String, _
        ByVal PlaceholderPartial As String, _
        ByVal PlaceholderComplete As String, _
        ByVal PlaceholderInvalid As String, _
        ByVal PlaceHolderTemplate As String, _
        ByVal RegexPattern As String, _
        ByVal RegexFilter As String)

    If inputTextBox Is Nothing Then
        Err.Raise vbObjectError + 101, "clsTextboxMask", "TextBox cannot be Nothing"
        Exit Sub
    End If

    If maskType < enumTypeMask.[_First] Or maskType > enumTypeMask.[_Last] Then
        Err.Raise vbObjectError + 102, "clsTextboxMask", "Invalid mask type"
        Exit Sub
    End If

    If IsControlInCollection(inputTextBox.Name) Then
        Err.Raise vbObjectError + 103, "clsTextboxMask", "The item has already been created"
        Exit Sub
    End If

    Dim itemCls     As clsTextboxMask
    Set itemCls = New clsTextboxMask

    With itemCls
        Set .TextBox = inputTextBox
        .Mask = textMask
        If .LenMask > 0 Then .TextBox.maxLength = .LenMask

        Dim labelName As String
        labelName = inputTextBox.Name & "_holder"

        ' Check if a label with this name already exists
        On Error Resume Next
        Dim existingLabel As MSForms.Control
        Set existingLabel = .TextBox.Parent.Controls(labelName)
        On Error GoTo 0

        If Not existingLabel Is Nothing Then .TextBox.Parent.Controls.Remove existingLabel.Name
        Set .LabelPlaceholder = .TextBox.Parent.Controls.Add("Forms.Label.1", labelName, True)

        With .LabelPlaceholder
            .Left = itemCls.TextBox.Left + 8
            .Width = itemCls.TextBox.Width
            .Height = itemCls.TextBox.Height * 0.6
            .Top = itemCls.TextBox.Top - .Height
            .BackStyle = fmBackStyleTransparent
            '.BorderStyle = fmBorderStyleSingle
            .Font.Name = itemCls.TextBox.Font.Name
            .Font.Size = itemCls.TextBox.Font.Size * 0.9
        End With

        ' Use default colors if 0 is passed as parameter
        .BorderColorValid = IIf(BorderColorValid = 0, Me.BorderColorValid, BorderColorValid)
        .BorderColorInvalid = IIf(BorderColorInvalid = 0, Me.BorderColorInvalid, BorderColorInvalid)

        .IsDecimal = allowDecimal
        .IsNegative = allowNegative

        .Max = maxValue
        .Min = minValue

        .VisibleLabelPlaceholder = showPlaceholder
        .CurrentMaskType = maskType
        ' Set the regular expression for this item
        .RegexPattern = RegexPattern
        .RegexFilter = RegexFilter

        ' Initialize placeholder texts if provided
        .PlaceholderEmpty = PlaceholderEmpty
        .PlaceholderPartial = PlaceholderPartial
        .PlaceholderComplete = PlaceholderComplete
        .PlaceholderInvalid = PlaceholderInvalid
        .PlaceHolderTemplate = PlaceHolderTemplate

        ' Set placeholder color
        .LabelPlaceholder.ForeColor = IIf(PlaceholderColor = 0, Me.PlaceholderColor, PlaceholderColor)

        .FormatValue = IIf(FormatValue = vbNullString, maskType, FormatValue)

        If mItems Is Nothing Then Set mItems = New Collection
        Call mItems.Add(itemCls, inputTextBox.Name)
        Call .IsValid
        Set .Items = mItems
        .UpdatePlaceholder
    End With
End Sub
' #####################################################################################################################
' Function: IsValid
' #####################################################################################################################
' Purpose:
'   Checks the correctness of data entered in the textbox field
'
' Return Value:
'   Boolean - True if data is correct, otherwise False
' #####################################################################################################################

Public Function IsValid() As Boolean
    IsValid = IsValidInput()
End Function
' #####################################################################################################################
' Procedure: Clear
' #####################################################################################################################
' Purpose:
'   Clears the textbox field
' #####################################################################################################################

Public Sub Clear()
    If Me.TextBox Is Nothing Then Exit Sub
    Me.Value = vbNullString
End Sub
' #####################################################################################################################
' Procedure: SetFocus
' #####################################################################################################################
' Purpose:
'   Sets focus on the textbox field
' #####################################################################################################################

Public Sub SetFocus()
    If Me.TextBox Is Nothing Then Exit Sub
    Me.TextBox.SetFocus
End Sub
' #####################################################################################################################
' Procedure: RemoveItem
' #####################################################################################################################
' Purpose:
'   Removes the textbox mask element and related components
' #####################################################################################################################

Public Sub RemoveItem()
    On Error Resume Next
    If IsControlInCollection(Me.TextBox.Name) Then
        If Not Me.LabelPlaceholder Is Nothing Then Call Me.TextBox.Parent.Controls.Remove(Me.LabelPlaceholder.Name)
        Call Me.Items.Remove(Me.TextBox.Name)
    End If
    On Error GoTo 0
 
    Set Me.LabelPlaceholder = Nothing
    Set Me.TextBox = Nothing
    Set Me.Items = Nothing
End Sub

Private Sub mTextBox_change()
    If Me.LabelPlaceholder Is Nothing Then Exit Sub
    Call UpdatePlaceholder
    Call IsValidInput
End Sub

Private Sub mTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Me.CurrentMaskType
        Case enumTypeMask.tDateFix, enumTypeMask.tTimeFix, enumTypeMask.tOtherFix:
            Call KeyAsciiFixLenText(KeyAscii)
        Case enumTypeMask.tNumeric
            Call NumericValue(KeyAscii)
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

' #####################################################################################################################
' Procedure: KeyAsciiFixLenText
' #####################################################################################################################
' Purpose:
'   Handles key press events for fixed length textboxes with input masks
'   Validates input character against the mask at the current position
'   Automatically appends fixed characters from the mask when appropriate
'
' Parameters:
'   KeyAscii - ASCII code of the pressed key (modified by reference)
' #####################################################################################################################

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
                Call NumericValue(KeyAscii)
            Case "@"
                Call LatinLetters(KeyAscii)
            Case "A"
                Call AlphaNumeric(KeyAscii)
            Case VBA.ChrW$(1041) ' Cyrillic letters
                Call UniCodeLetters(KeyAscii)
            Case VBA.ChrW$(1073) ' Cyrillic letters and digits
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
    haveSimvol = (VBA.InStr(1, Me.ConstSimvolsMasks, sVal, vbBinaryCompare) > 0)
End Function

Private Sub NumericValue(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57:    ' ASCII codes for digits 0-9
        Case 45:
            If Me.IsNegative Then
                KeyAscii = IIf(Me.Value <> vbNullString, 0, 45)
            Else
                KeyAscii = 0
            End If
        Case 44:
            If Me.IsDecimal Then
                KeyAscii = IIf(InStr(1, Me.Value, ",") > 0, 0, 44)
            Else
                KeyAscii = 0
            End If
        Case 46:
            If Me.IsDecimal Then
                KeyAscii = IIf(InStr(1, Me.Value, ",") > 0, 0, 44)
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
            Case VBA.ChrW$(1041):    ' Cyrillic letters
                IsValidMaskInput = IsValidCyrillicLetter(Asc(inputChar))
            Case VBA.ChrW$(1073):    ' Cyrillic letters and digits
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
            If InStr(1, Me.ConstSimvolsMasks, maskChar, vbBinaryCompare) > 0 Then
                IsValidMaskInput = True
                Exit Function
            End If
        Next i
    End If
End Function

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

Private Function IsValidInput() As Boolean
    If mTextBox Is Nothing Then
        IsValidInput = False
        Exit Function
    End If

    Select Case Me.CurrentMaskType
        Case enumTypeMask.tDateFix, enumTypeMask.tTimeFix
            If Me.LenMask = Me.LenValue Then
                If IsDate(Me.Value) Then
                    Me.Value = VBA.Format(VBA.CDate(Me.Value), Me.FormatValue)
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
            If IsValidInput And Me.Mask <> vbNullString Then IsValidInput = IsValidMaskInput()
        Case enumTypeMask.tRegex
            ' For regular expressions, use the appropriate validation
            IsValidInput = IsValidRegexInput()
        Case Else
            IsValidInput = (Me.LenValue = Me.LenMask)
    End Select
    Call ColorContr(IsValidInput)
End Function
Private Sub ColorContr(ByRef Value As Boolean)
    If Me.TextBox Is Nothing Then Exit Sub
    Me.TextBox.BorderColor = IIf(Value, Me.BorderColorValid, Me.BorderColorInvalid)
End Sub

Private Function IsControlInCollection(ByVal controlName As String) As Boolean
    Dim tempItem    As clsTextboxMask
    On Error Resume Next
    Set tempItem = GetItemByName(controlName)
    On Error GoTo 0
    If Not tempItem Is Nothing Then IsControlInCollection = True
End Function

Private Sub KeyAsciiRegex(ByRef KeyAscii As MSForms.ReturnInteger)
    ' Allow backspace key
    If KeyAscii = 8 Then Exit Sub

    ' Initialize regex object if not already done
    If mRegex Is Nothing Then
        Call InitializeRegex
        If mRegex Is Nothing Then Exit Sub
    End If

    ' Check if the new character is valid according to the pattern
    On Error GoTo ErrorHandler
    mRegex.pattern = Me.RegexFilter
    If Not mRegex.Test(VBA.ChrW$(KeyAscii)) Then KeyAscii = 0    ' Reject the character
    Exit Sub

ErrorHandler:
    ' If there's an error with the regex, allow the character to avoid blocking input
    ' This could happen if the regex pattern is invalid
End Sub

Private Sub InitializeRegex()
    On Error GoTo ErrorHandler
    Set mRegex = CreateObject("VBScript.RegExp")
    With mRegex
        .IgnoreCase = False
        .Global = False
    End With
    Exit Sub
ErrorHandler:
    Set mRegex = Nothing
End Sub
' #####################################################################################################################
' Procedure: UpdatePlaceholder
' #####################################################################################################################
' Purpose:
'   Updates the placeholder display depending on the textbox field state
' #####################################################################################################################

Public Sub UpdatePlaceholder()
    If Me.LabelPlaceholder Is Nothing Then Exit Sub
 
    With Me.LabelPlaceholder
        .Caption = processPlaceholderStatus()
        If .Caption <> vbNullString Then .Caption = .Caption & " "
        .Caption = .Caption & processPlaceholderTemplate()
    End With
End Sub

Private Function processPlaceholderStatus() As String
    Dim currentState As String
    If Me.LenValue = 0 Then
        ' Field is empty
        If Me.PlaceholderEmpty <> vbNullString Then currentState = Me.PlaceholderEmpty
    ElseIf Me.LenValue = Me.LenMask Then
        ' Field is completely filled
        If IsValidInput() Then
            If Me.PlaceholderComplete <> vbNullString Then currentState = Me.PlaceholderComplete
        Else
            If Me.PlaceholderInvalid <> vbNullString Then currentState = Me.PlaceholderInvalid
        End If
    ElseIf Me.LenValue < Me.LenMask Then
        ' Field is partially filled
        If IsValidInput() Then
            If Me.PlaceholderPartial <> vbNullString Then currentState = Me.PlaceholderPartial
        Else
            If Me.PlaceholderInvalid <> vbNullString Then currentState = Me.PlaceholderInvalid
        End If
    End If
    processPlaceholderStatus = currentState
End Function

Private Function processPlaceholderTemplate() As String
    Dim result      As String
    result = Me.PlaceHolderTemplate
    If result = vbNullString Then Exit Function

    ' Replace markers with actual values
    result = VBA.Replace(result, "{mask}", Me.Mask)
    result = VBA.Replace(result, "{filled}", Me.LenValue)
    result = VBA.Replace(result, "{remaining}", Me.LenMask - Me.LenValue)
    result = VBA.Replace(result, "{holder}", Me.PlaceholderMask)
    If Me.LenMask > 0 Then
        result = VBA.Replace(result, "{percent}", VBA.Int((Me.LenValue / Me.LenMask) * 100) & "%")
    Else
        result = VBA.Replace(result, "{percent}", "0%")
    End If

    processPlaceholderTemplate = result
End Function

Private Sub class_initialize()
    mSimvolsMasks = "#*@A" & VBA.ChrW$(1041) & VBA.ChrW$(1073) ' Default mask symbols
    mBorderColorValid = &H80000006                             ' Default valid color (usually system window text color)
    mBorderColorInvalid = &HC0C0FF                             ' Default invalid color (light red)
    mPlaceholderColor = &H808080                               ' Default holder color (gray)
End Sub

Private Sub Class_Terminate()
    Set mTextBox = Nothing
    Set mLabelPlaceholder = Nothing
    Set mItems = Nothing
    Set mRegex = Nothing
End Sub
