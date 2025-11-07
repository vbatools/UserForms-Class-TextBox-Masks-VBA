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
' Version:       1.0.2
' Creation Date: 11.05.2025 10:30
' Update Date:   11.05.2025 22:05
' License:       Apache License
' Description:   Класс для создания текстовых полей с масками ввода в VBA.
'              Поддерживает различные типы масок: цифровые, буквенные, даты,
'              телефонные номера и т.д. Обеспечивает проверку ввода, отображение
'              заполнителя и визуальное обозначение состояния заполнения поля.
'
'              Класс позволяет создавать валидированные текстовые поля с заданными
'              ограничениями на вводимые данные. Поддерживает следующие типы масок:
'              - Цифровые (с возможностью задания диапазона, знака, дробности)
'              - Даты (с проверкой корректности даты и диапазона)
'              - Время (с проверкой корректности времени)
'              - Текстовые с фиксированной длиной (с различными типами символов)
'
'              Каждое поле с маской имеет визуальные индикаторы: цвет границы
'              изменяется в зависимости от валидности введенных данных, а также
'              отображается подсказка-заполнитель с ожидаемым форматом.
' ======================================================================================
' Enum:          enumTypeMask
' Description:   Перечисление типов масок, поддерживаемых классом
' Values:
'   tOtherFix - Текстовая маска с фиксированной длиной (буквы, цифры, специальные символы)
'   tDateFix  - Маска для ввода даты с фиксированной длиной
'   tTimeFix  - Маска для ввода времени с фиксированной длиной
'   tNumeric  - Цифровая маска с возможностью задания диапазона, знака и дробности
'   [_First]  - Вспомогательное значение для определения начала перечисления
'   [_Last]   - Вспомогательное значение для определения конца перечисления
Public Enum enumTypeMask
    tOtherFix = 1
    tDateFix
    tTimeFix
    tNumeric
    [_First] = tOtherFix
    [_Last] = tNumeric
End Enum

Private Const SIMVOLS_MASKS As String = "#*@A��"
Private Const COLOR_DEFAULT_ON As Long = &H80000006
Private Const COLOR_DEFAULT_OFF As Long = &HC0C0FF
Private Const COLOR_DEFAULT_HOLDER As Long = &H808080

Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private mLabelHolder As MSForms.Label

Private mItems      As Collection

Private mMask       As String
Private mFormatValue As String

Private mBorderColorOn As Long
Private mBorderColorOff As Long

Private mIsDecemal  As Boolean
Private mIsNegative As Boolean

Private mMin        As Single
Private mMax        As Single

Private mCurrentMaskType As enumTypeMask

' Property:      constSimvolsMasks
' Description:   Возвращает строку символами масок, используемых в классе
' Returns:       Строка "#*@A" содержащая символы масок:
'                # - цифры, @ - латинские буквы, * - любые символы,
'                A - латинские буквы и цифры,  - кириллические буквы,
'                - кириллические буквы и цифры
Public Property Get constSimvolsMasks() As String
    constSimvolsMasks = SIMVOLS_MASKS
End Property

' Property:      Min
' Description:   Свойство для получения и установки минимального значения
'                для числовых масок. Используется при валидации ввода.
' Returns/Parameters:
'   Value - Значение минимального допустимого числа для числовых масок
Public Property Get Min() As Single
    Min = mMin
End Property

Public Property Let Min(ByVal Value As Single)
    mMin = Value
End Property

' Property:      Max
' Description:   Свойство для получения и установки максимального значения
'                для числовых масок. Используется при валидации ввода.
' Returns/Parameters:
'   Value - Значение максимального допустимого числа для числовых масок
Public Property Get Max() As Single
    Max = mMax
End Property

Public Property Let Max(ByVal Value As Single)
    mMax = Value
End Property

' Property:      visibleLabelHolder
' Description:   Свойство для получения и установки видимости подсказки-заполнителя
'                для текстового поля. Подсказка отображает маску ввода.
' Returns/Parameters:
'   Value - Логическое значение (True/False), определяющее видимость подсказки
Public Property Get visibleLabelHolder() As Boolean
    If Me.LabelHolder Is Nothing Then Exit Property
    visibleLabelHolder = Me.LabelHolder.Visible
End Property

Public Property Let visibleLabelHolder(ByVal Value As Boolean)
    If Me.LabelHolder Is Nothing Then Exit Property
    Me.LabelHolder.Visible = Value
End Property

' Property:      IsNegative
' Description:   Свойство для получения и установки разрешения ввода отрицательных
'                значений для числовых масок.
' Returns/Parameters:
'   Value - Логическое значение (True/False), разрешающее ввод отрицательных чисел
Public Property Get IsNegative() As Boolean
    IsNegative = mIsNegative
End Property

Public Property Let IsNegative(ByVal Value As Boolean)
    mIsNegative = Value
End Property

' Property:      IsDecemal
' Description:   Свойство для получения и установки разрешения ввода десятичных
'                (дробных) значений для числовых масок.
' Returns/Parameters:
'   Value - Логическое значение (True/False), разрешающее ввод дробных чисел
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
' Description:   Свойство для получения строки-заполнителя, которая отображается
'                в подсказке до тех пор, пока текстовое поле не будет полностью заполнено
' Returns:       Строка, содержащая текущее значение поля плюс оставшуюся часть маски
Public Property Get PlaceHolder() As String
    If Me.TextBox Is Nothing Then Exit Property
    If Me.LenMask > Me.LenValue Then PlaceHolder = Me.Value & VBA.Mid$(Me.Mask, Me.LenValue + 1, Me.LenMask - Me.LenValue)
End Property

' Property:      CurrentMaskType
' Description:   Свойство для получения и установки текущего типа маски ввода
'                для данного экземпляра класса
' Returns/Parameters:
'   maskType - Значение из перечисления enumTypeMask, определяющее тип маски
Public Property Get CurrentMaskType() As enumTypeMask
    CurrentMaskType = mCurrentMaskType
End Property

Public Property Let CurrentMaskType(ByVal maskType As enumTypeMask)
    mCurrentMaskType = maskType
End Property

' Property:      LenValue
' Description:   Свойство для получения текущей длины введенного значения
'                в текстовом поле
' Returns:       Целое число, представляющее количество символов в текущем значении
Public Property Get LenValue() As Integer
    If Me.TextBox Is Nothing Then Exit Property
    LenValue = VBA.Len(Me.Value)
End Property

' Property:      Value
' Description:   Свойство для получения и установки значения текстового поля
' Returns/Parameters:
'   Value - Строка, представляющая значение текстового поля
Public Property Let Value(ByRef Value As String)
    If Me.TextBox Is Nothing Then Exit Property
    Me.TextBox.Value = Value
End Property

Public Property Get Value() As String
    If Me.TextBox Is Nothing Then Exit Property
    Value = Me.TextBox.Value
End Property

' Property:      Mask
' Description:   Свойство для получения и установки маски ввода для текстового поля
' Returns/Parameters:
'   Mask - Строка, определяющая маску ввода с использованием специальных символов:
'          # (цифры), @ (латинские буквы), * (любые символы), A (латинские буквы и цифры),
'          (кириллические буквы),  (кириллические буквы и цифры)
Public Property Let Mask(ByRef Mask As String)
    mMask = Mask
End Property

Public Property Get Mask() As String
    Mask = mMask
End Property

' Property:      LenMask
' Description:   Свойство для получения длины маски ввода
' Returns:       Целое число, представляющее количество символов в маске
Public Property Get LenMask() As Integer
    LenMask = VBA.Len(Me.Mask)
End Property

' Property:      formatValue
' Description:   Свойство для получения и установки формата отображения значения
'                после ввода (используется для дат и чисел)
' Returns/Parameters:
'   formatValue - Строка формата, определяющая, как будет отображаться значение
Public Property Let formatValue(ByRef formatValue As String)
    mFormatValue = formatValue
End Property

Public Property Get formatValue() As String
    formatValue = mFormatValue
End Property

' Property:      borderColorOn
' Description:   Свойство для получения и установки цвета границы текстового поля
'                при валидном вводе
' Returns/Parameters:
'   Color - Значение цвета (Long), используемое для границы при валидном вводе
Public Property Let borderColorOn(ByRef Color As Long)
    mBorderColorOn = Color
End Property

Public Property Get borderColorOn() As Long
    borderColorOn = mBorderColorOn
End Property

' Property:      borderColorOff
' Description:   Свойство для получения и установки цвета границы текстового поля
'                при невалидном вводе
' Returns/Parameters:
'   Color - Значение цвета (Long), используемое для границы при невалидном вводе
Public Property Let borderColorOff(ByRef Color As Long)
    mBorderColorOff = Color
End Property

Public Property Get borderColorOff() As Long
    borderColorOff = mBorderColorOff
End Property

' Property:      Count
' Description:   Свойство для получения количества элементов в коллекции
'                масок ввода
' Returns:       Байт, представляющий количество элементов в коллекции Items
Public Property Get Count() As Byte
    Count = Me.Items.Count
End Property

' Property:      RemainingChars
' Description:   Свойство для получения количества оставшихся символов
'                до полного заполнения текстового поля по маске
' Returns:       Целое число, представляющее количество оставшихся символов
Public Property Get RemainingChars() As Integer
    If Me.TextBox Is Nothing Then
        RemainingChars = Me.LenMask
        Exit Property
    End If
    RemainingChars = Me.LenMask - Me.LenValue
End Property

Public Property Get getItemByIndex(ByVal index As Integer) As clsTextboxMask
    On Error GoTo endGetItem
    Set getItemByIndex = Me.Items(index)
    Exit Property
endGetItem:
    Err.Clear
End Property

' Property: getItemByName
' Purpose: Gets an item from the collection by name
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
            "Version: 1.0.2" & vbNewLine & _
            "Date of creation: 05.11.2025 10:30" & vbNewLine & _
            "Date of update: 05.11.2025 22:05" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: Class for creating textboxes with input masks in VBA" & vbNewLine & _
            "Supports various mask types: digital, letters, dates, phone numbers, etc." & vbNewLine & _
            "Provides input validation, placeholder display, and visual indication" & vbNewLine & _
            "of field fill status."
End Property

' Method:        addItemNumeric
' Description:   Добавляет новый элемент с числовым типом маски в коллекцию
' Parameters:
'   TextBox         - Объект текстового поля для применения маски
'   snMin           - Минимальное допустимое значение
'   snMax           - Максимальное допустимое значение
'   IsDecemal       - Флаг разрешения ввода дробных чисел
'   IsNegative      - Флаг разрешения ввода отрицательных чисел
'   visibleLabelHolder - (Опционально) Флаг видимости подсказки-заполнителя (по умолчанию True)
'   formatNumeric   - (Опционально) Формат отображения числа (по умолчанию "#.0")
'   borderColorOn   - (Опционально) Цвет границы при валидном вводе (по умолчанию COLOR_DEFAULT_ON)
'   borderColorOff  - (Опционально) Цвет границы при невалидном вводе (по умолчанию COLOR_DEFAULT_OFF)
'   foreColorHolder - (Опционально) Цвет текста подсказки (по умолчанию COLOR_DEFAULT_HOLDER)
Public Sub addItemNumeric(ByRef TextBox As MSForms.TextBox, _
        ByVal snMin As Single, _
        ByVal snMax As Single, _
        ByVal IsDecemal As Boolean, _
        ByVal IsNegative As Boolean, _
        Optional visibleLabelHolder As Boolean = True, _
        Optional formatNumeric As String = "#.0", _
        Optional borderColorOn As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorOff As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, String$(WorksheetFunction.Max(VBA.Len(snMin), VBA.Len(snMax)), "#"), _
            enumTypeMask.tNumeric, formatNumeric, IsDecemal, IsNegative, borderColorOn, borderColorOff, _
            snMin, snMax, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItemFixLenDate
' Description:   Добавляет новый элемент с маской даты фиксированной длины в коллекцию
' Parameters:
'   TextBox         - Объект текстового поля для применения маски
'   Mask            - Строка маски для даты (например, "##.##.####")
'   minDate         - Минимально допустимая дата
'   maxDate         - Максимально допустимая дата
'   formatDate      - (Опционально) Формат отображения даты (по умолчанию "dd.mm.yyyy")
'   visibleLabelHolder - (Опционально) Флаг видимости подсказки-заполнителя (по умолчанию True)
'   borderColorOn   - (Опционально) Цвет границы при валидном вводе (по умолчанию COLOR_DEFAULT_ON)
'   borderColorOff  - (Опционально) Цвет границы при невалидном вводе (по умолчанию COLOR_DEFAULT_OFF)
'   foreColorHolder - (Опционально) Цвет текста подсказки (по умолчанию COLOR_DEFAULT_HOLDER)
Public Sub addItemFixLenDate(ByRef TextBox As MSForms.TextBox, ByVal Mask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional formatDate As String = "dd.mm.yyyy", _
        Optional visibleLabelHolder As Boolean = True, _
        Optional borderColorOn As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorOff As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, Mask, enumTypeMask.tDateFix, formatDate, False, False, borderColorOn, borderColorOff, _
            minDate, maxDate, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItemFixLenTime
' Description:   Добавляет новый элемент с маской времени фиксированной длины в коллекцию
' Parameters:
'   TextBox         - Объект текстового поля для применения маски
'   Mask            - Строка маски для времени (например, "##:##")
'   minDate         - Минимально допустимое время
'   maxDate         - Максимально допустимое время
'   formatDate      - (Опционально) Формат отображения времени (по умолчанию "hh:mm")
'   visibleLabelHolder - (Опционально) Флаг видимости подсказки-заполнителя (по умолчанию True)
'   borderColorOn   - (Опционально) Цвет границы при валидном вводе (по умолчанию COLOR_DEFAULT_ON)
'   borderColorOff  - (Опционально) Цвет границы при невалидном вводе (по умолчанию COLOR_DEFAULT_OFF)
'   foreColorHolder - (Опционально) Цвет текста подсказки (по умолчанию COLOR_DEFAULT_HOLDER)
Public Sub addItemFixLenTime(ByRef TextBox As MSForms.TextBox, ByVal Mask As String, _
        ByVal minDate As Date, _
        ByVal maxDate As Date, _
        Optional formatDate As String = "hh:mm", _
        Optional visibleLabelHolder As Boolean = True, _
        Optional borderColorOn As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorOff As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, Mask, enumTypeMask.tTimeFix, formatDate, False, False, borderColorOn, borderColorOff, _
            0, 1, visibleLabelHolder, foreColorHolder)
End Sub

' Method:        addItemFixLen
' Description:   Добавляет новый элемент с текстовой маской фиксированной длины в коллекцию
' Parameters:
'   TextBox         - Объект текстового поля для применения маски
'   Mask            - Строка маски для текста (например, "###@@@")
'   visibleLabelHolder - (Опционально) Флаг видимости подсказки-заполнителя (по умолчанию True)
'   borderColorOn   - (Опционально) Цвет границы при валидном вводе (по умолчанию COLOR_DEFAULT_ON)
'   borderColorOff  - (Опционально) Цвет границы при невалидном вводе (по умолчанию COLOR_DEFAULT_OFF)
'   foreColorHolder - (Опционально) Цвет текста подсказки (по умолчанию COLOR_DEFAULT_HOLDER)
Public Sub addItemFixLen(ByRef TextBox As MSForms.TextBox, ByVal Mask As String, _
        Optional visibleLabelHolder As Boolean = True, _
        Optional borderColorOn As XlRgbColor = COLOR_DEFAULT_ON, _
        Optional borderColorOff As XlRgbColor = COLOR_DEFAULT_OFF, _
        Optional foreColorHolder As XlRgbColor = COLOR_DEFAULT_HOLDER)
    Call addItem(TextBox, Mask, enumTypeMask.tOtherFix, vbNullString, False, False, borderColorOn, borderColorOff, _
            0, 0, visibleLabelHolder, foreColorHolder)
End Sub

Private Sub addItem(ByRef TextBox As MSForms.TextBox, _
        ByVal Mask As String, _
        ByVal maskType As enumTypeMask, _
        ByVal formatValue As String, _
        ByVal IsDecemal As Boolean, _
        ByVal IsNegative As Boolean, _
        ByVal borderColorOn As XlRgbColor, _
        ByVal borderColorOff As XlRgbColor, _
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

        .borderColorOn = borderColorOn
        .borderColorOff = borderColorOff

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
' Description:   Проверяет валидность введенного значения в текстовом поле
'                в соответствии с установленной маской
' Returns:       Логическое значение (True/False), указывающее на валидность ввода
Public Function IsValid() As Boolean
    IsValid = IsValidInput()
End Function

' Method:        Clear
' Description:   Очищает значение текстового поля
Public Sub Clear()
    If Me.TextBox Is Nothing Then Exit Sub
    Me.Value = vbNullString
End Sub

' Method:        SetFocus
' Description:   Устанавливает фокус на текстовом поле
Public Sub SetFocus()
    If Me.TextBox Is Nothing Then Exit Sub
    Me.TextBox.SetFocus
End Sub

' Method:        RemoveItem
' Description:   Удаляет элемент маски ввода из коллекции и очищает связанные элементы управления
'                (подсказку-заполнитель и текстовое поле)
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
            Case "�"
                Call UniCodeLetters(KeyAscii)
            Case "�"
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
            .BorderColor = mBorderColorOn
        Else
            .BorderColor = mBorderColorOff
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Function: IsControlInCollection
' Purpose:  Проверяет, существует ли элемент управления с указанным именем в коллекции
'           элементов класса. Используется для предотвращения дублирования элементов
'           при добавлении новых масок ввода.
' Parameters:
'   controlName - Строка, содержащая имя элемента управления для проверки
' Returns: True, если элемент управления с таким именем уже существует в коллекции,
'          False в противном случае
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
