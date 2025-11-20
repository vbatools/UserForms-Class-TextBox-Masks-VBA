# Руководство для разработчиков clsTextboxMask

## Введение

Это руководство предназначено для разработчиков, которые хотят понять внутреннюю архитектуру класса `clsTextboxMask`, модифицировать его или создавать расширения. В документе рассматриваются внутренние механизмы работы класса, его структура и рекомендации по расширению функциональности.

## Архитектура класса

### Структура класса

Класс `clsTextboxMask` построен по принципу управления коллекцией элементов маски. Основные компоненты:

- **Основные свойства**: Хранят настройки для конкретного текстового поля
- **Коллекция Items**: Управляет несколькими элементами маски
- **Методы обработки событий**: Обрабатывают ввод данных и обновляют состояние полей
- **Методы валидации**: Проверяют корректность введенных данных

### Основные внутренние переменные

```vba
Private mSimvolsMasks As String              ' Символы маски
Private mBorderColorValid As Long            ' Цвет границы при корректном вводе
Private mBorderColorInvalid As Long          ' Цвет границы при некорректном вводе
Private WithEvents mTextBox As MSForms.TextBox ' Текстовое поле с обработкой событий
Private mLabelPlaceholder As MSForms.Label    ' Метка плейсхолдера
Private mItems As Collection                  ' Коллекция элементов маски
Private mMask As String                       ' Маска ввода
Private mFormatValue As String                ' Формат значения
Private mRegexPattern As String               ' Паттерн регулярного выражения
Private mRegexFilter As String                ' Фильтр регулярного выражения
Private mRegex As Object                      ' Объект регулярного выражения
```

## Внутренние механизмы

### Обработка событий

Класс использует `WithEvents` для отслеживания изменений в текстовом поле:

```vba
Private WithEvents mTextBox As MSForms.TextBox
```

События, которые обрабатываются:
- `mTextBox_change` - обновляет плейсхолдер и проверяет валидность
- `mTextBox_KeyPress` - контролирует вводимые символы

### Алгоритм обработки нажатия клавиш

В методе `mTextBox_KeyPress` реализована логика проверки вводимых символов:

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

### Валидация данных

Метод `IsValidInput()` проверяет корректность введенных данных в зависимости от типа маски:

- Для числовых полей проверяется диапазон значений
- Для дат проверяется корректность формата и диапазон
- Для текстовых масок проверяется соответствие символам маски
- Для регулярных выражений используется объект RegExp

## Расширение функциональности

### Добавление новых типов масок

Для добавления нового типа маски необходимо:

1. Добавить значение в перечисление `enumTypeMask`:

```vba
Public Enum enumTypeMask
    tOtherFix = 1
    tDateFix
    tTimeFix
    tNumeric
    tVariableLen
    tRegex
    tNewType  ' Новый тип маски
    [_First] = tOtherFix
    [_Last] = tNewType
End Enum
```

2. Обновить обработку в методе `mTextBox_KeyPress`:

```vba
Private Sub mTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Me.CurrentMaskType
        ' ... существующие случаи ...
        Case enumTypeMask.tNewType
            Call KeyAsciiNewType(KeyAscii)
    End Select
End Sub
```

3. Создать метод для обработки нового типа:

```vba
Private Sub KeyAsciiNewType(ByRef KeyAscii As MSForms.ReturnInteger)
    ' Логика обработки нового типа маски
End Sub
```

4. Обновить метод валидации:

```vba
Private Function IsValidInput() As Boolean
    ' ... существующая логика ...
    Select Case Me.CurrentMaskType
        ' ... существующие случаи ...
        Case enumTypeMask.tNewType
            ' Валидация для нового типа
    End Select
End Function
```

### Создание пользовательских символов маски

Для добавления новых символов маски:

1. Обновить константу `mSimvolsMasks` в `class_initialize`:

```vba
Private Sub class_initialize()
    mSimvolsMasks = "#*@A" & VBA.ChrW$(1041) & VBA.ChrW$(1073) & "N"  ' Добавляем символ "N"
    ' ... остальные настройки ...
End Sub
```

2. Добавить обработку нового символа в `KeyAsciiFixLenText`:

```vba
Private Sub KeyAsciiFixLenText(ByRef KeyAscii As MSForms.ReturnInteger)
    ' ... существующая логика ...
    Select Case endLetter
        ' ... существующие случаи ...
        Case "N"  ' Новый символ маски
            ' Обработка символа N
    End Select
End Sub
```

## Настройка производительности

### Оптимизация обработки событий

Для повышения производительности при работе с большим количеством полей:

1. Используйте оптимизированные методы проверки:

```vba
' Вместо многократного вызова функций, кэшируйте значения
Private Function IsValidInput() As Boolean
    Static cachedValue As String
    Static cachedResult As Boolean
    
    If Me.Value <> cachedValue Then
        cachedValue = Me.Value
        ' Выполнить проверку и сохранить результат
        cachedResult = PerformValidation()
    End If
    
    IsValidInput = cachedResult
End Function
```

2. Ограничьте частоту обновления плейсхолдеров:

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

## Обработка ошибок

### Внутренние обработчики ошибок

Класс использует несколько подходов для обработки ошибок:

1. Проверка параметров в методе `AddField`:

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

2. Обработка ошибок при работе с регулярными выражениями:

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

## Тестирование и отладка

### Модульное тестирование

Для тестирования функциональности класса рекомендуется создать тестовые сценарии:

```vba
' Модуль тестирования: modTextboxMaskTests
Sub TestNumericField()
    Dim mask As New clsTextboxMask
    Dim tb As MSForms.TextBox
    Set tb = CreateTestTextBox()
    
    Call mask.AddFieldNumeric(tb, 0, 100, True)
    
    ' Тестируем ввод допустимого значения
    tb.Value = "50.5"
    Debug.Assert mask.IsValid = True, "Допустимое значение должно быть принято"
    
    ' Тестируем ввод недопустимого значения
    tb.Value = "150"
    Debug.Assert mask.IsValid = False, "Недопустимое значение должно быть отклонено"
    
    Debug.Print "Тест числового поля пройден"
End Sub

Sub TestTextField()
    Dim mask As New clsTextboxMask
    Dim tb As MSForms.TextBox
    Set tb = CreateTestTextBox()
    
    Call mask.AddFieldText(tb, "###-###")
    
    ' Тестируем ввод допустимого значения
    tb.Value = "123-456"
    Debug.Assert mask.IsValid = True, "Допустимое значение должно быть принято"
    
    ' Тестируем ввод недопустимого значения
    tb.Value = "123-abc"
    Debug.Assert mask.IsValid = False, "Недопустимое значение должно быть отклонено"
    
    Debug.Print "Тест текстового поля пройден"
End Sub

Private Function CreateTestTextBox() As MSForms.TextBox
    ' Создание тестового текстового поля
    Dim tb As MSForms.TextBox
    Set tb = New MSForms.TextBox
    Set CreateTestTextBox = tb
End Function
```

### Отладка валидации

Для отладки процесса валидации можно добавить логирование:

```vba
Private Function IsValidInput() As Boolean
    ' Логирование для отладки
    #If DEBUG_MODE Then
        Debug.Print "Проверка валидации: " & Me.TextBox.Name
        Debug.Print "Значение: " & Me.Value
        Debug.Print "Тип маски: " & Me.CurrentMaskType
    #End If
    
    ' Основная логика валидации
    Select Case Me.CurrentMaskType
        ' ... основная логика ...
    End Select
End Function
```

## Рекомендации по использованию

### Лучшие практики

1. **Инициализация в нужное время**:
   - Инициализируйте маски в событии `UserForm_Initialize`, а не в `Activate`
   - Убедитесь, что все текстовые поля существуют перед добавлением масок

2. **Управление памятью**:
   - Используйте метод `RemoveItem` для удаления элементов маски
   - Освобождайте ссылки на объекты при закрытии формы:

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

3. **Обработка исключений**:
   - Оборачивайте вызовы методов в блоки обработки ошибок
   - Проверяйте, что элементы существуют перед обращением к ним

### Рекомендации по производительности

1. **Оптимизация при большом количестве полей**:
   - Используйте одну коллекцию для нескольких полей
   - Избегайте частого обновления плейсхолдеров

2. **Эффективное использование регулярных выражений**:
   - Кэшируйте объекты RegExp
   - Избегайте сложных регулярных выражений в реальном времени

## Расширение класса

### Создание производных классов

Можно создать специализированные классы на основе `clsTextboxMask`:

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

' Делегирование свойств
Public Property Get IsValid() As Boolean
    IsValid = baseMask.IsValid
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = baseMask.TextBox
End Property

' ... другие свойства и методы
```

### Добавление пользовательских валидаторов

Для добавления пользовательских функций валидации:

```vba
' Добавляем делегат для пользовательской валидации
Public Type ValidationDelegate
    ValidateProc As String  ' Имя процедуры валидации
End Type

Private customValidator As ValidationDelegate

Public Sub SetCustomValidator(validatorName As String)
    customValidator.ValidateProc = validatorName
End Sub

Private Function IsValidInput() As Boolean
    ' ... стандартная валидация ...
    
    ' Пользовательская валидация
    If customValidator.ValidateProc <> "" Then
        IsValidInput = Application.Run(customValidator.ValidateProc, Me.Value)
    End If
End Function
```

## Заключение

Класс `clsTextboxMask` представляет собой гибкую и расширяемую архитектуру для создания валидированных текстовых полей в VBA. Его модульная структура позволяет легко добавлять новые типы масок и функции валидации.

При разработке расширений класса рекомендуется придерживаться следующих принципов:
- Поддерживать совместимость с существующим API
- Обеспечивать корректную обработку ошибок
- Учитывать производительность при работе с большим количеством полей
- Писать тесты для новых функций

Эти рекомендации помогут вам эффективно использовать и расширять функциональность класса в ваших проектах.