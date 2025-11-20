# Техническая документация для clsTextboxMask

## Обзор класса

Класс `clsTextboxMask` - это мощный инструмент для VBA, который позволяет создавать текстовые поля с масками ввода в Excel и других приложениях Office. Он обеспечивает валидацию ввода, отображение плейсхолдеров и визуальное указание статуса заполнения поля.

### Основные возможности
- Поддержка различных типов масок ввода (цифры, даты, время, текст, регулярные выражения)
- Валидация ввода в реальном времени
- Отображение плейсхолдеров с различными статусами (пустое, частично заполненное, полностью заполненное, неверное)
- Визуальная индикация корректности ввода через цвет границы
- Поддержка числовых значений с ограничениями по диапазону, знаку и наличию десятичных знаков
- Поддержка переменной длины текста
- Поддержка валидации через регулярные выражения
- Поддержка настройки цвета плейсхолдера в зависимости от статуса поля
- Поддержка шаблонов плейсхолдера с маркерами

## Архитектура класса

### Перечисление enumTypeMask
Определяет типы поддерживаемых масок:
- `tOtherFix` (1) - Фиксированная маска с различными символами
- `tDateFix` (2) - Фиксированная маска для дат
- `tTimeFix` (3) - Фиксированная маска для времени
- `tNumeric` (4) - Числовая маска с возможностью ограничения диапазона
- `tVariableLen` (5) - Маска с переменной длиной
- `tRegex` (6) - Маска на основе регулярных выражений

### Основные свойства класса

| Свойство | Тип | Описание |
|----------|-----|----------|
| `TextBox` | MSForms.TextBox | Ссылка на текстовое поле, к которому применяется маска |
| `LabelPlaceholder` | MSForms.Label | Ссылка на метку плейсхолдера, которая отображает подсказки |
| `Mask` | String | Маска ввода, определяющая допустимые символы |
| `Value` | String | Текущее значение текстового поля |
| `CurrentMaskType` | enumTypeMask | Тип текущей маски |
| `Min` | Single | Минимальное значение для числовых полей |
| `Max` | Single | Максимальное значение для числовых полей |
| `IsDecimal` | Boolean | Разрешены ли десятичные значения |
| `BorderColorValid` | Long | Цвет границы при корректном вводе |
| `BorderColorInvalid` | Long | Цвет границы при некорректном вводе |
| `PlaceholderEmptyColor` | Long | Цвет текста плейсхолдера для пустого поля |
| `PlaceholderPartialColor` | Long | Цвет текста плейсхолдера для частично заполненного поля |
| `PlaceholderCompleteColor` | Long | Цвет текста плейсхолдера для полностью заполненного поля |
| `PlaceholderInvalidColor` | Long | Цвет текста плейсхолдера для поля с некорректными данными |
| `PlaceholderEmpty` | String | Текст плейсхолдера для пустого поля |
| `PlaceholderPartial` | String | Текст плейсхолдера для частично заполненного поля |
| `PlaceholderComplete` | String | Текст плейсхолдера для полностью заполненного поля |
| `PlaceholderInvalid` | String | Текст плейсхолдера для поля с некорректными данными |

## Подробное описание методов

### AddFieldNumeric
Добавляет числовое поле с заданными параметрами валидации.

**Синтаксис:**
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

**Параметры:**
- `inputTextBox` - текстовое поле, к которому применяется маска
- `minValue` - минимальное допустимое значение
- `maxValue` - максимальное допустимое значение
- `allowDecimal` - разрешение на ввод десятичных значений
- `showPlaceholder` - отображение плейсхолдера (опционально)
- `numberFormat` - формат отображения числа (опционально)
- `BorderColorValid` - цвет границы при корректном вводе (опционально)
- `BorderColorInvalid` - цвет границы при некорректном вводе (опционально)
- `PlaceholderEmptyColor` - цвет плейсхолдера для пустого поля (опционально)
- `PlaceholderEmpty` - текст плейсхолдера для пустого поля (опционально)
- `PlaceholderPartialColor` - цвет плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderPartial` - текст плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderCompleteColor` - цвет плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderComplete` - текст плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderInvalidColor` - цвет плейсхолдера для поля с некорректными данными (опционально)
- `PlaceholderInvalid` - текст плейсхолдера для поля с некорректными данными (опционально)
- `PlaceHolderTemplate` - шаблон плейсхолдера (опционально)

### AddFieldDate
Добавляет поле ввода даты с заданными параметрами валидации.

**Синтаксис:**
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

**Параметры:**
- `inputTextBox` - текстовое поле, к которому применяется маска
- `dateMask` - маска ввода даты
- `minDate` - минимальная допустимая дата
- `maxDate` - максимальная допустимая дата
- `dateFormat` - формат отображения даты (опционально)
- `showPlaceholder` - отображение плейсхолдера (опционально)
- `BorderColorValid` - цвет границы при корректном вводе (опционально)
- `BorderColorInvalid` - цвет границы при некорректном вводе (опционально)
- `PlaceholderEmptyColor` - цвет плейсхолдера для пустого поля (опционально)
- `PlaceholderEmpty` - текст плейсхолдера для пустого поля (опционально)
- `PlaceholderPartialColor` - цвет плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderPartial` - текст плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderCompleteColor` - цвет плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderComplete` - текст плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderInvalidColor` - цвет плейсхолдера для поля с некорректными данными (опционально)
- `PlaceholderInvalid` - текст плейсхолдера для поля с некорректными данными (опционально)
- `PlaceHolderTemplate` - шаблон плейсхолдера (опционально)

### AddFieldTime
Добавляет поле ввода времени с заданными параметрами валидации.

**Синтаксис:**
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

**Параметры:**
- `inputTextBox` - текстовое поле, к которому применяется маска
- `timeMask` - маска ввода времени
- `minTime` - минимальное допустимое время
- `maxTime` - максимальное допустимое время
- `timeFormat` - формат отображения времени (опционально)
- `showPlaceholder` - отображение плейсхолдера (опционально)
- `BorderColorValid` - цвет границы при корректном вводе (опционально)
- `BorderColorInvalid` - цвет границы при некорректном вводе (опционально)
- `PlaceholderEmptyColor` - цвет плейсхолдера для пустого поля (опционально)
- `PlaceholderEmpty` - текст плейсхолдера для пустого поля (опционально)
- `PlaceholderPartialColor` - цвет плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderPartial` - текст плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderCompleteColor` - цвет плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderComplete` - текст плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderInvalidColor` - цвет плейсхолдера для поля с некорректными данными (опционально)
- `PlaceholderInvalid` - текст плейсхолдера для поля с некорректными данными (опционально)
- `PlaceHolderTemplate` - шаблон плейсхолдера (опционально)

### AddFieldText
Добавляет текстовое поле с заданной маской ввода.

**Синтаксис:**
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

**Параметры:**
- `inputTextBox` - текстовое поле, к которому применяется маска
- `textMask` - маска ввода текста
- `showPlaceholder` - отображение плейсхолдера (опционально)
- `BorderColorValid` - цвет границы при корректном вводе (опционально)
- `BorderColorInvalid` - цвет границы при некорректном вводе (опционально)
- `PlaceholderEmptyColor` - цвет плейсхолдера для пустого поля (опционально)
- `PlaceholderEmpty` - текст плейсхолдера для пустого поля (опционально)
- `PlaceholderPartialColor` - цвет плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderPartial` - текст плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderCompleteColor` - цвет плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderComplete` - текст плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderInvalidColor` - цвет плейсхолдера для поля с некорректными данными (опционально)
- `PlaceholderInvalid` - текст плейсхолдера для поля с некорректными данными (опционально)
- `PlaceHolderTemplate` - шаблон плейсхолдера (опционально)

### AddFieldVariableLength
Добавляет поле с переменной длиной текста.

**Синтаксис:**
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

**Параметры:**
- `inputTextBox` - текстовое поле, к которому применяется маска
- `maxLength` - максимальная длина текста
- `textMask` - маска ввода текста (опционально)
- `showPlaceholder` - отображение плейсхолдера (опционально)
- `BorderColorValid` - цвет границы при корректном вводе (опционально)
- `BorderColorInvalid` - цвет границы при некорректном вводе (опционально)
- `PlaceholderEmptyColor` - цвет плейсхолдера для пустого поля (опционально)
- `PlaceholderEmpty` - текст плейсхолдера для пустого поля (опционально)
- `PlaceholderPartialColor` - цвет плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderPartial` - текст плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderCompleteColor` - цвет плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderComplete` - текст плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderInvalidColor` - цвет плейсхолдера для поля с некорректными данными (опционально)
- `PlaceholderInvalid` - текст плейсхолдера для поля с некорректными данными (опционально)
- `PlaceHolderTemplate` - шаблон плейсхолдера (опционально)

### AddFieldRegex
Добавляет поле с валидацией через регулярное выражение.

**Синтаксис:**
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

**Параметры:**
- `inputTextBox` - текстовое поле, к которому применяется маска
- `RegexPattern` - паттерн регулярного выражения для валидации
- `RegexFilter` - фильтр регулярного выражения
- `showPlaceholder` - отображение плейсхолдера (опционально)
- `BorderColorValid` - цвет границы при корректном вводе (опционально)
- `BorderColorInvalid` - цвет границы при некорректном вводе (опционально)
- `PlaceholderEmptyColor` - цвет плейсхолдера для пустого поля (опционально)
- `PlaceholderEmpty` - текст плейсхолдера для пустого поля (опционально)
- `PlaceholderPartialColor` - цвет плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderPartial` - текст плейсхолдера для частично заполненного поля (опционально)
- `PlaceholderCompleteColor` - цвет плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderComplete` - текст плейсхолдера для полностью заполненного поля (опционально)
- `PlaceholderInvalidColor` - цвет плейсхолдера для поля с некорректными данными (опционально)
- `PlaceholderInvalid` - текст плейсхолдера для поля с некорректными данными (опционально)
- `PlaceHolderTemplate` - шаблон плейсхолдера (опционально)

### IsValid
Проверяет корректность введенных данных в текстовом поле.

**Синтаксис:**
```vba
Public Function IsValid() As Boolean
```

**Возвращаемое значение:**
- `Boolean` - True, если данные корректны, иначе False

### Clear
Очищает текстовое поле.

**Синтаксис:**
```vba
Public Sub Clear()
```

### SetFocus
Устанавливает фокус на текстовое поле.

**Синтаксис:**
```vba
Public Sub SetFocus()
```

### RemoveItem
Удаляет элемент маски текстового поля и связанные с ним компоненты.

**Синтаксис:**
```vba
Public Sub RemoveItem()
```

## Символы маски

При создании текстовых масок используются следующие символы:

| Символ | Описание |
|--------|----------|
| `#` | Цифры (0-9) |
| `@` | Латинские буквы (A-Z, a-z) |
| `A` | Латинские буквы и цифры (A-Z, a-z, 0-9) |
| `Б` | Кирилические буквы |
| `б` | Кириллические буквы и цифры |
| `*` | Любые символы |

## Шаблоны плейсхолдеров

Класс поддерживает использование маркеров в шаблонах плейсхолдеров:
- `{mask}` - отображает маску
- `{filled}` - количество заполненных символов
- `{remaining}` - количество оставшихся символов
- `{holder}` - плейсхолдер с маской
- `{RegexPattern}` - паттерн регулярного выражения
- `{RegexFilter}` - фильтр регулярного выражения
- `{percent}` - процент заполнения

## Обработка событий

Класс автоматически обрабатывает события текстового поля:
- `Change` - обновляет плейсхолдер и проверяет валидность
- `KeyPress` - контролирует вводимые символы в соответствии с маской

## Примеры использования

### 1. Числовое поле с ограничениями:
```vba
Dim numField As New clsTextboxMask
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True)
```

### 2. Поле даты:
```vba
Dim dateField As New clsTextboxMask
Call dateField.AddFieldDate(inputTextBox:=Me.TextBox2, _
                           dateMask:="##.##.####", _
                           minDate:=#1/1/2020#, _
                           maxDate:=#12/31/2030#, _
                           dateFormat:="dd.mm.yyyy")
```

### 3. Текстовое поле с маской:
```vba
Dim textField As New clsTextboxMask
Call textField.AddFieldText(inputTextBox:=Me.TextBox3, _
                           textMask:="+7(*##) @# A# #Б#")  ' Буквы-цифры
```

### 4. Поле с регулярным выражением:
```vba
Dim regexField As New clsTextboxMask
Call regexField.AddFieldRegex(inputTextBox:=Me.TextBox6, _
                             RegexPattern:="^[A-Z]{2}\d{4}$", _
                             RegexFilter:="[A-Z0-9]")
```

## Внутренняя реализация

Класс использует коллекцию Items для хранения всех созданных элементов маски. Каждый элемент связан с текстовым полем и имеет свои собственные настройки валидации и отображения.

При инициализации класса устанавливаются значения по умолчанию:
- `mSimvolsMasks = "#*@A" & VBA.ChrW$(1041) & VBA.ChrW$(1073)` - символы маски
- `mBorderColorValid = &H800008` - цвет границы при корректном вводе
- `mBorderColorInvalid = &HC0C0FF` - цвет границы при некорректном вводе
- Цвета плейсхолдеров для разных статусов

## Зависимости

- MSForms.TextBox
- MSForms.Label
- VBScript.RegExp (для валидации через регулярные выражения)