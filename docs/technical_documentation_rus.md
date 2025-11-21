# Техническая документация для clsTextboxMask

## Содержание
1. [Обзор класса](#обзор-класса)
2. [Архитектура класса](#архитектура-класса)
3. [Свойства](#свойства)
4. [Методы](#методы)
5. [События](#события)
6. [Константы и перечисления](#константы-и-перечисления)
7. [Детали реализации](#детали-реализации)
8. [Зависимости](#зависимости)

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

## Свойства

### Основные свойства
- `TextBox` - Получает или устанавливает текстовое поле, к которому применяется маска
- `LabelPlaceholder` - Получает или устанавливает метку плейсхолдера, которая отображает подсказки
- `Mask` - Получает или устанавливает маску ввода, определяющую допустимые символы
- `Value` - Получает или устанавливает текущее значение текстового поля
- `CurrentMaskType` - Получает или устанавливает тип текущей маски
- `Min` - Получает или устанавливает минимальное значение для числовых полей
- `Max` - Получает или устанавливает максимальное значение для числовых полей
- `IsDecimal` - Получает или устанавливает, разрешены ли десятичные значения
- `BorderColorValid` - Получает или устанавливает цвет границы при корректном вводе
- `BorderColorInvalid` - Получает или устанавливает цвет границы при некорректном вводе
- `PlaceholderEmptyColor` - Получает или устанавливает цвет текста плейсхолдера для пустого поля
- `PlaceholderPartialColor` - Получает или устанавливает цвет текста плейсхолдера для частично заполненного поля
- `PlaceholderCompleteColor` - Получает или устанавливает цвет текста плейсхолдера для полностью заполненного поля
- `PlaceholderInvalidColor` - Получает или устанавливает цвет текста плейсхолдера для поля с некорректными данными
- `PlaceholderEmpty` - Получает или устанавливает текст плейсхолдера для пустого поля
- `PlaceholderPartial` - Получает или устанавливает текст плейсхолдера для частично заполненного поля
- `PlaceholderComplete` - Получает или устанавливает текст плейсхолдера для полностью заполненного поля
- `PlaceholderInvalid` - Получает или устанавливает текст плейсхолдера для поля с некорректными данными

### Дополнительные свойства
- `PlaceholderMask` - Получает текущую маску плейсхолдера, показывающую оставшиеся символы для заполнения
- `PlaceHolderTemplate` - Получает или устанавливает шаблон плейсхолдера с маркерами
- `VisibleLabelPlaceholder` - Получает или устанавливает видимость метки плейсхолдера
- `Items` - Получает коллекцию всех элементов маски текстового поля
- `Count` - Получает количество элементов в коллекции
- `RemainingChars` - Получает количество оставшихся символов для заполнения
- `FormatValue` - Получает или устанавливает формат отображения значения
- `LenValue` - Получает длину текущего значения
- `lenMask` - Получает длину маски
- `Version` - Получает информацию о версии класса

## Методы

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

## События

### mTextBox_Change
Событие, вызываемое при изменении значения текстового поля.

### mTextBox_KeyPress
Событие, вызываемое при нажатии клавиши в текстовом поле.

## Константы и перечисления

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

## Детали реализации

### Символы маски
При создании текстовых масок используются следующие символы:

| Символ | Описание |
|--------|----------|
| `#` | Цифры (0-9) |
| `@` | Латинские буквы (A-Z, a-z) |
| `A` | Латинские буквы и цифры (A-Z, a-z, 0-9) |
| `Б` | Кирилические буквы |
| `б` | Кириллические буквы и цифры |
| `*` | Любые символы |

### Шаблоны плейсхолдеров
Класс поддерживает использование маркеров в шаблонах плейсхолдеров:
- `{mask}` - отображает маску
- `{filled}` - количество заполненных символов
- `{remaining}` - количество оставшихся символов
- `{holder}` - плейсхолдер с маской
- `{RegexPattern}` - паттерн регулярного выражения
- `{RegexFilter}` - фильтр регулярного выражения
- `{percent}` - процент заполнения

### Обработка событий
Класс автоматически обрабатывает события текстового поля:
- `Change` - обновляет плейсхолдер и проверяет валидность
- `KeyPress` - контролирует вводимые символы в соответствии с маской

## Зависимости

- MSForms.TextBox
- MSForms.Label
- VBScript.RegExp (для валидации через регулярные выражения)