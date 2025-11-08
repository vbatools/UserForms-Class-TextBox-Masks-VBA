# UserForms-Class-TextBox-Masks

![User Forms Example](User_Forms.png)

## Описание
Этот проект содержит класс `clsTextboxMask` для Microsoft VBA, который позволяет создавать текстовые поля с масками ввода в UserForms. Класс обеспечивает проверку ввода, отображение заполнителя и визуальные индикаторы состояния заполнения поля.

## Возможности
- Поддержка различных типов масок ввода:
  - Числовые маски (с диапазоном, знаком и опцией десятичных чисел)
  - Маски дат (с проверкой даты и проверкой диапазона)
  - Маски времени (с проверкой времени)
  - Маски текста фиксированной длины (с различными типами символов)
  - Маски текста переменной длины (с дополнительной проверкой шаблона)
  - Маски на основе регулярных выражений (пользовательские шаблоны проверки)
- Визуальные индикаторы проверки (изменение цвета границы в зависимости от корректности ввода)
- Отображение подсказки-заполнителя с ожидаемым форматом
- Поддержка различных типов символов в масках:
  - `#` - цифры
  - `@` - латинские буквы
  - `*` - любые символы
  - `A` - латинские буквы и цифры
  - `Б` - кириллические буквы
  - `б` - кириллические буквы и цифры

## Установка
1. Скопируйте файл `clsTextboxMask.cls` ваш VBA проект
2. Импортируйте его в редактор VBA (например, в Excel или Word)

## Использование
Класс предоставляет несколько методов для добавления различных типов масок:

### Numeric Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox1, 0, 100, True, False)
```

Using named arguments (walrus operator equivalent in VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox:=TextBox1, snMin:=0, snMax:=100, IsDecemal:=True, IsNegative:=False)
```

Using all named arguments with optional parameters:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox:=TextBox1, snMin:=0, snMax:=100, IsDecemal:=True, IsNegative:=False, _
                                visibleLabelHolder:=True, formatNumeric:="#.0", borderColorValid:=&H8000006, _
                                borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

### Date Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox1, "##.##.####", #1/1/2000#, #12/31/2030#, "dd.mm.yyyy")
```

Using named arguments (walrus operator equivalent in VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox:=TextBox1, Mask:="##.##.####", minDate:=#1/1/2000#, maxDate:=#12/31/2030#, formatDate:="dd.mm.yyyy")
```

Using all named arguments with optional parameters:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox:=TextBox1, Mask:="##.##.####", minDate:=#1/1/2000#, maxDate:=#12/31/2030#, _
                                   formatDate:="dd.mm.yyyy", visibleLabelHolder:=True, borderColorValid:=&H8000006, _
                                   borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

### Time Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox1, "##:##", #0:00:00#, #23:59:59#, "hh:mm")
```

Using named arguments (walrus operator equivalent in VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox:=TextBox1, Mask:="##:##", minDate:=#0:00:00#, maxDate:=#23:59:59#, formatDate:="hh:mm")
```

Using all named arguments with optional parameters:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox:=TextBox1, Mask:="##:##", minDate:=#0:00:00#, maxDate:=#23:59:59#, _
                                   formatDate:="hh:mm", visibleLabelHolder:=True, borderColorValid:=&H8000006, _
                                   borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

### Fixed-Length Text Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox1, "###@@@")
```

Using named arguments (walrus operator equivalent in VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox:=TextBox1, Mask:="###@@@")
```

Using all named arguments with optional parameters:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox:=TextBox1, Mask:="###@@@", visibleLabelHolder:=True, _
                               borderColorValid:=&H8000006, borderColorNoValid:=&HC0C0FF, _
                               foreColorHolder:=&H808080)
```

### Variable-length Text Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemVariableLen(TextBox1, 20, "###@@@")
```

Using named arguments (walrus operator equivalent in VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemVariableLen(TextBox:=TextBox1, maxLength:=20, textMask:="###@@@")
```

Using all named arguments with optional parameters:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemVariableLen(TextBox:=TextBox1, maxLength:=20, textMask:="###@@@", _
                                  visibleLabelHolder:=True, borderColorValid:=&H8000006, _
                                  borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

### Regular Expression-based Mask
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemRegex(TextBox1, "^[A-Z]{3}\d{3}$")
```

Using named arguments (walrus operator equivalent in VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemRegex(TextBox:=TextBox1, regexPattern:="^[A-Z]{3}\d{3}$")
```

Using all named arguments with optional parameters:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemRegex(TextBox:=TextBox1, regexPattern:="^[A-Z]{3}\d{3}$", _
                            visibleLabelHolder:=True, borderColorValid:=&H8000006, _
                            borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

## Параметры
- `TextBox` - объект текстового поля, к которому применяется маска
- `Mask` - строка маски ввода
- `Min/Max` - минимальное и максимальное разрешенные значения (для числовых и датовых масок)
- `IsDecimal` - разрешить ввод десятичных чисел
- `IsNegative` - разрешить ввод отрицательных чисел
- `formatValue` - формат отображения значения (для дат и чисел)
- `visibleLabelHolder` - видимость подсказки-заполнителя
- `borderColorOn/borderColorOff` - цвета границ для корректного и некорректного ввода

## Свойства
- `Value` - текущее значение текстового поля
- `Mask` - маска ввода
- `LenValue` - длина текущего значения
- `LenMask` - длина маски
- `RemainingChars` - количество оставшихся символов до полного заполнения
- `IsValid` - проверка корректности ввода

## Автор
VBATools

## Version
1.0.4

## Лицензия
Apache License