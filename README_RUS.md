# UserForms-Class-TextBox-Masks

## Описание
Этот проект содержит класс `clsTextboxMask` для Microsoft VBA, который позволяет создавать текстовые поля с масками ввода в UserForms. Класс обеспечивает валидацию ввода, отображение подсказок-заполнителей и визуальные индикаторы состояния заполнения поля.

## Возможности
- Поддержка различных типов масок ввода:
  - Цифровые маски (с возможностью задания диапазона, знака, дробности)
  - Маски для дат (с проверкой корректности даты и диапазона)
  - Маски для времени (с проверкой корректности времени)
  - Текстовые маски с фиксированной длиной (с различными типами символов)
- Визуальные индикаторы валидности ввода (цвет границы изменяется в зависимости от валидности данных)
- Отображение подсказки-заполнителя с ожидаемым форматом
- Поддержка различных типов символов в масках:
  - `#` - цифры
  - `@` - латинские буквы
  - `*` - любые символы
  - `A` - латинские буквы и цифры
  - `Б` - кириллические буквы
  - `б` - кириллические буквы и цифры

## Установка
1. Скопируйте файл `clsTextboxMask.cls` в ваш VBA проект
2. Импортируйте его в редактор VBA (например, в Excel или Word)

## Использование
Класс предоставляет несколько методов для добавления различных типов масок:

### Числовая маска
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox1, 0, 100, True, False)
```

Использование именованных аргументов (эквивалент моржевого оператора в VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox:=TextBox1, snMin:=0, snMax:=100, IsDecemal:=True, IsNegative:=False)
```

Использование всех именованных аргументов с дополнительными параметрами:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemNumeric(TextBox:=TextBox1, snMin:=0, snMax:=100, IsDecemal:=True, IsNegative:=False, _
                                visibleLabelHolder:=True, formatNumeric:="#.0", borderColorValid:=&H8000006, _
                                borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```
### Маска даты
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox1, "##.##.####", #1/1/2000#, #12/31/2030#, "dd.mm.yyyy")
```

Использование именованных аргументов (эквивалент моржевого оператора в VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox:=TextBox1, Mask:="##.##.####", minDate:=#1/1/2000#, maxDate:=#12/31/2030#, formatDate:="dd.mm.yyyy")
```

Использование всех именованных аргументов с дополнительными параметрами:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenDate(TextBox:=TextBox1, Mask:="##.##.####", minDate:=#1/2000#, maxDate:=#12/31/2030#, _
                                   formatDate:="dd.mm.yyyy", visibleLabelHolder:=True, borderColorValid:=&H8000006, _
                                   borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

### Маска времени
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox1, "##:##", #0:00:00#, #23:59:59#, "hh:mm")
```

Использование именованных аргументов (эквивалент моржевого оператора в VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox:=TextBox1, Mask:="##:##", minDate:=#0:00:00#, maxDate:=#23:59:59#, formatDate:="hh:mm")
```

Использование всех именованных аргументов с дополнительными параметрами:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLenTime(TextBox:=TextBox1, Mask:="##:##", minDate:=#0:00:00#, maxDate:=#23:59:59#, _
                                   formatDate:="hh:mm", visibleLabelHolder:=True, borderColorValid:=&H8000006, _
                                   borderColorNoValid:=&HC0C0FF, foreColorHolder:=&H808080)
```

### Текстовая маска фиксированной длины
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox1, "###@@@")
```

Использование именованных аргументов (эквивалент моржевого оператора в VBA):
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox:=TextBox1, Mask:="###@@@")
```

Использование всех именованных аргументов с дополнительными параметрами:
```vba
Dim textboxMask As New clsTextboxMask
Call textboxMask.addItemFixLen(TextBox:=TextBox1, Mask:="###@@@", visibleLabelHolder:=True, _
                               borderColorValid:=&H800006, borderColorNoValid:=&HC0C0FF, _
                               foreColorHolder:=&H808080)
```

## Параметры
- `TextBox` - объект текстового поля для применения маски
- `Mask` - строка маски ввода
- `Min/Max` - минимальное и максимальное допустимые значения (для числовых и дат)
- `IsDecimal` - разрешение ввода дробных чисел
- `IsNegative` - разрешение ввода отрицательных чисел
- `formatValue` - формат отображения значения (для дат и чисел)
- `visibleLabelHolder` - видимость подсказки-заполнителя
- `borderColorOn/borderColorOff` - цвета границы при валидном и невалидном вводе

## Свойства
- `Value` - текущее значение текстового поля
- `Mask` - маска ввода
- `LenValue` - длина текущего значения
- `LenMask` - длина маски
- `RemainingChars` - количество оставшихся символов до полного заполнения
- `IsValid` - проверка валидности ввода

## Автор
VBATools

## Версия
1.0.3

## Лицензия
Apache License