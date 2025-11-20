# Руководство по использованию clsTextboxMask

## Введение

Класс `clsTextboxMask` предоставляет мощный инструмент для создания текстовых полей с масками ввода в VBA. В этом руководстве вы найдете пошаговые инструкции по установке, настройке и использованию класса в ваших проектах.

## Установка

### Требования
- Microsoft Excel (2010 или новее)
- Включенная поддержка VBA
- Разрешение на использование объектов MSForms

### Установка класса
1. Откройте ваш Excel файл с VBA проектом (Alt+F11)
2. В окне VBA редактора выберите "File" → "Import File"
3. Выберите файл `clsTextboxMask.cls`
4. Класс будет добавлен ваш проект

## Быстрый старт

### Простой пример использования

Создайте UserForm и добавьте на него TextBox. Затем используйте следующий код:

```vba
Dim maskField As New clsTextboxMask
Call maskField.AddFieldText(Me.TextBox1, "###-##-##")
```

Этот код создаст поле для ввода номера телефона в формате "123-45-67".

### Пример с числовым полем

```vba
Dim numField As New clsTextboxMask
Call numField.AddFieldNumeric(inputTextBox:=Me.TextBox1, _
                             minValue:=0, _
                             maxValue:=100, _
                             allowDecimal:=True)
```

## Подробные примеры использования

### 1. Создание поля ввода даты

```vba
Private Sub UserForm_Initialize()
    Dim dateField As New clsTextboxMask
    Call dateField.AddFieldDate(inputTextBox:=Me.TextBoxDate, _
                               dateMask:="##.##.####", _
                               minDate:=#1/1/2020#, _
                               maxDate:=#12/31/2030#, _
                               dateFormat:="dd.mm.yyyy")
End Sub
```

Этот код создаст поле для ввода даты в формате "дд.мм.гггг" с ограничением на диапазон дат с 1 января 2020 по 31 декабря 2030 года.

### 2. Создание поля ввода времени

```vba
Private Sub UserForm_Initialize()
    Dim timeField As New clsTextboxMask
    Call timeField.AddFieldTime(inputTextBox:=Me.TextBoxTime, _
                               timeMask:="##:##", _
                               minTime:=#0:00:00#, _
                               maxTime:=#23:59#, _
                               timeFormat:="hh:mm")
End Sub
```

Этот код создаст поле для ввода времени в формате "чч:м".

### 3. Создание поля ввода телефона

```vba
Private Sub UserForm_Initialize()
    Dim phoneField As New clsTextboxMask
    Call phoneField.AddFieldText(inputTextBox:=Me.TextBoxPhone, _
                                textMask:="+7(###) ###-##-##")
End Sub
```

Этот код создаст поле для ввода российского номера телефона с автоматическим форматированием.

### 4. Создание числового поля с ограничениями

```vba
Private Sub UserForm_Initialize()
    Dim numField As New clsTextboxMask
    Call numField.AddFieldNumeric(inputTextBox:=Me.TextBoxNumber, _
                                 minValue:=-10, _
                                 maxValue:=100, _
                                 allowDecimal:=True)
End Sub
```

Этот код создаст поле для ввода чисел от -10 до 100, включая десятичные значения.

### 5. Создание поля с переменной длиной

```vba
Private Sub UserForm_Initialize()
    Dim varField As New clsTextboxMask
    Call varField.AddFieldVariableLength(inputTextBox:=Me.TextBoxName, _
                                        maxLength:=50)
End Sub
```

Этот код создаст поле для ввода текста с максимальной длиной 50 символов.

### 6. Создание поля с регулярным выражением

```vba
Private Sub UserForm_Initialize()
    Dim emailField As New clsTextboxMask
    Call emailField.AddFieldRegex(inputTextBox:=Me.TextBoxEmail, _
                                 RegexPattern:="^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", _
                                 RegexFilter:="[a-zA-Z0-9._%+-@]")
End Sub
```

Этот код создаст поле для ввода email с валидацией через регулярное выражение.

## Настройка внешнего вида

### Изменение цветов границ

```vba
Private Sub UserForm_Initialize()
    Dim field As New clsTextboxMask
    Call field.AddFieldText(inputTextBox:=Me.TextBox1, _
                           textMask:="###-###", _
                           BorderColorValid:=RGB(0, 128, 0), _
                           BorderColorInvalid:=RGB(255, 0, 0))
End Sub
```

### Настройка плейсхолдеров

```vba
Private Sub UserForm_Initialize()
    Dim field As New clsTextboxMask
    Call field.AddFieldText(inputTextBox:=Me.TextBox1, _
                           textMask:="###-###", _
                           PlaceholderEmpty:="Введите код", _
                           PlaceholderEmptyColor:=RGB(128, 128, 128), _
                           PlaceholderComplete:="Код введен", _
                           PlaceholderCompleteColor:=RGB(0, 128, 0))
End Sub
```

## Работа с коллекцией полей

Класс позволяет управлять несколькими полями одновременно:

```vba
Private Sub UserForm_Initialize()
    Dim formMasks As New clsTextboxMask
    
    ' Добавляем несколько полей
    Call formMasks.AddFieldText(Me.TextBox1, "###-##-##")
    Call formMasks.AddFieldDate(Me.TextBox2, "##.##.####", #1/1/2000#, #12/31/2030#)
    Call formMasks.AddFieldNumeric(Me.TextBox3, 0, 100, False)
    
    ' Проверяем валидность всех полей
    Dim isValid As Boolean
    isValid = True
    
    Dim i As Integer
    For i = 1 To formMasks.Count
        If Not formMasks.GetItemByIndex(i).IsValid Then
            isValid = False
            Exit For
        End If
    Next i
    
    MsgBox "Все поля корректны: " & isValid
End Sub
```

## Практические советы

### 1. Обработка событий формы

Чтобы реагировать на изменения в полях с масками, используйте события текстовых полей:

```vba
Private Sub TextBox1_Change()
    Dim field As clsTextboxMask
    Set field = clsTB.GetItemByName(TextBox1.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            ' Поле заполнено корректно
            TextBox1.BackColor = RGB(240, 255, 240) ' Светло-зеленый
        Else
            ' Поле заполнено некорректно
            TextBox1.BackColor = RGB(255, 240, 240) ' Светло-красный
        End If
    End If
End Sub
```

### 2. Управление фокусом

```vba
Private Sub CommandButton1_Click()
    ' Установить фокус на определенное поле
    Dim field As clsTextboxMask
    Set field = clsTB.GetItemByName(TextBox1.Name)
    If Not field Is Nothing Then field.SetFocus
End Sub
```

### 3. Очистка полей

```vba
Private Sub CommandButton2_Click()
    ' Очистить все поля
    Dim formMasks As clsTextboxMask
    Set formMasks = New clsTextboxMask
    
    Dim i As Integer
    For i = 1 To formMasks.Count
        formMasks.GetItemByIndex(i).Clear
    Next i
End Sub
```

### 4. Удаление элементов маски

```vba
Private Sub CommandButton3_Click()
    ' Удалить конкретное поле
    Dim field As clsTextboxMask
    Set field = clsTB.GetItemByName(TextBox1.Name)
    If Not field Is Nothing Then field.RemoveItem
End Sub
```

## Распространенные ошибки и решения

### Ошибка: "The item has already been created"

Эта ошибка возникает при попытке добавить маску к текстовому полю, которое уже имеет маску. Решение:

```vba
' Проверяем, существует ли уже элемент
Dim existingField As clsTextboxMask
Set existingField = clsTB.GetItemByName(TextBox1.Name)

If existingField Is Nothing Then
    ' Элемент не существует, можно добавлять
    Call clsTB.AddFieldText(TextBox1, "###-###")
Else
    ' Элемент уже существует, можно обновить его свойства
    existingField.Mask = "###-###"
End If
```

### Ошибка: "TextBox cannot be Nothing"

Убедитесь, что текстовое поле существует и не равно Nothing перед добавлением маски:

```vba
If Not Me.TextBox1 Is Nothing Then
    Call maskField.AddFieldText(Me.TextBox1, "###-###")
End If
```

## Примеры реальных сценариев

### Форма регистрации пользователя

```vba
Private Sub UserForm_Initialize()
    Dim formMasks As New clsTextboxMask
    
    ' Поле для email
    Call formMasks.AddFieldRegex(Me.TextBoxEmail, _
                                "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", _
                                "[a-zA-Z0-9._%+-@]", _
                                True, , , , "Email", , "Частично", , "OK", , "Неверный email")
    
    ' Поле для телефона
    Call formMasks.AddFieldText(Me.TextBoxPhone, "+7(###) ###-##-##", _
                               True, , , , "Телефон", , "Частично", , "OK", , "Неверный формат")
    
    ' Поле для возраста
    Call formMasks.AddFieldNumeric(Me.TextBoxAge, 18, 100, False, _
                                  True, , , , "Возраст", , "Частично", , "OK", , "18-100 лет")
    
    ' Поле для даты рождения
    Dim birthDate As Date
    birthDate = Date - 365 * 18 ' 18 лет назад
    Call formMasks.AddFieldDate(Me.TextBoxBirth, "##.##.####", _
                               birthDate - 365 * 50, Date, _
                               "dd.mm.yyyy", True, , , , "ДР", , "Частично", , "OK", , "дд.м.гггг")
End Sub

Private Sub CommandButtonSubmit_Click()
    Dim formMasks As clsTextboxMask
    Set formMasks = New clsTextboxMask
    
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To formMasks.Count
        If Not formMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            MsgBox "Поле " & formMasks.GetItemByIndex(i).TextBox.Name & " заполнено некорректно"
            formMasks.GetItemByIndex(i).SetFocus
            Exit Sub
        End If
    Next i
    
    If allValid Then
        MsgBox "Все поля заполнены корректно! Форма может быть отправлена."
    End If
End Sub
```

## Расширенные возможности

### Использование шаблонов плейсхолдеров

Шаблоны плейсхолдеров позволяют динамически отображать информацию о состоянии поля:

```vba
Call maskField.AddFieldText(Me.TextBox1, "####-####-####", _
                           True, , , , , , , "Шаблон: {holder} Осталось: {remaining}")
```

Доступные маркеры:
- `{mask}` - отображает маску
- `{filled}` - количество заполненных символов
- `{remaining}` - количество оставшихся символов
- `{holder}` - плейсхолдер с маской
- `{RegexPattern}` - паттерн регулярного выражения
- `{RegexFilter}` - фильтр регулярного выражения
- `{percent}` - процент заполнения

### Кастомизация масок

Вы можете создавать сложные маски с комбинацией различных символов:

```vba
' Маска для автомобильного номера: A123AA123
Call maskField.AddFieldText(Me.TextBoxCarNumber, "@###@@###", _
                           True, , , , "Номер", , "Частично", , "OK", , "A123AA123")

' Маска с кириллическими буквами: А123БВ456
Call maskField.AddFieldText(Me.TextBoxCyrillic, "Б###ББ###", _
                           True, , , , "Номер", , "Частично", , "OK", , "А123БВ456")
```

## Заключение

Класс `clsTextboxMask` предоставляет мощный и гибкий инструмент для создания валидированных текстовых полей в VBA. С его помощью вы можете улучшить пользовательский интерфейс своих приложений, обеспечивая корректный ввод данных и упрощая процесс валидации.

Используйте предоставленные примеры как отправную точку для создания собственных решений с использованием этого класса.