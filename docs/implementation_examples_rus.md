# Примеры реализации clsTextboxMask

## Введение

В этом документе представлены различные примеры реализации класса `clsTextboxMask` в реальных сценариях. Эти примеры помогут вам понять, как использовать класс в ваших проектах и как адаптировать его под конкретные задачи.

## Пример 1: Форма ввода персональных данных

### Описание
Форма для ввода персональных данных с валидацией всех полей.

### Реализация

```vba
' UserForm: frmPersonalData
' Controls: TextBoxName, TextBoxPhone, TextBoxEmail, TextBoxBirthDate, TextBoxPassport
' CommandButton: cmdSubmit, cmdClear

Dim personalMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set personalMasks = New clsTextboxMask
    
    ' Поле для ФИО - только буквы и пробелы, до 50 символов
    Call personalMasks.AddFieldVariableLength(TextBoxName, 50, "@@@@@@@@@@@@@@@@@@", _
                                             True, , , , "ФИО", RGB(128, 128, 128), "Частично", RGB(165, 102, 41), _
                                             "Введено", RGB(0, 128, 0), "Неверно", RGB(255, 0, 0))
    
    ' Поле для телефона - формат +7(XXX) XXX-XX-XX
    Call personalMasks.AddFieldText(TextBoxPhone, "+7(###) ###-##-##", _
                                   True, RGB(0, 128, 0), RGB(255, 0, 0), , "Тел.", , "Частично", , "OK", , "Ошибка")
    
    ' Поле для email с валидацией через регулярное выражение
    Call personalMasks.AddFieldRegex(TextBoxEmail, _
                                    "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", _
                                    "[a-zA-Z0-9._%+-@]", _
                                    True, , , , "Email", , "Частично", , "OK", , "Неверный email")
    
    ' Поле для даты рождения
    Dim minBirthDate As Date
    minBirthDate = Date - 365 * 100 ' 100 лет назад
    Call personalMasks.AddFieldDate(TextBoxBirthDate, "##.##.####", _
                                   minBirthDate, Date - 365 * 18, "dd.mm.yyyy", _
                                   True, , , "ДР", , "Частично", , "OK", , "дд.мм.ггг")
    
    ' Поле для паспорта - формат XXXX XXXXXX
    Call personalMasks.AddFieldText(TextBoxPassport, "#### ######", _
                                   True, , , "Паспорт", , "Частично", , "OK", , "ХХХХ ХХХХХХ")
End Sub

Private Sub TextBoxName_Change()
    UpdateFieldStatus TextBoxName
End Sub

Private Sub TextBoxPhone_Change()
    UpdateFieldStatus TextBoxPhone
End Sub

Private Sub TextBoxEmail_Change()
    UpdateFieldStatus TextBoxEmail
End Sub

Private Sub TextBoxBirthDate_Change()
    UpdateFieldStatus TextBoxBirthDate
End Sub

Private Sub TextBoxPassport_Change()
    UpdateFieldStatus TextBoxPassport
End Sub

Private Sub UpdateFieldStatus(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = personalMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(240, 255, 240) ' Светло-зеленый
        Else
            textBox.BackColor = RGB(255, 240, 240) ' Светло-красный
        End If
    End If
End Sub

Private Sub cmdSubmit_Click()
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To personalMasks.Count
        Dim currentField As clsTextboxMask
        Set currentField = personalMasks.GetItemByIndex(i)
        
        If Not currentField.IsValid Then
            allValid = False
            MsgBox "Поле '" & GetFieldName(currentField.TextBox.Name) & "' заполнено некорректно!"
            currentField.SetFocus
            Exit Sub
        End If
    Next i
    
    If allValid Then
        MsgBox "Все данные введены корректно! Форма может быть отправлена.", vbInformation
        ' Здесь можно добавить код для сохранения данных
    End If
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To personalMasks.Count
        personalMasks.GetItemByIndex(i).Clear
    Next i
End Sub

Private Function GetFieldName(controlName As String) As String
    Select Case controlName
        Case "TextBoxName": GetFieldName = "ФИО"
        Case "TextBoxPhone": GetFieldName = "Телефон"
        Case "TextBoxEmail": GetFieldName = "Email"
        Case "TextBoxBirthDate": GetFieldName = "Дата рождения"
        Case "TextBoxPassport": GetFieldName = "Паспорт"
        Case Else: GetFieldName = controlName
    End Select
End Function
```

## Пример 2: Форма ввода финансовых данных

### Описание
Форма для ввода финансовых данных с ограничениями по диапазону и форматированием.

### Реализация

```vba
' UserForm: frmFinancialData
' Controls: TextBoxSalary, TextBoxBonus, TextBoxTax, TextBoxTotal
' CommandButton: cmdCalculate, cmdReset

Dim financialMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set financialMasks = New clsTextboxMask
    
    ' Поле для зарплаты - от 100 до 1000000, с десятичными значениями
    Call financialMasks.AddFieldNumeric(TextBoxSalary, 1000, 1000000, True, _
                                       True, "#,##0.00", , , , "Зарплата", , "Частично", , "OK", , "1,000.00 - 1,000,000.00")
    
    ' Поле для премии - от 0 до 500000, с десятичными значениями
    Call financialMasks.AddFieldNumeric(TextBoxBonus, 0, 50000, True, _
                                       True, "#,##0.00", , , , "Премия", , "Частично", , "OK", , "0.00 - 500,000.00")
    
    ' Поле для налога - от 0 до 100 (проценты), с десятичными значениями
    Call financialMasks.AddFieldNumeric(TextBoxTax, 0, 100, True, _
                                       True, "0.00%", , , , "Налог", , "Частично", , "OK", , "0% - 100%")
End Sub

Private Sub TextBoxSalary_Change()
    ValidateAndFormatField TextBoxSalary
End Sub

Private Sub TextBoxBonus_Change()
    ValidateAndFormatField TextBoxBonus
End Sub

Private Sub TextBoxTax_Change()
    ValidateAndFormatField TextBoxTax
End Sub

Private Sub ValidateAndFormatField(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = financialMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(240, 255, 240)
        Else
            textBox.BackColor = RGB(255, 240, 240)
        End If
    End If
End Sub

Private Sub cmdCalculate_Click()
    ' Проверяем валидность всех полей
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To financialMasks.Count
        If Not financialMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Вычисляем итоговое значение
        Dim salary As Double, bonus As Double, tax As Double
        salary = CDbl(financialMasks.GetItemByName("TextBoxSalary").Value)
        bonus = CDbl(financialMasks.GetItemByName("TextBoxBonus").Value)
        tax = CDbl(financialMasks.GetItemByName("TextBoxTax").Value)
        
        Dim total As Double
        total = (salary + bonus) * (1 - tax / 100)
        
        ' Форматируем итоговое значение
        TextBoxTotal.Value = Format(total, "#,##0.00")
        TextBoxTotal.BackColor = RGB(240, 255, 240)
        
        MsgBox "Расчет завершен успешно!", vbInformation
    Else
        MsgBox "Пожалуйста, проверьте правильность заполнения всех полей.", vbExclamation
    End If
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer
    For i = 1 To financialMasks.Count
        financialMasks.GetItemByIndex(i).Clear
    Next i
    TextBoxTotal.Value = ""
    TextBoxTotal.BackColor = RGB(255, 255, 255)
End Sub
```

## Пример 3: Форма ввода технических параметров

### Описание
Форма для ввода технических параметров оборудования с различными типами масок.

### Реализация

```vba
' UserForm: frmTechnicalParams
' Controls: TextBoxSerial, TextBoxMAC, TextBoxIP, TextBoxVersion, TextBoxDate
' CommandButton: cmdValidate, cmdExport

Dim techMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set techMasks = New clsTextboxMask
    
    ' Серийный номер - формат XXXX-XXXX-XXXX
    Call techMasks.AddFieldText(TextBoxSerial, "????-????-????", _
                               True, , , , "Серийный №", , "Частично", , "OK", , "XXXX-XXXX-XXXX")
    
    ' MAC-адрес - формат XX:XX:XX:XX:XX:XX
    Call techMasks.AddFieldText(TextBoxMAC, "AA:AA:AA:AA", _
                               True, , , , "MAC-адрес", , "Частично", , "OK", , "XX:XX:XX:XX:XX:XX")
    
    ' IP-адрес - валидация через регулярное выражение
    Call techMasks.AddFieldRegex(TextBoxIP, _
                                "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$", _
                                "[0-9\.]", _
                                True, , , , "IP-адрес", , "Частично", , "OK", , "XXX.XXX.XXX.XXX")
    
    ' Версия ПО - формат X.X.X
    Call techMasks.AddFieldText(TextBoxVersion, "A.A", _
                               True, , , , "Версия", , "Частично", , "OK", , "X.X.X")
    
    ' Дата установки
    Call techMasks.AddFieldDate(TextBoxDate, "##.##.####", _
                               #1/1/2000#, Date, "dd.mm.yyyy", _
                               True, , , "Дата", , "Частично", , "OK", , "дд.м.гггг")
End Sub

Private Sub TextBoxSerial_Change()
    UpdateTechFieldStatus TextBoxSerial
End Sub

Private Sub TextBoxMAC_Change()
    UpdateTechFieldStatus TextBoxMAC
End Sub

Private Sub TextBoxIP_Change()
    UpdateTechFieldStatus TextBoxIP
End Sub

Private Sub TextBoxVersion_Change()
    UpdateTechFieldStatus TextBoxVersion
End Sub

Private Sub TextBoxDate_Change()
    UpdateTechFieldStatus TextBoxDate
End Sub

Private Sub UpdateTechFieldStatus(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = techMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(220, 255, 220)
            textBox.ForeColor = RGB(0, 100, 0)
        Else
            textBox.BackColor = RGB(255, 220, 220)
            textBox.ForeColor = RGB(150, 0, 0)
        End If
    End If
End Sub

Private Sub cmdValidate_Click()
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To techMasks.Count
        Dim currentField As clsTextboxMask
        Set currentField = techMasks.GetItemByIndex(i)
        
        If Not currentField.IsValid Then
            allValid = False
            MsgBox "Поле '" & GetTechFieldName(currentField.TextBox.Name) & "' содержит неверные данные!"
            currentField.SetFocus
            Exit Sub
        End If
    Next i
    
    If allValid Then
        MsgBox "Все параметры введены корректно!", vbInformation
    End If
End Sub

Private Sub cmdExport_Click()
    ' Проверяем валидность перед экспортом
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To techMasks.Count
        If Not techMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Здесь можно добавить код для экспорта данных
        MsgBox "Данные успешно экспортированы!", vbInformation
    Else
        MsgBox "Невозможно экспортировать данные с неверными параметрами.", vbExclamation
    End If
End Sub

Private Function GetTechFieldName(controlName As String) As String
    Select Case controlName
        Case "TextBoxSerial": GetTechFieldName = "Серийный номер"
        Case "TextBoxMAC": GetTechFieldName = "MAC-адрес"
        Case "TextBoxIP": GetTechFieldName = "IP-адрес"
        Case "TextBoxVersion": GetTechFieldName = "Версия ПО"
        Case "TextBoxDate": GetTechFieldName = "Дата установки"
        Case Else: GetTechFieldName = controlName
    End Select
End Function
```

## Пример 4: Форма ввода данных для отчета

### Описание
Форма с комбинированными полями для ввода данных отчета с валидацией и автоматическим заполнением.

### Реализация

```vba
' UserForm: frmReportData
' Controls: TextBoxReportID, TextBoxPeriod, TextBoxAmount, TextBoxCurrency, TextBoxDescription
' CommandButton: cmdGenerate, cmdSave, cmdLoad

Dim reportMasks As clsTextboxMask

Private Sub UserForm_Initialize()
    Set reportMasks = New clsTextboxMask
    
    ' ID отчета - формат RPT-XXXXX
    Call reportMasks.AddFieldText(TextBoxReportID, "RPT-#####", _
                                 True, , , , "ID отчета", , "Частично", , "OK", , "RPT-XXXXX")
    
    ' Период отчета - формат ММ.ГГГГ
    Call reportMasks.AddFieldText(TextBoxPeriod, "##.####", _
                                 True, , , "Период", , "Частично", , "OK", , "ММ.ГГГГ")
    
    ' Сумма - числовое поле с ограничениями
    Call reportMasks.AddFieldNumeric(TextBoxAmount, 0, 99999.99, True, _
                                    True, "#,##0.00", , , , "Сумма", , "Частично", , "OK", , "0.00 - 9,999,999.99")
    
    ' Код валюты - 3 буквы
    Call reportMasks.AddFieldText(TextBoxCurrency, "AAA", _
                                 True, , , "Валюта", , "Частично", , "OK", , "USD/EUR/RUB")
    
    ' Описание - переменная длина до 200 символов
    Call reportMasks.AddFieldVariableLength(TextBoxDescription, 200, , _
                                           True, , , , "Описание", , "Частично", , "OK", , "Максимум 200 символов")
End Sub

Private Sub TextBoxReportID_Change()
    UpdateReportFieldStatus TextBoxReportID
End Sub

Private Sub TextBoxPeriod_Change()
    UpdateReportFieldStatus TextBoxPeriod
End Sub

Private Sub TextBoxAmount_Change()
    UpdateReportFieldStatus TextBoxAmount
End Sub

Private Sub TextBoxCurrency_Change()
    UpdateReportFieldStatus TextBoxCurrency
End Sub

Private Sub TextBoxDescription_Change()
    ' Обновляем счетчик оставшихся символов для описания
    Dim field As clsTextboxMask
    Set field = reportMasks.GetItemByName(TextBoxDescription.Name)
    
    If Not field Is Nothing Then
        Dim remaining As Integer
        remaining = field.RemainingChars
        
        ' Обновляем подсказку для поля описания
        field.PlaceholderPartial = "Осталось: " & remaining & " символов"
        
        If remaining < 10 Then
            field.PlaceholderPartialColor = RGB(255, 0, 0) ' Красный цвет при малом количестве символов
        Else
            field.PlaceholderPartialColor = RGB(128, 128, 128) ' Стандартный цвет
        End If
        
        field.UpdatePlaceholder
        
        UpdateReportFieldStatus TextBoxDescription
    End If
End Sub

Private Sub UpdateReportFieldStatus(textBox As MSForms.TextBox)
    Dim field As clsTextboxMask
    Set field = reportMasks.GetItemByName(textBox.Name)
    
    If Not field Is Nothing Then
        If field.IsValid Then
            textBox.BackColor = RGB(245, 255, 245)
        Else
            textBox.BackColor = RGB(255, 245, 245)
        End If
    End If
End Sub

Private Sub cmdGenerate_Click()
    ' Проверяем валидность всех полей
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To reportMasks.Count
        If Not reportMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Генерируем отчет
        Dim reportID As String, period As String, amount As String, currency As String, description As String
        reportID = reportMasks.GetItemByName("TextBoxReportID").Value
        period = reportMasks.GetItemByName("TextBoxPeriod").Value
        amount = reportMasks.GetItemByName("TextBoxAmount").Value
        currency = reportMasks.GetItemByName("TextBoxCurrency").Value
        description = reportMasks.GetItemByName("TextBoxDescription").Value
        
        MsgBox "Отчет " & reportID & " за период " & period & " на сумму " & amount & " " & currency & " успешно сгенерирован!" & vbCrLf & vbCrLf & _
               "Описание: " & description, vbInformation
    Else
        MsgBox "Пожалуйста, проверьте все поля перед генерацией отчета.", vbExclamation
    End If
End Sub

Private Sub cmdSave_Click()
    ' Проверяем валидность перед сохранением
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Integer
    For i = 1 To reportMasks.Count
        If Not reportMasks.GetItemByIndex(i).IsValid Then
            allValid = False
            Exit For
        End If
    Next i
    
    If allValid Then
        ' Здесь можно добавить код для сохранения данных
        MsgBox "Данные отчета успешно сохранены!", vbInformation
    Else
        MsgBox "Невозможно сохранить данные с неверными значениями.", vbExclamation
    End If
End Sub

Private Sub cmdLoad_Click()
    ' Здесь можно добавить код для загрузки данных
    MsgBox "Функция загрузки данных будет реализована позже.", vbInformation
End Sub
```

## Заключение

Эти примеры демонстрируют гибкость и мощность класса `clsTextboxMask`. С помощью этого класса можно создавать сложные формы с валидацией, автоматическим форматированием и пользовательским интерфейсом. Класс поддерживает различные типы масок и позволяет настраивать внешний вид полей в зависимости от их состояния.

Каждый пример можно адаптировать под конкретные требования вашего проекта. Главное - понимать принципы работы с классом и правильно применять различные типы масок в зависимости от типа данных, которые нужно вводить.