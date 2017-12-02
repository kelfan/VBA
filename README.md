# VBA
learning materials for Excel VBA

# say hello
```vb
Sub SayHello()

MsgBox "hello world"

End Sub
```

# cell value add 1
```vb
Sub Lower()

Range("e3").Value = Range("e3").Value - 1

End Sub

Sub Higher()

Range("e3").Value = Range("e3").Value + 1

End Sub
```

# select move upside down
```vb
Sub MoveUp()

Selection.Offset(-1, 0).Select

End Sub

Sub MoveDown()

Selection.Offset(1, 0).Select

End Sub

Sub MoveRight()

Selection.Offset(0, 1).Select

End Sub

Sub MoveLeft()

Selection.Offset(0, -1).Select

End Sub

```

# declare variable/Dim counter as Integer
```vb
Dim counter as Integer
```

# generate a list with loop
```vb
Sub ListWithLoop()

Dim counter As Integer

For counter = 0 To 10

Selection.Value = counter

Selection.Offset(1, 0).Select

Next counter

End Sub
```

# number add string
```vb
idx = "A" & counter
```

# range selection by variable
```vb
Dim MyRange as String
MyRange = "A1:D11"
Range(MyRange).Select
```

```vb
Dim Copyrange As String

Startrow = 1
Lastrow = 11
Let Copyrange = "A" & Startrow & ":" & "D" & Lastrow
Range(Copyrange).Select
```

# Random Numbers in a Range of cells
```vb
Sub qwerty()
    Dim r As Range
    Set r = Range("G8:H34")
    With r
        .Formula = "=randbetween(0,2)"
        .Copy
        .PasteSpecial (xlPasteValues)
    End With
End Sub
```

# get random cell
```vb
Sub GetRandomCell()

    Dim RNG As Range
    Set RNG = Range("A1:p4")

    Dim randomCell As Long
        randomCell = Int(Rnd * RNG.Cells.Count) + 1

    With RNG.Cells(randomCell)
        .Select
        .Interior.Color = vbYellow
    End With

End Sub
```

# check each row cell
```vb
Dim rng As Range
Dim row As Range
Dim cell As Range

Set rng = Range("A1:C2")

For Each row In rng.Rows
  For Each cell in row.Cells
    'Do Something
  Next cell
Next row
```

# set row cell value
```vb
Sub CycleThrough()
 Dim Counter As Integer
 For Counter = 1 To 20
 Worksheets("Sheet1").Cells(Counter, 3).Value = Counter
 Next Counter
End Sub

```

# get set cell value | if then | compare string
```vb
Sub check()

    Dim Counter As Integer

    For Counter = 2 To 15

        Dim value As String
        value = Worksheets("If not and fired").Cells(Counter, 1).Text
        If StrComp(value, "sunny", vbTextCompare) = 0 Then
            Worksheets("If not and fired").Cells(Counter, 5).value = "Play"
            Worksheets("If not and fired").Cells(Counter, 6).value = "1"
        ElseIf StrComp(value, "rainy", vbTextCompare) = 0 Then
            Worksheets("If not and fired").Cells(Counter, 5).value = "Not Play"
            Worksheets("If not and fired").Cells(Counter, 6).value = "2"
        End If

        value = Worksheets("If not and fired").Cells(Counter, 3).Text
        If StrComp(value, "high", vbTextCompare) = 0 Then
            Worksheets("If not and fired").Cells(Counter, 5).value = "Not Play"
            Worksheets("If not and fired").Cells(Counter, 6).value = "3"
        End If

    Next Counter

End Sub
```

# shortcut/ ctrl + 拖动 = 复制


# call a function
```vb
Sub mySecondMacro()

'runs myfirstmacro
Call myFirstMacro

End Sub

Sub myFirstMacro()

'this is my first macro
MsgBox ("hello")

End Sub
```

# Option Explicit = this means all variable must be declared

# variables 变量
```vb
Option Explicit 'this means all variable must be declared
Sub variables()
'this is single line variable declaration
Dim int1 As Integer, int2 As Integer, xdate2 As Date, xstr As String
Dim int6, int7, int8 As Integer 'warning: int6 int7 will be variant not integer

'give value
Dim myvar As Integer
myvar = 8

'this is a constant variable
Const num As Integer = 9

Dim var_byte As Byte
var_byte = 255 '256 will be overflow

Dim vbool As Boolean
vbool = False ' or 0..555 or true

Dim vint As Integer 'this can store -32,768 to
vint = 5.7
MsgBox (vint) 'implicit rounding apply

Dim vcurrency As Currency
vcurrency = 4566.88

Dim vlong As Long
vlong = 2147731423#

Dim vsingle As Single
vsingle = -2.5333

Dim vdouble As Double
vdouble = -5.00001

Dim vdate As Date
vdate = "12/31/9999"

Dim vstr As String, str2 As String  '0-2billion characters --->10 Byte of memory
vstr = "my name is xxx"
str2 = 100
MsgBox (str2 - vstr) 'result is a number


Dim vvariant As Variant 'this can numbers up to data type
vvariant = "2342342"


End Sub
```

# Scope of variable = 变量的范围 public>dim top>static>dim
module 1
```vb
Public q As Integer
Dim z As Integer

Sub Sub1()

Dim x As Integer
Static y As Integer

x = x + 100
y = y + 100
z = z + 100
q = q + 100

MsgBox ("x in sub 1 = " & x) 'dies when sub1 ends
MsgBox ("y in sub 1 = " & y) 'lives after sub1 ends but not seen in sub2
MsgBox ("z in sub 1 = " & z) 'lives after sub1 ends and seen sub2
MsgBox ("q in sub 1 = " & q) 'lives after sub1 ends and seen sub2 and seen across modules

Call Sub2
Call GlobalVariable

End Sub

Sub Sub2()

MsgBox ("x in sub2 = " & x) 'no value
MsgBox ("y in sub2 = " & y) 'no value
MsgBox ("z in sub2 = " & z) 'has a value because declared "dim" at the top
MsgBox ("q in sub2 = " & q) 'has a value because declared public at the top

End Sub
```
module2
```vb
Sub GlobalVariable()

MsgBox ("x in second module = " & x) 'no Value
MsgBox ("y in second module = " & y) 'no Value
MsgBox ("z in second module = " & z) 'no Value
MsgBox ("q in second module = " & q) 'has a Value because declare public

End Sub
```

# parameter = 带参数的函数
```vb
Sub mySecondMacro()

'runs myfirstmacro
Call myFirstMacro("hello", "world", 9090909)

End Sub

Sub myFirstMacro(strVar As String, strVar2 As String, num As Long)

'this is my first macro
MsgBox (strVar & " - " & strVar2 & " - " & num)

End Sub
```

# Function = 可以在单元格里调用
```vb
Option Explicit
Sub TestFunctions()

Dim x As Integer
Dim y As Double
x = Return1()
'MsgBox (x)

y = ConvertToCelsius(100)
MsgBox (y)

End Sub

Function ConvertToCelsius(TempFahrenheit As Double) As Double

ConvertToCelsius = (TempFahrenheit - 32) * 5 / 9

'Dim z As Double
'z = (TempFahrenheit - 32) * 5 / 9

End Function


Function Return1() As Integer

Return1 = 1

End Function
```
# Event 当Excel改变时触发的事件
thisWorkbook 整个Excel相关,例如切换sheet页面
```vb
Private Sub Workbook_Open()

MsgBox ("you opened Excel")

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)

MsgBox ("your new sheet is " & Sh.Name)

End Sub
```
sheet1 页面内变化触发
```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

MsgBox ("new selection " & Target.Address)

End Sub
```

# class & property 类和属性
class modules "basketballTeam"
```vb
Option Explicit

Private teamName As String

Public Property Get Name() As String
Name = teamName
End Property


Public Property Let Name(param_name As String)
teamName = param_name
End Property
```
module "classes"
```vb
Sub TestClasses()

Dim bbteam As basketballTeam
Set bbteam = New basketballTeam 'this instantiate the object

bbteam.Name = "Lakers" 'use Let
MsgBox (bbteam.Name) 'use Get

MsgBox (Application.Name) 'output Microsoft Excel

End Sub
```

# object variables
```vb
'''declare object variables
Sub objectVarable()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    MsgBox (ws.Name)

    Dim ws2 As Worksheet
    Set ws2 = Sheets("sheet2")
    MsgBox (ws2.Name)

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    MsgBox (wb.Name)
End Sub
```

# macros recorder = 可以记录步骤为命令,然后下次直接运行重复的命令
记录
  developer -> record macro -> 可以记录名字,快捷键,范围和描述
运行
  developer -> macros -> 选中名字 -> run
记录的例子如下
```vb
Sub MoveRange()
'
' MoveRange Macro
' this is the move range macro
'

'
    Sheets("Sheet1").Select
    Range("A5:B7").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.NumberFormatLocal = "yyyy/m/d"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[699]C)"
    Range("A2:B2").Select
    Selection.Font.Bold = True
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2:B2").Select
    Range("B2").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
```

# relational|logical operators & if statement
```vb
'relational operators - define a relationship between 2 things
'6 relational operators: 4=4 4<>5 5>4 6<8 7>=7 9<=10

'logical operators - Return true/false
'3 logical operators: And, Or, and Not
' 4=4 and 5=7 -->false
' 4=4 or 5=7 -->true
' not(4=4) -->false


Option Explicit
Function getTaxRate(salary As Double) As Double

    If salary > 5000 Then
        getTaxRate = 0.25
    End If

End Function

Function getTaxRateElse(salary As Double) As Double

    If salary > 5000 Then
        getTaxRateElse = 0.25
        Else
        getTaxRateElse = 0.1
    End If

End Function

Function getTaxRateElseAND(salary As Double) As Double

    If salary > 5000 And salary < 40000 Then
        getTaxRateElseAND = 0.25
        Else
        getTaxRateElseAND = 0.1
    End If

End Function

Function getTaxRateElseIf(salary As Double) As Double

    If salary > 5000 And salary < 40000 Then
        getTaxRateElseIf = 0.25
    ElseIf salary >= 40000 And salary < 90000 Then
        getTaxRateElseIf = 0.35
    ElseIf salary >= 90000 Then
        getTaxRateElseIf = 0.45
    Else
        getTaxRateElseIf = 0.1
    End If

End Function

Function getTaxRateNestedIf(salary As Double, haskids As String) As Double

    If salary > 5000 And salary < 40000 Then
        ' this is a nested if
        If haskids = "yes" Then
            getTaxRateNestedIf = 0.15
        Else
            getTaxRateNestedIf = 0.25
        End If
    ElseIf salary >= 40000 And salary < 90000 Then
        If haskids = "yes" Then
            getTaxRateNestedIf = 0.28
        Else
            getTaxRateNestedIf = 0.35
        End If
    ElseIf salary >= 90000 Then
        If haskids = "yes" Then
            getTaxRateNestedIf = 0.42
        Else
            getTaxRateNestedIf = 0.45
        End If
    Else
        getTaxRateNestedIf = 0
    End If

End Function
```

# select case statement
```vb
Option Explicit
Function getTaxRateSelectCase(salary As Double) As Double

    Select Case salary
        Case Is > 5000
            getTaxRateSelectCase = 0.25
        Case Else
            getTaxRateSelectCase = 0
    End Select

End Function

Function getTaxRateSelectCaseTo(salary As Double) As Double

    Select Case salary
        Case 5000 To 40000
            getTaxRateSelectCaseTo = 0.25
        Case Else
            getTaxRateSelectCaseTo = 0.1
    End Select

End Function

Function getTaxRateSelectCaseNested(salary As Double, haskids As String) As Double

    Select Case salary
        Case 5000 To 40000
            Select Case haskids
                Case "yes"
                    getTaxRateSelectCaseNested = 0.15
                Case Else
                    getTaxRateSelectCaseNested = 0.25
            End Select
        Case 40000 To 90000
            Select Case haskids
                Case "yes"
                    getTaxRateSelectCaseNested = 0.28
                Case Else
                    getTaxRateSelectCaseNested = 0.35
            End Select
        Case Is > 90000
            Select Case haskids
                Case "yes"
                    getTaxRateSelectCaseNested = 0.42
                Case Else
                    getTaxRateSelectCaseNested = 0.45
            End Select
        Case Else
            getTaxRateSelectCaseNested = 0
    End Select

End Function
```

# do while loop
```vb
Option Explicit
Sub TestDoWhile()

    Sheets("loops").Select
    Cells.ClearContents

    Do While Range("a1").Value < 10
        Range("a1").Value = Range("a1").Value + 1
    Loop

End Sub

Sub TestDoWhile2()

    Sheets("loops").Select
    Cells.ClearContents

    Do
        Range("a1").Value = Range("a1").Value + 1
    Loop While Range("a1").Value < 10

End Sub

```

# sheet 操作
```vb
 Sheets("loops").Select
```
# Range 操作 一个单元格
```vb
Sub TestDoUntil()

Sheets("loops").Select
Cells.ClearContents

'loops stops when condition is true
Do Until Range("a1").Value >= 10
    Range("a1").Value = Range("a1").Value + 1
Loop

End Sub
```

# cells 操作 row & columns
```vb
Sub DoWhileLoopsRowColumn()

    Dim num As Integer, xrow As Long, xcol As Long

    Sheets("loops").Select
    Cells.ClearContents

    num = 10
    xrow = 1
    xcol = 1

    Do While xrow <= 5
        Cells(xrow, xcol).Value = num
        xrow = xrow + 1 ' increment the row variable
        xcol = xcol + 1 ' increment the Column Variable
        num = num + 1 'increment mumber
    Loop

End Sub
```

# delete blank row 删除空白行
```vb
Sub DeleteBlankRows()

Dim lastrow As Long, xrow As Long
xrow = 1

'find last cell in column A with Data
lastrow = Range("A1000000").End(xlUp).Row

'delete row from first until the last row
Do Until xrow = lastrow

    If Cells(xrow, 1).Value = "" Then
        Cells(xrow, 1).Select
        Selection.EntireRow.Delete

        xrow = xrow - 1 'because a row is deleted
        lastrow = lastrow - 1 'because a row is deleted
    End If

    xrow = xrow + 1
Loop

End Sub
```

# do until blank cell 求last row number or last Column number
```vb
Sub DoUntilBlankCell()

Dim xrow As Long, xcol As Long, lastCol As Long
xrow = 1
xcol = 1

    Do Until Cells(xrow, xcol).Value = ""
        Cells(xrow, xcol).Select
        xcol = xcol + 1
    Loop
    lastCol = xcol - 1

End Sub
```

```vb
Sub DoUntilBlankCell()

Dim xrow As Long, xcol As Long, lastrow As Long
xrow = 1
xcol = 1

    Do Until Cells(xrow, xcol).Value = ""
        Cells(xrow, xcol).Select
        xrow = xrow + 1
    Loop
    lastrow = xrow - 1

End Sub
```

# do until loop
```vb
Option Explicit

Sub TestDoUntil()

Sheets("loops").Select
Cells.ClearContents

'loops stops when condition is true
Do Until Range("a1").Value >= 10
    Range("a1").Value = Range("a1").Value + 1
Loop

End Sub


Sub DoUntilBlankCell()

Dim xrow As Long, xcol As Long, lastrow As Long
xrow = 1
xcol = 1

    Do Until Cells(xrow, xcol).Value = ""
        Cells(xrow, xcol).Select
        xrow = xrow + 1
    Loop
    lastrow = xrow - 1

End Sub

Sub DeleteBlankRows()

Dim lastrow As Long, xrow As Long
xrow = 1

'find last cell in column A with Data
lastrow = Range("A1000000").End(xlUp).Row

'delete row from first until the last row
Do Until xrow = lastrow

    If Cells(xrow, 1).Value = "" Then
        Cells(xrow, 1).Select
        Selection.EntireRow.Delete

        xrow = xrow - 1 'because a row is deleted
        lastrow = lastrow - 1 'because a row is deleted
    End If

    xrow = xrow + 1
Loop

End Sub
```

# for Next loop
```vb
Option Explicit
Sub TestForNext()

    Dim i As Long
    Sheets("Next").Select
    Cells.ClearContents

    For i = 1 To 10
        Cells(i, 1).Value = i
    Next i

End Sub

Sub TestForNext2()

    Dim i As Long
    Sheets("Next").Select
    Cells.ClearContents

    For i = 0 To 10
        Cells(i + 1, 1).Value = i
    Next i

End Sub

Sub TestForNext3()

    Dim i As Long
    Sheets("Next").Select
    Cells.ClearContents

    For i = 1 To 20 Step 2
        Cells(i, 1).Value = i
    Next i

End Sub


Sub TestForNext4()

    Dim i As Long
    Sheets("Next").Select
    Cells.ClearContents

    For i = 20 To 1 Step -2
        Cells(i, 1).Value = i
    Next i

End Sub

Sub ForNextLoopAddSheets()

    Dim numberOfSheets As Integer, counter As Integer

    numberOfSheets = Application.InputBox("how many worksheets do you want to add?", "add worksheets", , , , , , 1)

    If numberOfSheets = False Then
        Exit Sub 'end if user clicked CANCEL
    Else
        'add worksheets
        For counter = 1 To numberOfSheets
            Worksheets.Add 'add a worksheet
        Next counter
    End If

End Sub
```


# select Range highlight condition 选择范围内高亮适合的
```vb
Sub ForEachLoopRange()

    Dim rng As Range
    Dim rcell As Range
    Set rng = Selection

    For Each rcell In rng
        rcell.Value = rcell.Address
    Next rcell

    For Each rcell In rng
        rcell.Select

        If rcell.Value > 200 Then
            With Selection.Interior
                .Color = 65535
            End With
        Else
            With Selection.Interior
                .Pattern = xlNone
            End With
        End If

    Next rcell

End Sub
```

# for each loop
```vb
Option Explicit
Sub ForEachLoopWorksheets()

    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Select

        If ws.Name = "loops" Then
            MsgBox (ws.Name)
        End If

    Next ws

End Sub

Sub ForEachLoopRange()

    Dim rng As Range
    Dim rcell As Range
    Set rng = Selection

    For Each rcell In rng
        rcell.Value = rcell.Address
    Next rcell

    For Each rcell In rng
        rcell.Select

        If rcell.Value > 200 Then
            With Selection.Interior
                .Color = 65535
            End With
        Else
            With Selection.Interior
                .Pattern = xlNone
            End With
        End If

    Next rcell

End Sub
```

# Array populated & loop
```vb
'Option Base 1 'change Array start From 1
Option Explicit
Sub StaticArray()

    Dim names1(2) As String 'names(0) names(1) names(2)
    Dim names2(2, 2) As String 'String(0 to 2, 0 to 2)
    Dim names3(2, 2, 2) As String 'String(0 to 2, 0 to 2, 0 to 2)

    names1(1) = "HI"
    MsgBox (names1(1))

End Sub

Sub StaticArrayPopulateAndLoop()

    Dim names(2) As String

    names(0) = "Bob"
    names(1) = "Mary"
    names(2) = "George"

    MsgBox ("Ubound(names,1) = " & UBound(names, 1))
    ' loop through the Array
    For i = 0 To UBound(names, 1) 'give the lastIndex Of Array, 1 is the dimension Of the Array

        Cells(i + 1, 1).Value = names(i)

    Next i

End Sub
```

# Array/acquire a column Of Data 获取一列数据
```vb
Sub populate1DArrayFromWorksheet()

    Dim months(11) As String
    Dim i As Integer
    Dim xrow As Long

    i = 0 'Variable for the Index Of the Array
    xrow = 2 'Variable for the row # on worksheet

    Do Until Cells(xrow, 1).Value = ""
        months(i) = Cells(xrow, 1).Value 'this populates the array

        i = i + 1
        xrow = xrow + 1
    Loop

    For i = 0 To UBound(months, 1)
        If months(i) = MonthName(Month(Date)) Then
            MsgBox ("the Current month is " & MonthName(Month(Date)))
        End If
    Next i

End Sub
```

# Array/ currency change example
```vb
Function ConvertToUsd(foreignCurrencySymbol As String, amount As Double) As Double

    Dim ExchangeRates(3, 2) As Variant, i As Integer

    ExchangeRates(0, 0) = "Canada"
    ExchangeRates(0, 1) = "CAD"
    ExchangeRates(0, 2) = "1.05"

    ExchangeRates(1, 0) = "Euro Zone"
    ExchangeRates(1, 1) = "EUR"
    ExchangeRates(1, 2) = "1.2"

    ExchangeRates(2, 0) = "Japan"
    ExchangeRates(2, 1) = "JPY"
    ExchangeRates(2, 2) = "0.012"

    ExchangeRates(3, 0) = "Mexico"
    ExchangeRates(3, 1) = "MXN"
    ExchangeRates(3, 2) = "0.07"

    For i = 0 To UBound(ExchangeRates, 1)
        If foreignCurrencySymbol = ExchangeRates(i, 1) Then 'check the second Index/dimension
            ConvertToUsd = amount * ExchangeRates(i, 2) 'multiply by the third Index/dimension
        End If
    Next i

End Function
```

# Array/ acquire table values
```vb
Sub Populate2DArrayFromExcel()
    Dim ExchangeRates(3, 2) As Variant, xrow As Long, xcol As Long, _
        rowIndex As Long, colIndex As Long, i As Long, j As Long
    rowIndex = 0
    colIndex = 0
    xrow = 10
    xcol = 5

    'outer loop down rows
    Do Until Cells(xrow, xcol).Value = ""

        'inner loop across columns
        Do Until Cells(xrow, xcol).Value = ""
            ExchangeRates(rowIndex, colIndex) = Cells(xrow, xcol)
            colIndex = colIndex + 1
            xcol = xcol + 1
        Loop

        xcol = 5 'reset after done with row loop
        colIndex = 0 'reset after done with row loop

        rowIndex = rowIndex + 1
        xrow = xrow + 1
    Loop

''''''print the Array
xrow = 14
xcol = 10
For i = 0 To UBound(ExchangeRates, 1)
    For j = 0 To UBound(ExchangeRates, 2)
        Cells(xrow, xcol).Value = ExchangeRates(i, j)
        xcol = xcol + 1
    Next
    xcol = 10
    xrow = xrow + 1
Next

End Sub
```

# Array/dynamic Array 列表的大小随使用而变大变小
```vb
Sub OneDDynamicArray()

    Dim city() As String ' with dynamic array there is no Size (i.e. upper bound) in parentheses
    Dim xrow As Long, i As Long
    i = 0
    xrow = 17

    ReDim city(0) ' resize Array to hold 1 String

    Do Until Cells(xrow, 5).Value = ""
        If Cells(xrow, 5).Value = "CA" Then
            city(i) = Cells(xrow, 4).Value
            i = i + 1 ' increase upper bound Of the city Array
            ReDim Preserve city(i) 'resize Array to new upper bound
            'preserve ensure the stored Value will not be reset in redim the Array
        End If

        xrow = xrow + 1
    Loop

    'resize city Array To eliminate the unused last element
    ReDim Preserve city(i - 1)

    'For i = 0 To UBound(city)
        'city(i)
    'Next i

End Sub
```

# msgbox
parameter
    https://msdn.microsoft.com/en-us/library/aa445082(v=vs.60).aspx

```vb
Option Explicit
Sub msgboxExamples()

    Dim x As Integer, response As Integer
    x = 9
    'https://msdn.microsoft.com/en-us/library/aa445082%28v=vs.60%29.aspx?f=255&MSPPError=-2147217396
    'The buttons argument settings
    Call msgbox("hi" & " how are you? x = " & x, vbRetryCancel)
    Call msgbox("hi" & " how are you? x = " & x, 5)
    Call msgbox("hi" & " how are you? x = " & x, vbRetryCancel + vbQuestion) 'add a question Symbol
    Call msgbox("hi" & " how are you? x = " & x, vbRetryCancel + vbQuestion + vbDefaultButton1) 'add first Button as default Button
    Call msgbox("hi" & " how are you? x = " & x, vbRetryCancel + vbQuestion + vbDefaultButton1 + vbSystemModal) 'change the Windows style into system style
    Call msgbox("hi" & " how are you? x = " & x, 2 + vbQuestion + vbDefaultButton1 + vbSystemModal) 'three Button in prompt
    Call msgbox("hi" & " do you want to try? x = " & x, 2 + 16 + vbDefaultButton1 + vbSystemModal) 'change to X Symbol
    response = msgbox("hi" & " how are you? x = " & x, 2 + 16 + vbDefaultButton1 + vbSystemModal)

    'https://msdn.microsoft.com/en-us/library/aa445082%28v=vs.60%29.aspx?f=255&MSPPError=-2147217396
    'Return Values
    If response = 3 Then
        msgbox ("you clicked abort")
    ElseIf response = 4 Then
        msgbox ("you clicked retry")
    Else
        msgbox ("you clicked ignore")
    End If


End Sub
```

# Inputbox  
```vb
'Value Meaning
'0   A Formula
'1   A Number
'2   Text (a string)
'4   A logical value (True or False)
'8   A cell reference, as a Range object
'16  An error value, such as #N/A
'64  An array of values

Sub InputboxDemo()

    Dim numberOfSheets As Integer

    'parameter (displayed string, title, default value, ,,,, number of input type)
    numberOfSheets = Application.InputBox("how many worksheets do you want to add?", "add worksheets", 777, 100, 500, , , 1)

End Sub

Sub FindMaxInRange()

    Dim numberRange As Range
    Dim c As Range, maxvalue As Double, maxaddress As String

    'if the user presses cancel to to the calceled label
    On Error GoTo canceled

    Set numberRange = Application.InputBox("Enter a range of cells to find-max:", "find max", , , , , , 8) '8   A cell reference, as a Range object

    maxvalue = numberRange.Cells(1, 1).Value
    maxaddress = numberRange.Cells(1, 1).Address

    'loop cells in range
    For Each c In numberRange.Cells
        If c.Value > maxvalue Then
            maxvalue = c.Value
            maxaddress = c.Address
        End If
    Next c

    'dispaly max value and its addresss
    msgbox ("the max value in the range is " & maxvalue & " at " & maxaddress)

canceled:

End Sub

```

# find max number in a range
```vb
Sub FindMaxInRange()

    Dim numberRange As Range
    Dim c As Range, maxvalue As Double, maxaddress As String

    'if the user presses cancel to to the calceled label
    On Error GoTo canceled

    Set numberRange = Application.InputBox("Enter a range of cells to find-max:", "find max", , , , , , 8) '8   A cell reference, as a Range object

    maxvalue = numberRange.Cells(1, 1).Value
    maxaddress = numberRange.Cells(1, 1).Address

    'loop cells in range
    For Each c In numberRange.Cells
        If c.Value > maxvalue Then
            maxvalue = c.Value
            maxaddress = c.Address
        End If
    Next c

    'dispaly max value and its addresss
    msgbox ("the max value in the range is " & maxvalue & " at " & maxaddress)

canceled:

End Sub
```

# Event 
```vb
Private Sub Workbook_Open()

MsgBox ("you opened Excel")

End Sub


Private Sub Workbook_SheetActivate(ByVal Sh As Object)

MsgBox ("your new sheet is " & Sh.Name)

End Sub
```

# Event/ Change cells Color 
```vb
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    Dim i As Integer
    
    'msgbox (Sh.Name)
    If Sh.Name = "foreach" Then
    
        'msgbox ("you select row " & Target.Row & " column " & Target.Column)
    
        Cells.Interior.ColorIndex = xlNone 'set all cells to no color
        'loop to Color down a Column
        For i = 1 To Target.Row
            Cells(i, Target.Column).Interior.ColorIndex = 6
        Next i
    
        'loop To color down a row
        For i = 1 To Target.Column
            Cells(Target.Row, i).Interior.ColorIndex = 6
        Next i

    End If
End Sub
```

# error/出错 
1. 名字里带有空格 

# Selection Of Data in Range through formula
formulas -> Name Manager -> new -> Name & refer to `=OFFSET('pivotTable '!$+'pivotTable '!$1:$1048576A$1,0,0,COUNTA('pivotTable '!$A:$A),COUNTA('pivotTable '!$1:$1))` // offset(起始点,0,0,行数,列数)

# PivotTable/ Create A PivotTable 
formulas -> Name Manager -> new -> Name `Data` & refer to `=OFFSET('pivotTable '!$+'pivotTable '!$1:$1048576A$1,0,0,COUNTA('pivotTable '!$A:$A),COUNTA('pivotTable '!$1:$1))` // offset(起始点,0,0,行数,列数)

```vb
Sub MakeAPivotTable()

    Dim pt As PivotTable
    Dim cachePT As PivotCache
    
    Sheets("pivotTable").Select
    ActiveSheet.PivotTables("MyPT").TableRange2.Clear 'delete old PivotTable
    
    'sets source Of data for Pivot Table
    Set cachePT = ActiveWorkbook.PivotCaches.Create(xlDatabase, Range("Data")) 'Create(类型,数据源) Data 是之前通过PivotTable设置的数据源
    
    
    'Create PT
    Set pt = ActiveSheet.PivotTables.Add(cachePT, Range("K1"), "MyPT") '在K1的地方显示Pivot Table,名字为MyPT
    
    With pt
        'set the orientation Of the fields
        .PivotFields("Date").Orientation = xlRowField 'pick all rows under Column Of "Date" field
        .PivotFields("product").Orientation = xlRowField
        .PivotFields("Name").Orientation = xlPageField
        .PivotFields("price").Orientation = xlDataField
        
        'set To classic View
        .RowAxisLayout xlTabularRow
        
        'set format for price
        .DataBodyRange.NumberFormat = "#,##0.00"
        
        'add a calculated field for commission
        .CalculatedFields.Add "Commission", "=price*.1"
        .PivotFields("Commission").Orientation = xlDataField
        
        
    End With

End Sub
```

# PivotTable/ filter the display Of items 
```vb
Sub MakeAPivotTable()

    Dim pt As PivotTable
    Dim cachePT As PivotCache
    
    Dim pf As PivotField
    Dim pi As PivotItem
    
    Sheets("pivotTable").Select
    ActiveSheet.PivotTables("MyPT").TableRange2.Clear 'delete old PivotTable
    
    'sets source Of data for Pivot Table
    Set cachePT = ActiveWorkbook.PivotCaches.Create(xlDatabase, Range("Data")) 'Create(类型,数据源) Data 是之前通过PivotTable设置的数据源
    
    
    'Create PT
    Set pt = ActiveSheet.PivotTables.Add(cachePT, Range("K1"), "MyPT") '在K1的地方显示Pivot Table,名字为MyPT
    
    With pt
        'set the orientation Of the fields
        .PivotFields("Date").Orientation = xlRowField 'pick all rows under Column Of "Date" field
        .PivotFields("product").Orientation = xlRowField
        .PivotFields("Name").Orientation = xlPageField
        .PivotFields("price").Orientation = xlDataField
        
        'set To classic View
        .RowAxisLayout xlTabularRow
        
        'set format for price
        .DataBodyRange.NumberFormat = "#,##0.00"
        
        'add a calculated field for commission
        .CalculatedFields.Add "Commission", "=price*.1"
        .PivotFields("Commission").Orientation = xlDataField
    End With
    
    
    '''''TURN ON only certain items
    Set pf = pt.PivotFields("name") 'sets the Pivot field To Change To the name field
    With pf
        'loop over all the names in the name field
        For Each pi In pf.PivotItems
            If pi.Name = "Bob" Or pi.Name = "Ann" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    

    Set pf = pt.PivotFields("product") 'sets the Pivot field To Change To the product field
    With pf
        'loop over all the names in the product field
        For Each pi In pf.PivotItems
            If pi.Name = "basketball" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next pi
    End With
End Sub
```

# with 语句
With 语句可以对某个对象执行一系列的语句，而不用重复指出对象的名称。例如，要改变一个对象的多个属性，可以在 With 控制结构中加上属性的赋值语句，这时候只是引用对象一次而不是在每个属性赋值时都要引用它。下面的例子显示了如何使用 With 语句来给同一个对象的几个属性赋值。
```vb
With MyLabel
   .Height = 2000
   .Width = 2000
   .Caption = "This is MyLabel"
End With
```
注意 当程序一旦进入 With 块，object 就不能改变。因此不能用一个 With 语句来设置多个不同的对象。

# references = 外部引用的类库 类似 import/include
developer -> visual basic -> tools -> references
developer .] visual basic .] tools .] references

# object browser = 查看所有的对象/类
developer .] visual basic .] bar[object browser]

# 调试工具
- local window = 查看全部变量
- step into
- watch window = 只关心要看的变量
  - right click -> add watch
- breakpoint

https://www.youtube.com/watch?v=SpnWO6GOvrY&list=PL3A6U40JUYCi4njVx59-vaUxYkG0yRO4m&index=11
