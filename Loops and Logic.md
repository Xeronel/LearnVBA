## <a name="pagetop" href="#pagetop">Loops and Logic</a> ##
|Section    			|Description												|
|:---------------------:|-----------------------------------------------------------|
|[For](#for)		|Repeat a block of code a certain number of times				|
|[Do Until](#dountil)	|Repeat a block of code until a condition is met			|
|[Do While](#dowhile)	|Repeat a block of code while a condition is true			|
|[If Then Else](#if)	|If a condition is met run a block of code					|
|[Select Case](#selcase)|Choose which block of code to run based on somethings value|
----------

### <a name=for href=#for>For</a> ###
A **For** loop repeats a block of code the specified number of times.

```VB
Sub Example()
    Dim i As Integer    'Create a variable to use as a loop counter

    ' This will repeat the code between 'For' and 'Next' until 'i' reaches 10
    ' 'i' will start at 1 and increase by the step each time next is reached
    For i = 1 To 10 Step 1
        ' Remember i = 1 to 10, so each time this code is run 'i' will increase by one
        ' The end result is the first row counting up from 1 to 10
        Cells(1, i).Value = i   'Set the value of Row 1, Column number 'i' to the value of i
    Next
End Sub
```

Result:
>![Result](./images/For_Result.jpg)

--
For loops are very useful if you want to repeat a block of code on every row or column.

Here is an example that gives rows alternating colors
```VB
Sub Example()
    Dim i As Integer                    'Create a variable to use as a loop iterator
    Dim Usedrange As Range              'Create a variable to access cells on our worksheet
    Set Usedrange = Sheet1.Usedrange    'Set the variable UsedRange to the range containing data on Sheet1

    ' This will repeat the code between 'For' and 'Next' until 'i' reaches
    ' the number of rows in UsedRange
    For i = 1 To Usedrange.Rows.Count Step 1
        'Color even numbered rows light blue and odd numbered rows dark blue
        If i Mod 2 = 0 Then
            Usedrange.Rows(i).Interior.Color = RGB(184, 204, 228)   'Light Blue
        ElseIf i Mod 2 = 1 Then
            Usedrange.Rows(i).Interior.Color = RGB(79, 129, 189)    'Dark Blue
        End If
    Next
End Sub
```

Result:
>![Result](./images/For_Result2.jpg)

<sub>[Go to top](#pagetop)</sub>

----------
### <a name=dountil href=#dountil>Do Until</a> ###
A **Do Until** loop repeats a block of code until a condition is met.

```VB
Sub Example()
    Range("A1").Select                                  'Select cell A1
    Do Until Selection.Value = ""                       'Repeat the code between 'Do' and 'Loop' until a blank cell is selected
        Selection.Interior.Color = RGB(64, 128, 200)    'Color the selected cell blue
        Selection.Offset(1, 0).Select                   'Select the cell below our current selection
    Loop
End Sub
```

Result:
>![Result](./images/DoUntil_Result.jpg)

<sub>[Go to top](#pagetop)</sub>

----------
### <a name=dowhile href=#dowhile>Do While</a> ###
A **Do While** loop repeats a block of code while a condition is true.

```VB
Sub Example()
    Range("A1").Select                                  'Select cell A1
    Do While Selection.Value = "1"                      'Repeat the code between 'Do' and 'Loop' while the value of the selected cell is 1
        Selection.Interior.Color = RGB(64, 128, 200)    'Color the selected cell blue
        Selection.Offset(1, 0).Select                   'Select the cell below our current selection
    Loop                                                'Start next loop iteration
End Sub
```

Result:
>![Result](./images/DoWhile_Result.jpg)

<sub>[Go to top](#pagetop)</sub>

----------
### <a name=if href=#if>If, Then, Else</a> ###
An **If** statment executes a block of code if a statement is true or false.

```VB
Sub Example()
    Const B As Boolean = True   'Create a constant with it's value set to true

    If B = True Then                'If B is True Then
        Range("A1").Value = "True"  'Set A1 to "True"
    End If                          'End our if statement
End Sub
```

Result:
>![Result](./images/If_Result.jpg)

--
You can use an **Else** statement to define what happens if the boolean operation is not true.

```VB
Sub Example()
    Const B As Boolean = False   'Create a constant with it's value set to false

    If B = True Then                    'if B is True then
        Range("A1").Value = "True"      'Set A1 to "True"
    Else                                'if B is not True then
        Range("A1").Value = "False"     'Set A1 to "False"
    End If                              'Let the computer know our if statement is done
End Sub
```

Result:
>![Result](./images/If_Result2.jpg)

--
You an also use an **ElseIf** statement to have an additional logic check if your original boolean operation is not true.

```VB
Sub Example()
    Const i As Integer = 2                      'Create a constant with it's value set to 1

    If i = 1 Then                               'If i = 1 then
        Range("A1").Value = "i = 1"             'Set A1 to "i = 1"
    ElseIf i = 2 Then                           'If i is not 1 and i = 2 then
        Range("A1").Value = "i = 2"             'Set A1 to "i = 2"
    Else                                        'If i is not 1 or 2 then
        Range("A1").Value = "i isn't 1 or 2"    'Set A1 to "i isn't 1 or 2"
    End If                                      'Lets the computer know to end our if statment
End Sub
```

Result:
>![Result](./images/If_Result3.jpg)

<sub>[Go to top](#pagetop)</sub>

----------
