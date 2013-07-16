## <a name="pagetop" href="#pagetop">Loops and Logic</a> ##
|Section    			|Description												|
|:---------------------:|-----------------------------------------------------------|
|[For Next](#for)		|Repeat a block of code a certain number of times			|
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

----------