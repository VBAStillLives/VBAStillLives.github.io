---
layout: default
permalink: /tutorial/excel/excel-vba-conditionallogic/
---

Conditional logic is a fundamental concept you'll need to effectively program in any software language.  

Simply put, conditional logic is the process of decision making by conditions.  We use this all the time in our daily lives.  

> **If** there is a burger restaurant within 1 mile of me, **then** let's get hamburgers, **or else** let's grab tacos. 

The above example is conditional logic in action.  Your decision to get hamburgers is based on the *condition* that there is a hamburger restaurant within 1 mile of your current location.  If there is no hamburger restaurant within 1 mile of your location, then you'll eat tacos instead.  

In a work setting, it might play out as follows: 

> **If** sales comes in higher than budget, **then** color cell green, **if** sales figures are equal to budget, **then** color cell yellow, **else** color cell red.  

Get it?  Conditional logic forms the basis for how you to tell your script what to do when a condition is true. 

There are two major ways you can use conditional logic in VBA. 

1. IF statement - Most frequently used (by me anyway).  Best used in circumstances where you are evaluating one statement at a time.  
2. SELECT CASE statement - Best used where you are evaluating many different possibilities in your conditional logic. 


**Here is a look at an IF statement:**

```
Option Explicit

Sub If_Example()

'declare and define variable strCellContent
Dim strCellContent As String

'defined by the value contained in Sheet1, cell A1
strCellContent = ThisWorkbook.Worksheets("Sheet1").Range("A1").Value

'if the variable strCellContent = x, then print y, else print z
If strCellContent = "x" Then
    MsgBox "y"
Else
    MsgBox "z"
End If


End Sub
``` 

If the value of Sheet1, cell A1 = x, then you'll get a message box that reads the value "y."  If it's anything other than "x," you'll get a message box that reads the value "z."

Very simple. Play around with it and modify the statement to test different conditions. 


**Now let's look at a SELECT CASE statement**

```VBA
Option Explicit

Sub Case_Example()

'declare and define variable strCellContent
Dim strCellContent As String

'defined by the value contained in worksheet1, cell A1
strCellContent = ThisWorkbook.Worksheets("Sheet1").Range("A1").Value

'select cases statement evaluating multiple conditions
Select Case strCellContent
    Case Is = "x"
        MsgBox "a"
    Case Is = "y"
        MsgBox "b"
    Case Is = "z"
        MsgBox "c"
    Case Else
        MsgBox "not x, y, or z"
End Select


End Sub	 
```

This script is mostly set up the same way.  Take a look at the SELECT CASE statement and how many conditions I'm evaluating.  The SELECT CASE statement let's me evaluate multiple conditions very quickly an easily.  I can add **CASE IS** statements for all 26 letters of the alphabet if I wanted.  And the convenient **CASE ELSE** is a catch all for anything that doesn't match.  

[You can download the sample file for this lesson here](/assets/files/Conditional_Logic.xlsm)


**What you've learned**

We covered the basics of conditional logic, and two different types of conditional logic statements in VBA - the IF statement, and the SELECT CASE statement.  

**What's next**

The next lesson will tie all the previous concepts together and format an Excel worksheet as one might be expected to do in an actual working environment. 

**Navigation**

[Excel Tutorial Homepage](/Excel-VBA-Tutorial/)
