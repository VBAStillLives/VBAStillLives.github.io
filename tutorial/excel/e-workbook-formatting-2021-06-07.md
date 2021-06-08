---
layout: default
permalink: /tutorial/excel/excel-vba-formatting/
---


For this lesson, it's important that you first [download the sample workbook for this lesson](/assets/files/Format_Workbook.xlsm) to follow along. 

We are starting with a single-page workbook, with a very simple data set of consolidated financial results.  The starting workbook should look like this: 

![Sample Worksheet](/assets/images/consolidated_financial_results_excel.png)


Save a copy of this so that if by mistake your code does something you didn't expect, you always have the original as a backup.  This is important to do in the real world as **there is no undo button for actions performed by VBA**.  That's right.  Anything you do in VBA can't be undone in the traditional sense.  There is no CTRL+Z or clicking "undo" in the edit menu.  Keep that in mind. 

In real life, your financial statements are going to be more complex than this, but the same concepts will apply.  Imagine your boss wants you to do perform three separate tasks: 

1. Create second worksheet that includes all data for all the years there was a loss in income. 
2. Create a third worksheet that shows years where the net income changed by equal to or greater than 20%, either up or down. 
3. On the first spreadsheet, highlight any of the annual figures' columns in red where operating expenses grew by more than 30%, either up or down.

Let's do this all in VBA. 

Before we begin, I want to introduce the **Immediate Window**.  The Immediate Window is a helpful tool that is part of the VBA IDE.  It let's us look at results of our code like in a similar, but more efficient, way than the message box.  It let's us see the results of our code while running it and without interrupting our process.  

**How to set up your Immediate Window**

1. [Open your VBA IDE](/tutorial/excel/Excel-VBA-Setting-Up-Dev-Environment/)
2. Click **View** > click **Immediate Window**
![Click View](/assets/images/immediate_window_excel.png)
3. Confirm the **Immediate Window** is visible by looking near the bottom of your IDE Window.
![Immediate Window](/assets/images/immediate_window_display_excel.png)

All code samples below will illustrate use of the Immediate Window.

Now let's get to coding!!

**Create second worksheet that includes all data for all the years there was a loss in income.**

First thing to do is to create a module. Let's name it **modLossYears**.  Copy / paste or type the code below: 

```VBA
Option Explicit

Sub LossYears()

Dim wb As Workbook
Dim ws As Worksheet
Dim wsSource As Worksheet
Dim wsDest As Worksheet
Dim boolLossYearsExists As Boolean
Dim i As Integer, j As Integer

'define wb object variable
Set wb = ThisWorkbook


'define boolean variable
boolLossYearsExists = False


'loop through existing worksheets to see if there is one named "Loss Years"
'only do this if worksheet count is > 1 (to save time)
If wb.Worksheets.Count > 1 Then
    For Each ws In wb.Worksheets
        If ws.Name = "LossYears" Then
            boolLossYearsExists = True
            Exit For
        End If
    Next ws
End If


'if boolLossYearExists = false, then add worksheet with a name "LossYears"
If boolLossYearsExists = False Then
    wb.Sheets.Add.Name = "LossYears"
End If


'define wsSource and wsSource object variables
Set wsDest = wb.Worksheets("LossYears")
Set wsSource = wb.Worksheets("FinPerformance")


'define counter variable
j = 1

'our data columns go from year column 2 to column 12, net income row is 6
'loop through columns 2 through 12, if net income < 0 then copy/paste to wsDest (LossYears)
For i = 2 To 12
    If wsSource.Cells(6, i) < 0 Then
        wsSource.Range(wsSource.Cells(3, i), wsSource.Cells(6, i)).Copy wsDest.Cells(1, j)
        
        'iterate counter variable by one
        j = j + 1
        
        Debug.Print "value of j = " & j
    End If
Next i



End Sub
```


Notice the following line:

```VBA
Debug.Print "value of j = " & j
```

This line will "print" the value of "j" in the Immediate Window.  You can follow the variable along without interrupting your code.  It's a neat feature and helps with coding productivity and quality. 

Notice the "value of j = " in quotations.  This will be printed in the Immediate Window as a literal value.  The "& j" is a way to concatenate the variable "j" to the literal string that precedes it. 

Step through the entire code block to see what happens!  Make some changes or tweaks of your own to reinforce the lesson or expand upon the code's purpose.


**Create a third worksheet that shows years where the net income changed by equal to or greater than 20%, either up or down**

```VBA
Option Explicit

Sub LargeChangeNetIncome()

Dim wb As Workbook
Dim ws As Worksheet
Dim wsSource As Worksheet
Dim wsDest As Worksheet
Dim boolWsExist As Boolean
Dim i As Integer, j As Integer
Dim dResult As Double


'define wb object variable
Set wb = ThisWorkbook


'define boolean variable
boolWsExist = False


'loop through existing worksheets to see if there is one named "Change20Percent"
'only do this if worksheet count is > 1 (to save time)
If wb.Worksheets.Count > 1 Then
    For Each ws In wb.Worksheets
        If ws.Name = "Change20Percent" Then
            boolWsExist = True
            Exit For
        End If
    Next ws
End If


'if boolWsExist = false, then add worksheet with a name "Change20Percent"
If boolWsExist = False Then
    wb.Sheets.Add.Name = "Change20Percent"
End If


'define wsSource and wsSource object variables
Set wsDest = wb.Worksheets("Change20Percent")
Set wsSource = wb.Worksheets("FinPerformance")

'define counter variable
j = 1

'define result variable - we will store the results of the change calculation here for later calculation
dResult = 0

'our data columns go from year column 2 to column 11, net income row is 6 - we stop at column 11 since we can't calculate change before then
'loop through columns 2 through 12, if net income < 0 then copy/paste to wsDest (LossYears)
For i = 2 To 11
    
    'perform the change formula, store results in variable
    dResult = (wsSource.Cells(6, i + 1) / wsSource.Cells(6, i)) - 1
    
    'if the change result is negative, turn it positive (just for calculation purposes), and evaluate if it is >= .2 (20%)
    If dResult < 0 Then
        dResult = dResult * -1
    End If
    
    'if change is greater than 20% year over year, move to new worksheet
    If dResult >= 0.2 Then
        wsSource.Range(wsSource.Cells(3, i), wsSource.Cells(6, i)).Copy wsDest.Cells(1, j)
        
        'add change % from previous year to worksheet, row 4
        wsDest.Cells(4, j).Value = dResult
        
        'iterate counter variable by one
        j = j + 1
        
        Debug.Print "value of j = " & j
        Debug.Print "value of dResult = " & dResult
    End If
    
Next i


End Sub
```

Step through this code to see what it does, line by line.  Tweak anything you think can expand upon this code to reinforce this lesson. 

**On the first spreadsheet, highlight any of the annual figures' columns in red where operating expenses grew by more than 30%, either up or down.**

```VBA


Option Explicit




Sub HighlightLossYears()



Dim wb As Workbook
Dim ws As Worksheet
Dim i As Integer, j As Integer
Dim dResult As Double


'define wb object variable
Set wb = ThisWorkbook

'define ws object variable
Set ws = wb.Worksheets("FinPerformance")


'our data columns go from year column 2 to column 11, net income row is 6 - we stop at column 11 since we cant calculate change before then
'loop through columns 2 through 11, if opex changes by more than 30% then highlight red
For i = 2 To 11
    
    'perform the change formula, store results in variable
    dResult = (ws.Cells(5, i + 1) / ws.Cells(5, i)) - 1
    
    'if the change result is negative, turn it positive (just for calculation purposes), and evaluate if it is > .3 (20%)
    If dResult < 0 Then
        dResult = dResult * -1
    End If
    
    'if change is greater than 20% year over year, move to new worksheet
    If dResult > 0.3 Then
        ws.Range(ws.Cells(3, i), ws.Cells(6, i)).Interior.ColorIndex = 22
    End If
Next i


End Sub
```


Notice the following line of code: 

```VBA
If dResult > 0.3 Then
    ws.Range(ws.Cells(3, i), ws.Cells(6, i)).Interior.ColorIndex = 22
End If
```

How did I know that ColorIndex #22 = red?  Easy!! I used the Immediate Window. See how I did it below: 

![Immediate Window Color Finder](/assets/images/immediate_window_interrogation.png)


I clicked in the cell, highlighted it red, then I went on over to the Immediate Window and added the following code: 

```VBA
?ActiveCell.Interior.ColorIndex
```

Through the Immediate Window, I was able to "read" what the ColorIndex was and apply it in my actual code. It is important to note that you start any and all code in the Immediate Window with a question mark "?."  And it makes sense, really.  You are after all interrogating the program to see what color the cell is. 

**What else can you do?**

My code stopped short in a few areas on purpose.  See if you can do the following:

1. Rearrange tab orders
2. Write code that adds row labels to the worksheet with losses and the worksheet with a >= 20% variance.  

This is great practice!


**What you've learned**

A lot.  This particular lesson offered much more in the ways of code and how you can manipulate Excel with VBA.  Keep practicing and drilling into the code for your particular needs.  Google is your friend.  Also check out sites like Stackoverflow or MrExcel.  These are great sites where you can a ton of info on things you're looking for related to VBA. It is important to note that VBA can do so much more than this.  What I've shown you are the basics in a cohesive and logical order so that you can grow from here, but barely scratches the surface. 


**Note:** I realize this is a simple spreadsheet with simple tasks you probably could have done faster if you'd done it manually.  However, the point here was to expose you to different facets of Excel so that you could apply it in the real world.  If this was a monthly process for a board deck or another process to support month end, it really helps to have something like this automated in VBA.