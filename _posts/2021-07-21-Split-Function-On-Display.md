---
layout: post
title:  "Split Function on Display"
date:   2021-07-21 20:44
categories: Excel VBA
---

# Split Function in Excel VBA

Alright guys, this is the first post since completing phase 1 of the tutorials for [Excel](https://www.vbalives.com/Excel-VBA-Tutorial/) and [Access](https://www.vbalives.com/Access-VBA-Tutorial/) VBA.  I suggest you check those out if you need a quick and basic primer.  


Take a minute and download [this workbook](/assets/files/LoopThroughColumn.xlsm).  The code is also included directly in this post as well. 

The setup we have is a common one.  I've simplified the example here to only include one column, but this example and how we work through it will apply to real world scenarios for multiple-column spreadsheets as well.  The scenario plays out as follows - you have a spreadsheet with combined information in one cell, and you'll need to separate it and do some formatting to make use of the information.  This example I'm working with may be simple, but we'll use it to uncover a lot of interesting tips and tricks with Excel VBA, as well as provide some self-learning opportunities and ideas for future posts!

In this example we have a list of first and last names listed in a row in column A:

![Excel List of Names](/assets/images/e-post-list-of-names.png)


Our imaginary boss has asked us to separate the first and last names into their own columns, to highlight each row where the last name begins with "S" (for some reason - just go with it for now), and then she wants you to organize that list alphabetically (A-Z) by last name. Got it? Good. Let's get to work. 

```VBA
Option Explicit

'script purpose: separate first and last name into different columns; highlight row where last name starts with letter "S"; sort in descending order by first name

'Do not change the name of "Sheet1" without adjusting your code where I set wsSource.  Code will not function properly.

Sub HighlightSort()

Dim i As Integer
Dim lRow As Long
Dim ws As Worksheet
Dim wsSource As Worksheet
Dim wsFinal As Worksheet
Dim boolWorksheetExists As Boolean
Dim strFullName As String
Dim strFirstName As String
Dim strLastName As String
Dim strFirstLetter As String


'existing source worksheet
Set wsSource = ThisWorkbook.Worksheets("Sheet1")

'==============
'results worksheet, check if exists, if not then create, else set wsFinal to existing "Final" worksheet
'==============
boolWorksheetExists = False

For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "Final" Then boolWorksheetExists = True
Next ws

If boolWorksheetExists = False Then
    Set wsFinal = ThisWorkbook.Worksheets.Add
    wsFinal.Name = "Final"
Else
    Set wsFinal = ThisWorkbook.Worksheets("Final")
End If

'clear worksheet (in case of previous formatting; best practices); format final worksheet with column names
'the WITH block lets me code without having to type the same objects every time.
'Notice I start each line in the "with" block with a period "." - this acts like I typed "wsFinal" before each line.
'WITH blocks are completely optional

With wsFinal
    .Cells.Clear
    .Range("A1").Value = "First Name"
    .Range("B1").Value = "Last Name"
End With
'==============

'get last row in source worksheet
lRow = wsSource.Range("A" & wsSource.Rows.Count).End(xlUp).Row


'loop through rows in column 1 (column A) to highlight last names that start with the letter "S"
For i = 2 To lRow
    'variabilize the string for simpler coding
    strFullName = wsSource.Cells(i, 1).Value
    
    'extract first name and last name - notice the use of the split function - great opportunity for some Google searching and self-learning!!
     strFirstName = Split(strFullName, " ")(0)
     strLastName = Split(strFullName, " ")(1)

    'get first letter of last name
    strFirstLetter = Left(strLastName, 1)
    
    'place separated first and last name in "Final" worksheet
    With wsFinal
        .Cells(i, 1).Value = strFirstName
        .Cells(i, 2).Value = strLastName
    End With
    
    'highlight rows where first letter of last name = "S" in "Final" worksheet
    If UCase(strFirstLetter) = "S" Then
        wsFinal.Range(wsFinal.Cells(i, 1), wsFinal.Cells(i, 2)).Interior.ColorIndex = 6
    End If
    
Next i


'order the list of names, alphabetically, by last name; autofit columns for neatness

''NOTE: I got this code by using the macro recorder!! You can modify this code to your use.  The code below can certainly be more elegant, and much shorter,
'''but in a pinch, you can use the recorded code so long as it is flexible and fits your need...ALWAYS TEST THE CODE.  It might work the first time, but then break after a new condition is added - this case, a new name.  Try it to see if it works.  It did in my case.
''''the code block for sorting can be reduced to 2 lines - opportunity for self-learning!!

'=======================
'macro recorder code for sorting (slightly modified)
'=======================
wsFinal.Range("A1:B1").AutoFilter
ActiveWorkbook.Worksheets("Final").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Final").AutoFilter.Sort.SortFields.Add2 Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Final").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
'=======================


'autofit code - do this last to account for the drop down arrow width
wsFinal.Cells.EntireColumn.AutoFit

'activate the "Final" worksheet so that it appears when you look at your workbook
wsFinal.Activate

'optional - this tells you your code is complete
MsgBox "Complete!", vbInformation, "VBA Lives"



'bonus round / self learning opportunities
'- add more names to the list in sheet 1, see what happens
'- what other formatting can you do? based on what criteria? try it!

End Sub

```

I highly suggest you type this code yourself.  It helps commit the various commands to memory.  When coding it is important to remember that there are usually so many ways to do any one task.  But you need to start somewhere, and you need to have those methods at your disposal.  There is no shortcut to memory but practice, and sooner or later this will become second nature. 

When you've executed the code and it executes correctly, you should get results that appear like this: 
![Results](/assets/images/split-function-result-0721.png)


As always, I include a ton of documentation and notes inside the code itself. Any text that follows a single quote is a comment in VBA and will not be executed by the compiler.  It is a best practice to document blocks of code, and not only will your technically inclined support staff or colleagues thank you for doing this, but you'll thank your past self when you come back and look at some code you've written a year later. It helps to more quickly understand what your intention was when you were coding the script. 



A few tings  to note: 

I used the macro-recorder for the part of the code that sorts the list by alphabetical order.  The macro-recorder generally gives us garbage code, but that doesn't mean it isn't useful.  We will often have to modify the code, but it works in giving us the appropriate commands when we are in a pinch.  The code for sorting, for example, can be done in one or two lines.  And I could have refined what the macro-recorder gave me.  For this example though, I wanted you to see me use it in action.  

Also, check out the **with** block.  The **with** block allows us to access the objects and commands of parent objects without having to retype a parent object over and over again. 

For example:

```VBA
With ThisWorkbook.Worksheets("Sheet1")
	.range("A1").value = "hi"
	.range("A1").colorindex = 10
End With

'versus

ThisWorkbook.Worksheets("Sheet1").range("A1").value = "hi"
ThisWorkbook.Worksheets("Sheet1").range("A1").colorindex = 10
```

See the difference?  I could even take it a step further.  Notice how each line in the **with** block references range(A1)?

```VBA
With ThisWorkbook.Worksheets("Sheet1").Range("A1")
	.value = "hi"
	.colorindex = 10
End With
```

See the difference, again? I don't want to get hung up on this.  It's a way to organize and shorten the code you write, but not necessary.


 
## Self-Learning Opportunities

1. Try to use the macro-recorder yourself.  
2. Can you make the highlight color dependent on a user selection i.e. the user wants to higlight using red instead of yellow? 
3. Can you make the condition for highlighting a row dependent on a user selection i.e. the user wants to highlight last names that start with "B?"

