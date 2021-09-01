---
layout: post
title:  "Get Your Computer to Talk with Excel VBA"
date:   2021-09-1 5:21
categories: VBA
---

# Excel Can Make Your Computer Talk!


Turn the volume up on your PC and let's have some fun!  A little known and seldom used feature of VBA is that you can use it to get your computer to actually speak.  

I know.  I know.  When would this ever be useful in a work setting?  Probably never.  But who says you can't have fun once in a while? 

While it is true that you may never use the "Speak" function in VBA in your career, it is also true that you will probably use the **InputBox** feature in your code.  The **InputBox** allows you to collect short-hand inputs from the user without having to spend the extra time creating your own userform.  We will use the InputBox to prompt the user to enter the words they'd like to have spoken by the computer in this short program. 


```VBA
Option Explicit

Sub ExcelSpeak()


Dim strMessage As String


'input box will appear requiring you to enter text
strMessage = InputBox("Enter message:", "VBA Speak")

'computer will speak the message you wrote in the InputBox
Application.Speech.Speak strMessage


End Sub
``` 


