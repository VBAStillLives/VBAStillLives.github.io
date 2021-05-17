---
layout: default
permalink: /tutorial/excel/excel-vba-variables/
---

Variables are an important and foundational aspect of VBA.  You need to understand them and what they do. 

Ultimately, and simply, variables are used to store a value for later use in a program, and to make coding simpler and more streamlined. 

Let us dive into an example of a variable and how it is used. 

```
Option Explicit

Sub VBA_VariableExample

'Variable declaration - we are declaring using the Dim keyword
Dim i as Integer

'variable definition i.e. i = 1
i = 1

MsgBox i

End Sub
``` 


**Option Explicit**
*Option Explicit* requires that you declare what data type your variable will represent. 

You might recall that in the first part of this tutorial, I had you configure your IDE so that it required variable declaration.  This is signified by the *Option Explicit* at the top of your module.  Declaring variable data types helps organize your code and helps tremendously with debugging later on.  Imagine a case where a generic variable is used to represent a name, a number, and a date.  I have seen this happen in code before.  If the person had defined their data type, it would have been easier to debug.  It turned out that they were trying to add "yes + 2” which did not work for obvious reasons :) 

**Variable Declaration**

The key words *Option Explicit* forces you to declare variables, but how do you do it? 

**You declare variables by using the Dim keyword**.  

Examples:

```
Dim i as Integer
```

Another example:

```
Dim strFullName as string
```


**Variable Names**

You want to name your variables something intuitive and descriptive that gives you information about the variable's data type and value. 

For example:

```
Dim intNumOfEmployees as Integer
```

In a simple 5-line program it might not make much of a difference.  You can see your variable declaration and executable code in one glance.  But now imagine you have written an 800-line program.  Just by looking at something like *intNumOfEmployees* you can tell that it is an integer, and you can infer what it is storing with the "NumOfEmployees."  It makes your code more readable and understandable when you go back to look at it for edits or additions. 


**Readability and Formatting**

This cannot be understated.  Format the name of your variables to conform with best practices.  Always try to include a brief description of the data type, and always capitalize the first letter of each word in a variable name. 

Example: 

```
Dim intProduceCount as Integer
Dim intstaplercount as Integer
```

See the difference? In VBA, nothing stops you from using all lowercase letters to name a variable.  Take it upon yourself to write and format readable code.  You will thank your past self when you inevitably need to revisit your code. 

Variable names must be contiguous (no spaces) and cannot start with a number.  You can use underscores if you want to have the effect of a space between words or to make your code more readable. 


**Variable Data Types**

There is a variety of data types.  In my experience, the most frequently used data types for variable declarations in VBA are string, integer, and long, in no particular order.


Here is a list of data types you will find in VBA: 

```
Dim strString As String
Dim bByte As Byte
Dim iInteger As Integer
Dim lLong As Long
Dim lLongLong As LongLong
Dim dDouble As Double
Dim sSingle As Single
Dim bBoolean As Boolean
Dim vVariant As Variant
Dim dateDate As Date
Dim currCurrency As Currency
```

**Variable Definition**

To define a variable you simply write the variable name, the equal sign, and a value.  Strings **must** be surrounded by double quotes. You define a variable *after* you've declared it and assigned a data type.  Example below: 

```
Dim strFirstName as String
Dim intNumOfEmployees as Integer

strFirstName = "Jon"
intNumOfEmployees = 5
```



**Workbook Sample**
In this workbook sample, we simply use the code from the "Hell World" program in the previous lesson and build upon it with a variable. 

```
Option Explicit

Sub HelloWorld()


Dim strMsg As String


strMsg = "Hello World!"


MsgBox strMsg


End Sub
```

[Download the workbook for this lesson](/assets/files/HelloWorld_variables.xlsm) 

After you open the downloadable workbook and the IDE (Alt + F11), you will notice there are two modules.  One module contains our "Hello World" program using a variable for our message.  The other module contains variable declarations of each data type with some notes for your reference. 

**What you've learned**

* How a variable is defined and used in software
* Variable data types
* How to declare a variable
* How to define a variable


**Extra Practice**

After download and open the workbook that goes with this tutorial lesson, go ahead and change the value of the **strMsg** variable another value and see if you can get it to work.  Change it to **I love VBA.** and see if that works.  If it does, change it to another value of your choosing to see if it works. 


**References and Guides**

[Microsoft's summary of data types can be really helpful](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary).  Check it out!

[Excel Tutorial Homepage](/Excel-VBA-Tutorial/)