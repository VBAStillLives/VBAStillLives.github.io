---
layout: default
permalink: /tutorial/excel/Excel-VBA-Setting-Up-Dev-Environment/
---

After reviewing the [VBA Orientation](https://vbastilllives.github.io/tutorial/excel/Excel-VBA-GettingStarted/), we are now ready to set up our development environment.  By the end of this section, you'll have navigated to and set up your IDE, made an important configuration, and created your firt module. 

**Steps**

1. Open an Excel workbook (easy enough!)

2. Navigate to the **File** tab
![Navigate to File](/assets/images/1e_click_file_getting_started.png)

3. Select **Options**
![Select Options](/assets/images/2e_click_options_getting_started.png)

4. Select **Customize Ribbon**, then enable the **Developer Tab** and click **OK**
![Developer Tab](/assets/images/3e_click_dev_tab_gettingstarted.png)

5.  After completing the previous step, you should be able to see a new tab in your Excel workbook titled **Developer**

6. Navigate to the **Developer** tab, then click the **Visual Basic** button in the ribbon
![Navigate to Visual Basic](/assets/images/4e_dev_tab_getting_started.png)

7. You should see an IDE screen that appears like this (or something similar)
![VBA IDE](/assets/images/5e_ide_getting_started.png)

8. Within the IDE navigation bar, click **Tools > Options** 
![Navigate to Options](/assets/images/6e_nav_options_getting_started.png)

9. Once the Options Toolbox opens up, select **Require Variable Declarations**.  After you have enabled **Require Variable Declarations**, select **OK**.  This will require that you declare all your variables going forward (more on this later), and this rule will apply to all Modules going forward. If this step was completed successfully, you should see the words **Option Explicit** on the top of every module you open.  This might not make sense now, but it is a configuration I want you to set and I'll explain why and what it does in a later post. 
![Require Variable Declarations](/assets/images/7e_var_declaration_getting_started.png)

10. Navigate to **Insert** in the IDE navigation bar and select **Module**.  This is where you will type code!! We did it!!
![VBA Module](/assets/images/8e_insert_module_getting_started.png)

You'll notice in the left hand navigation pane there is a new directory with a module in it.  You'll see where you can rename the module to something more meaningful to you and your project, and you'll also see a blank white IDE with the words **Option Explicit** near the top as discussed in step 9. 
![Module](/assets/images/9e_ide_loaded_getting_started.png)




**Quick launch of the IDE**

You could skip steps 2 - 7 with a simple shortcut key, ALT + F11.  Pressing ALT+F11 will pull up your IDE screen.  You'll still want to do steps 8 - 10 if it's your first time loading up your IDE.  Every subsequent time you can simply press ALT + F11, insert a module, and go straight to coding. 
![Alt + F11](/assets/images/altf11.png)

**What you've accomplished**
* Set up the Developer Tab
* Opened your IDE
* Configured your IDE to require variable declarations
* Learned how to insert a module where you can begin to write code

**What's next?**

You are now ready to create your first program in VBA.  Let's do it!

[Excel Tutorial Homepage](/Excel-VBA-Tutorial/)

[First VBA Program]

