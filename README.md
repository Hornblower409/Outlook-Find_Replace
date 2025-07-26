# Outlook-Find_Replace
Outlook VBA Macro to Find and Replace text in the current Selection of an Outlook Item. This is an attempt to compensate for the fact that Outlook does not have the ability to "save" a Find/Replace so you can use it again.

## Install
Click on the "ReplaceInSelectionModule.bas" file link in the files list above. Once it is open, click the "..." (More Actions) menu icon in the upper right and choose "Download". After it has downloaded, open the Outlook VBA Editor (Alt+F11) and from the VBA Editor Menu choose:

File -> Import
{Point it to the downloaded .bas file}.

Tools -> References ...
{Find "Microsoft Word ... Object Library" in the list and check it}

Debug -> Compile ...

File -> Save VbaProject.OTM

For help on using the VBA Editor, Self Signing and running Macros, and adding Macros to your Quick Access Toolbar or Ribbon see the Slipstick Systems web site article: [How to use Outlook's VBA Editor](https://www.slipstick.com/developer/how-to-use-outlooks-vba-editor/)

## Create a new Find/Replace macro
Using the VBA Editor, copy/paste a duplicate of the Main Sub, rename the Sub and modify the two paramaters in the Word_ReplaceInSelection function call. e.g.

Public Sub ReplaceInSelection_XXwithYY()
    Word_ReplaceInSelection "XX", "YY"   
End Sub
