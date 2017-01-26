# word-vba
VBA macros for Microsoft Word (Windows)

This repository contains macros written in Word VBA. 

## Installing macros
To install macros, you need to be able to open the macro editor in Microsoft Word. In Windows, you can do this by pressing Alt+F11 while you're in a document. If you've configured Word to show the Developer tab, you can also use the Code area of the ribbon to manage macros.

Here's a primer on how to use the VBA editing environment and work with macros:

[Getting Started with VBA in Word 2010](https://msdn.microsoft.com/en-us/library/office/ff604039(v=office.14).aspx)

## List of macros
Here's a brief description of what's in this repo:

**`change-plaintext-to-normal.mac`**

    This was written in response to a query to on the Word-Pc list. If you open up a pure-text file in Word, Word formats the text in the file as the style named `Plain Text`. The poster wondered if there was a way to have Word instead show everythign in `Normal`. There might be a better way to do that, but this macro seems to do the trick. It relies on the fact that Word runs `AutoNew`, `AutoOpen`, and `AutoClose` macros if they exist when (respectively) you create, open, and close a document.









