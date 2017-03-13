# word-vba
This repository contains macros written in Word VBA. Mike Pope is the host.

## Versions of Word
I use Word 2013 for Windows, but much of this applies to earlier versions of Word. Most of it probably applies to Word for the Mac, but I cannot guarantee that.

## Installing macros
To install macros, you need to be able to open the macro editor in Microsoft Word. In Windows, you can do this by pressing Alt+F11 while you're in a document. If you've configured Word to show the Developer tab, you can also use the Code area of the ribbon to manage macros.

Here's a primer on how to use the VBA editing environment and work with macros:

[Getting Started with VBA in Word 2010](https://msdn.microsoft.com/en-us/library/office/ff604039(v=office.14).aspx)

## List of macros
Here's a brief description of what's in this repo:

**`change-plaintext-to-normal.mac`**

> This was written in response to a query to on the Word-Pc list. If you open up a pure-text file in Word, Word formats the text in the file as the style named `Plain Text`. The poster wondered if there was a way to have Word instead show everythign in `Normal`. There might be a better way to do that, but this macro seems to do the trick. It relies on the fact that Word runs `AutoNew`, `AutoOpen`, and `AutoClose` macros if they exist when (respectively) you create, open, and close a document.


**`show-hide-revisions.mac`**

> This macro toggles revision marks. In the language of the Word UI, it's the equivalent of the following:
>
> **Review (tab)** &gt; **Tracking** &gt; [**All Markup**|**No Markup**]
>
> This macro is particularly suitable for mapping to a keystroke. I use  `Alt+V,A`
>
> More info: [http://mikepope.com/blog/DisplayBlog.aspx?permalink=2407](http://mikepope.com/blog/DisplayBlog.aspx?permalink=2407)


**`toggle-smart-quotes.mac`**

>Turns on/off the feature in Word that converts straight quotes to curly quotes. Most useful when this has a keyboard shortcut assigned to it.


### Managing the Styles pane
**`set-styles-panel-to-all-alphabetical.mac`**

>Changes the display in the Styles pane to a) show all styles in b) alphabetical order.


### HTML-related macros
**`convert-all-links.mac`**

>Walks through all Word hyperlinks and converts them to &lt;a&gt; elements (with `target="_blank"`).

>An interesting feature of this macro is that it walks *backward* through the collection of links. The issue is that as each hyperlink is converted, it stops being a hyperlink and is thus removed from the collection. This confuses any code that is walking the collection.   


**`make-bold.mac`


>Adds `<b></b>` tags around the selected text.


**`make-code.mac`


>Adds `<code></code>` tags around the selected text.

