# bkmk4pdftk
While pdftk is great at adding bookmarks to a pdf, many (including me) find it's format tedious. This project tries to generate a pdftk format bookmark file from a table in google sheets using GAS (Google Sheet Scripts). Simply recreate the table of contents in one tab, run a script, and bam, you can download the pdftk file to add onto your pdf. 

Basically the polar opposite of [this](https://www.excelforum.com/excel-general/1056038-is-there-a-way-to-convert-the-index-bookmarks-of-a-pdf-file-to-excel.html). With Google Sheets (cause idk vba and I suck as a coder). 

# What does it do?
I find it easier to copy the Table of Contents from a pdf [by Alt + Click -> Drag to select the text and copying it on a table](https://i.imgur.com/CYP0NnK.gif). 
Then, use the macro to gather the data and export it to a pdftk bookmark formatted text file. See gif below for what this macro does.

![see this](https://i.imgur.com/lQX2AQy.gif)



# How to Use It?
Click on this [link](https://docs.google.com/spreadsheets/d/12ZUxQpqf-XuwGNzusuPX34YKGFjB3KfXHapLq6fJTOw/edit?usp=sharing) to view and duplicate the spreadsheet with the macro scripts. 

<details>
        <summary>Or follow these instructions to create your own spreadsheet with the GAS. </summary>
Create a **new** Google spreadsheet, click Tools > Script editor... then copy and paste the contents of the interday.gs file (see above) into the script editor and save. Return to the spreadsheet and refresh the page (Note: actually click the refresh button or select it from the menu; the keyboard shortcut is overriden on Google Sheets, at least in Google Chrome). A couple seconds after the page reloads you should see a "Fitbit" menu at the top.

</details>
