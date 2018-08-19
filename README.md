# stixls
Poor man’s kanban (but much more).

What is it:
An excel spreadsheet, with macros, no special references needed, all in VBA, not an add-in

What it does:
- Renders sticky notes from a table, and also updates the table from the notes (if enabled).
- Stricky speaking, it is a visualisation tool, can be used for a kanban or any use of sticky notes.
- You get one “Board” to render the notes on,  but you can define and iterate through several "layouts" to show the data arranged in various ways, hiding them etc. See the demo board.
- Smartly, (if enabled) the notes can pick values from its position to update the [Data] sheet. For example if a note is moved to a “completed” column, the next refresh would update the data field. See the template_data sheet in the demo board to understand how this works.
- Can render and export all layouts in a single
sheet ready for shsring.


Why:
- because even the coolest cloud/mobile apps will not be allowed in corporate environments easily.
- most kanban apps I’ve seen are too rigid with the arrangements of notes (just allow grouping by stages of the task) 
- most postit/whiteboard apps doesn’t link to data or export anything structured out of them.



Features
- two-way updating.. notes to board and board to note
- works in Excel standard, not an add in, etc.
- separated app from data
- version control, in the back it saves snapshots (except when using the fast buttons)
- autoarrange feature gives you a sort of data driven visualisation, but still saves the shape position, so you can manually tweak
- it renders in Excel online and ios app, so technically people can see and update the data or postit from a mobile. No macros on those, unfortunately. 
- export features (File Menu)
- adding a row of data will generate a note next time you do “Table to Board” and a new (copy/paste) note will add a row and generally there are many controls to detect issues, not perfect but works well enough.
- Many rendering options: colour, visible, size, bold font, red text... based on conditions you can setup in the layout, and ultimately you can define in a column in excel with whatever formulae you see fit.
- It can add an image on the top right corner of the note based on the value of a field. You need to paste the image in the sheet [Icons] and ensure the top right corner of the image is over a cell that has that value (e.g. "Red", or "Peter"). The script will pick the value from the [Data] sheet, and then search for it in the [Icons] sheet, and will take care of the rest.

Quick guide to get you going:

1. Open the spreadsheet, you should see a new Add-ins tab in the ribbon (once enabled the content etc). This spreadsheet is not an Add-in btw, but that is how Excel names a new a tab. This ribbon is created and deleted each time the xls is loaded/closed. Needless to say you need to enable the content. You are strongly suggested to check the code to your comfort before running in sensitive envirnments, the spreadsheet is given "as is", no warranties at all on any effect on your equipment.

2. Go to that Add-ins tab and select File-> New Board, it will create a Demo board, fully functional.

3. Play with it, start with using the buttons left /right /reload and autoarrange. The magnifier glass takes you to the corresponding note/row of the row/note selected.

4. Try updating a note in [Board] and hitting refresh (the one with two arrows forming a circle), you should see the row having updated

5. Try updating a row and hitting  [Table to Board]>[Refresh Fully from Table] 

6. Try adding a row and the running [Refresh board from Table] to see what happens. Ctrl+P while shape is selected would attempt to place it as per template (if configured).

7. If you move the cell to the completed column, it should update the status and colour once you hit Refresh button.

8. Try copy-paste a note and update the reference number to a unique string.  Then hit [Board to Table] it will create a new row but tou still need to complete fields, the macro will tell you that.

9. Try copy-paste a row into a new consecutive row (or insert one), try [Refresh Fully From Table] it should give you errors that the row is a duplicate and where is the issue. In short, you need to delete the SHID value and put a unique new value in the REV value in that row. Then hit [Refresh Fully From Table] and it should create a nice shiny note.

10. Note that each board is self contained with their data, main board, layout and templates. There is true separation between the application and the configuration... in Excel... sorcery

11 Eventually, read the documentation.. but no hurry, there is no documentation other than this doc yet.

Some  concepts:
1. Board workbooks have only one [Board] sheet.
2. The application renders in the [Board] sheet any number of layouts, as configured in the Layout_config sheet. Layouts are combination of positioning, colouring, size , visibility , font rules.
3. Each layout is based on one template but one template can be used by many layouts, for example, one Kanban template can be made into a layout that applies a RAG colour based on task status and includes lots of detail text, and another layout that has no colour coding and minimal task descriptions and hiding certain notes.
4. Each template has two sheets, one for the look and feel and one for the data (_data). The template has a range named after the template name, and the template  name needs to be the same as the main template sheet. 
5. The script parses each note as follows: 
      5.1 the first word to be the unique reference (REF column in Data tab) , 
      5.2 (space)
      5.3 Title field (configurable the column where it is saved) 
      5.4 Then each line following the Field:Value format , 
      5.5 Anything not recognised goes to a notes field (configurable)
      5.6 Appended to the note text, it would add the value stored in the equivalent cell of the [template] data sheet. ie a note whose top left corner is in board.cell(2,5) will be parsed and will have appended the text value of sheets(“TemplateName_data”).range(“TemplateName”).cells(2,5). Technically this is incorrect as the range is defined in the [TemplateName] tab but the macro reuses that range for both templatename and templatename_data sheets. 
6. A mandatory hard coded field is SHID which stores the shape ID of the note.
7. Some fields accept labels for colours or size. Check the code, eventually I will document here what are the options. 
8. The retouch menu allows to apply “filters” to the current board but are not permanent. If you want them to be permanent then you need to use a layout 
9. The configuration is stored at app_config, then board_config, then layout_config with increasingly limited scope.

