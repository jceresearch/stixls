# stixls
Poor man’s kanban (but much more).

What is it:
An excel spreadsheet, with macros, no special references needed, all in VBA, not an add-in

What it does:
1. Renders sticky notes from a table, and also updates the table from the notes (if enabled).
2. Stricky speaking it is a visualisation tool, can be used for kanban or any layout
3. You get one “Board” to render the notes,  but you can define and iterate through several layouts to show the data arranged in various ways, hiding them etc. See the demo board.
4. Smartly,
 (if enabled) the notes can pick values from its position to update the [Data] sheet. For example if a note is moved to a “completed” column, the next refresh would update the data field. See the template_data sheet in the demo board to understand how this works.
5. Can render and export all layouts in a single
sheet ready for shsring.


Why:
- because even the coolest cloud/mobile apps wont be allowed mmin corporate environments easily.
- most kanban apps I’ve seen are too rigid with the arrangements of notes (stages of the task) 
- most postit/whiteboard apps doesn’t link to data or export anything structured out of them.



Features
- two-way updating.. notes to board and board to note
- works in Excel standard, not an add in, etc.
- separated app from data
- version control, in the back it saves snapshots (except when using the fast buttons)
- autoarrange feature gives you a sort of data driven visualisation, but still saves the shape position, so you can manually tweak
- it renders in Excel online and ios app, so technically people can see and update the data or postit from a mobile. No macros on those, unfortunately. 
- export feature 
-  adding a row of data will generate a note next time you do “table to board” and a new (copy/paste) note will add a row and generally there are many controls to detect issues, not perfect but works well enough.
- lots of rendering options, colour, hide, size, bold font, red text... based on conditions you can setup in the layout, and ultimately you can define in a column in excel with whatever formulae you see fit.
- It can add an icon on the top right corner od the note based on the value of a field. See demo board, [Icons] sheet.


How to use:
a) open the spreadsheet, you should see a new Add-ins tab in the ribbon (once enabled the content etc). This spreadsheet is not an Add-in btw, but that is how Excel names a new a tab.

b) go to that Add-ins tab and select File-> New Board, it will create a demo board, fully functional.

c) play with it, the buttons left right, reload and autoarrange should be your first port of call

d) then try adding a note by adding a row and the running “refresh board from Table”

e) then try copy-paste a note and update the reference number to a umique string.  Then hit refresg data from board. it will create a new row but tou still need to
complete fields, the macro will tell you that.

e) Note that each board is self contained with their data, main board, layout and templates. There is true separation between the application and the configuration... in Excel... sorcery


f) eventually, read the documentation.. but no hurry, there is no documentation yet.

Some  concepts:
- Board workbooks have only one [Board] sheet.
- The application renders in the Board any number of layouts, as configured in the Layout_config sheet. Layouts are combination of positioning, colouring, size , visibility , font rules.
- Each layout is based on one template but one template can be used by many layouts, for example, one Kanban template can be made into a layout that applies a RAG colour based on task status and includes lots of detail text, and another layout that has no colour coding and minimal task descriptions and hiding certain notes.
- Each template has two sheets, one for the look and feel and one for the data (_data). The template has a range named after the template name, and the template  name needs to be the same as the main template sheet. 
- The script parses each note as this: expects the first word to be the unique reference (REF column in Data tab) , the  the Title field (configurable) then each line following the Field:Value format , anything not recognised goes to a notes field (configurable)
Also would pick the value stored in the equivalent cell of the [template]_data sheet ie a note whose top left corner is in board.cell(2,5) will be parsed and appended tge value of sheets(“templatename_data”).range(“templatename”).cells(2,5). Technically this is incorrect as the range is defined in the [templatename] tab but the macro reuses that range for both template sheets.
      
- the other mandatory hard coded field is SHID which stores the shape ID of the note.

- Some fields accept labels for colours or size. 

- The retouch menu allows to apply “filters” to the current board, while if you want them to be permanent then you need to use a layout 

- The configuration is stored at app_config, then board_config, then layout_config with increasingly limited scope.

