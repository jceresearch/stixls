# stixls
Poor man’s kanban tool, in Excel (but much more).

What is it:
An excel file, with macros, no special references needed, all in VBA, not an add in

What it does:
1. Renders sticky notes from a table, and also updates the table from the notes.
2. Stricky speaking it is a visualisation tool, can be used for kanban or any layout
3. You get one “Board” to render the notes,  but you can define and iterate through several layouts to show the data arranged in various ways, hiding them etc. See the demo board.
4. Smartly (if allowed) the notes can pick values from its position to update the data table. For example if a note is moved to the “completed” column, the next refresh would update the data field. See the template_data sheet in the demo board to understand how this works.
5. Can iterate and generate a stand alone static export of the layouts ready for shsring.


Why:
- because even the coolest cloud/mobile apps wont be allowed in corporate environments easily.
- most kanban apps I’ve seen are too rigid with the arrangements of notes (stages of the task) 
- most postit/whiteboard apps doesn’t link to data or export anything structured out of them.



Features
- two way updating.. notes to board, board to note
- works in excel standard, not an add in, etc.
- separated app from data
- version control, in the back it saves snapshots (except when using the fast buttons)
- autoarrange feature gives you a sort of data driven visualisation, but still saves the shape position, so you can manually tweak
- it does render in excel online and ios, so technically people can see and update the data or postit from a mobile. No macros in O365 unfortunately. 
- export feature 
-  adding a row of data will generate a note next time you do “table to board” and a new (copy/paste) note will add a row and generally there are many controls to detect issues, not perfect but works well enough.
- lots of rendering options, colour, hide, size, bold font, red text... based on conditions you can setup in the layout, and ultimately you can define in a column in excel with whatever formulae you see fit.
- 


How to use:
a) open the spreadsheet, you should see a new Add-ins tab in the ribbon (once enabled the content etc). This spreadsheet is not an Add-in btw, but that is how Excel names a new a tab.

b) go to that Add-ins tab and select File-> New Board, it will create a demo board, fully functional.

c) play with it, the buttons left right, reload and autoarrange should be your first port of call

d) Note that each board is self contained with their data, main board, layout and templates. There is true separation between the application and the configuration... in Excel... sorcery

e) eventually, try to read the documentation.. but not to worry, there is no documentation yet.

Some  concepts:
- Boards have only one board tab
- The application renders there any number of layouts. These are combination of positioning, colouring, size , visibility , font rules
- Each layout is based on one template, one template can be used by many layouts, for example, one Kanban template can be made into a layout that applies a RAG colour based on task status and includes lots of detail text, and another layout that has no colour coding and minimal task descriptions
- Each template has two sheets, one for the look and feel and one for the data. The template has a range named after the template name, and the template  name needs to be the same as the template sheet. 
- 
