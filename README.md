# stixls
Poor man’s kanban (but much more).

What is it:
- An excel spreadsheet, with VBA macros, no special references needed, not an add-in

You are strongly suggested to check the code to your comfort before running in sensitive envirnments, the spreadsheet is published "as is", no warranties at all on any effect on your data or equipment.

What it does:
- Renders sticky notes from a table, and also updates the table from the notes (if enabled).
- Stricky speaking, it is a visualisation tool, can be used for a kanban or any use of sticky notes.
- You get one “Board” to render the notes on,  but you can define and iterate through several "layouts" to show the data arranged in various ways, hiding them etc. See the demo board.
- Smartly, (if enabled) the notes can pick values from its position to update the [Data] sheet. For example if a note is moved to a “completed” column, the next refresh would update the data field. See the [Template1Data] sheet in the demo board to understand how this works.
- Can render and export all layouts in a single
sheet ready for shsring.


Why:
- because even the coolest cloud/mobile apps will not be allowed in most corporate environment.
- most kanban apps I’ve seen are too rigid with the arrangements of notes (just allow grouping by stages of the task) 
- most postit/whiteboard apps doesn’t link to data or export anything structured out of them.



Features
- Automatically creates a toolbar,  menu and icons to navigate the options. 
- Configurable, you dont need to tweak VBA. That said you need to be quite comfortable with Excel and its inners.
- Two-way updating (notes > board and board > note).
- Works in Excel standard, it is not an add in
- Separate app (the one you download here) from data (Board workbooks).
- Version control, it saves snapshots when doing "save to table" menu option (Note: not when using the fast buttons)
- Autoarrange notes feature gives you data driven positioning, but still saves each shape position, so you can manually tweak
- Renders in Excel 365/iOS app, so users can see/update the data or postit from a mobile.  
- Export features (File Menu) to generate static rendering of all the layouts and the data.
- If you add a row of data it will create  a note next time you do “Table to Board” 
- If you copy>paste a note, it will add a row, thought you need to ensure all fields are valid, not perfect.
- Many rendering options: colour, visible, size, bold font, red text... based on conditions you can setup in the layout, and ultimately you can define in a column in excel with whatever formulae you see fit.
- Add an image on the top right corner of the note based on the value of a field. You need to paste the image in the sheet [Icons] and ensure the top left corner of the image is over a cell that has that value (e.g. "Red", or "Peter"). The script will pick the value from the [Data] sheet, and then search for it in the [Icons] sheet, and will take care of the rest.

Quick guide to get you going:

1. Open the spreadsheet, you should see a new Add-ins tab in the ribbon (once enabled the content etc).  
2. Go to that Add-ins tab and select File-> New Board, it will create a Demo board, fully functional.
3. Navigate with the buttons left /right /reload and autoarrange. 
4. Select a note and click the magnifier glass to take you to the corresponding note/row 
5. Select a row in teh [Data] tab, and click the magnifier glass to go to the corresponding note.
6. Update the text of a note in [Board] and hit save button (the one with arrows curving down), you should see the corresponding row updated. 
7. Try updating a row and hitting  [Table to Board]>[Refresh Fully from Table] or the revert button (the one with arrow curving upwards), you should see the note reflecting that change.
8. Try adding a row (copy>paste from existing) and the running [Refresh board from Table] to see what happens. 
    You should get errors if the REF and ID are duplicates.
    Eventually you should get a new shape
    Select the shape and hit Ctrl+P , the macro will attempt to place it. See below how this is meant to work.
9. If you move the cell to the completed column, it should update the status and colour once you hit Refresh button.
10. Try copy-paste a note and update the reference number to a unique string.  Then hit [Board to Table] it will create a new row but tou still need to complete fields, the macro will tell you that.
11. Try the export feature under the custom File menu. You can create a full render of all the layouts and the data.

Some  concepts:
0. This spreadsheet is not an Add-in, even if the toolbar name says so. This toolbar/ribbon is created and deleted each time the xls is loaded/closed. 
1. Board workbooks have only one [Board] sheet.
2. The application renders in the [Board] sheet any number of "layouts". Layouts are a combination of the notes positioning, criteria for colouring, size , visibility , font rules, name, and other parameters that compose one "way of showing" the notes. They are configured in the [Layout_config] sheet
3. Each "layout" needs to refer to one "template".  One template can be used by many layouts, for example, you can define one Kanban template, but have several layouts: One for RAG status, another where there is no colour coding,  the note has lots of data, another that hides "low priority" notes.
4. Each template has two sheets, one for the look and feel [TemplateName] and one for the data [TemplateNameData]. The template has to have the actual displayable area as a range named after the template name, and the template name needs to be the same as the main template sheet. See the demo board to see how it is meant to work. 
5. The script parses each note as follows:
    1. the first word to be the unique reference (REF column in Data tab) , 
    2. (space)
    3. Title field (configurable the column where it is saved) 
    4. Then each line following the Field:Value format  
    5. Anything not recognised goes to a notes field (configurable)
    6. Scripts appends the value stored in the equivalent cell of the [TemplateNameData]  sheet. ie a note whose top left corner is in board.cell(2,5) will be parsed and will have appended the text value of sheets(“TemplateNameData”).range(“TemplateName”).cells(2,5). Technically this is incorrect as the range is defined in the [TemplateName] tab but the macro reuses that range for both templatename and templatename_data sheets.  Therefore, values picked by positioning overrides the value on the note itself as they are picked last.
6. SHID is mandatory column, where the script store the shape ID of the note.
7. Some fields accept labels for colours or size. Check the code, eventually I will document here what are the options. 
8. The retouch menu allows to apply “filters” to the current board but are not permanent. If you want them to be permanent then you need to use a layout 
9. The configuration is stored at app_config, then board_config, then layout_config with increasingly limited scope.
10. The sheet DataXY saves the cordinates of the shape, it also saves the size and zorder btw. Currently it would remember changes in size but does not handle zorder yet.
11. Auto positioning of shapes works based on matching the content of the field LayoutAutoArrange in [LayoutConfig], looks for that field and tries to find that value somewhere in the [TemplateNameData] sheet. Things are a bit more complex than that, but that is the gist.
