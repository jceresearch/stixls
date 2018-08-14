# stixls
kanban for Excel

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
- Each template has two sheets, one for the look and feel and one for the data. The template has a range named after the template name, and the template  name needs to be the same as the template sheet n
