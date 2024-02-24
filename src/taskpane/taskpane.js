/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// let count = 0;

// Links taskpane buttons to a function
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    // document.getElementById("test-btn").onclick = testmsg;
    document.getElementById("header-btn").onclick = createHeader;
    document.getElementById("cover-btn").onclick = insertCoverPage;
    // document.getElementById("test-btn").onmousedown = testmsg;
    /*document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("add-style").onclick = () => tryCatch(addStyle);
    document.getElementById("count").onclick = () => tryCatch(getCount);
    document.getElementById("add-style-list").onclick = () => tryCatch(addStyleList);
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("insert-control").onclick = () => tryCatch(insertControl);
    document.getElementById("find-content-controls").onclick = () => tryCatch(findContentControls);
    document.getElementById("update-document-details").onclick = () => tryCatch(updateDocumentDetails);
    document.getElementById("create-header").onclick = () => tryCatch(createHeader);*/
    //document.getElementById("create-header-json").onclick = () => createHeaderFromJson;
  }
});

/*async function setStyleSectionHeader(event) {
  await Word.run(async (context) => {
    //Get selected range and expand it to include the whole first and last paragraphs
    var selection = context.document.getSelection().getRange();
    var firstParagraph = selection.paragraphs.getFirstOrNullObject();
    var lastParagraph = selection.paragraphs.getLastOrNullObject();
    var updatedSelection = selection.expandTo(firstParagraph.getRange()).expandTo(lastParagraph.getRange());
    //load the paragraphs and await sync
    updatedSelection.paragraphs.load();
    await context.sync();

    //console.log(updatedSelection.text)
    updatedSelection.style = "test2";
    await context.sync();

    //Move the cursor to the end of the selection
    updatedSelection.paragraphs.getLast().getNextOrNullObject().select("Start");
  });
  event.completed();
}
Office.actions.associate("setStyleSectionHeader", setStyleSectionHeader);*/

// Placeholder for ribbon buttons
async function placeholder(event) {
  event.completed();
}
Office.actions.associate("placeholder", placeholder);

// Example function for ribbon buttons
async function test(event) {
  await Word.run(async (context) => {
    let selection = context.document.getSelection(); // Get cursor location or highlighted text
    context.load(selection)
    await context.sync(); // Wait for Word to return the selection
    // Print "WARNING" at end of cursor/highlight location
    selection.insertText("WARNING", Word.InsertLocation.end);
    // Set color to red
    selection.font.color = "red";
    await context.sync();
  });
  event.completed();
}
Office.actions.associate("test", test);

// Example function for taskpane buttons
async function testmsg() {
  await Word.run(async (context) => {
    const body = context.document.body; // Get the entire document body
    // Print "Test." at the start of the document
    body.insertParagraph("Test.", Word.InsertLocation.start);
    await context.sync();
  });
}

// Function to insert an attachment placeholder at the current selection (mouse cursor) position
async function insertAttachment(event) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const contentControls = body.contentControls;
      contentControls.load("items"); // Load the items property of the contentControls collection
      await context.sync();

      const figureCount = contentControls.items.length + 1;
      const attachmentPlace = context.document.getSelection().insertContentControl();
      attachmentPlace.title = "Attachment";
     
      const paragraph = attachmentPlace.insertParagraph("Figure " + figureCount + ", [Figure Title]", "After");
      paragraph.font.name = "Arial"; // Set font to Arial
      paragraph.font.size = 12; // Set font size to 12

      await context.sync();
    });
    event.completed();
  } catch (error) {
    console.error("Error occurred:", error);
    event.completed({ error: "An error occurred while inserting the attachment." });
  }
}

// Register the function with Office actions
Office.actions.associate("insertAttachment", insertAttachment);


// Inserts a Note template into the document
async function note(event) {
  await Word.run(async (context) => {
    let selection = context.document.getSelection(); // Get cursor location or highlighted text
    context.load(selection)
    await context.sync(); // Wait for Word to return the selection
    selection.insertBreak("Line", "After"); // Line Break
    const box = selection.insertTable(1, 1, "After"); // Insert 1x1 table for border
    box.getBorder(Word.BorderLocation.outside).type = "Single"; // Set border style to single
    box.getBorder(Word.BorderLocation.all).width = 1; // Thicken the border to 1pt
    const head = box.getRange("Start").insertParagraph("NOTE", "After"); // Add head
    // Format head
    head.alignment = "Centered";
    head.font.bold = true;
    const desc = head.getRange("After").insertText("Type here.", "Start"); // Add description after head
    selection.insertBreak("Line", "After"); // Line Break
    await context.sync();
  });
  event.completed();
}
Office.actions.associate("note", note);

// Inserts a Caution template into the document
async function caution(event) {
  await Word.run(async (context) => {
    let selection = context.document.getSelection(); // Get cursor location or highlighted text
    context.load(selection)
    await context.sync(); // Wait for Word to return the selection
    selection.insertBreak("Line", "After"); // Line Break
    const box = selection.insertTable(1, 1, "After"); // Insert 1x1 table for border
    box.getBorder(Word.BorderLocation.outside).type = "Double"; // Set border style to single
    box.getBorder(Word.BorderLocation.all).width = 1; // Thicken the border to 1pt
    const head = box.getRange("Start").insertParagraph("CAUTION", "After"); // Add head
    // Format head
    head.alignment = "Centered";
    head.font.bold = true;
    const desc = head.getRange("After").insertText("Type here.", "Start"); // Add description after head
    selection.insertBreak("Line", "After"); // Line Break
    await context.sync();
  });
  event.completed();
}
Office.actions.associate("caution", caution);

// Inserts a Warning template into the document
async function warning(event) {
  await Word.run(async (context) => {
    let selection = context.document.getSelection(); // Get cursor location or highlighted text
    context.load(selection)
    await context.sync(); // Wait for Word to return the selection
    selection.insertBreak("Line", "After"); // Line Break
    const box = selection.insertTable(1, 1, "After"); // Insert 1x1 table for border
    box.getBorder(Word.BorderLocation.outside).type = "Triple"; // Set border style to single
    box.getBorder(Word.BorderLocation.all).width = 1; // Thicken the border to 1pt
    const head = box.getRange("Start").insertParagraph("WARNING", "After"); // Add head
    // Format head
    head.alignment = "Centered";
    head.font.bold = true;
    const desc = head.getRange("After").insertText("Type here.", "Start"); // Add description after head
    selection.insertBreak("Line", "After"); // Line Break
    await context.sync();
  });
  event.completed();
}
Office.actions.associate("warning", warning);

async function beginBullet(event) {
  await Word.run(async (context) => {
    let selection = context.document.getSelection().paragraphs; // Get cursor location or highlighted text
    context.load(selection)
    await context.sync(); // Wait for Word to return the selection

    if(!selection.items[0].isListItem) {
      const list = selection.items[0].startNewList();
      list.load("$none");
      await context.sync();
    }

  });
  event.completed();
}
Office.actions.associate("beginBullet", beginBullet);

async function beginNumber(event) {
  await Word.run(async (context) => {
    let selection = context.document.getSelection().paragraphs; // Get cursor location or highlighted text
    context.load(selection)
    await context.sync(); // Wait for Word to return the selection

    if(!selection.items[0].isListItem) {
      const list = selection.items[0].startNewList();
      list.load("$none");
      await context.sync();

      let level = 0;
      list.setLevelNumbering(0, Word.ListNumbering.arabic, [level, "."]);
      list.setLevelStartingNumber(0, 1);
      
      level += 1;
      list.setLevelNumbering(1, "LowerLetter", [level, "."]);
      list.setLevelStartingNumber(1, 1);

      list.load("levelTypes");
      await context.sync();
    }

  });
  event.completed();
}
Office.actions.associate("beginNumber", beginNumber);

//Inserts preformatted header 
async function createHeader() {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    if (sections.items.length > 0) {
      const firstSection = sections.items[0];
      const header = firstSection.getHeader("primary");

      // Clear existing content in the header
      header.clear();
      await context.sync();

      // Insert a table into the header
      const table = header.insertTable(1, 3, "start", [
        ["Procedure #", "Procedure Title", "Revision #\n"]
      ]);
 
      await context.sync();

      // Set table style if needed
      table.font.bold = true;
      table.getCell(0, 1).horizontalAlignment = "Centered";
      table.getCell(0, 1).getBorder(Word.BorderLocation.right).type = "None";
      table.getCell(0, 1).getBorder(Word.BorderLocation.left).type = "None";
      table.getCell(0, 2).horizontalAlignment = "Right";
      //table.getCell(0, 2).insertParagraph("").inserText("Page ").insertField("PAGE").insertText(" of ").insertField("NUMPAGES");
      await context.sync();
    }
  });
}
// Function to insert a cover page at the beginning of the document
function insertCoverPage() {
  Word.run(function (context) {
      // Cover page content
      var coverPageContent = `
      <div style="text-align: center; font-family: Arial; font-size: 12pt; line-height: 1;">
          <p style="font-weight: bold; margin-bottom: 0.5em;">[Procedure Title]</p>
          <p style="font-weight: bold; margin-bottom: 0.5em;">[Procedure Number]</p>
          <p style="font-weight: bold; margin-bottom: 1.5em;">[Reactivity Statement]</p> <br>
          <p style="margin-bottom: 0.5em;">Revision #</p>
          <p style="margin-bottom: 0.5em;">[Safety or Quality Classification]</p>
          <p style="margin-bottom: 1.5em;">Level of Use: </p> 
          <br><br><br>
          <!-- Additional Information (optional) -->
          <p style="margin-bottom: 0.5em;">Effective Date: </p>
          <p style="margin-bottom: 0;">Responsible Organization: </p>
          <p style="margin-bottom: 0.5em;">Prepared By: </p>
          <p style="margin-bottom: 0.5em;">Approved By: </p>
          <br>
      </div>
      <div style="page-break-after: always;">
          <br>
      </div>
      `;


      // Insert the cover page content at the beginning of the document
      context.document.body.insertHtml(coverPageContent, "start");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error.message);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}
/** Default helper for invoking an action and handling errors. */
/*async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

async function addStyle() {
  // Adds a new style.
  await Word.run(async (context) => {
    const newStyleName = "NewStyle";
    const newStyleType = Word.StyleType.list;
    context.document.addStyle(newStyleName, newStyleType);
    await context.sync();

    console.log(newStyleName + " has been added to the style list.");

    //edit style properties
    const style = context.document.getStyles().getByNameOrNullObject(newStyleName);
    style.load();
    await context.sync();

    //edit font properties
    const font = style.font;
    //font.color = "#99FF66";
    font.size = 20;
    await context.sync();
    console.log(`Successfully updated font properties of the '${newStyleName}' style.`);

    //edit paragraph format
    //style.paragraphFormat.leftIndent = 30;
    //style.paragraphFormat.alignment = Word.Alignment.centered;
    //await context.sync();
    //console.log(`Successfully the paragraph format of the '${newStyleName}' style.`);
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    //Get selected range and expand it to include the whole first and last paragraphs
    var selection = context.document.getSelection().getRange();
    var firstParagraph = selection.paragraphs.getFirstOrNullObject();
    var lastParagraph = selection.paragraphs.getLastOrNullObject();
    var updatedSelection = selection.expandTo(firstParagraph.getRange()).expandTo(lastParagraph.getRange());
    //load the paragraphs and await sync
    updatedSelection.paragraphs.load();
    await context.sync();

    //console.log(updatedSelection.text)
    updatedSelection.style = "test2";

    await context.sync();

    //Move the cursor to the end of the selection
    updatedSelection.paragraphs.getLast().getNextOrNullObject().select("Start");
  });
}

async function getCount() {
  // Gets the number of styles.
  await Word.run(async (context) => {
    const styles = context.document.getStyles();
    const count = styles.getCount();
    await context.sync();
    document.getElementById("count-output").innerHTML = count.value;
    //console.log(`Number of styles: ${count.value}`);
  });
}

// Import styles to add
import { stylesToAdd } from "./styleList.js";

function addStyleList() {
  //define styles to be added

  //call the forEach function to add styles
  stylesToAdd.forEach(addCustomStyles);

  //define the function
  async function addCustomStyles(value) {
    //create variables for each iteration
    const newStyleType = value.type;
    const newStyleName = value.name;
    //paragraph characteristics
    const leftIndent = value.leftIndent;

    //font characteristics
    const fontColor = value.fontColor;
    const fontSize = value.fontSize;
    //document.getElementById("style-buttons").innerHTML = newStyleName;

    //add style
    await Word.run(async (context) => {
      context.document.addStyle(newStyleName, newStyleType);
      await context.sync();
      console.log(newStyleName + " has been added to the style list.");
    });

    //define style
    await Word.run(async (context) => {
      //load style
      const style = context.document.getStyles().getByNameOrNullObject(newStyleName);
      style.load("$all");
      await context.sync();
      console.log("Style loaded.");

      //define font style (color working, size not)
      const font = style.font;
      font.color = fontColor;
      font.size = fontSize;
      await context.sync();
      console.log("Font has been formatted.");

      //define paragraph style
      style.paragraphFormat.leftIndent = leftIndent;
      style.paragraphFormat.alignment = Word.Alignment.centered;
      console.log(`Paragraph has been formatted`);
      await context.sync();
    });
  }
}

async function insertControl() {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const wordContentControl = range.insertContentControl();

    wordContentControl.tag = "OT-DocTitle";
    wordContentControl.title = "Document Title";
    wordContentControl.insertText("Procedure #1", "Replace");
    wordContentControl.cannotEdit = true;
    //wordContentControl.cannotDelete = true;
    await context.sync();
  });
}

async function findContentControls() {
  await Word.run(async (context) => {
    var nameControls = context.document.contentControls;
    nameControls.load();
    await context.sync();
    //console.log(nameControls);
    //for testing
    document.getElementById("controls-number").innerHTML = nameControls.items.length;
  });
}

async function updateDocumentDetails() {
  // Adds title and colors to odd and even content controls and changes their appearance.
  await Word.run(async (context) => {
    // Get the complete sentence (as range) associated with the insertion point.
    let evenContentControls = context.document.contentControls.getByTag("OT-DocTitle");
    evenContentControls.load("length");

    await context.sync();

    for (let i = 0; i < evenContentControls.items.length; i++) {
      // Change a few properties and append a paragraph
      evenContentControls.items[i].set({
        //insert properties to set here
        cannotEdit: false,
      });
      var newName = document.getElementById("content-control-input").value;
      //change the text
      evenContentControls.items[i].insertText(newName, "Replace");
      evenContentControls.items[i].set({
        //insert properties to set here
        cannotEdit: true,
      });
    }
    await context.sync();
  });
}

async function createHeader() {
  await Word.run(async (context) => {
    var headerTables = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).tables;
    headerTables.load();
    await context.sync();
    //console.log(headerTables.items.length);

    // If no tables in header, create table
    if (headerTables.items.length == 0) {
      context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).clear();
      context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).insertTable(3, 3, "Start");
      await context.sync();

      // Load the table
      const headerTable = context.document.sections
        .getFirst()
        .getHeader(Word.HeaderFooterType.primary)
        .tables.getFirst();
      headerTable.load();
      await context.sync();

      //Insert Logo
      const logoCell = headerTable.getCell(0, 0).body;
      logoCell.load();
      await context.sync();
      logoCell.insertText("Insert Logo Here", "Replace");

      //Insert Document Name
      const range = headerTable.getCell(0, 1).body.getRange("Content");
      var cellData = range.insertContentControl();

      cellData.tag = "OT-DocTitle";
      cellData.title = "Document Title";
      cellData.insertText("Procedure #1", "Replace");
      cellData.cannotEdit = true;
      //cellData.cannotDelete = true;
      await context.sync();

      await context.sync();
      return;
    }

    //header.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.", Word.InsertLocation.start);
  });
}

/* import { headerArray } from "./header.js";

async function createHeaderFromJson() {
  await Word.run(async (context) => {
    var headerTables = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).tables;
    headerTables.load();
    await context.sync();
    //console.log(headerTables.items.length);

    // If no tables in header, create table
    if (headerTables.items.length ==0) {
      context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).clear();
      context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).insertTable(3,3,"Start");
      await context.sync();

      // Load the table
      const headerTable = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).tables.getFirst();
      headerTable.load();
      await context.sync();
      console.log("header loaded");


      //Iterate through headerArray
      headerArray.forEach(addHeaderEntry);
      async function addHeaderEntry(headerCell) {

        //Load the cell
        const selectedCell = headerTable.getCell(headerCell.row, headerCell.cell).body;
        selectedCell.load();
        await context.sync();
        console.log("cell loaded");
*/
/*         
        //Insert Document Name
        const range = headerTable.getCell(0,1).body.getRange("Content");
        var wordContentControl = range.insertContentControl(); 
        
        wordContentControl.tag = headerCell.tag;  
        console.log(headerCell.tag);
        wordContentControl.title = headerCell.title;   
        console.log(headerCell.title);
        wordContentControl.insertText("Procedure #1", 'Replace');

 */
/*
        wordContentControl.cannotEdit = headerCell.cannotEdit;
        //wordContentControl.cannotdelete = headerCell.cannotDelete;
        
        await context.sync(); 

      }
    return;
    };

  });
} */
