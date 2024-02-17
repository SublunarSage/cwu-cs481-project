/*
Office.onReady((info) => {});

async function setStyleSectionHeader(event) {
  await Word.run(async (context) => {
    //Get selected range and expand it to include the whole first and last paragraphs
    var selection = context.document.getSelection().getRange();
    var firstParagraph = selection.paragraphs.getFirstOrNullObject();
    var lastParagraph = selection.paragraphs.getLastOrNullObject();
    var updatedSelection = selection.
      expandTo(firstParagraph.getRange()).
      expandTo(lastParagraph.getRange());
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
  
  Office.actions.associate("setStyleSectionHeader", setStyleSectionHeader);
  */