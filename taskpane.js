Office.onReady(() => {
  document.getElementById("insertText").onclick = () => {
    Word.run(async (context) => {
      context.document.body.insertParagraph("Texte inséré par le complément", Word.InsertLocation.end);
      await context.sync();
    });
  };

  document.getElementById("boldText").onclick = () => {
    Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.bold = true;
      await context.sync();
    });
  };

  document.getElementById("italicText").onclick = () => {
    Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.italic = true;
      await context.sync();
    });
  };
});
