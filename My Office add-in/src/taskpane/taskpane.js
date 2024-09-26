/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { base64Image } from "../../base64Image";


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph)
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle)
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle)
    document.getElementById("change-font").onclick = () => tryCatch(changeFont)
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange)
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextOutsideRange)
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText)
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

const TEXT_TO_INSERT_INTO_PARAGRAPH = "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.";

async function insertParagraph() {
  await Word.run(async (context) => {

    const docBody = context.document.body;
    docBody.insertParagraph(TEXT_TO_INSERT_INTO_PARAGRAPH,
      Word.InsertLocation.start);// "Start" | "End" will also work

    await context.sync();
  });
}

const applyStyle = async () => {
  await Word.run(async (context) => {
    const getFirstParagraph = context.document.body.paragraphs.getFirst();
    getFirstParagraph.styleBuiltIn = Word.Style.intenseReference;

    await context.sync();
  })
}

const applyCustomStyle = async () => {
  await Word.run(async (context) => {
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";

    await context.sync();
  });
}

const changeFont = async () => {
  await Word.run(async (context) => {
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
      color: "Red",
    })
    await context.sync();
  })
};

const insertTextIntoRange = async () => {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalrange = doc.getSelection();
    originalrange.insertText(" (M365)", Word.InsertLocation.end);

    originalrange.load("text");
    await context.sync();

    doc.body.insertParagraph("Original range: " + originalrange.text, Word.InsertLocation.end);

    await context.sync();
  })
}

const insertTextOutsideRange = async () => {
  await Word.run(async (context) => {
    const doc = context.document;
    const selectedRange = doc.getSelection();
    selectedRange.insertText("---Inserted Before Text---", Word.InsertLocation.before);

    selectedRange.load("text");
    await context.sync();

    doc.body.insertParagraph("Current text of original range: " + selectedRange.text, Word.InsertLocation.end);

    await context.sync();
  });
};

const replaceText = async () => {
  await Word.run(async (context) => {
    const doc = context.document;
    const selectedRange = doc.getSelection();
    selectedRange.insertText("many", Word.InsertLocation.replace);

    selectedRange.load("text");
    await context.sync();

    doc.body.insertParagraph("Replaced word is: " + selectedRange.text, Word.InsertLocation.end);

    await context.sync();
  });
}

const insertImage = async () => {
  await Word.run(async (context) => {
    context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);

    await context.sync();
  })
}

const insertHTML = async () => {
  await Word.run(async (context) => {
    const blanckParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
    blanckParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);
    await context.sync();
  })
}

const insertTable = async () => {
  await Word.run(async (context) => {
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext()
    const tableData = [
      ["Name", "ID", "Birth City"],
      ["Pedro", "1", "Pravets"],
      ["Ivan", "2", "Botevgrad"]
    ];

    secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);

    await context.sync();
  })
}

const tryCatch = async (callback) => {
  try {
    await callback();
  } catch (error) {
    console.error(`Error in the tryCatch block: ${error}`)
  }
}