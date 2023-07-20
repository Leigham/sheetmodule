import SheetManager from "./SheetManager";
import fs from "node:fs";
async function main() {
  try {
    const sheetManager = await SheetManager.getInstance(
      JSON.parse(fs.readFileSync("./credentials.json").toString())
    );
    const doc = await sheetManager.createNewDocument("test", []);
    if (!doc.id) throw new Error("No document ID");
    console.log(
      `Created document with ID ${
        doc.id
      }\n URL: ${await sheetManager.getSheetURL(doc.id)}`
    );
    const insert = await sheetManager.addSheetValues(doc.id, [
      {
        sheetName: "New Sheet",
        headers: ["test", "test", "test"],
        rows: [
          ["test1", 1, true],
          ["test2", 1, true],
          ["test3", 1, true],
        ],
      },
    ]);
    const sheetInfo = await sheetManager.getSheetInfo(doc.id);
    const values = await sheetManager.getSheetValues(doc.id, 1);

    const headers = await sheetManager.getSheetHeaders(doc.id, 1);
    const filtered = await sheetManager.getSheetValuesByFilter(
      doc.id,
      1,
      "A",
      "dropdown_1"
    );
  } catch (error) {
    console.log(error);
  }
}
main();
