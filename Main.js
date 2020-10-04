const wishbook_sheet = SpreadsheetApp.getActive().getSheetByName("Wishbook");
const wishes = wishbook_sheet.getRange("A3:I20")

function onEdit(e) {
  var active = e.source.getActiveSheet();
  if (active.getName() == wishbook_sheet.getName()) wishbook_main();
}

function wishbook_main() {
  sort();
  reset_format();  // Colors from https://yagisanatode.com/2019/08/06/google-apps-script-hexadecimal-color-codes-for-google-docs-sheets-and-slides-standart-palette/
  conditional_format();
}

function sort() {
  wishbook_sheet.getRange("A3:I20").sort([
    { column: 5, ascending: true  },  // Completed: asc
    { column: 6, ascending: true  },  // Managed/blocked: asc
    { column: 4, ascending: false }   // Priority: desc
  ]);
}

function reset_format() {
  wishes
    .clearFormat()
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("top");
  wishbook_sheet.getRange("A3:A20")  // Name column
    .setBackground("#d9ead3");
  wishbook_sheet.getRange("B3:I20")  // Body table
    .setBackground("#fff2cc");
  wishbook_sheet.getRange("D3:E20")  // Priority & completed
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
  wishbook_sheet.getRange("G3:I20")  // Completed when and in how many days
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
}

function conditional_format() {
  const wishes_values = wishes.getValues();

  for (i = 0; i < wishes_values.length; i++) {
    const wish = wishes_values[i];
    if (wish[0] != "") {
      const wish_row = i + 3;
      const wish_rng = wishbook_sheet.getRange(wish_row, 1, 1, 9);

      // Color depending on priorities
      const prio = wish[3];
           if (prio == 3) wish_rng.setBackground("#f4cccc");  // Red
      else if (prio == 2) wish_rng.setBackground("#fce5cd");  // Orange
      else if (prio == 1) wish_rng.setBackground("#cfe2f3");  // Blue

      // If blocked/managed, color grey
      const managed = wish[5];
      if (managed != "") wish_rng.setBackground("#d9d9d9");  // Grey

      // If completed, color grey and strikethrough
      const completed = wish[4];
      if (completed) wish_rng
        .setBackground("#d9d9d9")  // Grey
        .setFontLine("line-through");  // Strikethrough, seriously 🤷‍♂️?! https://stackoverflow.com/a/37931177
    }
  }
}
