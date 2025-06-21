Office.onReady(() => {
  document.getElementById("define-field-a").onclick = () => defineField("fieldA");
  document.getElementById("define-field-b").onclick = () => defineField("fieldB");
  document.getElementById("define-result-field").onclick = () => defineField("resultField");
  document.getElementById("calculate-sum").onclick = calculateSum;
});

async function defineField(tagName) {
  await Word.run(async (context) => {
    const existing = context.document.contentControls.getByTag(tagName);
    context.load(existing);
    await context.sync();

    // מחק שדה קודם אם קיים
    existing.items.forEach(item => item.delete(true));

    // צור שדה חדש
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = tagName;
    cc.title = getFieldTitle(tagName);
    cc.appearance = "BoundingBox";
    cc.color = getFieldColor(tagName);

    await context.sync();
  });
}

async function calculateSum() {
  showLoader(true);

  setTimeout(async () => {
    await Word.run(async (context) => {
      const fieldA = context.document.contentControls.getByTag("fieldA").getFirst();
      const fieldB = context.document.contentControls.getByTag("fieldB").getFirst();
      const resultField = context.document.contentControls.getByTag("resultField").getFirst();

      fieldA.load("text");
      fieldB.load("text");

      await context.sync();

      const a = parseFloat(fieldA.text);
      const b = parseFloat(fieldB.text);
      const sum = (!isNaN(a) ? a : 0) + (!isNaN(b) ? b : 0);

      resultField.insertText(sum.toString(), Word.InsertLocation.replace);

      await context.sync();
      showResult("הסכום הוא: " + sum);
    }).catch(err => {
      console.error(err);
      showResult("שגיאה בחישוב הסכום");
    });

    showLoader(false);
  }, 3000);
}

function showLoader(show) {
  document.getElementById("loader").style.display = show ? "block" : "none";
  document.getElementById("result").innerText = "";
}

function showResult(text) {
  document.getElementById("loader").style.display = "none";
  document.getElementById("result").innerText = text;
}

function getFieldTitle(tag) {
  switch (tag) {
    case "fieldA": return "שדה A";
    case "fieldB": return "שדה B";
    case "resultField": return "סכום";
    default: return "";
  }
}

function getFieldColor(tag) {
  switch (tag) {
    case "fieldA": return "blue";
    case "fieldB": return "green";
    case "resultField": return "orange";
    default: return "black";
  }
}
