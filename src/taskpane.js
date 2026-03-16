Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    loadSettings();
    document.getElementById("format-btn").onclick = runFormatter;
  }
});

function showError(message) {
  const errorDiv = document.getElementById("error-msg");
  errorDiv.innerText = message;
  errorDiv.style.display = "block";
}

function clearError() {
  const errorDiv = document.getElementById("error-msg");
  errorDiv.style.display = "none";
  errorDiv.innerText = "";
}

async function runFormatter() {
  clearError();
  saveSettings();

  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();

      // Explicitly load the text AND the type of selection
      range.load("text, type");
      await context.sync();

      // User selected nothing
      const rawCode = range.text;
      if (!rawCode || rawCode.trim() === "") {
        throw new Error("Please highlight the code you want to format.");
      }

      // User selected unsupported frormat
      if (range.type === "Image" || range.type === "Shape") {
        throw new Error("Invalid selection. Please highlight text only.");
      }

      const isInline = document.getElementById("inline-check").checked;
      // If inline is checked, we ensure they didn't highlight multiple lines
      if (isInline && (rawCode.includes("\n") || rawCode.includes("\r"))) {
        throw new Error(
          "Inline code cannot span multiple lines. Uncheck 'Inline Code' for blocks.",
        );
      }

      const langString = document.getElementById("lang-select").value;
      const selectedEngine = getLanguageEngine(langString);
      const tokens = Prism.tokenize(rawCode, selectedEngine);

      const selectedThemeName = document.getElementById("theme-select").value;
      const theme = getTheme(selectedThemeName);

      if (isInline) {
        const htmlPayload = buildInlineHTML(tokens, theme);
        range.insertHtml(htmlPayload, Word.InsertLocation.replace);
      } else {
        const ooxmlPayload = buildOOXML(tokens, theme);
        range.insertOoxml(ooxmlPayload, Word.InsertLocation.replace);
      }

      await context.sync();
    });
  } catch (error) {
    console.error("LazyCode Error:", error);

    if (error instanceof Error) {
      showError(error.message);
    } else {
      showError("An unexpected error occurred while communicating with Word.");
    }
  }
}
