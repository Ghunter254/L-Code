/**
 * Escapes XML special characters to prevent Word from crashing on
 * C++ operators like < and >
 */
function escapeXml(unsafeStr) {
  return unsafeStr
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Prism sometimes returns nested arrays for complex tokens (like directives).
 * This recursively flattens them into a single raw string so we can escape it safely.
 */
function extractText(content) {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) return content.map(extractText).join("");
  if (content && content.content) return extractText(content.content);
  return "";
}

/**
 * Translates a Prism.js token array into a Word-compatible OOXML string.
 */
function buildOOXML(tokens, theme) {
  let runXml = "";

  for (const token of tokens) {
    let type = "text";
    let content = "";

    if (typeof token === "string") {
      content = token;
    } else {
      type = token.type;
      content = extractText(token.content);
    }

    const color = theme.tokens[type] || theme.defaultText;
    content = escapeXml(content);

    const lines = content.split(/\r\n|\r|\n/);

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].replace(/ /g, "&#xA0;");

      if (line.length > 0) {
        runXml += `
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>
                        <w:color w:val="${color}"/>
                        <w:sz w:val="20"/> <w:noProof/> 
                    </w:rPr>
                    <w:t xml:space="preserve">${line}</w:t>
                </w:r>`;
      }

      if (i < lines.length - 1) {
        runXml += `<w:r><w:br/></w:r>`;
      }
    }
  }

  return `<?xml version="1.0" standalone="yes"?>
    <?mso-application progid="Word.Document"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:tbl>
                <w:tblPr>
                  <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="${theme.border}"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="${theme.border}"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${theme.border}"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="${theme.border}"/>
                  </w:tblBorders>
                </w:tblPr>
                <w:tr>
                  <w:trPr><w:cantSplit/></w:trPr>
                  <w:tc>
                    <w:tcPr>
                      <w:tcMar>
                        <w:top w:w="113" w:type="dxa"/> 
                        <w:left w:w="113" w:type="dxa"/>
                        <w:bottom w:w="113" w:type="dxa"/>
                        <w:right w:w="113" w:type="dxa"/>
                      </w:tcMar>
                      <w:shd w:val="clear" w:color="auto" w:fill="${theme.background}"/>
                    </w:tcPr>
                    <w:p>
                      <w:pPr>
                        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                      </w:pPr>
                      ${runXml}
                    </w:p>
                  </w:tc>
                </w:tr>
              </w:tbl>
            </w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>`;
}

/**
 * Translates Prism tokens into an HTML string specifically for inline text.
 * Uses Microsoft-specific CSS to disable the spellchecker natively.
 */
function buildInlineHTML(tokens, theme) {
  let html = `<span lang="zxx" style="mso-no-proof: yes; font-family: Consolas, 'Courier New', monospace;">`;

  for (const token of tokens) {
    let type = "text";
    let content = "";

    if (typeof token === "string") {
      content = token;
    } else {
      type = token.type;
      content = extractText(token.content);
    }

    const color = theme.tokens[type] || theme.defaultText;
    content = escapeXml(content);
    content = content.replace(/ /g, "&nbsp;");

    // Reinforce the spellcheck bypass on every single token
    html += `<span lang="zxx" style="mso-no-proof: yes; color: #${color};">${content}</span>`;
  }

  html += `</span>`;
  return html;
}
