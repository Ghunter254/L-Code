````markdown
# L-Code 🚀

**Professional syntax highlighting for Microsoft Word.**

L-Code is a lightweight developer tool that bridges the gap between IDEs and formal documentation. It hijacks the Prism.js lexing engine to translate source code into native **Office Open XML (OOXML)**, providing IDE-grade syntax highlighting that Word treats as native content—no red squiggly lines, no layout breakage.

![L-Code Demo](assets/demo.gif)

## 🎯 The Problem

Word's default behavior treats code as standard text, leading to:

- **Red Spellcheck Squiggles:** Word tries to "fix" your variable names.
- **Collapsing Whitespace:** Indentation is often lost or inconsistent.
- **Visual Noise:** Copy-pasting from IDEs often brings hidden, clunky HTML that breaks document flow.

## ✨ Features

- **Native OOXML Injection:** Generates raw Microsoft Open XML to ensure code blocks stay anchored and styled.
- **Multi-Engine Support:** Dedicated lexers for **C++, Go, Python, and JavaScript**.
- **Thematic Control:** - `VS Code Dark+`: For digital-first documentation.
  - `Academic Light`: Optimized for printed lab reports (ink-friendly).
- **Smart Inline Formatting:** A custom "Ghost" mode for variables inside sentences that won't break line-height.
- **Persistent Preferences:** Remembers your language and theme selection across sessions.

## 🛠️ Tech Stack

- **API:** Office.js (Word API)
- **Engine:** Headless Prism.js (Custom Tokenizer)
- **Markup:** Microsoft Flat OPC (OOXML) & HTML5
- **Persistence:** Browser LocalStorage

## 🚀 Installation & Local Development

1. **Clone & Setup:**
   ```bash
   git clone [https://github.com/Ghunter254/L-Code.git](https://github.com/Ghunter254/L-Code.git)
   cd L-Code
   npm init -y
   ```
````

2. **Install Dev Certs:**

```powershell
npx office-addin-dev-certs install

```

3. **Start Secure Server:**

```powershell
npx http-server -p 3000 --ssl --cert "C:\Users\ypaul\.office-addin-dev-certs\localhost.crt" --key "C:\Users\ypaul\.office-addin-dev-certs\localhost.key" --cors

```

4. **Sideload to Word:**

```powershell
npx office-addin-debugging start manifest.xml desktop

```

---

**Developed with ❤️ for the Engineering Community.**
