const VSCodeDark = {
  background: "1E1E1E",
  border: "454545",
  defaultText: "D4D4D4",
  tokens: {
    keyword: "569CD6",
    function: "DCDCAA",
    string: "CE9178",
    number: "B5CEA8",
    comment: "6A9955",
    operator: "D4D4D4",
    punctuation: "D4D4D4",
    directive: "C586C0",
  },
};

const AcademicLight = {
  background: "F8F9FA", // Very subtle off-white/gray
  border: "D1D5DB", // Clean gray border
  defaultText: "24292E", // Near black for high contrast readability
  tokens: {
    keyword: "0000FF", // Classic pure blue
    function: "795E26", // Warm brown/yellow
    string: "A31515", // Brick red
    number: "098658", // Deep green
    comment: "008000", // Standard green
    operator: "24292E",
    punctuation: "24292E",
    directive: "AF00DB", // Purple
  },
};

/**
 * Returns the requested theme object.
 */
function getTheme(themeName) {
  return themeName === "light" ? AcademicLight : VSCodeDark;
}
