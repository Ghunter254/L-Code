function getLanguageEngine(langString) {
  const prismLangMap = {
    cpp: Prism.languages.cpp,
    go: Prism.languages.go,
    javascript: Prism.languages.javascript,
    python: Prism.languages.python,
  };

  return (selectedEngine = prismLangMap[langString]);
}
