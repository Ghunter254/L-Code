function saveSettings() {
  const settings = {
    lang: document.getElementById("lang-select").value,
    theme: document.getElementById("theme-select").value,
    inline: document.getElementById("inline-check").checked,
  };
  localStorage.setItem("lcode_settings", JSON.stringify(settings));
}

function loadSettings() {
  const saved = localStorage.getItem("lcode_settings");
  if (!saved) return;

  const settings = JSON.parse(saved);
  document.getElementById("lang-select").value = settings.lang;
  document.getElementById("theme-select").value = settings.theme;
  document.getElementById("inline-check").checked = settings.inline;
}
