// Main popup logic for JP-VI Translator Browser Extension

// UI Element References
const tabBtns = document.querySelectorAll(".tab-btn");
const tabPanes = document.querySelectorAll(".tab-pane");
const sourceTextEl = document.getElementById("source-text");
const ttsBtnEl = document.getElementById("tts-btn");
const autoTranslateSwitch = document.getElementById("auto-translate-switch");
const translateBtn = document.getElementById("translate-btn");
const resultCard = document.getElementById("result-card");
const translationOutput = document.getElementById("translation-output");
const copyBtn = document.getElementById("copy-btn");
const clearHistoryBtn = document.getElementById("clear-history-btn");
const historyList = document.getElementById("history-list");
const toastEl = document.getElementById("toast");
const toastMsgEl = document.getElementById("toast-message");

// Language Swap Element References
const swapLangBtn = document.getElementById("swap-lang-btn");
const sourceLangTitle = document.getElementById("source-lang-title");
const targetLangTitle = document.getElementById("target-lang-title");
const inputLangLabel = document.getElementById("input-lang-label");
const outputLangLabel = document.getElementById("output-lang-label");
const ttsOutputBtn = document.getElementById("tts-output-btn");

// Separate Hiragana Element References
const inputHiraganaContainer = document.getElementById("input-hiragana-container");
const inputHiraganaOutput = document.getElementById("input-hiragana-output");
const outputHiraganaContainer = document.getElementById("output-hiragana-container");
const outputHiraganaOutput = document.getElementById("output-hiragana-output");

// Settings UI Element References
const phoneticSelect = document.getElementById("phonetic-select");

// State Variables
let isAutoTranslate = localStorage.getItem("jp_vi_ext_auto_translate") !== "false"; // Default to true
let lastTranslatedText = "";
let isJaToVi = true; // true: Japanese -> Vietnamese, false: Vietnamese -> Japanese
let phoneticMode = localStorage.getItem("jp_vi_ext_phonetic_mode") || "hiragana";
let lastHiragana = "";
let lastRomaji = "";

// Theme Toggle Reference & State
const themeToggleBtn = document.getElementById("theme-toggle-btn");
let currentTheme = localStorage.getItem("jp_vi_ext_theme") || "dark";

// Apply theme immediately
applyTheme(currentTheme);

function applyTheme(theme) {
  if (theme === "light") {
    document.body.classList.add("light-theme");
    if (themeToggleBtn) {
      themeToggleBtn.innerHTML = '<i class="fa-solid fa-moon"></i>';
      themeToggleBtn.title = "Chuyển sang chế độ tối";
    }
  } else {
    document.body.classList.remove("light-theme");
    if (themeToggleBtn) {
      themeToggleBtn.innerHTML = '<i class="fa-solid fa-sun"></i>';
      themeToggleBtn.title = "Chuyển sang chế độ sáng";
    }
  }
}

// Initialize Extension
document.addEventListener("DOMContentLoaded", () => {
  initApp();
  // Get text selection from active browser tab
  tryGetActiveTabSelection();
});

// Try to grab selected text from the browser tab
function tryGetActiveTabSelection() {
  if (typeof chrome === "undefined" || !chrome.tabs) return;

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const activeTab = tabs[0];
    if (!activeTab || !activeTab.id) return;

    // Avoid scripting on chrome:// or edge:// system pages
    if (activeTab.url && (activeTab.url.startsWith("chrome:") || activeTab.url.startsWith("edge:") || activeTab.url.startsWith("about:") || activeTab.url.startsWith("https://chrome.google.com"))) {
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId: activeTab.id },
        func: () => window.getSelection().toString()
      },
      (results) => {
        if (chrome.runtime.lastError) {
          console.warn("Script execution failed:", chrome.runtime.lastError.message);
          return;
        }

        if (results && results[0] && results[0].result) {
          const selectedText = results[0].result.trim();
          if (selectedText) {
            sourceTextEl.value = selectedText;
            ttsBtnEl.disabled = false;

            if (isAutoTranslate) {
              translateText(selectedText);
            }
          }
        }
      }
    );
  });
}

// Main initialization function
function initApp() {
  // Set UI state based on saved preferences
  if (autoTranslateSwitch) {
    autoTranslateSwitch.checked = isAutoTranslate;
  }
  if (phoneticSelect) {
    phoneticSelect.value = phoneticMode;
  }

  // Theme Toggle Button Click Event
  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", () => {
      currentTheme = currentTheme === "light" ? "dark" : "light";
      localStorage.setItem("jp_vi_ext_theme", currentTheme);
      applyTheme(currentTheme);
    });
  }

  // Tab Navigation Click Event
  tabBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetTab = btn.getAttribute("data-tab");
      
      tabBtns.forEach((b) => b.classList.remove("active"));
      tabPanes.forEach((p) => p.classList.remove("active"));
      
      btn.classList.add("active");
      document.getElementById(targetTab).classList.add("active");

      if (targetTab === "tab-history") {
        renderHistory();
      }
    });
  });

  // Phonetic Selection Dropdown Event
  if (phoneticSelect) {
    phoneticSelect.addEventListener("change", (e) => {
      phoneticMode = e.target.value;
      localStorage.setItem("jp_vi_ext_phonetic_mode", phoneticMode);
      if (lastHiragana || lastRomaji) {
        displayPhonetics(lastHiragana, lastRomaji);
      }
    });
  }

  // Language Swap Button Click Event
  swapLangBtn.addEventListener("click", () => {
    isJaToVi = !isJaToVi;
    updateLanguageUI();
    
    // Clear current inputs and outputs on swap to avoid confusion
    sourceTextEl.value = "";
    translationOutput.textContent = "";
    inputHiraganaOutput.textContent = "";
    outputHiraganaOutput.textContent = "";
    inputHiraganaContainer.classList.add("hidden");
    outputHiraganaContainer.classList.add("hidden");
    resultCard.classList.add("hidden");
    ttsBtnEl.disabled = true;
    lastTranslatedText = "";
    lastHiragana = "";
    lastRomaji = "";
  });

  // Translate Button Click (Manual translation)
  translateBtn.addEventListener("click", () => {
    const text = sourceTextEl.value.trim();
    if (text) {
      translateText(text);
    } else {
      showToast("Vui lòng nhập hoặc bôi đen văn bản cần dịch.");
    }
  });

  // Source text area input event to enable speech button
  sourceTextEl.addEventListener("input", (e) => {
    const text = e.target.value.trim();
    ttsBtnEl.disabled = text.length === 0;
  });

  // Auto-translate Toggle Switch Event
  autoTranslateSwitch.addEventListener("change", (e) => {
    isAutoTranslate = e.target.checked;
    localStorage.setItem("jp_vi_ext_auto_translate", isAutoTranslate);
  });

  // TTS Speaker Button (Input Card - Japanese Text in JA->VI)
  ttsBtnEl.addEventListener("click", () => {
    const text = sourceTextEl.value.trim();
    if (text) {
      speakJapanese(text);
    }
  });

  // TTS Speaker Button (Output Card - Japanese Text in VI->JA)
  ttsOutputBtn.addEventListener("click", () => {
    const text = translationOutput.textContent.trim();
    if (text) {
      speakJapanese(text);
    }
  });

  // Copy Translation Button Click Event
  copyBtn.addEventListener("click", () => {
    const text = translationOutput.textContent;
    if (text) {
      copyToClipboard(text);
    }
  });

  // Copy Input Text Button Click Event
  const copyInputBtn = document.getElementById("copy-input-btn");
  if (copyInputBtn) {
    copyInputBtn.addEventListener("click", () => {
      const text = sourceTextEl.value.trim();
      if (text) {
        copyToClipboard(text, "Đã sao chép văn bản gốc!");
      }
    });
  }

  // Clear All History Button Click Event
  clearHistoryBtn.addEventListener("click", () => {
    localStorage.removeItem("jp_vi_ext_history");
    renderHistory();
    showToast("Đã xóa toàn bộ lịch sử.");
  });

  // Initial UI Setup
  updateLanguageUI();
  renderHistory();
}

// Update UI text based on current translation direction
function updateLanguageUI() {
  if (isJaToVi) {
    sourceLangTitle.textContent = "JP";
    targetLangTitle.textContent = "VI";
    inputLangLabel.innerHTML = '<i class="fa-solid fa-location-dot"></i> JP';
    outputLangLabel.innerHTML = '<i class="fa-solid fa-circle-check"></i> VI';
    sourceTextEl.placeholder = "Quét chọn văn bản tiếng Nhật trên trang web hoặc nhập tại đây...";
    
    ttsBtnEl.classList.remove("hidden");
    ttsOutputBtn.classList.add("hidden");
  } else {
    sourceLangTitle.textContent = "VI";
    targetLangTitle.textContent = "JP";
    inputLangLabel.innerHTML = '<i class="fa-solid fa-location-dot"></i> VI';
    outputLangLabel.innerHTML = '<i class="fa-solid fa-circle-check"></i> JP';
    sourceTextEl.placeholder = "Quét chọn văn bản tiếng Việt trên trang web hoặc nhập tại đây...";
    
    ttsBtnEl.classList.add("hidden");
    ttsOutputBtn.classList.remove("hidden");
  }
}

// Display phonetics based on mode (Hiragana, Romaji, or both)
function displayPhonetics(hiraganaText, romajiText) {
  lastHiragana = hiraganaText || "";
  lastRomaji = romajiText || "";

  let hiragana = lastHiragana;
  let romaji = lastRomaji;

  if (typeof wanakana !== "undefined") {
    if (!hiragana && romaji) {
      hiragana = wanakana.toHiragana(romaji);
    } else if (hiragana && !romaji) {
      romaji = wanakana.toRomaji(hiragana);
    }
  }

  if (hiragana && !lastHiragana) lastHiragana = hiragana;
  if (romaji && !lastRomaji) lastRomaji = romaji;

  let displayText = "";

  if (phoneticMode === "romaji") {
    displayText = romaji;
  } else if (phoneticMode === "both") {
    displayText = `${hiragana} / ${romaji}`;
  } else {
    displayText = hiragana;
  }

  const inputOutput = document.getElementById("input-hiragana-output");
  const outputOutput = document.getElementById("output-hiragana-output");

  if (inputOutput) inputOutput.textContent = displayText;
  if (outputOutput) outputOutput.textContent = displayText;

  if (displayText.trim()) {
    if (isJaToVi) {
      inputHiraganaContainer.classList.remove("hidden");
      outputHiraganaContainer.classList.add("hidden");
    } else {
      outputHiraganaContainer.classList.remove("hidden");
      inputHiraganaContainer.classList.add("hidden");
    }
  } else {
    inputHiraganaContainer.classList.add("hidden");
    outputHiraganaContainer.classList.add("hidden");
  }
}

// Translate text between languages using Google Translate API
async function translateText(text) {
  if (!text) return;
  
  lastTranslatedText = text;
  lastHiragana = "";
  lastRomaji = "";
  
  translationOutput.innerHTML = '<span class="loading"><i class="fa-solid fa-circle-notch fa-spin"></i> Đang dịch...</span>';
  resultCard.classList.remove("hidden");
  inputHiraganaContainer.classList.add("hidden");
  outputHiraganaContainer.classList.add("hidden");

  try {
    const sl = isJaToVi ? "ja" : "vi";
    const tl = isJaToVi ? "vi" : "ja";
    
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sl}&tl=${tl}&dt=t&dt=rm&q=${encodeURIComponent(text)}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error("Mất kết nối mạng hoặc lỗi dịch vụ.");
    
    const data = await response.json();
    
    let translation = "";
    let romaji = "";
    
    if (data && data[0]) {
      data[0].forEach((item) => {
        if (item[0]) {
          translation += item[0];
        }
        
        // Parse Romaji: index 3 is source transliteration, index 2 is target transliteration
        if (item[3]) {
          romaji = item[3];
        } else if (item[2]) {
          romaji = item[2];
        }
      });
    }

    if (translation) {
      translationOutput.textContent = translation;
      displayPhonetics(null, romaji);

      // Add to History
      const jaText = isJaToVi ? text : translation;
      const viText = isJaToVi ? translation : text;
      saveToHistory(jaText, viText, lastHiragana);
    } else {
      translationOutput.textContent = "Không tìm thấy bản dịch phù hợp.";
    }
  } catch (error) {
    console.error("Translation error:", error);
    translationOutput.textContent = "Lỗi khi kết nối dịch thuật: " + error.message;
  }
}

// Speak Japanese aloud (Text to Speech)
function speakJapanese(text) {
  if (!text) return;
  
  window.speechSynthesis.cancel();

  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = "ja-JP";

  const voices = window.speechSynthesis.getVoices();
  const jaVoice = voices.find(v => v.lang.startsWith("ja") || v.lang === "ja-JP");
  if (jaVoice) {
    utterance.voice = jaVoice;
  }
  
  window.speechSynthesis.speak(utterance);
}

// Copy Text to Clipboard
async function copyToClipboard(text, successMsg = "Đã sao chép bản dịch!") {
  try {
    await navigator.clipboard.writeText(text);
    showToast(successMsg);
  } catch (err) {
    const tempTextarea = document.createElement("textarea");
    tempTextarea.value = text;
    tempTextarea.style.position = "fixed";
    document.body.appendChild(tempTextarea);
    tempTextarea.select();
    try {
      document.execCommand("copy");
      showToast(successMsg);
    } catch (e) {
      showToast("Không thể tự động sao chép.");
    }
    document.body.removeChild(tempTextarea);
  }
}

// Display dynamic toast message
let toastTimeout;
function showToast(message) {
  clearTimeout(toastTimeout);
  toastMsgEl.textContent = message;
  toastEl.classList.remove("hidden");
  
  toastTimeout = setTimeout(() => {
    toastEl.classList.add("hidden");
  }, 2000);
}

// Save translation search to LocalStorage
function saveToHistory(ja, vi, hiragana) {
  let history = JSON.parse(localStorage.getItem("jp_vi_ext_history") || "[]");
  
  if (history.length > 0 && history[0].ja === ja && history[0].vi === vi) return;

  history = history.filter(item => !(item.ja === ja && item.vi === vi));

  history.unshift({
    id: Date.now(),
    ja,
    vi,
    hiragana
  });

  if (history.length > 20) {
    history.pop();
  }

  localStorage.setItem("jp_vi_ext_history", JSON.stringify(history));
}

// Render history list inside UI
function renderHistory() {
  const history = JSON.parse(localStorage.getItem("jp_vi_ext_history") || "[]");
  
  if (history.length === 0) {
    historyList.innerHTML = `
      <div class="empty-history">
        <i class="fa-regular fa-folder-open"></i>
        <p>Chưa có lịch sử dịch thuật</p>
      </div>
    `;
    clearHistoryBtn.classList.add("hidden");
    return;
  }

  clearHistoryBtn.classList.remove("hidden");
  historyList.innerHTML = "";

  history.forEach((item) => {
    const itemEl = document.createElement("div");
    itemEl.className = "history-item";
    
    itemEl.innerHTML = `
      <div class="history-item-header">
        <span class="jp-text" title="${item.ja.replace(/"/g, '&quot;')}">${item.ja}</span>
        <div class="actions">
          <button class="icon-btn play-hist-btn" title="Nghe tiếng Nhật"><i class="fa-solid fa-volume-high"></i></button>
          <button class="icon-btn copy-hist-btn" title="Sao chép bản dịch"><i class="fa-solid fa-copy"></i></button>
          <button class="icon-btn delete-hist-btn" title="Xóa"><i class="fa-solid fa-xmark"></i></button>
        </div>
      </div>
      <div class="history-vi-text">${item.vi}</div>
    `;

    itemEl.querySelector(".play-hist-btn").addEventListener("click", (e) => {
      e.stopPropagation();
      speakJapanese(item.ja);
    });

    itemEl.querySelector(".copy-hist-btn").addEventListener("click", (e) => {
      e.stopPropagation();
      copyToClipboard(item.vi);
    });

    itemEl.querySelector(".delete-hist-btn").addEventListener("click", (e) => {
      e.stopPropagation();
      deleteHistoryItem(item.id);
    });

    historyList.appendChild(itemEl);
  });
}

// Delete single item from history
function deleteHistoryItem(id) {
  let history = JSON.parse(localStorage.getItem("jp_vi_ext_history") || "[]");
  history = history.filter(item => item.id !== id);
  localStorage.setItem("jp_vi_ext_history", JSON.stringify(history));
  renderHistory();
}
