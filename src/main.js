// Main Application Logic for JP-VI Translator Add-in
import './style.css';
import * as wanakana from 'wanakana';

// UI Element References
const hostBadgeEl = document.getElementById("host-badge");
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
let isAutoTranslate = true;
let isSelectionHandlerRegistered = false;
let debounceTimer = null;
let lastTranslatedText = "";
let isJaToVi = true; // true: Japanese -> Vietnamese, false: Vietnamese -> Japanese
let phoneticMode = localStorage.getItem("jp_vi_phonetic_mode") || "hiragana";
let lastHiragana = "";
let lastRomaji = "";

// Theme Toggle Reference & State
const themeToggleBtn = document.getElementById("theme-toggle-btn");
let currentTheme = localStorage.getItem("jp_vi_theme") || "dark";

// Apply theme immediately on script load
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

// Initialize Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    hostBadgeEl.textContent = "Excel Connection";
    hostBadgeEl.className = "badge connected-excel";
    setupOfficeHandlers();
  } else if (info.host === Office.HostType.PowerPoint) {
    hostBadgeEl.textContent = "PowerPoint Connection";
    hostBadgeEl.className = "badge connected-ppt";
    setupOfficeHandlers();
  } else if (info.host === Office.HostType.Word) {
    hostBadgeEl.textContent = "Word Connection";
    hostBadgeEl.className = "badge connected-word";
    setupOfficeHandlers();
  } else {
    // Fallback for browser testing
    hostBadgeEl.textContent = "Browser Mode";
    hostBadgeEl.className = "badge";
    // Enable manual testing input in browser
    sourceTextEl.placeholder = "Nhập văn bản cần dịch...";
  }

  // Common UI initialization
  initApp();
});

// Setup Office.js event handlers
function setupOfficeHandlers() {
  // Initial check of selection
  handleTranslationFromSelection();

  // Register selection change handler if auto-translate is on
  if (isAutoTranslate) {
    registerSelectionChangeHandler();
  }
}

// Register Selection Change Handler
function registerSelectionChangeHandler() {
  if (isSelectionHandlerRegistered) return;
  
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        isSelectionHandlerRegistered = true;
        console.log("Selection change handler registered.");
      } else {
        console.error("Failed to register handler:", result.error.message);
      }
    }
  );
}

// Unregister Selection Change Handler
function unregisterSelectionChangeHandler() {
  if (!isSelectionHandlerRegistered) return;

  Office.context.document.removeHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    { handler: onSelectionChanged },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        isSelectionHandlerRegistered = false;
        console.log("Selection change handler removed.");
      } else {
        console.error("Failed to unregister handler:", result.error.message);
      }
    }
  );
}

// Handle debounced selection changes
function onSelectionChanged() {
  if (!isAutoTranslate) return;

  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    handleTranslationFromSelection();
  }, 400); // 400ms debounce to prevent spamming
}

// Get text selection and trigger translation
function handleTranslationFromSelection() {
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const text = result.value ? result.value.toString().trim() : "";
        if (text && text !== lastTranslatedText) {
          sourceTextEl.value = text;
          ttsBtnEl.disabled = false;
          translateText(text);
        }
      }
    }
  );
}

// Main initialization function
function initApp() {
  // Theme Toggle Button Click Event
  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", () => {
      currentTheme = currentTheme === "light" ? "dark" : "light";
      localStorage.setItem("jp_vi_theme", currentTheme);
      applyTheme(currentTheme);
    });
    // Refresh theme in case the element was not ready when first applied
    applyTheme(currentTheme);
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
    phoneticSelect.value = phoneticMode;
    phoneticSelect.addEventListener("change", (e) => {
      phoneticMode = e.target.value;
      localStorage.setItem("jp_vi_phonetic_mode", phoneticMode);
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
      showToast("Vui lòng nhập hoặc chọn văn bản cần dịch.");
    }
  });

  // Auto-translate Toggle Switch Event
  autoTranslateSwitch.addEventListener("change", (e) => {
    isAutoTranslate = e.target.checked;
    if (typeof Office !== "undefined" && Office.context) {
      if (isAutoTranslate) {
        registerSelectionChangeHandler();
        handleTranslationFromSelection();
      } else {
        unregisterSelectionChangeHandler();
      }
    }
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
    localStorage.removeItem("jp_vi_history");
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
    sourceTextEl.placeholder = "Quét chọn văn bản tiếng Nhật trong tài liệu hoặc nhập tại đây...";
    
    // JA -> VI: Show source speech button, hide target speech button
    ttsBtnEl.classList.remove("hidden");
    ttsOutputBtn.classList.add("hidden");
  } else {
    sourceLangTitle.textContent = "VI";
    targetLangTitle.textContent = "JP";
    inputLangLabel.innerHTML = '<i class="fa-solid fa-location-dot"></i> VI';
    outputLangLabel.innerHTML = '<i class="fa-solid fa-circle-check"></i> JP';
    sourceTextEl.placeholder = "Quét chọn văn bản tiếng Việt trong tài liệu hoặc nhập tại đây...";
    
    // VI -> JA: Hide source speech button, show target speech button (for Japanese output)
    ttsBtnEl.classList.add("hidden");
    ttsOutputBtn.classList.remove("hidden");
  }
}



// Translate text between languages
// Display phonetics based on mode (Hiragana, Romaji, or both)
function displayPhonetics(hiraganaText, romajiText) {
  // Store the raw inputs in state
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

  // Update state with computed values if empty
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

// Translate text between languages
async function translateText(text) {
  if (!text) return;
  
  lastTranslatedText = text;
  lastHiragana = "";
  lastRomaji = "";
  
  // Show loading indicator
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

      // Add to History (normalize so 'ja' is always Japanese and 'vi' is always Vietnamese in the record)
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
  
  // Cancel current speech if speaking
  window.speechSynthesis.cancel();

  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = "ja-JP";

  // Find Japanese voice
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
    // Fallback using temp textarea
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
  let history = JSON.parse(localStorage.getItem("jp_vi_history") || "[]");
  
  // Prevent duplicate consecutive entries
  if (history.length > 0 && history[0].ja === ja && history[0].vi === vi) return;

  // Remove matching existing entry to put it at the top
  history = history.filter(item => !(item.ja === ja && item.vi === vi));

  // Unshift to top of array
  history.unshift({
    id: Date.now(),
    ja,
    vi,
    hiragana
  });

  // Limit to 20 items
  if (history.length > 20) {
    history.pop();
  }

  localStorage.setItem("jp_vi_history", JSON.stringify(history));
}

// Render history list inside UI
function renderHistory() {
  const history = JSON.parse(localStorage.getItem("jp_vi_history") || "[]");
  
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

    // Event binding
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

// Delete single item from history list
function deleteHistoryItem(id) {
  let history = JSON.parse(localStorage.getItem("jp_vi_history") || "[]");
  history = history.filter(item => item.id !== id);
  localStorage.setItem("jp_vi_history", JSON.stringify(history));
  renderHistory();
}
