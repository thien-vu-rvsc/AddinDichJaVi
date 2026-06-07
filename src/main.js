// Main Application Logic for JP-VI Translator Add-in
import './style.css';
import * as wanakana from 'wanakana';
import { CreateMLCEngine } from "@mlc-ai/web-llm";

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
const engineSelect = document.getElementById("engine-select");
const webllmSettingsGroup = document.getElementById("webllm-settings-group");
const webllmModelSelect = document.getElementById("webllm-model-select");
const webllmInitBtn = document.getElementById("webllm-init-btn");
const webllmProgressContainer = document.getElementById("webllm-progress-container");
const webllmStatusText = document.getElementById("webllm-status-text");
const webllmPercentText = document.getElementById("webllm-percent-text");
const webllmProgressBar = document.getElementById("webllm-progress-bar");

// Ollama UI Element References
const ollamaSettingsGroup = document.getElementById("ollama-settings-group");
const ollamaStatusLed = document.getElementById("ollama-status-led");
const ollamaStatusText = document.getElementById("ollama-status-text");
const ollamaModelSelect = document.getElementById("ollama-model-select");
const ollamaRefreshBtn = document.getElementById("ollama-refresh-btn");

// State Variables
let isAutoTranslate = true;
let isSelectionHandlerRegistered = false;
let debounceTimer = null;
let lastTranslatedText = "";
let isJaToVi = true; // true: Japanese -> Vietnamese, false: Vietnamese -> Japanese

// Translation Engine & Model State
let currentEngine = localStorage.getItem("jp_vi_engine") || "google";
let webllmModel = localStorage.getItem("jp_vi_webllm_model") || "Qwen2.5-1.5B-Instruct-q4f16_1-MLC";
let webllmEngine = null;
let isWebllmLoading = false;
let ollamaModel = localStorage.getItem("jp_vi_ollama_model") || "qwen2.5:1.5b";
let isOllamaConnected = false;

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
      } else if (targetTab === "tab-settings") {
        if (currentEngine === "ollama") {
          checkOllamaConnection();
        }
      }
    });
  });

  // Engine Selection Dropdown Event
  if (engineSelect) {
    engineSelect.value = currentEngine;
    engineSelect.addEventListener("change", (e) => {
      currentEngine = e.target.value;
      localStorage.setItem("jp_vi_engine", currentEngine);
      updateEngineUI();
      if (currentEngine === "ollama") {
        checkOllamaConnection();
      }
    });
  }

  // WebLLM Model Selection Dropdown Event
  if (webllmModelSelect) {
    webllmModelSelect.value = webllmModel;
    webllmModelSelect.addEventListener("change", (e) => {
      webllmModel = e.target.value;
      localStorage.setItem("jp_vi_webllm_model", webllmModel);
      // Reset engine if model changes so it has to reinitialize
      webllmEngine = null;
      if (webllmInitBtn) {
        webllmInitBtn.disabled = false;
        webllmInitBtn.innerHTML = '<i class="fa-solid fa-download"></i> Tải & Khởi tạo Mô hình';
      }
    });
  }

  // WebLLM Initialize Button Event
  if (webllmInitBtn) {
    webllmInitBtn.addEventListener("click", () => {
      initWebLLM();
    });
  }

  // Ollama Model Selection Dropdown Event
  if (ollamaModelSelect) {
    ollamaModelSelect.value = ollamaModel;
    ollamaModelSelect.addEventListener("change", (e) => {
      ollamaModel = e.target.value;
      localStorage.setItem("jp_vi_ollama_model", ollamaModel);
    });
  }

  // Ollama Refresh Button Event
  if (ollamaRefreshBtn) {
    ollamaRefreshBtn.addEventListener("click", () => {
      checkOllamaConnection();
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

  // Clear All History Button Click Event
  clearHistoryBtn.addEventListener("click", () => {
    localStorage.removeItem("jp_vi_history");
    renderHistory();
    showToast("Đã xóa toàn bộ lịch sử.");
  });

  // Initial UI Setup
  updateLanguageUI();
  updateEngineUI();
  renderHistory();
}

// Update UI based on active engine
function updateEngineUI() {
  if (currentEngine === "webllm") {
    if (webllmSettingsGroup) webllmSettingsGroup.classList.remove("hidden");
    if (ollamaSettingsGroup) ollamaSettingsGroup.classList.add("hidden");
    if (sourceTextEl) {
      sourceTextEl.placeholder = isJaToVi 
        ? "Quét chọn văn bản tiếng Nhật hoặc tự nhập để dịch offline..."
        : "Quét chọn văn bản tiếng Việt hoặc tự nhập để dịch offline...";
    }
    
    // Update init button state if engine is already loaded
    if (webllmEngine) {
      if (webllmInitBtn) {
        webllmInitBtn.disabled = true;
        webllmInitBtn.innerHTML = '<i class="fa-solid fa-circle-check"></i> Mô hình đã sẵn sàng';
      }
    } else if (isWebllmLoading) {
      if (webllmInitBtn) {
        webllmInitBtn.disabled = true;
        webllmInitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang tải...';
      }
    } else {
      if (webllmInitBtn) {
        webllmInitBtn.disabled = false;
        webllmInitBtn.innerHTML = '<i class="fa-solid fa-download"></i> Tải & Khởi tạo Mô hình';
      }
    }
  } else if (currentEngine === "ollama") {
    if (webllmSettingsGroup) webllmSettingsGroup.classList.add("hidden");
    if (ollamaSettingsGroup) ollamaSettingsGroup.classList.remove("hidden");
    if (sourceTextEl) {
      sourceTextEl.placeholder = isJaToVi 
        ? "Quét chọn văn bản tiếng Nhật hoặc tự nhập để dịch offline (CPU)..."
        : "Quét chọn văn bản tiếng Việt hoặc tự nhập để dịch offline (CPU)...";
    }
  } else {
    if (webllmSettingsGroup) webllmSettingsGroup.classList.add("hidden");
    if (ollamaSettingsGroup) ollamaSettingsGroup.classList.add("hidden");
    if (sourceTextEl) {
      sourceTextEl.placeholder = isJaToVi
        ? "Quét chọn văn bản tiếng Nhật trong tài liệu hoặc nhập tại đây..."
        : "Quét chọn văn bản tiếng Việt trong tài liệu hoặc nhập tại đây...";
    }
  }
}

// Initialize WebLLM Engine
async function initWebLLM() {
  if (isWebllmLoading || webllmEngine) return;

  isWebllmLoading = true;
  if (webllmInitBtn) {
    webllmInitBtn.disabled = true;
    webllmInitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang khởi tạo...';
  }
  if (webllmProgressContainer) webllmProgressContainer.classList.remove("hidden");
  if (webllmStatusText) webllmStatusText.textContent = "Đang chuẩn bị...";
  if (webllmPercentText) webllmPercentText.textContent = "0%";
  if (webllmProgressBar) webllmProgressBar.style.width = "0%";

  try {
    webllmEngine = await CreateMLCEngine(webllmModel, {
      initProgressCallback: (report) => {
        const percent = Math.round(report.progress * 100);
        if (webllmPercentText) webllmPercentText.textContent = `${percent}%`;
        if (webllmProgressBar) webllmProgressBar.style.width = `${percent}%`;
        
        // Clean up status text description to make it short and clean
        let status = report.text;
        if (status.includes("Fetch")) {
          status = "Đang tải dữ liệu mô hình...";
        } else if (status.includes("Loading")) {
          status = "Đang nạp mô hình vào GPU...";
        } else if (status.includes("compile") || status.includes("Shader")) {
          status = "Đang biên dịch Shader WebGPU...";
        } else if (status.includes("Finish")) {
          status = "Đã hoàn thành!";
        }
        if (webllmStatusText) webllmStatusText.textContent = status;
      }
    });

    isWebllmLoading = false;
    if (webllmInitBtn) {
      webllmInitBtn.innerHTML = '<i class="fa-solid fa-circle-check"></i> Mô hình đã sẵn sàng';
    }
    if (webllmStatusText) webllmStatusText.textContent = "Mô hình đã sẵn sàng ngoại tuyến!";
    
    // Auto-hide progress bar after 3 seconds
    setTimeout(() => {
      if (webllmProgressContainer) webllmProgressContainer.classList.add("hidden");
    }, 3000);
    
    showToast("Đã tải và khởi tạo mô hình thành công!");
  } catch (error) {
    console.error("WebLLM Init Error:", error);
    isWebllmLoading = false;
    if (webllmInitBtn) {
      webllmInitBtn.disabled = false;
      webllmInitBtn.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Lỗi, thử lại';
    }
    if (webllmStatusText) webllmStatusText.textContent = "Lỗi: " + error.message;
    if (webllmProgressBar) webllmProgressBar.style.width = "0%";
    showToast("Khởi tạo mô hình thất bại. Vui lòng kiểm tra thiết bị hỗ trợ WebGPU.");
  }
}

// Check connection and fetch models from local Ollama
async function checkOllamaConnection() {
  if (!ollamaStatusLed || !ollamaStatusText || !ollamaModelSelect) return;

  // Show testing status
  ollamaStatusLed.className = "status-led led-orange";
  ollamaStatusText.textContent = "Đang kết nối...";

  try {
    const response = await fetch("/api/ollama/api/tags");
    if (!response.ok) throw new Error("Không phản hồi từ Ollama.");

    const data = await response.json();
    isOllamaConnected = true;
    ollamaStatusLed.className = "status-led led-green";
    ollamaStatusText.textContent = "Đã kết nối";

    // Populate dynamic models
    const availableModels = data.models || [];
    
    // Clear current select
    ollamaModelSelect.innerHTML = "";

    if (availableModels.length === 0) {
      // If connected but no models downloaded, add defaults
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "Chưa tải model (Chạy cmd: ollama run gemma4:e2b)";
      ollamaModelSelect.appendChild(opt);
    } else {
      availableModels.forEach((model) => {
        const opt = document.createElement("option");
        opt.value = model.name;
        opt.textContent = `${model.name} (${(model.size / (1024 * 1024 * 1024)).toFixed(2)} GB)`;
        ollamaModelSelect.appendChild(opt);
      });

      // Restore previously selected model if it exists in the list
      const restoredModel = localStorage.getItem("jp_vi_ollama_model");
      if (restoredModel && availableModels.some(m => m.name === restoredModel)) {
        ollamaModel = restoredModel;
      } else {
        // Fallback to first model in list
        ollamaModel = availableModels[0].name;
        localStorage.setItem("jp_vi_ollama_model", ollamaModel);
      }
      ollamaModelSelect.value = ollamaModel;
    }
  } catch (error) {
    console.error("Ollama connection error:", error);
    isOllamaConnected = false;
    ollamaStatusLed.className = "status-led led-red";
    ollamaStatusText.textContent = "Chưa kết nối";
    
    // Fallback static option list if disconnected
    ollamaModelSelect.innerHTML = `
      <option value="qwen2.5:1.5b">Qwen 2.5 1.5B (986MB)</option>
      <option value="gemma4:e2b">Gemma 4 E2B (1.6GB)</option>
      <option value="gemma4:e4b">Gemma 4 E4B (2.8GB)</option>
      <option value="llama3.2:1b">Llama 3.2 1B (1.2GB)</option>
    `;
    ollamaModelSelect.value = ollamaModel;
  }
}

// Update UI text based on current translation direction
function updateLanguageUI() {
  if (isJaToVi) {
    sourceLangTitle.textContent = "Tiếng Nhật";
    targetLangTitle.textContent = "Tiếng Việt";
    inputLangLabel.innerHTML = '<i class="fa-solid fa-location-dot"></i> Tiếng Nhật';
    outputLangLabel.innerHTML = '<i class="fa-solid fa-circle-check"></i> Tiếng Việt';
    sourceTextEl.placeholder = "Quét chọn văn bản tiếng Nhật trong tài liệu hoặc nhập tại đây...";
    
    // JA -> VI: Show source speech button, hide target speech button
    ttsBtnEl.classList.remove("hidden");
    ttsOutputBtn.classList.add("hidden");
  } else {
    sourceLangTitle.textContent = "Tiếng Việt";
    targetLangTitle.textContent = "Tiếng Nhật";
    inputLangLabel.innerHTML = '<i class="fa-solid fa-location-dot"></i> Tiếng Việt';
    outputLangLabel.innerHTML = '<i class="fa-solid fa-circle-check"></i> Tiếng Nhật';
    sourceTextEl.placeholder = "Quét chọn văn bản tiếng Việt trong tài liệu hoặc nhập tại đây...";
    
    // VI -> JA: Hide source speech button, show target speech button (for Japanese output)
    ttsBtnEl.classList.add("hidden");
    ttsOutputBtn.classList.remove("hidden");
  }
}

// Helper to extract and parse JSON from LLM response
function extractJSON(text) {
  if (!text) return null;
  const trimmed = text.trim();
  
  // Try direct parsing
  try {
    return JSON.parse(trimmed);
  } catch (e) {
    // Ignore and proceed
  }

  // Regex to match anything between the first { and the last }
  try {
    const jsonMatch = trimmed.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {
    console.error("Failed to parse extracted JSON block:", e);
  }

  return null;
}

// Translate text between languages
async function translateText(text) {
  if (!text) return;
  
  lastTranslatedText = text;
  
  // Show loading indicator
  translationOutput.innerHTML = '<span class="loading"><i class="fa-solid fa-circle-notch fa-spin"></i> Đang dịch...</span>';
  resultCard.classList.remove("hidden");
  inputHiraganaContainer.classList.add("hidden");
  outputHiraganaContainer.classList.add("hidden");

  // WebLLM Local Translation
  if (currentEngine === "webllm") {
    if (!webllmEngine) {
      translationOutput.innerHTML = '<span style="color: #ff5252; font-size: 13px;"><i class="fa-solid fa-triangle-exclamation"></i> Vui lòng vào phần <strong>Cài Đặt</strong> tải và khởi tạo mô hình trước khi dịch.</span>';
      return;
    }

    try {
      const systemPrompt = isJaToVi
        ? `You are a professional Japanese to Vietnamese translator. Translate the input Japanese text to Vietnamese and provide the Hiragana reading of the Japanese text. You must return a JSON object with this schema: { "translation": "Vietnamese translation", "hiragana": "Hiragana reading of the Japanese text (convert Kanji to Hiragana, keep Hiragana/Katakana as is)" }`
        : `You are a professional Vietnamese to Japanese translator. Translate the input Vietnamese text to Japanese and provide the Hiragana reading of the Japanese translation. You must return a JSON object with this schema: { "translation": "Japanese translation", "hiragana": "Hiragana reading of the Japanese translation (convert Kanji to Hiragana, keep Hiragana/Katakana as is)" }`;

      const messages = [
        { role: "system", content: systemPrompt },
        { role: "user", content: text }
      ];

      const reply = await webllmEngine.chat.completions.create({
        messages
      });

      const rawContent = reply.choices[0].message.content.trim();
      const data = extractJSON(rawContent);

      let translation = "";
      let hiragana = "";

      if (data) {
        translation = data.translation || "";
        hiragana = data.hiragana || "";
      } else {
        // Fallback: use raw content as translation if not JSON
        translation = rawContent;
      }

      if (translation) {
        translationOutput.textContent = translation;

        // Update Hiragana UI based on current translation direction
        if (hiragana) {
          if (isJaToVi) {
            inputHiraganaOutput.textContent = hiragana;
            inputHiraganaContainer.classList.remove("hidden");
            outputHiraganaContainer.classList.add("hidden");
          } else {
            outputHiraganaOutput.textContent = hiragana;
            outputHiraganaContainer.classList.remove("hidden");
            inputHiraganaContainer.classList.add("hidden");
          }
        } else {
          inputHiraganaContainer.classList.add("hidden");
          outputHiraganaContainer.classList.add("hidden");
        }

        // Add to History
        const jaText = isJaToVi ? text : translation;
        const viText = isJaToVi ? translation : text;
        saveToHistory(jaText, viText, hiragana);
      } else {
        translationOutput.textContent = "Không thể tạo bản dịch phù hợp.";
      }
    } catch (error) {
      console.error("WebLLM translation error:", error);
      translationOutput.textContent = "Lỗi khi chạy mô hình cục bộ: " + error.message;
    }
    return;
  }

  // Ollama Local Translation (CPU/GPU-accelerated)
  if (currentEngine === "ollama") {
    try {
      const systemPrompt = isJaToVi
        ? `You are a professional Japanese to Vietnamese translator. Translate the input Japanese text to Vietnamese and provide the Hiragana reading of the Japanese text. You must return a JSON object with this schema: { "translation": "Vietnamese translation", "hiragana": "Hiragana reading of the Japanese text (convert Kanji to Hiragana, keep Hiragana/Katakana as is)" }`
        : `You are a professional Vietnamese to Japanese translator. Translate the input Vietnamese text to Japanese and provide the Hiragana reading of the Japanese translation. You must return a JSON object with this schema: { "translation": "Japanese translation", "hiragana": "Hiragana reading of the Japanese translation (convert Kanji to Hiragana, keep Hiragana/Katakana as is)" }`;

      const response = await fetch("/api/ollama/api/chat", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: ollamaModel,
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: text }
          ],
          format: "json", // Constrains output to JSON
          stream: false
        })
      });

      if (!response.ok) {
        throw new Error(`Ollama trả về mã lỗi: ${response.status}`);
      }

      const replyData = await response.json();
      const rawContent = replyData.message.content.trim();
      const data = extractJSON(rawContent);

      let translation = "";
      let hiragana = "";

      if (data) {
        translation = data.translation || "";
        hiragana = data.hiragana || "";
      } else {
        translation = rawContent;
      }

      if (translation) {
        translationOutput.textContent = translation;

        // Update Hiragana UI based on current translation direction
        if (hiragana) {
          if (isJaToVi) {
            inputHiraganaOutput.textContent = hiragana;
            inputHiraganaContainer.classList.remove("hidden");
            outputHiraganaContainer.classList.add("hidden");
          } else {
            outputHiraganaOutput.textContent = hiragana;
            outputHiraganaContainer.classList.remove("hidden");
            inputHiraganaContainer.classList.add("hidden");
          }
        } else {
          inputHiraganaContainer.classList.add("hidden");
          outputHiraganaContainer.classList.add("hidden");
        }

        // Add to History
        const jaText = isJaToVi ? text : translation;
        const viText = isJaToVi ? translation : text;
        saveToHistory(jaText, viText, hiragana);
      } else {
        translationOutput.textContent = "Không nhận được phản hồi phù hợp từ mô hình local.";
      }
    } catch (error) {
      console.error("Ollama translation error:", error);
      translationOutput.innerHTML = `<span style="color: #ff5252; font-size: 13px;"><i class="fa-solid fa-triangle-exclamation"></i> Lỗi kết nối Ollama: ${error.message}. Hãy đảm bảo Ollama đang chạy và bạn đã tải mô hình bằng lệnh: <code style="display:block;background:rgba(0,0,0,0.3);padding:4px;margin-top:4px;user-select:all;">ollama run ${ollamaModel}</code></span>`;
    }
    return;
  }

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
      
      // Convert Romaji to Hiragana using WanaKana
      let hiragana = "";
      if (romaji && typeof wanakana !== "undefined") {
        hiragana = wanakana.toHiragana(romaji);
      }
      
      // Update Hiragana UI based on current translation direction
      if (hiragana) {
        if (isJaToVi) {
          // JA -> VI: Japanese is the source, show Hiragana under input card text area
          inputHiraganaOutput.textContent = hiragana;
          inputHiraganaContainer.classList.remove("hidden");
          outputHiraganaContainer.classList.add("hidden");
        } else {
          // VI -> JA: Japanese is the target, show Hiragana under output card translated text
          outputHiraganaOutput.textContent = hiragana;
          outputHiraganaContainer.classList.remove("hidden");
          inputHiraganaContainer.classList.add("hidden");
        }
      } else {
        inputHiraganaContainer.classList.add("hidden");
        outputHiraganaContainer.classList.add("hidden");
      }

      // Add to History (normalize so 'ja' is always Japanese and 'vi' is always Vietnamese in the record)
      const jaText = isJaToVi ? text : translation;
      const viText = isJaToVi ? translation : text;
      saveToHistory(jaText, viText, hiragana);
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
async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    showToast("Đã sao chép bản dịch!");
  } catch (err) {
    // Fallback using temp textarea
    const tempTextarea = document.createElement("textarea");
    tempTextarea.value = text;
    tempTextarea.style.position = "fixed";
    document.body.appendChild(tempTextarea);
    tempTextarea.select();
    try {
      document.execCommand("copy");
      showToast("Đã sao chép bản dịch!");
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
