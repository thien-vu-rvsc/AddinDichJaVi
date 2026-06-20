// Content script to detect text selection on web pages and send it to the side panel

document.addEventListener("mouseup", () => {
  const selectedText = window.getSelection().toString().trim();
  if (selectedText) {
    // Send selected text to runtime (side panel)
    chrome.runtime.sendMessage({ type: "TEXT_SELECTED", text: selectedText }).catch((err) => {
      // Ignore errors when side panel is not open or extension context is invalidated
    });
  }
});
