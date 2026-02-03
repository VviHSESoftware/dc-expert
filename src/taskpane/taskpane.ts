/* global Excel, Office */
import { getSelectionContext } from "./office_helpers";
import { promptStream, dcexpert } from "../functions/functions";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initApp();
  }
});

function initApp() {
  // Elements
  const sendBtn = document.getElementById("send-btn");
  const userInput = document.getElementById("user-input");
  const messagesList = document.getElementById("messages-list");
  const settingsBtn = document.getElementById("settings-btn");
  const settingsModal = document.getElementById("settings-modal");
  const saveSettingsBtn = document.getElementById("save-settings");
  const closeSettingsBtn = document.getElementById("close-settings");
  const helpBtn = document.getElementById("help-btn");
  const helpModal = document.getElementById("help-modal");
  const closeHelpBtn = document.getElementById("close-help");

  // Load Settings
  loadSettings();

  // Event Listeners
  sendBtn.onclick = handleSend;
  const analyzeBtn = document.getElementById("analyze-btn");

  userInput.onkeydown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  if (analyzeBtn) {
    analyzeBtn.onclick = handleAnalyze;
  }

  settingsBtn.onclick = () => settingsModal.classList.remove("hidden");
  closeSettingsBtn.onclick = () => settingsModal.classList.add("hidden");
  helpBtn.onclick = () => helpModal.classList.remove("hidden");
  closeHelpBtn.onclick = () => helpModal.classList.add("hidden");
  saveSettingsBtn.onclick = saveSettings;

  async function handleSend() {
    const text = (userInput as HTMLTextAreaElement).value.trim();
    if (!text) return;

    (userInput as HTMLTextAreaElement).value = "";
    addMessage("user", text);

    const context = await getSelectionContext();
    const fullPrompt = `${context}\n\nUser Question: ${text}`;

    const assistantMsgDiv = addMessage("assistant", "...");
    let fullResponse = "";

    try {
      const settings = getSettings();
      if (!settings.apiKey) {
        assistantMsgDiv.innerText = "Ошибка: пожалуйста, укажите API ключ в настройках.";
        return;
      }
      // Create a dummy invocation for the streaming function
      // Since promptStream is designed for Excel Custom Functions,
      // we adapt it here for our UI.
      const invocation: any = {
        setResult: (result: string) => {
          fullResponse = result;
          assistantMsgDiv.innerText = fullResponse;
          scrollToBottom();
        },
        onCanceled: () => {},
      };

      promptStream(
        fullPrompt,
        settings.model,
        settings.apiKey,
        settings.systemPrompt ||
          "Вы — экспертный ассистент Excel по имени DC expert. Помогайте пользователю с анализом данных, формулами и задачами в таблицах.",
        settings.provider,
        invocation
      );
    } catch (error) {
      assistantMsgDiv.innerText = `Error: ${error.message}`;
    }
  }

  async function handleAnalyze() {
    const text = (userInput as HTMLTextAreaElement).value.trim();
    if (!text) {
      addMessage("system", "Пожалуйста, введите запрос для анализа.");
      return;
    }

    try {
      // Get selected cell content
      let cellContent = "";
      await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("values");
        await context.sync();

        const values = selectedRange.values;
        if (values && values.length > 0 && values[0].length > 0) {
          cellContent = String(values[0][0] || "");
        }
      });

      if (!cellContent) {
        addMessage("system", "Пожалуйста, выделите ячейку с данными для анализа.");
        return;
      }

      addMessage("user", `Анализ ячейки: ${cellContent}\n\nЗапрос: ${text}`);

      const assistantMsgDiv = addMessage("assistant", "Анализирую...");

      const settings = getSettings();
      if (!settings.apiKey) {
        assistantMsgDiv.innerText = "Ошибка: пожалуйста, укажите API ключ в настройках.";
        return;
      }

      // Use the dcexpert function directly
      const result = await dcexpert(cellContent, text);
      assistantMsgDiv.innerText = result;

      // Insert result back to the cell
      await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.values = [[result]];
        await context.sync();
      });

      addMessage("system", "✅ Результат вставлен в выделенную ячейку");
    } catch (error) {
      addMessage("system", `Ошибка: ${error instanceof Error ? error.message : "Неизвестная ошибка"}`);
    }
  }

  function addMessage(role: "user" | "assistant" | "system", text: string) {
    const div = document.createElement("div");
    div.className = `message ${role}`;
    div.innerText = text;
    messagesList.appendChild(div);
    scrollToBottom();
    return div;
  }

  function scrollToBottom() {
    messagesList.scrollTop = messagesList.scrollHeight;
  }

  function getSettings() {
    const providerEl = document.getElementById("setting-provider") as HTMLSelectElement;
    const apiKeyEl = document.getElementById("setting-key") as HTMLInputElement;
    const modelEl = document.getElementById("setting-model") as HTMLInputElement;
    const proxyEl = document.getElementById("setting-proxy") as HTMLInputElement;
    const systemPromptEl = document.getElementById("setting-system-prompt") as HTMLTextAreaElement;

    return {
      provider: providerEl ? providerEl.value : "nebius",
      apiKey: apiKeyEl ? apiKeyEl.value : "",
      model: modelEl ? modelEl.value : "meta-llama/Meta-Llama-3.1-70B-Instruct",
      proxy: proxyEl ? proxyEl.value : "",
      systemPrompt: systemPromptEl ? systemPromptEl.value : "",
    };
  }

  function saveSettings() {
    const settings = getSettings();
    localStorage.setItem("copilot_settings", JSON.stringify(settings));
    settingsModal.classList.add("hidden");
  }

  function loadSettings() {
    const saved = localStorage.getItem("copilot_settings");
    if (saved) {
      const settings = JSON.parse(saved);
      (document.getElementById("setting-provider") as HTMLSelectElement).value = settings.provider || "nebius";
      (document.getElementById("setting-key") as HTMLInputElement).value = settings.apiKey || "";
      (document.getElementById("setting-model") as HTMLInputElement).value =
        settings.model || "meta-llama/Meta-Llama-3.1-70B-Instruct";
      (document.getElementById("setting-proxy") as HTMLInputElement).value = "https://excelprx.hsecontest.ru";
      (document.getElementById("setting-system-prompt") as HTMLTextAreaElement).value =
        settings.systemPrompt ||
        "Вы — экспертный ассистент Excel по имени DC expert. Помогайте пользователю с анализом данных, формулами и задачами в таблицах.";
    } else {
      // Default proxy - run ./start-proxy.sh to start local proxy server
      (document.getElementById("setting-proxy") as HTMLInputElement).value = "https://excelprx.hsecontest.ru";
      (document.getElementById("setting-system-prompt") as HTMLTextAreaElement).value =
        "Вы — экспертный ассистент Excel по имени DC expert. Помогайте пользователю с анализом данных, формулами и задачами в таблицах.";
    }
  }
}
