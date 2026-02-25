/* global Excel, Office */
import { getSelectionContext } from "./office_helpers";
import { dcexpert, chatStream, Message } from "../functions/functions";
import { marked } from "marked";

let conversationHistory: Message[] =[];

Office.onReady((info) => {
  initApp();
});

function initApp() {
  const sendBtn = document.getElementById("send-btn");
  const userInput = document.getElementById("user-input") as HTMLTextAreaElement;
  const messagesList = document.getElementById("messages-list");
  const settingsBtn = document.getElementById("settings-btn");
  const settingsModal = document.getElementById("settings-modal");
  const saveSettingsBtn = document.getElementById("save-settings");
  const closeSettingsBtn = document.getElementById("close-settings");
  const helpBtn = document.getElementById("help-btn");
  const helpModal = document.getElementById("help-modal");
  const closeHelpBtn = document.getElementById("close-help");
  const analyzeBtn = document.getElementById("analyze-btn");
  const clearContextBtn = document.getElementById("clear-selection");

  loadSettings();

  sendBtn.onclick = handleSend;
  userInput.onkeydown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  if (analyzeBtn) analyzeBtn.onclick = handleAnalyze;

  settingsBtn.onclick = () => settingsModal.classList.remove("hidden");
  closeSettingsBtn.onclick = () => settingsModal.classList.add("hidden");
  helpBtn.onclick = () => helpModal.classList.remove("hidden");
  closeHelpBtn.onclick = () => helpModal.classList.add("hidden");
  saveSettingsBtn.onclick = saveSettings;

  if (clearContextBtn) {
    clearContextBtn.onclick = () => {
      conversationHistory =[];
      messagesList.innerHTML = "";
      addMessage("system", "Контекст диалога очищен.");
    };
  }

  async function handleSend() {
    const text = userInput.value.trim();
    if (!text) return;

    userInput.value = "";
    await addMessage("user", text);

    const context = await getSelectionContext();
    const isSelectionEmpty = context === "No data selected.";
    const userMessageContent = isSelectionEmpty ? text : `${context}\n\nUser Question: ${text}`;

    conversationHistory.push({ role: "user", content: userMessageContent });

    const assistantMsgDiv = await addMessage("assistant", "<span class='loading-dots'>Обрабатываю</span>");
    let fullResponse = "";

    try {
      const settings = getSettings();
      if (!settings.serviceToken) {
        assistantMsgDiv.innerText = "Ошибка: пожалуйста, укажите токен доступа в настройках.";
        return;
      }

      const invocation = {
        setResult: async (result: string) => {
          fullResponse = result;
          assistantMsgDiv.innerHTML = await marked.parse(fullResponse);
          enhanceCodeBlocks(assistantMsgDiv);
          scrollToBottom();
        },
        onCanceled: () => {},
      };

      const systemMessage: Message = {
        role: "system",
        content: settings.systemPrompt,
      };

      const messagesToSend =[systemMessage, ...conversationHistory];

      await chatStream(messagesToSend, invocation);
      conversationHistory.push({ role: "assistant", content: fullResponse });
    } catch (error) {
      assistantMsgDiv.innerText = `Ошибка: ${error.message}`;
      conversationHistory.pop();
    }
  }

  async function handleAnalyze() {
    const text = userInput.value.trim();
    if (!text) {
      await addMessage("system", "Пожалуйста, введите запрос для анализа.");
      return;
    }

    try {
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
        await addMessage("system", "Пожалуйста, выделите ячейку с данными для анализа.");
        return;
      }

      await addMessage("user", `Анализ ячейки: ${cellContent}\n\nЗапрос: ${text}`);
      const assistantMsgDiv = await addMessage("assistant", "<span class='loading-dots'>Анализирую</span>");

      const settings = getSettings();
      if (!settings.serviceToken) {
        assistantMsgDiv.innerText = "Ошибка: пожалуйста, укажите токен доступа в настройках.";
        return;
      }

      const result = await dcexpert(cellContent, text);
      assistantMsgDiv.innerHTML = await marked.parse(result);

      await Excel.run(async (context) => {
        const ranges = context.workbook.getSelectedRanges();
        ranges.load("areas");
        await context.sync();

        if (ranges.areas && ranges.areas.items && ranges.areas.items.length > 0) {
          const firstRange = ranges.areas.items[0];
          const parsedTable = tryParseMarkdownTable(result);

          if (parsedTable) {
            const rowsCount = parsedTable.length;
            const colsCount = parsedTable[0].length;
            const insertRange = firstRange.getCell(0, 0).getResizedRange(rowsCount - 1, colsCount - 1);
            insertRange.values = parsedTable;
          } else {
            firstRange.getCell(0, 0).values = [[result]];
          }
        }
        await context.sync();
      });

      await addMessage("system", "✅ Результат вставлен в выделенную ячейку");
    } catch (error) {
      await addMessage("system", `Ошибка: ${error instanceof Error ? error.message : "Неизвестная ошибка"}`);
    }
  }

  async function addMessage(role: "user" | "assistant" | "system", text: string) {
    const div = document.createElement("div");
    div.className = `message ${role}`;
    if (role === "system") {
      div.innerText = text;
    } else {
      div.innerHTML = await marked.parse(text);
    }
    messagesList.appendChild(div);
    scrollToBottom();
    return div;
  }

  function scrollToBottom() {
    messagesList.scrollTop = messagesList.scrollHeight;
  }

  function getSettings() {
    const tokenEl = document.getElementById("setting-token") as HTMLInputElement;
    const systemPromptEl = document.getElementById("setting-system-prompt") as HTMLTextAreaElement;

    return {
      serviceToken: tokenEl ? tokenEl.value : "",
      systemPrompt: systemPromptEl && systemPromptEl.value ? systemPromptEl.value : "Вы — экспертный ассистент Excel. Давай краткие и точные ответы на русском языке.",
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
      (document.getElementById("setting-token") as HTMLInputElement).value = settings.serviceToken || "";
      const sp = document.getElementById("setting-system-prompt") as HTMLTextAreaElement;
      if (sp && settings.systemPrompt) sp.value = settings.systemPrompt;
    }
  }

  function enhanceCodeBlocks(container: HTMLElement) {
    const preElements = container.querySelectorAll("pre");
    preElements.forEach((pre) => {
      const codeElement = pre.querySelector("code");
      if (codeElement && (window as any).hljs) {
        (window as any).hljs.highlightElement(codeElement);
      }
      if (!pre.querySelector(".copy-btn")) {
        const copyBtn = document.createElement("button");
        copyBtn.className = "copy-btn";
        copyBtn.innerText = "Copy";
        copyBtn.onclick = () => {
          if (codeElement) {
            navigator.clipboard.writeText(codeElement.innerText).then(() => {
              copyBtn.innerText = "Copied!";
              setTimeout(() => { copyBtn.innerText = "Copy"; }, 2000);
            });
          }
        };
        pre.appendChild(copyBtn);
      }
    });
  }

  function tryParseMarkdownTable(text: string): string[][] | null {
    const lines = text.trim().split("\n");
    const tableLines = lines.filter((line) => line.trim().startsWith("|") && line.trim().endsWith("|"));
    if (tableLines.length < 2) return null;
    const separatorLine = tableLines[1];
    if (!/^\|[\s\-:|]+\|$/.test(separatorLine.trim())) return null;

    const resultTable: string[][] =[];
    for (let i = 0; i < tableLines.length; i++) {
      if (i === 1) continue;
      const line = tableLines[i].trim();
      const rowContent = line.substring(1, line.length - 1);
      const cells = rowContent.split("|").map((cell) => cell.trim());
      if (resultTable.length > 0 && cells.length !== resultTable[0].length) return null;
      resultTable.push(cells);
    }
    return resultTable.length > 0 ? resultTable : null;
  }
}