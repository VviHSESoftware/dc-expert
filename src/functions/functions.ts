/* global CustomFunctions */

export type Message = { role: "user" | "system" | "assistant"; content: string };

// ВАЖНО: Укажите здесь URL вашего боевого прокси-сервера. 
// Для локальной разработки оставьте localhost.
const PROXY_URL = "https://excelprx.hsecontest.ru/api/chat"; 

/**
 * Получает сохраненный сервисный токен из LocalStorage
 */
function getStoredToken(): string {
  const saved = localStorage.getItem("copilot_settings");
  if (saved) {
    const parsed = JSON.parse(saved);
    return parsed.serviceToken || "";
  }
  return "";
}

/**
 * Внутренняя функция для стриминга ответов в чат (не торчит в Excel напрямую).
 */
export async function chatStream(
  messages: Message[],
  invocation: { setResult: (res: string) => Promise<void> | void; onCanceled: () => void }
): Promise<void> {
  try {
    const token = getStoredToken();
    if (!token) {
      throw new Error("Не указан сервисный токен. Откройте настройки плагина (шестеренка) и введите токен.");
    }

    const response = await fetch(PROXY_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({ messages, stream: true })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error || `Ошибка сервера: ${response.status}`);
    }

    const reader = response.body?.getReader();
    const decoder = new TextDecoder();
    let fullResponse = "";

    if (reader) {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;

        const chunk = decoder.decode(value, { stream: true });
        const lines = chunk.split("\n");

        for (const line of lines) {
          if (line.startsWith("data: ")) {
            const data = line.slice(6);
            if (data.trim() === "[DONE]") continue;

            try {
              const parsed = JSON.parse(data);
              const content = parsed.choices?.[0]?.delta?.content || "";
              if (content) {
                fullResponse += content;
                await invocation.setResult(fullResponse);
              }
            } catch (e) {
              // Игнорируем ошибки парсинга неполных JSON-чанков
            }
          }
        }
      }
    }

    if (!fullResponse) {
      throw new Error("Получен пустой ответ от сервера.");
    }
  } catch (error) {
    console.error("Error in chatStream:", error);
    await invocation.setResult(`Ошибка: ${error instanceof Error ? error.message : "Неизвестная ошибка"}`);
  }
}

/**
 * Анализирует выделенные данные с помощью DC Expert.
 * @customfunction DCEXPERT
 * @param cellValue Значение или диапазон ячеек для анализа.
 * @param query Конкретный вопрос или задача для ИИ.
 * @returns Итоговый ответ ИИ в формате текста.
 */
export async function dcexpert(cellValue: any, query: string): Promise<string> {
  try {
    const token = getStoredToken();
    if (!token) {
      return "Ошибка: Не указан сервисный токен. Настройте плагин на панели DC Expert.";
    }

    const systemPrompt = `Ты — DC Expert, профессиональный ассистент Excel. 
Твоя задача: предоставлять краткие, точные и полезные ответы на запросы пользователя по данным из ячеек.
СТРОГИЕ ПРАВИЛА:
1. Выдавай ТОЛЬКО результат без лишних слов, приветствий и пояснений.
2. КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО использовать разметку Markdown (бектики \`, звездочки *, решетки #).
3. Весь текст должен быть простым и чистым, чтобы его можно было сразу использовать в Excel.
4. Ответ должен быть на русском языке.`;

    const messages: Message[] =[
      { role: "system", content: systemPrompt },
      { role: "user", content: `Данные из ячеек: ${JSON.stringify(cellValue)}\n\nЗапрос: ${query}` },
    ];

    const response = await fetch(PROXY_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({ messages, stream: false })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      return `Ошибка: ${errorData.error || response.statusText}`;
    }

    const data = await response.json();
    return data.choices?.[0]?.message?.content?.trim() || "Пустой ответ от сервера.";
  } catch (error) {
    console.error("Error in dcexpert:", error);
    return `Ошибка: ${error instanceof Error ? error.message : "Неизвестная ошибка"}`;
  }
}

CustomFunctions.associate("DCEXPERT", dcexpert);