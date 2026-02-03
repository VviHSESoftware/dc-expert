/* global CustomFunctions  */
import { OpenAI } from "openai";
import Anthropic from "@anthropic-ai/sdk";
import { MessageParam } from "@anthropic-ai/sdk/resources";

type Provider = "openai" | "anthropic" | "nebius";
type Message = { role: "user" | "system"; content: string };

interface AIClient {
  generateCompletion(messages: Message[], model: string): Promise<string>;
  generateStreamingCompletion(messages: Message[], model: string, onChunk: (chunk: string) => void): Promise<void>;
}

class OpenAIClient implements AIClient {
  private client: OpenAI;
  private proxy: string | undefined;

  constructor(apiKey: string, proxy?: string) {
    this.proxy = proxy;
    this.client = new OpenAI({ apiKey, dangerouslyAllowBrowser: true });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    if (this.proxy) {
      return this.makeProxiedRequest(messages, model, false);
    }

    const response = await this.client.chat.completions.create({ messages, model });
    return response.choices[0].message.content || "No content in response";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    if (this.proxy) {
      await this.makeProxiedRequest(messages, model, true, onChunk);
      return;
    }

    const stream = await this.client.chat.completions.create({ messages, model, stream: true });
    for await (const chunk of stream) {
      const content = chunk.choices[0]?.delta?.content || "";
      onChunk(content);
    }
  }

  // Proxy request handler
  private async makeProxiedRequest(
    messages: Message[],
    model: string,
    stream: boolean,
    onChunk?: (chunk: string) => void
  ): Promise<string> {
    try {
      // Hardcoded proxy - 168.90.196.95:8000:Xjyc9L:bEJrmk
      const targetUrl = "https://api.openai.com/v1/chat/completions";

      console.log("Making proxied request through local proxy server");

      const response = await fetch(`${this.proxy}/proxy/openai`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${this.client.apiKey}`,
          "x-target-url": targetUrl,
        },
        body: JSON.stringify({
          messages,
          model,
          stream,
        }),
      });

      if (!response.ok) {
        throw new Error(`Proxy request failed: ${response.status} ${response.statusText}`);
      }

      if (stream) {
        // Handle streaming response
        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        let fullResponse = "";

        if (reader && onChunk) {
          for (;;) {
            const { done, value } = await reader.read();
            if (done) break;

            const chunk = decoder.decode(value);
            const lines = chunk.split("\n");

            for (const line of lines) {
              if (line.startsWith("data: ")) {
                const data = line.slice(6);
                if (data === "[DONE]") continue;

                try {
                  const parsed = JSON.parse(data);
                  const content = parsed.choices?.[0]?.delta?.content || "";
                  if (content) {
                    fullResponse += content;
                    onChunk(content);
                  }
                } catch (e) {
                  // Ignore parsing errors
                }
              }
            }
          }
        }
        return fullResponse;
      } else {
        // Handle regular response
        const data = await response.json();
        return data.choices?.[0]?.message?.content || "No content in response";
      }
    } catch (error) {
      console.error("Proxy request failed:", error);
      throw new Error(
        `Proxy request failed: ${error instanceof Error ? error.message : "Unknown error"}. Make sure the proxy server is running (npm run proxy-start)`
      );
    }
  }
}

class NebiusClient implements AIClient {
  private apiKey: string;
  private proxy: string | undefined;

  constructor(apiKey: string, proxy?: string) {
    this.apiKey = apiKey;
    this.proxy = proxy;
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    if (this.proxy) {
      return this.makeProxiedRequest(messages, model, false);
    }

    const client = new OpenAI({
      apiKey: this.apiKey,
      baseURL: "https://api.studio.nebius.ai/v1/",
      dangerouslyAllowBrowser: true,
    });
    const response = await client.chat.completions.create({ messages, model });
    return response.choices[0].message.content || "";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    if (this.proxy) {
      await this.makeProxiedRequest(messages, model, true, onChunk);
      return;
    }

    const client = new OpenAI({
      apiKey: this.apiKey,
      baseURL: "https://api.studio.nebius.ai/v1/",
      dangerouslyAllowBrowser: true,
    });

    const stream = await client.chat.completions.create({ messages, model, stream: true });
    for await (const chunk of stream) {
      const content = chunk.choices[0]?.delta?.content || "";
      onChunk(content);
    }
  }

  // Proxy request handler
  private async makeProxiedRequest(
    messages: Message[],
    model: string,
    stream: boolean,
    onChunk?: (chunk: string) => void
  ): Promise<string> {
    try {
      const proxyServer = this.proxy;
      const targetUrl = "https://api.studio.nebius.ai/v1/chat/completions";

      console.log("Making proxied request through local proxy server");

      const response = await fetch(`${proxyServer}/proxy/nebius`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${this.apiKey}`,
          "x-target-url": targetUrl,
        },
        body: JSON.stringify({
          messages,
          model,
          stream,
        }),
      });

      if (!response.ok) {
        throw new Error(`Proxy request failed: ${response.status} ${response.statusText}`);
      }

      if (stream) {
        // Handle streaming response
        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        let fullResponse = "";

        if (reader && onChunk) {
          for (;;) {
            const { done, value } = await reader.read();
            if (done) break;

            const chunk = decoder.decode(value);
            const lines = chunk.split("\n");

            for (const line of lines) {
              if (line.startsWith("data: ")) {
                const data = line.slice(6);
                if (data === "[DONE]") continue;

                try {
                  const parsed = JSON.parse(data);
                  const content = parsed.choices?.[0]?.delta?.content || "";
                  if (content) {
                    fullResponse += content;
                    onChunk(content);
                  }
                } catch (e) {
                  // Ignore parsing errors
                }
              }
            }
          }
        }
        return fullResponse;
      } else {
        // Handle regular response
        const data = await response.json();
        return data.choices?.[0]?.message?.content || "No content in response";
      }
    } catch (error) {
      console.error("Proxy request failed:", error);
      throw new Error(
        `Proxy request failed: ${error instanceof Error ? error.message : "Unknown error"}. Make sure the proxy server is running (npm run proxy-start)`
      );
    }
  }
}

class AnthropicClient implements AIClient {
  private client: Anthropic;
  private proxy: string | undefined;

  constructor(apiKey: string, proxy?: string) {
    this.proxy = proxy;
    this.client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    if (this.proxy) {
      // Handle proxy format: host:port:user:pass
      return this.makeProxiedRequest(messages, model, false);
    }

    const systemMessage = messages.find((msg) => msg.role === "system");
    const userMessages = messages.filter((msg) => msg.role === "user");

    const response = await this.client.messages.create({
      messages: userMessages as MessageParam[],
      model,
      max_tokens: 1000,
      system: systemMessage?.content,
    });
    return response.content[0].type === "text" ? response.content[0].text : "";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    if (this.proxy) {
      // Handle proxy format: host:port:user:pass
      await this.makeProxiedRequest(messages, model, true, onChunk);
      return;
    }

    const systemMessage = messages.find((msg) => msg.role === "system");
    const userMessages = messages.filter((msg) => msg.role === "user");

    const stream = await this.client.messages.create({
      messages: userMessages as MessageParam[],
      model,
      max_tokens: 1000,
      stream: true,
      system: systemMessage?.content,
    });
    for await (const chunk of stream) {
      if (chunk.type === "content_block_delta" && chunk.delta.type === "text_delta") {
        onChunk(chunk.delta.text);
      }
    }
  }

  // Proxy request handler
  private async makeProxiedRequest(
    messages: Message[],
    model: string,
    stream: boolean,
    onChunk?: (chunk: string) => void
  ): Promise<string> {
    try {
      // Hardcoded proxy - 168.90.196.95:8000:Xjyc9L:bEJrmk
      const targetUrl = "https://api.anthropic.com/v1/messages";

      console.log("Making proxied request through local proxy server");

      const systemMessage = messages.find((msg) => msg.role === "system");
      const userMessages = messages.filter((msg) => msg.role === "user");

      const requestBody: any = {
        messages: userMessages as MessageParam[],
        model,
        max_tokens: 1000,
        stream,
      };

      if (systemMessage?.content) {
        requestBody.system = systemMessage.content;
      }

      const response = await fetch(`${this.proxy}/proxy/anthropic`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": this.client.apiKey,
          "x-target-url": targetUrl,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`Proxy request failed: ${response.status} ${response.statusText}`);
      }

      if (stream) {
        // Handle streaming response
        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        let fullResponse = "";

        if (reader && onChunk) {
          for (;;) {
            const { done, value } = await reader.read();
            if (done) break;

            const chunk = decoder.decode(value);
            const lines = chunk.split("\n");

            for (const line of lines) {
              if (line.startsWith("data: ")) {
                const data = line.slice(6);
                if (data === "[DONE]") continue;

                try {
                  const parsed = JSON.parse(data);
                  if (parsed.type === "content_block_delta" && parsed.delta?.type === "text_delta") {
                    const content = parsed.delta.text;
                    if (content) {
                      fullResponse += content;
                      onChunk(content);
                    }
                  }
                } catch (e) {
                  // Ignore parsing errors
                }
              }
            }
          }
        }
        return fullResponse;
      } else {
        // Handle regular response
        const data = await response.json();
        return data.content?.[0]?.text || "No content in response";
      }
    } catch (error) {
      console.error("Proxy request failed:", error);
      throw new Error(
        `Proxy request failed: ${error instanceof Error ? error.message : "Unknown error"}. Make sure the proxy server is running (npm run proxy-start)`
      );
    }
  }
}

function createAIClient(provider: Provider, apiKey: string, proxy?: string): AIClient {
  switch (provider) {
    case "openai":
      return new OpenAIClient(apiKey, proxy);
    case "anthropic":
      return new AnthropicClient(apiKey, proxy);
    case "nebius":
      return new NebiusClient(apiKey, proxy);
    default:
      throw new Error(`Unsupported provider: ${provider}`);
  }
}

/**
 * Helper to get settings from storage
 */
function getStoredSettings() {
  const saved = localStorage.getItem("copilot_settings");
  if (saved) {
    return JSON.parse(saved);
  }
  return {
    provider: "nebius",
    model: "meta-llama/Meta-Llama-3.1-70B-Instruct",
    apiKey: "",
    proxy: "http://localhost:8080/proxy/nebius",
    systemPrompt:
      "Вы — экспертный ассистент Excel по имени DC expert. Помогайте пользователю с анализом данных, формулами и задачами в таблицах.",
  };
}

/**
 * Generates a response based on the given prompt using the specified AI model and provider.
 * @customfunction PROMPT
 * @helpUrl https://dcexpert.liminity.se/help
 * @param message The prompt message to send to the AI. a
 * @param model The AI model to use for generating the response.
 * @param apiKey The API key for the AI service.
 * @param systemPrompt An optional system prompt to provide context for the AI.
 * @param provider The AI provider (Nebius).
 * @returns A promise that resolves to the generated response.
 */
export async function prompt(
  message: string,
  model: string,
  apiKey: string,
  systemPrompt: string,
  provider: string
): Promise<string> {
  try {
    if (!message || !model || !apiKey || !provider) {
      throw new Error("Missing required parameters");
    }

    if (
      provider.toLowerCase() !== "nebius" &&
      provider.toLowerCase() !== "openai" &&
      provider.toLowerCase() !== "anthropic"
    ) {
      throw new Error("Invalid provider.");
    }

    const settings = getStoredSettings();
    const client = createAIClient(provider.toLowerCase() as Provider, apiKey, settings.proxy);
    const messages: Message[] = systemPrompt
      ? [
          { role: "system", content: systemPrompt },
          { role: "user", content: message },
        ]
      : [{ role: "user", content: message }];

    const response = await client.generateCompletion(messages, model);
    if (!response) {
      throw new Error("Empty response from AI provider");
    }
    return response;
  } catch (error) {
    console.error("Error in prompt function:", error);
    if (error instanceof Error) {
      return `Error: ${error.message}`;
    } else {
      return "An unexpected error occurred";
    }
  }
}

/**
 * Simplified AI function. Uses settings from the taskpane (API Key, Model, Provider).
 * @customfunction DCEXPERT
 * @param data The cell value or data to process.
 * @param prompt The prompt or instruction for the AI.
 * @returns The AI response.
 */
export async function dcexpert(data: any, prompt: string): Promise<string> {
  try {
    const settings = getStoredSettings();

    if (!settings.apiKey) {
      return "Error: Please check your Copilot settings (API Key is missing).";
    }

    if (!data && !prompt) {
      return "Error: Missing input.";
    }

    // Combine data and prompt
    const fullMessage = data ? `${data}\n\nInstruction: ${prompt}` : prompt;

    // Create client directly - decoupled from prompt()
    const client = createAIClient(settings.provider.toLowerCase() as Provider, settings.apiKey, settings.proxy);

    const messages: Message[] = settings.systemPrompt
      ? [
          { role: "system", content: settings.systemPrompt },
          { role: "user", content: fullMessage },
        ]
      : [{ role: "user", content: fullMessage }];

    const response = await client.generateCompletion(messages, settings.model);
    return response || "No response";
  } catch (error) {
    return `Error: ${error instanceof Error ? error.message : "Unknown error"}`;
  }
}

/**
 * Generates a streaming response based on the given prompt using the specified AI model and provider.
 * @customfunction PROMPT_STREAM
 * @streaming
 * @helpUrl https://dcexpert.liminity.se/help
 * @param message The prompt message to send to the AI.
 * @param model The AI model to use for generating the response.
 * @param apiKey The API key for the AI service.
 * @param systemPrompt An optional system prompt to provide context for the AI.
 * @param provider The AI provider (Nebius).
 * @param invocation The streaming invocation object
 */
export function promptStream(
  message: string,
  model: string,
  apiKey: string,
  systemPrompt: string,
  provider: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  try {
    if (!message || !model || !apiKey || !provider) {
      throw new Error("Missing required parameters");
    }

    if (
      provider.toLowerCase() !== "nebius" &&
      provider.toLowerCase() !== "openai" &&
      provider.toLowerCase() !== "anthropic"
    ) {
      throw new Error("Invalid provider.");
    }

    const settings = getStoredSettings();
    const client = createAIClient(provider.toLowerCase() as Provider, apiKey, settings.proxy);
    let fullResponse = "";

    const messages: Message[] = systemPrompt
      ? [
          { role: "system", content: systemPrompt },
          { role: "user", content: message },
        ]
      : [{ role: "user", content: message }];

    client
      .generateStreamingCompletion(messages, model, (chunk) => {
        fullResponse += chunk;
        invocation.setResult(fullResponse);
      })
      .then(() => {
        if (!fullResponse) {
          throw new Error("Empty response from AI provider");
        }
      })
      .catch((error) => {
        console.error("Error in promptStream function:", error);
        if (error instanceof Error) {
          invocation.setResult(`Error: ${error.message}`);
        } else {
          invocation.setResult("An unexpected error occurred");
        }
      });

    invocation.onCanceled = () => {
      console.log("Stream cancelled by user");
    };
  } catch (error) {
    console.error("Error in promptStream function setup:", error);
    if (error instanceof Error) {
      invocation.setResult(`Error: ${error.message}`);
    } else {
      invocation.setResult("An unexpected error occurred");
    }
  }
}

CustomFunctions.associate("PROMPT", prompt);
CustomFunctions.associate("DCEXPERT", dcexpert);
CustomFunctions.associate("PROMPT_STREAM", promptStream);

// Test function
export function hello(name: string): string {
  return `Hello ${name}!`;
}
CustomFunctions.associate("HELLO", hello);
