# DC expert

DC expert is a free, open-source Excel add-in that allows you to use GPT and Anthropic AI models directly within Excel spreadsheets.

## Proxy Support

DC expert now supports proxy configurations for accessing AI APIs through corporate firewalls or restricted networks.

### Integrated Proxy Server

We've integrated a local proxy server that uses your specific proxy `168.90.196.95:8000:Xjyc9L:bEJrmk`:

#### Quick Start:
```bash
./start-proxy.sh
```

Or use npm scripts:
```bash
npm run proxy-start
```

#### Manual Setup:
1. Navigate to the proxy directory: `cd proxy`
2. Install dependencies: `npm install`
3. Start the server: `npm start`

#### Proxy Endpoints:
- **Health Check**: `http://localhost:8080/health`
- **OpenAI**: `http://localhost:8080/proxy/openai`
- **Nebius**: `http://localhost:8080/proxy/nebius`
- **Anthropic**: `http://localhost:8080/proxy/anthropic`

### Using in Excel Copilot

In the Excel Copilot settings, use these proxy URLs:
- **Nebius**: `http://localhost:8080/proxy/nebius`
- **OpenAI**: `http://localhost:8080/proxy/openai`
- **Anthropic**: `http://localhost:8080/proxy/anthropic`

The proxy server will automatically route requests through `168.90.196.95:8000:Xjyc9L:bEJrmk`.

## Features

- Integrate OpenAI, Anthropic, and Nebius AI models into your Excel workflows
- Use custom functions to generate AI responses based on cell inputs
- Support for both standard and streaming completions
- Proxy support for corporate networks
- Easy-to-use interface with clear documentation

## How to Use

### PROMPT Function

Use the PROMPT function for standard completions:

```excel
=PROMPT(message, model, apiKey, systemPrompt, provider)
```

### PROMPT_STREAM Function

For streaming responses, use the PROMPT_STREAM function:

```excel
=PROMPT_STREAM(message, model, apiKey, systemPrompt, provider)
```

### Parameters

- `message`: The prompt message to send to the AI.
- `model`: The AI model to use (see provider documentation for available models).
- `apiKey`: Your API key for the chosen provider.
- `systemPrompt`: A system prompt to provide context for the AI.
- `provider`: The AI provider to use ("openai" or "anthropic").

## Privacy

Your queries are sent directly to OpenAI/Anthropic servers. No data is stored by Liminity AB.

## Documentation

For more detailed information on available models:

- [OpenAI Models](https://platform.openai.com/docs/models)
- [Anthropic Models](https://docs.anthropic.com/en/docs/about-claude/models)

## Contributing

We welcome contributions to DC expert! As an open-source project, we appreciate any help, from bug reports to feature additions. Here's how you can contribute:

1. Fork the repository
2. Create a new branch for your feature or bug fix
3. Make your changes and commit them with a clear message
4. Push your changes to your fork
5. Create a pull request to the main repository

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE.md) file for details.

## About

DC expert is developed and maintained by [Liminity AB](https://liminity.se).
