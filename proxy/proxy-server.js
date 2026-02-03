// File: proxy/proxy-server.js
const express = require("express");
const http = require("http");
const https = require("https");

const app = express();
// Важно: Cloud провайдеры (Render/Heroku) сами задают PORT в env
const PORT = process.env.PORT || 8080;

app.use(express.json());

// CORS настройки - разрешаем доступ всем (можно ограничить доменом надстройки)
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept, Authorization, x-api-key, x-target-url, anthropic-version"
  );
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

// Конфигурация Upstream Proxy (Ваш 168.90...)
// Данные будут браться из Environment Variables сервера
const PROXY_CONFIG = {
  host: process.env.UPSTREAM_PROXY_HOST || "168.90.196.95",
  port: parseInt(process.env.UPSTREAM_PROXY_PORT || "8000"),
  user: process.env.UPSTREAM_PROXY_USER || "Xjyc9L", // Замените на реальные если тестируете локально
  pass: process.env.UPSTREAM_PROXY_PASS || "bEJrmk", // Замените на реальные если тестируете локально
};

app.get("/health", (req, res) => {
  res.json({ status: "ok", service: "Excel AI Proxy" });
});

app.all("/proxy/*", async (req, res) => {
  try {
    const targetUrl = req.headers["x-target-url"];

    if (!targetUrl) {
      return res.status(400).json({ error: "Header x-target-url is required" });
    }

    console.log(`[Proxy] Request to: ${targetUrl}`);

    const url = new URL(targetUrl);
    const proxyAuth = Buffer.from(`${PROXY_CONFIG.user}:${PROXY_CONFIG.pass}`).toString("base64");

    // Опции для подключения к ВАШЕМУ прокси (CONNECT метод)
    const proxyRequestOptions = {
      hostname: PROXY_CONFIG.host,
      port: PROXY_CONFIG.port,
      method: "CONNECT",
      path: `${url.hostname}:${url.port || 443}`,
      headers: {
        "Proxy-Authorization": `Basic ${proxyAuth}`,
        "User-Agent": "Node.js Backend",
      },
    };

    // Создаем туннель
    const proxyReq = http.request(proxyRequestOptions);

    proxyReq.on("connect", (proxyRes, socket) => {
      if (proxyRes.statusCode !== 200) {
        console.error(`Upstream proxy error: ${proxyRes.statusCode}`);
        socket.end();
        return res.status(502).json({ error: "Upstream proxy connection failed" });
      }

      // Туннель создан, делаем запрос к AI провайдеру
      const clientRequest = https.request(
        {
          host: url.hostname,
          port: url.port || 443,
          path: url.pathname + url.search,
          method: req.method,
          headers: {
            ...req.headers,
            host: url.hostname, // Важно перезаписать host
            connection: "close", // Избегаем зависания сокетов
          },
          socket: socket, // Используем сокет туннеля
          agent: false, // Отключаем встроенный агент Node.js
        },
        (apiRes) => {
          // Проксируем ответ обратно клиенту (Excel)
          res.status(apiRes.statusCode);
          Object.keys(apiRes.headers).forEach((key) => {
            res.setHeader(key, apiRes.headers[key]);
          });
          apiRes.pipe(res);
        }
      );

      clientRequest.on("error", (err) => {
        console.error("API Request Error:", err);
        if (!res.headersSent) res.status(500).json({ error: "API request failed" });
      });

      // Передаем тело запроса (JSON от Excel) в API
      if (req.body) {
        // Express уже распарсил body, нужно превратить обратно в string для отправки
        // Или использовать req.pipe(clientRequest) если убрать express.json() middleware
        // Для надежности пересобираем JSON:
        clientRequest.write(JSON.stringify(req.body));
      }
      clientRequest.end();
    });

    proxyReq.on("error", (err) => {
      console.error("Proxy Connection Error:", err);
      res.status(502).json({ error: "Cannot connect to upstream proxy" });
    });

    proxyReq.end();
  } catch (error) {
    console.error("Server Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Backend running on port ${PORT}`);
});
