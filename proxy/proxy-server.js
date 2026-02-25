const express = require('express');
const http = require('http');
const https = require('https');
const promClient = require('prom-client');

const app = express();
const PORT = process.env.PORT || 8080;

// === КОНФИГУРАЦИЯ БЭКЕНДА ===
const NEBIUS_API_KEY = process.env.NEBIUS_API_KEY || 'ВАШ_NEBIUS_КЛЮЧ';
const SERVICE_TOKEN = process.env.SERVICE_TOKEN || 'my-super-secret-token';
const MODEL_NAME = 'Qwen/Qwen3-Coder-480B-A35B-Instruct';

// Настройка корпоративного прокси
const PROXY_CONFIG = {
  host: process.env.UPSTREAM_PROXY_HOST || '168.90.196.95',
  port: parseInt(process.env.UPSTREAM_PROXY_PORT || '8000'),
  user: process.env.UPSTREAM_PROXY_USER || 'Xjyc9L',
  pass: process.env.UPSTREAM_PROXY_PASS || 'bEJrmk'
};

// === НАСТРОЙКА PROMETHEUS МЕТРИК ===
const collectDefaultMetrics = promClient.collectDefaultMetrics;
const Registry = promClient.Registry;
const register = new Registry();
collectDefaultMetrics({ register });

const requestCounter = new promClient.Counter({
  name: 'excel_ai_requests_total',
  help: 'Общее количество запросов от плагина Excel',
  labelNames: ['status', 'type']
});
register.registerMetric(requestCounter);

const requestDuration = new promClient.Histogram({
  name: 'excel_ai_request_duration_seconds',
  help: 'Длительность обработки запросов к ИИ',
  buckets:[0.5, 1, 2, 5, 10, 20, 30, 60]
});
register.registerMetric(requestDuration);

// Парсинг JSON
app.use(express.json({ limit: '10mb' }));

// Настройка CORS для Excel
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Эндпойнт для сбора метрик (для Prometheus/Grafana)
app.get('/metrics', async (req, res) => {
  res.setHeader('Content-Type', register.contentType);
  res.send(await register.metrics());
});

// Основной API эндпойнт для плагина
app.post('/api/chat', (req, res) => {
  const endTimer = requestDuration.startTimer();
  const isStream = !!req.body.stream;
  const reqType = isStream ? 'stream' : 'sync';

  // 1. Авторизация по сервисному токену
  const authHeader = req.headers['authorization'];
  if (!authHeader || authHeader !== `Bearer ${SERVICE_TOKEN}`) {
    requestCounter.inc({ status: '401', type: reqType });
    return res.status(401).json({ error: 'Неверный токен доступа. Обратитесь к администратору.' });
  }

  // 2. Формирование Payload для Nebius
  const payload = JSON.stringify({
    messages: req.body.messages ||[],
    model: MODEL_NAME,
    stream: isStream,
    max_tokens: 2000
  });

  const proxyAuth = Buffer.from(`${PROXY_CONFIG.user}:${PROXY_CONFIG.pass}`).toString('base64');

  // 3. Создание туннеля через корпоративный прокси
  const proxyReq = http.request({
    hostname: PROXY_CONFIG.host,
    port: PROXY_CONFIG.port,
    method: 'CONNECT',
    path: 'api.studio.nebius.ai:443',
    headers: { 'Proxy-Authorization': `Basic ${proxyAuth}` }
  });

  proxyReq.on('connect', (proxyRes, socket) => {
    if (proxyRes.statusCode !== 200) {
      console.error(`Upstream Proxy Error: ${proxyRes.statusCode}`);
      requestCounter.inc({ status: '502', type: reqType });
      res.status(502).json({ error: 'Ошибка корпоративного прокси' });
      return socket.end();
    }

    // 4. Отправка запроса в Nebius по установленному туннелю
    const apiReq = https.request({
      hostname: 'api.studio.nebius.ai',
      path: '/v1/chat/completions',
      method: 'POST',
      socket: socket,
      agent: false,
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${NEBIUS_API_KEY}`,
        'Content-Length': Buffer.byteLength(payload)
      }
    }, (apiRes) => {
      res.status(apiRes.statusCode);
      // Проброс заголовков (важно для Server-Sent Events / stream)
      Object.keys(apiRes.headers).forEach(k => res.setHeader(k, apiRes.headers[k]));
      
      requestCounter.inc({ status: apiRes.statusCode.toString(), type: reqType });
      
      apiRes.pipe(res);
      apiRes.on('end', () => endTimer());
    });

    apiReq.on('error', (e) => {
      console.error('API Error:', e.message);
      requestCounter.inc({ status: '500', type: reqType });
      if (!res.headersSent) res.status(500).json({ error: e.message });
      endTimer();
    });

    apiReq.write(payload);
    apiReq.end();
  });

  proxyReq.on('error', (e) => {
    console.error('Proxy Connect Error:', e.message);
    requestCounter.inc({ status: '502', type: reqType });
    if (!res.headersSent) res.status(502).json({ error: e.message });
    endTimer();
  });
  
  proxyReq.end();
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`[Excel Copilot Proxy] Запущен на порту ${PORT}`);
  console.log(`[Metrics] Доступны по адресу http://localhost:${PORT}/metrics`);
});