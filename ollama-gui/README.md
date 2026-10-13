# Ollama Chat GUI

A single-file HTML/JavaScript chat client for a local [Ollama](https://ollama.com) installation.
It has no build step and no npm dependencies. It can talk to Ollama directly, or go through
[AnythingLLM](https://anythingllm.com) or [Open WebUI](https://openwebui.com) when you want
document RAG, users, or API keys.

| File | Purpose |
|---|---|
| `index.html` | The GUI: model picker, system prompt, temperature, streaming chat, Stop button, history |
| `serve.py` | Optional zero-dependency Python server that serves the GUI and proxies `/api/*` and `/v1/*` to Ollama |
| `nginx.conf` | The same reverse-proxy setup for nginx (used by Docker Compose) |
| `docker-compose.yml` | Ollama + GUI in one command. AnythingLLM and Open WebUI are optional profiles |

---

## How the integration works

```
                         ┌───────────── Option A: direct ─────────────┐
 Browser (index.html) ───┤                                            ├──► Ollama :11434
                         └── Option B: same-origin proxy ─────────────┘
                              serve.py :8080  or  nginx :8080
                              (/api/*, /v1/* → Ollama)

 Browser (index.html) ──► AnythingLLM :3001 /api/v1/openai ──► Ollama   (Option C)
 Browser (index.html) ──► Open WebUI  :3000 /api           ──► Ollama   (Option D)
```

### 1. Ollama native REST API (default, needs no middleware)

Ollama has an HTTP API on port **11434**. The GUI uses two endpoints:

| Call | Endpoint | Used for |
|---|---|---|
| `GET /api/tags` | `{"models":[{"name":"llama3.2:latest", ...}]}` | Fill the Model dropdown |
| `POST /api/chat` | body `{"model", "messages":[{role,content}], "stream":true, "options":{"temperature"}}` | Chat |

`/api/chat` streams **newline-delimited JSON (NDJSON)**. Each line is
`{"message":{"role":"assistant","content":"<token>"},"done":false}`. The final line has
`"done":true` plus timing stats (`eval_count`, `eval_duration`), which the GUI uses to show tokens/sec.
The browser reads the stream with `fetch()` + `ReadableStream`, so text appears as it is generated.
The model has no memory between calls, so the GUI sends the whole conversation (`messages`)
each time.

### 2. OpenAI-compatible API (Ollama `/v1`, AnythingLLM, Open WebUI)

Set **Backend type → OpenAI-compatible** and the GUI uses:

| Call | Endpoint |
|---|---|
| List models | `GET {base}/models` → `{"data":[{"id": ...}]}` |
| Chat | `POST {base}/chat/completions` with `{"model","messages","stream":true,"temperature"}` |

These responses stream as **Server-Sent Events** (`data: {"choices":[{"delta":{"content":"..."}}]}`
lines ending with `data: [DONE]`). If you set an API key, it is sent as `Authorization: Bearer <key>`.

| Server | Base URL | Model value | API key |
|---|---|---|---|
| Ollama itself | `http://localhost:11434/v1` | Ollama model name | not needed |
| AnythingLLM | `http://localhost:3001/api/v1/openai` | **workspace slug** (e.g. `my-workspace`) | AnythingLLM → Settings → Developer API |
| Open WebUI | `http://localhost:3000/api` | Ollama model name | Open WebUI → Settings → Account → API keys |

### 3. CORS: why a proxy can help

Browsers block cross-origin requests unless the server allows them. Ollama allows `localhost` /
`127.0.0.1` origins by default. It **rejects** pages opened as `file://`, pages served from another
host, and pages served from a LAN IP. There are two fixes:

- **Allow the origin in Ollama:** set the `OLLAMA_ORIGINS` environment variable (see Step 2 below), or
- **Use the bundled proxy (recommended):** `serve.py` or nginx serves the page *and* forwards
  `/api/*` and `/v1/*` to Ollama. The browser then only ever talks to one origin, so CORS never
  comes up. You also don't need to expose port 11434 to your network. Both proxies turn off response
  buffering so tokens still stream.

AnythingLLM's developer API and Open WebUI both send CORS headers by default, so the GUI can call them directly.

---

## Step-by-step deployment

### Step 1: Install Ollama and pull a model

- **macOS / Windows:** download the installer from <https://ollama.com/download> and run it.
- **Linux:**
  ```bash
  curl -fsSL https://ollama.com/install.sh | sh
  ```

Then pull a model and check that the API answers:

```bash
ollama pull llama3.2          # ~2 GB; try qwen2.5, mistral, gemma3, etc.
curl http://localhost:11434/api/tags
```

### Step 2: Choose how to serve the GUI

#### Option A: Simplest (Python proxy, recommended)

Needs Python 3.8+ and nothing else.

```bash
cd ollama-gui
python3 serve.py
# → Ollama Chat GUI on http://127.0.0.1:8080  (proxying /api, /v1 -> http://localhost:11434)
```

Open <http://localhost:8080>. The Base URL field already points to the page's own address, so click
**Connect**, pick a model, and start chatting.

Environment variables: `OLLAMA_URL` (default `http://localhost:11434`), `PORT` (default `8080`),
`HOST` (default `127.0.0.1`; use `0.0.0.0` to reach the page from other machines on your LAN).

#### Option B: Open the HTML file directly (no server)

Double-click `index.html`. The page then has a `file://` origin, so you must allow it in Ollama
and restart Ollama:

| OS | How |
|---|---|
| Linux (systemd) | `sudo systemctl edit ollama` → add `[Service]` / `Environment="OLLAMA_ORIGINS=*"` → `sudo systemctl restart ollama` |
| macOS | `launchctl setenv OLLAMA_ORIGINS "*"`, then quit and reopen the Ollama app |
| Windows | Settings → System → Advanced → Environment variables → add `OLLAMA_ORIGINS` = `*`, then restart Ollama |
| Manual | `OLLAMA_ORIGINS="*" ollama serve` |

Set Base URL to `http://localhost:11434` and click **Connect**. `*` is fine on a personal machine.
On shared networks, list specific origins instead (e.g. `http://myhost:8080`).

#### Option C: Docker Compose (Ollama + GUI together)

Needs Docker with the Compose plugin.

```bash
cd ollama-gui
docker compose up -d
docker compose exec ollama ollama pull llama3.2
```

Open <http://localhost:8080>. nginx serves `index.html` and proxies `/api` and `/v1` to the `ollama`
container (see `nginx.conf`). Models persist in the `ollama` volume. For an NVIDIA GPU, uncomment
the `deploy:` block in `docker-compose.yml`.

> If Ollama is already installed on the host, either stop it first (port 11434 conflict) or remove
> the `ollama` service. In that case, point nginx at `host.docker.internal:11434` and add
> `extra_hosts: ["host.docker.internal:host-gateway"]` to the `gui` service.

### Step 3 (optional): Add AnythingLLM as middleware

Use this when you want the chat to answer from **your documents** (PDFs, docs, websites) or when you
need API-key access control.

1. Start it:
   - Docker: `docker compose --profile anythingllm up -d` (already set up to use the `ollama` container), **or**
   - Desktop app from <https://anythingllm.com/download>. In onboarding, choose **Ollama** as the LLM
     provider, base URL `http://127.0.0.1:11434`.
2. Open AnythingLLM (<http://localhost:3001> for Docker), create a **workspace**, and upload documents to it if you want RAG.
3. **Settings → Developer API → Generate New API Key.** Copy it.
4. In the GUI, set:
   - Backend type: **OpenAI-compatible**
   - Base URL: `http://localhost:3001/api/v1/openai`
   - API key: the key from step 3
5. Click **Connect**. The Model dropdown now lists your **workspace slugs**. Each one routes to Ollama
   with that workspace's documents, system prompt, and settings.

AnythingLLM also has a native endpoint, `POST /api/v1/workspace/{slug}/chat`
(`{"message":"...","mode":"chat"}`), documented at `http://localhost:3001/api/docs`. The GUI uses the
OpenAI-compatible endpoint instead so that one code path works with every backend.

### Step 4 (optional): Add Open WebUI as middleware

Use this for multi-user logins, server-side chat history, and per-user API keys.

1. `docker compose --profile openwebui up -d` → open <http://localhost:3000> and create the admin account.
2. **Settings → Account → API keys → Create new secret key.**
   (An admin may first need to enable API keys under Admin Panel → Settings.)
3. In the GUI: Backend type **OpenAI-compatible**, Base URL `http://localhost:3000/api`, paste the key, then **Connect**.

---

## Using the GUI

- **Connect** loads the models. Your settings are saved in the browser's `localStorage`.
- **Enter** sends a message, **Shift+Enter** adds a new line, and **Stop** cancels generation (the request is aborted).
- **System prompt** is sent as the first `system` message. **Temperature** controls randomness (0 = deterministic).
- **Clear chat** resets the conversation history the GUI sends with each request.
- Replies are shown as plain text. HTML and Markdown are not rendered, so model output can't inject script into the page.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `Failed to fetch` / connection failed | Ollama isn't running (`ollama serve`), or CORS is blocking the request. Use `serve.py`, or set `OLLAMA_ORIGINS` (Step 2B) |
| `HTTP 403` from Ollama | Origin not allowed. Same fix as above |
| `HTTP 404 model not found` | Run `ollama pull <model>` and click Connect again |
| `HTTP 502` from serve.py | `OLLAMA_URL` is wrong or Ollama is down |
| `HTTP 401/403` from AnythingLLM / Open WebUI | Missing or invalid API key |
| Reply appears all at once instead of streaming | A proxy is buffering. In nginx keep `proxy_buffering off` |
| Very slow first reply | The model is loading into memory. Later replies are faster |

## Security notes

- Ollama has **no authentication**. Don't expose port 11434 to the internet. Keep `serve.py` on
  `127.0.0.1` unless you trust your network, and put authentication in front of it if you need remote
  access (or use AnythingLLM / Open WebUI API keys).
- The API key field is stored in this browser's `localStorage` in plain text.
