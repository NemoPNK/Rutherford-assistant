const http = require("http");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { URL } = require("url");

const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const DATA_DIR = path.join(ROOT_DIR, "data");
const PORT = Number(process.env.PORT || 4321);
const PASSWORD = process.env.RUTHERFORD_PASSWORD || "rutherford9498rj";
const SESSION_COOKIE = "rutherford_session";
const SESSION_TTL_MS = 12 * 60 * 60 * 1000;
const MAX_LOG_LINES = 800;

const sessions = new Map();
const jobs = new Map();

const TASKS = {
  setup: {
    id: "setup",
    name: "Setup Rutherford",
    script: "LaRoche.ps1",
    includedPaths: ["LaRoche.ps1", "wallpaper.jpg", "preinstall"]
  },
  network: {
    id: "network",
    name: "Network Rutherford",
    script: "Network.ps1",
    includedPaths: ["Network.ps1"]
  }
};

const CONTENT_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".cmd": "text/plain; charset=utf-8",
  ".ps1": "text/plain; charset=utf-8",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png"
};

function safeJoin(root, relativePath) {
  const resolved = path.resolve(root, relativePath);
  const normalizedRoot = `${root}${path.sep}`;
  if (resolved !== root && !resolved.startsWith(normalizedRoot)) {
    throw new Error(`Forbidden path: ${relativePath}`);
  }
  return resolved;
}

function pruneSessions() {
  const now = Date.now();
  for (const [sessionId, session] of sessions.entries()) {
    if (session.expiresAt <= now) {
      sessions.delete(sessionId);
    }
  }
}

function createSession() {
  const sessionId = crypto.randomBytes(24).toString("hex");
  sessions.set(sessionId, {
    id: sessionId,
    createdAt: Date.now(),
    expiresAt: Date.now() + SESSION_TTL_MS
  });
  return sessionId;
}

function getSession(req) {
  pruneSessions();
  const cookies = parseCookies(req.headers.cookie || "");
  const sessionId = cookies[SESSION_COOKIE];
  if (!sessionId) {
    return null;
  }

  const session = sessions.get(sessionId);
  if (!session) {
    return null;
  }

  session.expiresAt = Date.now() + SESSION_TTL_MS;
  return session;
}

function deleteSession(req) {
  const cookies = parseCookies(req.headers.cookie || "");
  const sessionId = cookies[SESSION_COOKIE];
  if (sessionId) {
    sessions.delete(sessionId);
  }
}

function parseCookies(cookieHeader) {
  const result = {};
  for (const part of cookieHeader.split(";")) {
    const [rawName, ...rawValue] = part.trim().split("=");
    if (!rawName) {
      continue;
    }

    result[rawName] = decodeURIComponent(rawValue.join("=") || "");
  }
  return result;
}

function json(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function text(res, statusCode, body, contentType = "text/plain; charset=utf-8") {
  res.writeHead(statusCode, {
    "Content-Type": contentType,
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function notFound(res) {
  text(res, 404, "Not Found");
}

function unauthorized(res) {
  json(res, 401, { error: "Unauthorized" });
}

function badRequest(res, message) {
  json(res, 400, { error: message });
}

async function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;

    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > 1024 * 1024) {
        reject(new Error("Body too large"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

async function readJsonBody(req) {
  const raw = await readBody(req);
  if (!raw) {
    return {};
  }
  return JSON.parse(raw);
}

function timingSafePasswordMatch(candidate) {
  const actualHash = crypto.createHash("sha256").update(PASSWORD).digest();
  const candidateHash = crypto.createHash("sha256").update(candidate || "").digest();
  return crypto.timingSafeEqual(actualHash, candidateHash);
}

function generateId(prefix) {
  return `${prefix}_${crypto.randomBytes(6).toString("hex")}`;
}

async function collectRelativeFiles(relativePath) {
  const absolutePath = safeJoin(ROOT_DIR, relativePath);
  const stats = await fsp.stat(absolutePath);

  if (stats.isFile()) {
    return [relativePath];
  }

  const files = [];

  async function walk(currentAbsolute, currentRelative) {
    const entries = await fsp.readdir(currentAbsolute, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.name === ".DS_Store") {
        continue;
      }

      const entryAbsolute = path.join(currentAbsolute, entry.name);
      const entryRelative = path.posix.join(currentRelative.split(path.sep).join("/"), entry.name);

      if (entry.isDirectory()) {
        await walk(entryAbsolute, entryRelative);
      } else if (entry.isFile()) {
        files.push(entryRelative);
      }
    }
  }

  await walk(absolutePath, relativePath.split(path.sep).join("/"));
  return files;
}

async function createJob(taskId) {
  const task = TASKS[taskId];
  if (!task) {
    throw new Error(`Unknown task: ${taskId}`);
  }

  const files = [];
  for (const relativePath of task.includedPaths) {
    const collected = await collectRelativeFiles(relativePath);
    files.push(...collected);
  }

  const job = {
    id: generateId("job"),
    token: crypto.randomBytes(24).toString("hex"),
    taskId: task.id,
    taskName: task.name,
    entryScript: task.script,
    files: [...new Set(files)].sort(),
    status: "pending",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    logs: [],
    meta: {}
  };

  jobs.set(job.id, job);
  return job;
}

function publicJob(job) {
  return {
    id: job.id,
    taskId: job.taskId,
    taskName: job.taskName,
    entryScript: job.entryScript,
    status: job.status,
    createdAt: job.createdAt,
    updatedAt: job.updatedAt,
    logCount: job.logs.length,
    logs: job.logs,
    meta: job.meta,
    report: buildJobReportSummary(job)
  };
}

function updateJobStatus(job, status, message, meta) {
  const previousStatus = job.status;
  job.status = status;
  job.updatedAt = new Date().toISOString();

  if (meta && typeof meta === "object") {
    job.meta = { ...job.meta, ...meta };
  }

  if (!job.meta.startedAt && ["downloading", "running", "completed", "error"].includes(status)) {
    job.meta.startedAt = job.updatedAt;
  }

  if (status === "running" && previousStatus !== "running" && !job.meta.startedAt) {
    job.meta.startedAt = job.updatedAt;
  }

  if (["completed", "error"].includes(status)) {
    job.meta.finishedAt = job.updatedAt;
  }

  if (message) {
    job.meta.lastMessage = message;
    appendJobLogs(job, [message]);
  }
}

function appendJobLogs(job, lines) {
  if (!Array.isArray(lines) || lines.length === 0) {
    return;
  }

  const timestamp = new Date().toISOString();
  for (const line of lines) {
    if (typeof line !== "string" || !line.trim()) {
      continue;
    }

    job.logs.push({
      timestamp,
      line: line.replace(/\r/g, "")
    });
  }

  if (job.logs.length > MAX_LOG_LINES) {
    job.logs = job.logs.slice(job.logs.length - MAX_LOG_LINES);
  }

  job.updatedAt = new Date().toISOString();
}

function getServerOrigins(req) {
  const hostHeader = req.headers.host || `localhost:${PORT}`;
  const originFromRequest = `http://${hostHeader}`;
  const networkInterfaces = os.networkInterfaces();
  const origins = new Set([originFromRequest]);

  for (const addresses of Object.values(networkInterfaces)) {
    for (const address of addresses || []) {
      if (address.family === "IPv4" && !address.internal) {
        origins.add(`http://${address.address}:${PORT}`);
      }
    }
  }

  return [...origins];
}

function buildJobReportSummary(job) {
  const startedAt = job.meta.startedAt || null;
  const finishedAt = job.meta.finishedAt || null;
  let durationSeconds = null;

  if (startedAt && finishedAt) {
    durationSeconds = Math.max(0, Math.round((Date.parse(finishedAt) - Date.parse(startedAt)) / 1000));
  }

  return {
    available: ["completed", "error"].includes(job.status),
    startedAt,
    finishedAt,
    durationSeconds,
    computerName: job.meta.computerName || "",
    userName: job.meta.userName || "",
    lastMessage: job.meta.lastMessage || "",
    logCount: job.logs.length
  };
}

function buildJobReportText(job, req) {
  const report = buildJobReportSummary(job);
  const lines = [
    "Rutherford Assistant Report",
    `Generated: ${new Date().toISOString()}`,
    `Server: http://${req.headers.host}`,
    "",
    `Job ID: ${job.id}`,
    `Task: ${job.taskName}`,
    `Script: ${job.entryScript}`,
    `Status: ${job.status}`,
    `Created At: ${job.createdAt}`,
    `Started At: ${report.startedAt || "N/A"}`,
    `Finished At: ${report.finishedAt || "N/A"}`,
    `Duration (s): ${report.durationSeconds ?? "N/A"}`,
    `Computer Name: ${report.computerName || "N/A"}`,
    `User Name: ${report.userName || "N/A"}`,
    `Last Message: ${report.lastMessage || "N/A"}`,
    `Log Count: ${job.logs.length}`,
    "",
    "Logs:",
    ...job.logs.map((entry) => `[${entry.timestamp}] ${entry.line}`)
  ];

  return `${lines.join("\n")}\n`;
}

function buildLauncherContent(req, job) {
  const baseUrl = `http://${req.headers.host}`;
  const runnerUrl = `${baseUrl}/agent/runner.ps1?jobId=${encodeURIComponent(job.id)}&token=${encodeURIComponent(job.token)}`;
  const runnerTarget = `%TEMP%\\Rutherford-${job.id}.ps1`;

  return `@echo off
setlocal
cd /d "%~dp0"

if /I "%~1"=="elevated" goto elevated
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList 'elevated' -Verb RunAs"
exit /b

:elevated
echo Rutherford Assistant - ${job.taskName}
echo Downloading secure runner from ${baseUrl}
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -UseBasicParsing -Uri '${runnerUrl}' -OutFile '${runnerTarget}'"
if errorlevel 1 (
  echo Failed to download the runner.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "${runnerTarget}"
set EXIT_CODE=%ERRORLEVEL%

echo.
echo Task finished with exit code %EXIT_CODE%.
pause
exit /b %EXIT_CODE%
`;
}

function buildRunnerScript(req, job) {
  const baseUrl = `http://${req.headers.host}`;

  return `$ErrorActionPreference = "Stop"

$jobId = "${job.id}"
$jobToken = "${job.token}"
$baseUrl = "${baseUrl}"
$workDir = Join-Path $env:TEMP ("RutherfordAssistant\\" + $jobId)

function Send-JsonPost {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)]$Payload
    )

    try {
        $json = $Payload | ConvertTo-Json -Depth 6 -Compress
        Invoke-RestMethod -Method Post -Uri $Url -ContentType "application/json" -Body $json | Out-Null
    }
    catch {
        Write-Host "Unable to send data to server: $($_.Exception.Message)"
    }
}

function Send-Status {
    param(
        [Parameter(Mandatory=$true)][string]$Status,
        [string]$Message = "",
        [hashtable]$Meta = @{}
    )

    Send-JsonPost -Url ($baseUrl + "/api/jobs/" + $jobId + "/status?token=" + $jobToken) -Payload @{
        status = $Status
        message = $Message
        meta = $Meta
    }
}

function Send-Logs {
    param([string[]]$Lines)

    if (-not $Lines -or $Lines.Count -eq 0) {
        return
    }

    Send-JsonPost -Url ($baseUrl + "/api/jobs/" + $jobId + "/log?token=" + $jobToken) -Payload @{
        lines = $Lines
    }
}

try {
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null

    $meta = @{
        computerName = $env:COMPUTERNAME
        userName = $env:USERNAME
    }
    Send-Status -Status "downloading" -Message "Preparing local workspace." -Meta $meta

    $manifest = Invoke-RestMethod -UseBasicParsing -Uri ($baseUrl + "/api/jobs/" + $jobId + "/manifest?token=" + $jobToken)

    foreach ($file in $manifest.files) {
        $targetPath = Join-Path $workDir $file.path
        $targetDir = Split-Path -Path $targetPath -Parent

        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        $encodedPath = [System.Uri]::EscapeDataString($file.path)
        $downloadUrl = $baseUrl + "/api/jobs/" + $jobId + "/file?token=" + $jobToken + "&path=" + $encodedPath
        Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $targetPath
    }

    $scriptPath = Join-Path $workDir $manifest.entryScript

    if (-not (Test-Path $scriptPath)) {
        throw "Entry script not found: $scriptPath"
    }

    Send-Status -Status "running" -Message ("Running " + $manifest.entryScript)

    $buffer = New-Object System.Collections.Generic.List[string]
    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath *>&1 | ForEach-Object {
        $line = $_.ToString()
        Write-Host $line
        $buffer.Add($line)

        if ($buffer.Count -ge 20) {
            Send-Logs -Lines $buffer.ToArray()
            $buffer.Clear()
        }
    }

    if ($buffer.Count -gt 0) {
        Send-Logs -Lines $buffer.ToArray()
        $buffer.Clear()
    }

    if ($LASTEXITCODE -ne 0) {
        throw "Script exited with code $LASTEXITCODE"
    }

    Send-Status -Status "completed" -Message "Task completed successfully."
    exit 0
}
catch {
    $message = $_.Exception.Message
    Write-Host "ERROR: $message"
    Send-Logs -Lines @("ERROR: $message")
    Send-Status -Status "error" -Message $message
    exit 1
}
`;
}

async function serveStatic(req, res, pathname) {
  const relativePath = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
  const absolutePath = safeJoin(PUBLIC_DIR, relativePath);

  let stats;
  try {
    stats = await fsp.stat(absolutePath);
  } catch {
    notFound(res);
    return;
  }

  if (!stats.isFile()) {
    notFound(res);
    return;
  }

  const ext = path.extname(absolutePath).toLowerCase();
  res.writeHead(200, {
    "Content-Type": CONTENT_TYPES[ext] || "application/octet-stream",
    "Content-Length": stats.size,
    "Cache-Control": "no-store"
  });
  fs.createReadStream(absolutePath).pipe(res);
}

async function handleRequest(req, res) {
  try {
    const url = new URL(req.url, `http://${req.headers.host || `localhost:${PORT}`}`);
    const pathname = url.pathname;

    if (req.method === "GET" && (pathname === "/" || pathname.startsWith("/app.js") || pathname.startsWith("/styles.css"))) {
      await serveStatic(req, res, pathname);
      return;
    }

    if (req.method === "GET" && pathname === "/api/session") {
      const session = getSession(req);
      json(res, 200, {
        authenticated: Boolean(session),
        port: PORT,
        tasks: Object.values(TASKS).map((task) => ({
          id: task.id,
          name: task.name,
          script: task.script
        })),
        accessUrls: getServerOrigins(req)
      });
      return;
    }

    if (req.method === "POST" && pathname === "/api/login") {
      const body = await readJsonBody(req);
      if (!timingSafePasswordMatch(body.password || "")) {
        unauthorized(res);
        return;
      }

      const sessionId = createSession();
      res.setHeader("Set-Cookie", `${SESSION_COOKIE}=${encodeURIComponent(sessionId)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${Math.floor(SESSION_TTL_MS / 1000)}`);
      json(res, 200, { ok: true });
      return;
    }

    if (req.method === "POST" && pathname === "/api/logout") {
      deleteSession(req);
      res.setHeader("Set-Cookie", `${SESSION_COOKIE}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`);
      json(res, 200, { ok: true });
      return;
    }

    if (!getSession(req) && pathname.startsWith("/api/") && !pathname.includes("/manifest") && !pathname.includes("/file") && !pathname.includes("/log") && !pathname.includes("/status")) {
      unauthorized(res);
      return;
    }

    if (req.method === "GET" && pathname === "/api/jobs") {
      const list = [...jobs.values()]
        .sort((left, right) => right.createdAt.localeCompare(left.createdAt))
        .map(publicJob);

      json(res, 200, { jobs: list });
      return;
    }

    if (req.method === "POST" && pathname === "/api/jobs") {
      const body = await readJsonBody(req);
      if (!body.type || !TASKS[body.type]) {
        badRequest(res, "Invalid task type.");
        return;
      }

      const job = await createJob(body.type);
      json(res, 201, {
        job: publicJob(job),
        downloadUrl: `/downloads/${job.id}/launcher.cmd`
      });
      return;
    }

    const launcherMatch = pathname.match(/^\/downloads\/([^/]+)\/launcher\.cmd$/);
    if (req.method === "GET" && launcherMatch) {
      if (!getSession(req)) {
        unauthorized(res);
        return;
      }

      const job = jobs.get(launcherMatch[1]);
      if (!job) {
        notFound(res);
        return;
      }

      const launcherBody = buildLauncherContent(req, job);
      res.writeHead(200, {
        "Content-Type": "text/plain; charset=utf-8",
        "Content-Length": Buffer.byteLength(launcherBody),
        "Content-Disposition": `attachment; filename="Rutherford-${job.taskId}-${job.id}.cmd"`,
        "Cache-Control": "no-store"
      });
      res.end(launcherBody);
      return;
    }

    const reportMatch = pathname.match(/^\/api\/jobs\/([^/]+)\/report$/);
    if (req.method === "GET" && reportMatch) {
      if (!getSession(req)) {
        unauthorized(res);
        return;
      }

      const job = jobs.get(reportMatch[1]);
      if (!job) {
        notFound(res);
        return;
      }

      const reportBody = buildJobReportText(job, req);
      res.writeHead(200, {
        "Content-Type": "text/plain; charset=utf-8",
        "Content-Length": Buffer.byteLength(reportBody),
        "Content-Disposition": `attachment; filename="Rutherford-report-${job.id}.txt"`,
        "Cache-Control": "no-store"
      });
      res.end(reportBody);
      return;
    }

    if (req.method === "GET" && pathname === "/agent/runner.ps1") {
      const jobId = url.searchParams.get("jobId");
      const token = url.searchParams.get("token");
      const job = jobs.get(jobId);

      if (!job || token !== job.token) {
        unauthorized(res);
        return;
      }

      text(res, 200, buildRunnerScript(req, job), "text/plain; charset=utf-8");
      return;
    }

    const manifestMatch = pathname.match(/^\/api\/jobs\/([^/]+)\/manifest$/);
    if (req.method === "GET" && manifestMatch) {
      const job = jobs.get(manifestMatch[1]);
      const token = url.searchParams.get("token");
      if (!job || token !== job.token) {
        unauthorized(res);
        return;
      }

      json(res, 200, {
        id: job.id,
        taskId: job.taskId,
        taskName: job.taskName,
        entryScript: job.entryScript,
        files: job.files.map((filePath) => ({ path: filePath }))
      });
      return;
    }

    const fileMatch = pathname.match(/^\/api\/jobs\/([^/]+)\/file$/);
    if (req.method === "GET" && fileMatch) {
      const job = jobs.get(fileMatch[1]);
      const token = url.searchParams.get("token");
      const filePath = url.searchParams.get("path");

      if (!job || token !== job.token || !filePath) {
        unauthorized(res);
        return;
      }

      if (!job.files.includes(filePath)) {
        badRequest(res, "Unknown file path.");
        return;
      }

      const absolutePath = safeJoin(ROOT_DIR, filePath);
      const stats = await fsp.stat(absolutePath);
      const ext = path.extname(absolutePath).toLowerCase();

      res.writeHead(200, {
        "Content-Type": CONTENT_TYPES[ext] || "application/octet-stream",
        "Content-Length": stats.size,
        "Cache-Control": "no-store"
      });
      fs.createReadStream(absolutePath).pipe(res);
      return;
    }

    const logMatch = pathname.match(/^\/api\/jobs\/([^/]+)\/log$/);
    if (req.method === "POST" && logMatch) {
      const job = jobs.get(logMatch[1]);
      const token = url.searchParams.get("token");
      if (!job || token !== job.token) {
        unauthorized(res);
        return;
      }

      const body = await readJsonBody(req);
      appendJobLogs(job, Array.isArray(body.lines) ? body.lines : []);
      json(res, 200, { ok: true });
      return;
    }

    const statusMatch = pathname.match(/^\/api\/jobs\/([^/]+)\/status$/);
    if (req.method === "POST" && statusMatch) {
      const job = jobs.get(statusMatch[1]);
      const token = url.searchParams.get("token");
      if (!job || token !== job.token) {
        unauthorized(res);
        return;
      }

      const body = await readJsonBody(req);
      updateJobStatus(job, body.status || job.status, body.message || "", body.meta || {});
      json(res, 200, { ok: true });
      return;
    }

    notFound(res);
  } catch (error) {
    json(res, 500, {
      error: "Internal server error",
      details: error.message
    });
  }
}

async function ensureDataDir() {
  await fsp.mkdir(DATA_DIR, { recursive: true });
}

async function main() {
  await ensureDataDir();

  const server = http.createServer(handleRequest);
  server.listen(PORT, "0.0.0.0", () => {
    const origins = getServerOrigins({ headers: { host: `localhost:${PORT}` } });
    console.log("Rutherford Assistant server ready.");
    for (const origin of origins) {
      console.log(`- ${origin}`);
    }
  });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
