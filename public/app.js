const state = {
  authenticated: false,
  accessUrls: [],
  jobs: [],
  busy: false,
  error: "",
  flash: "",
  pollingId: null
};

const app = document.querySelector("#app");

function statusLabel(status) {
  const labels = {
    pending: "En attente",
    downloading: "Telechargement",
    running: "En cours",
    completed: "Termine",
    error: "Erreur"
  };

  return labels[status] || status;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function request(url, options = {}) {
  const response = await fetch(url, {
    credentials: "same-origin",
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {})
    },
    ...options
  });

  const contentType = response.headers.get("content-type") || "";
  const data = contentType.includes("application/json") ? await response.json() : await response.text();

  if (!response.ok) {
    const message = typeof data === "string" ? data : data.error || "Request failed";
    throw new Error(message);
  }

  return data;
}

async function loadSession() {
  const data = await request("/api/session", { method: "GET" });
  state.authenticated = Boolean(data.authenticated);
  state.accessUrls = data.accessUrls || [];
}

async function loadJobs() {
  if (!state.authenticated) {
    state.jobs = [];
    return;
  }

  const data = await request("/api/jobs", { method: "GET" });
  state.jobs = data.jobs || [];
}

function startPolling() {
  stopPolling();
  if (!state.authenticated) {
    return;
  }

  state.pollingId = window.setInterval(async () => {
    try {
      await loadJobs();
      render();
    } catch (error) {
      state.error = error.message;
      render();
    }
  }, 2000);
}

function stopPolling() {
  if (state.pollingId) {
    window.clearInterval(state.pollingId);
    state.pollingId = null;
  }
}

function renderLogin() {
  app.innerHTML = `
    <section class="dashboard">
      <aside class="panel">
        <p class="panel-label">Connexion</p>
        <h2>Acces protege</h2>
        <p>Entre le mot de passe operateur pour afficher les actions disponibles sur ce serveur local.</p>
        <form class="login-form" id="login-form">
          <label>
            Mot de passe
            <input type="password" name="password" autocomplete="current-password" required />
          </label>
          <button class="primary" type="submit">Se connecter</button>
        </form>
        ${state.error ? `<p class="error-text">${escapeHtml(state.error)}</p>` : ""}
      </aside>

      <section class="panel stack">
        <p class="panel-label">Acces reseau</p>
        <h2>Adresses a ouvrir depuis le PC Windows</h2>
        <div class="url-list">
          ${state.accessUrls.map((url) => `<code>${escapeHtml(url)}</code>`).join("")}
        </div>
        <div class="notice">
          L'operateur ouvre une de ces adresses dans son navigateur Windows, puis telecharge le lanceur correspondant a l'action voulue.
        </div>
      </section>
    </section>
  `;

  document.querySelector("#login-form")?.addEventListener("submit", onLogin);
}

function renderDashboard() {
  const jobMarkup = state.jobs.length
    ? state.jobs
        .map((job) => {
          const logs = job.logs.length
            ? job.logs.map((entry) => `[${entry.timestamp}] ${entry.line}`).join("\n")
            : "Aucun log pour l'instant.";
          const report = job.report || {};
          const reportMarkup = report.available
            ? `
              <div class="report-box">
                <h4>Rapport final</h4>
                <div class="job-meta">
                  <span><strong>Machine:</strong> ${escapeHtml(report.computerName || "N/A")}</span>
                  <span><strong>Utilisateur:</strong> ${escapeHtml(report.userName || "N/A")}</span>
                  <span><strong>Debut:</strong> ${escapeHtml(report.startedAt || "N/A")}</span>
                  <span><strong>Fin:</strong> ${escapeHtml(report.finishedAt || "N/A")}</span>
                  <span><strong>Duree:</strong> ${escapeHtml(report.durationSeconds != null ? `${report.durationSeconds} sec` : "N/A")}</span>
                  <span><strong>Dernier message:</strong> ${escapeHtml(report.lastMessage || "N/A")}</span>
                </div>
              </div>
            `
            : "";
          const actionsMarkup = report.available
            ? `
              <div class="job-actions">
                <a class="secondary" href="/api/jobs/${encodeURIComponent(job.id)}/report">Telecharger le rapport</a>
              </div>
            `
            : "";

          return `
            <article class="job-card">
              <header>
                <div>
                  <p class="panel-label">${escapeHtml(job.taskName)}</p>
                  <h3>${escapeHtml(job.entryScript)}</h3>
                </div>
                <span class="status-chip status-${escapeHtml(job.status)}">${escapeHtml(statusLabel(job.status))}</span>
              </header>
              <div class="job-meta">
                <span><strong>Job:</strong> ${escapeHtml(job.id)}</span>
                <span><strong>Cree:</strong> ${escapeHtml(job.createdAt)}</span>
                <span><strong>Maj:</strong> ${escapeHtml(job.updatedAt)}</span>
                <span><strong>Machine:</strong> ${escapeHtml(job.meta.computerName || "En attente")}</span>
                <span><strong>Utilisateur:</strong> ${escapeHtml(job.meta.userName || "En attente")}</span>
              </div>
              ${actionsMarkup}
              ${reportMarkup}
              <div class="logs">${escapeHtml(logs)}</div>
            </article>
          `;
        })
        .join("")
    : `<div class="empty">Aucun job pour le moment. Lance une action pour voir apparaitre les statuts et les logs ici.</div>`;

  app.innerHTML = `
    <section class="dashboard">
      <aside class="panel stack">
        <p class="panel-label">Actions</p>
        <h2>Piloter la preparation</h2>
        <div class="actions">
          <div class="action-card">
            <h3>Run Setup</h3>
            <p>Telecharge un lanceur Windows admin qui execute <code>LaRoche.ps1</code> avec les fichiers necessaires.</p>
            <button class="primary" data-task="setup" ${state.busy ? "disabled" : ""}>Telecharger le lanceur Setup</button>
          </div>
          <div class="action-card">
            <h3>Run Network</h3>
            <p>Telecharge un lanceur Windows admin qui execute <code>Network.ps1</code>.</p>
            <button class="secondary" data-task="network" ${state.busy ? "disabled" : ""}>Telecharger le lanceur Network</button>
          </div>
          <div class="action-card">
            <h3>Run Updates</h3>
            <p>Reserve pour la future integration de <code>update.ps1</code>.</p>
            <button class="ghost" type="button" disabled>Bientot disponible</button>
          </div>
        </div>
        ${state.flash ? `<div class="notice">${escapeHtml(state.flash)}</div>` : ""}
        ${state.error ? `<p class="error-text">${escapeHtml(state.error)}</p>` : ""}
      </aside>

      <section class="panel stack">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap;">
          <div>
            <p class="panel-label">Suivi</p>
            <h2>Jobs et logs</h2>
          </div>
          <button class="ghost" id="logout-button" type="button">Se deconnecter</button>
        </div>
        <p class="hint">
          Apres le clic, le navigateur telecharge un fichier <code>.cmd</code>. L'operateur doit l'ouvrir sur le PC Windows et accepter l'elevation administrateur.
        </p>
        <div class="meta-list">
          ${state.accessUrls.map((url) => `<code>${escapeHtml(url)}</code>`).join("")}
        </div>
        <div class="jobs">${jobMarkup}</div>
      </section>
    </section>
  `;

  document.querySelectorAll("[data-task]").forEach((button) => {
    button.addEventListener("click", onCreateJob);
  });
  document.querySelector("#logout-button")?.addEventListener("click", onLogout);
}

function render() {
  if (!state.authenticated) {
    renderLogin();
    return;
  }

  renderDashboard();
}

async function onLogin(event) {
  event.preventDefault();
  state.error = "";

  const form = event.currentTarget;
  const password = new FormData(form).get("password");

  try {
    await request("/api/login", {
      method: "POST",
      body: JSON.stringify({ password })
    });
    state.authenticated = true;
    await loadJobs();
    startPolling();
    render();
  } catch (error) {
    state.error = "Mot de passe invalide.";
    render();
  }
}

async function onLogout() {
  try {
    await request("/api/logout", { method: "POST", body: "{}" });
  } catch (error) {
    console.error(error);
  }

  stopPolling();
  state.authenticated = false;
  state.jobs = [];
  state.flash = "";
  state.error = "";
  render();
}

async function onCreateJob(event) {
  const task = event.currentTarget.dataset.task;
  state.busy = true;
  state.error = "";
  state.flash = "";
  render();

  try {
    const data = await request("/api/jobs", {
      method: "POST",
      body: JSON.stringify({ type: task })
    });

    state.flash = "Le lanceur est pret. Ouvre le fichier telecharge sur le PC Windows pour lancer l'action et surveiller les logs ici.";
    await loadJobs();
    render();
    const link = document.createElement("a");
    link.href = data.downloadUrl;
    link.download = "";
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    state.error = error.message;
    render();
  } finally {
    state.busy = false;
    render();
  }
}

async function bootstrap() {
  try {
    await loadSession();
    if (state.authenticated) {
      await loadJobs();
      startPolling();
    }
  } catch (error) {
    state.error = error.message;
  }

  render();
}

bootstrap();
