/* global JSZip */
const state = {
  token: "",
  user: null,
  zip: null,
  originalZipFile: null,
  files: [],
  selectedRepo: null,
  objectUrls: new Map(),
  previewUrl: "",
  previewSession: "",
};

const els = {
  tokenInput: document.getElementById("tokenInput"),
  rememberToken: document.getElementById("rememberToken"),
  loginBtn: document.getElementById("loginBtn"),
  authStatus: document.getElementById("authStatus"),
  zipInput: document.getElementById("zipInput"),
  zipStatus: document.getElementById("zipStatus"),
  projectSummary: document.getElementById("projectSummary"),
  fileTree: document.getElementById("fileTree"),
  codePreview: document.getElementById("codePreview"),
  selectedFileName: document.getElementById("selectedFileName"),
  selectedFileMeta: document.getElementById("selectedFileMeta"),
  staticPreviewBtn: document.getElementById("staticPreviewBtn"),
  openPreviewBtn: document.getElementById("openPreviewBtn"),
  runnerUrlInput: document.getElementById("runnerUrlInput"),
  runnerBtn: document.getElementById("runnerBtn"),
  runStatus: document.getElementById("runStatus"),
  previewFrame: document.getElementById("previewFrame"),
  loadReposBtn: document.getElementById("loadReposBtn"),
  repoSelect: document.getElementById("repoSelect"),
  newRepoName: document.getElementById("newRepoName"),
  newRepoPrivate: document.getElementById("newRepoPrivate"),
  createRepoBtn: document.getElementById("createRepoBtn"),
  targetFolderInput: document.getElementById("targetFolderInput"),
  repoStatus: document.getElementById("repoStatus"),
  commitMessageInput: document.getElementById("commitMessageInput"),
  pushBtn: document.getElementById("pushBtn"),
  pushProgress: document.getElementById("pushProgress"),
  pushStatus: document.getElementById("pushStatus"),
  logOutput: document.getElementById("logOutput"),
  expandAllBtn: document.getElementById("expandAllBtn"),
  downloadExtractedBtn: document.getElementById("downloadExtractedBtn"),
  loadCommitsBtn: document.getElementById("loadCommitsBtn"),
  rollbackLastBtn: document.getElementById("rollbackLastBtn"),
  rollbackSelectedBtn: document.getElementById("rollbackSelectedBtn"),
  commitSelect: document.getElementById("commitSelect"),
  rollbackStatus: document.getElementById("rollbackStatus"),
};

const TEXT_EXTENSIONS = new Set([
  "txt","md","markdown","html","htm","css","scss","sass","less","js","jsx","ts","tsx","mjs","cjs",
  "json","xml","svg","yml","yaml","toml","ini","env","example","gitignore","dockerfile","dockerignore",
  "py","php","rb","go","rs","java","kt","kts","c","h","cpp","hpp","cs","swift","dart","lua",
  "sh","bash","zsh","ps1","bat","cmd","sql","graphql","gql","vue","svelte","astro","jsx"
]);

const MIME_TYPES = {
  html: "text/html; charset=utf-8",
  htm: "text/html; charset=utf-8",
  css: "text/css; charset=utf-8",
  js: "text/javascript; charset=utf-8",
  mjs: "text/javascript; charset=utf-8",
  cjs: "text/javascript; charset=utf-8",
  json: "application/json; charset=utf-8",
  svg: "image/svg+xml",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  ico: "image/x-icon",
  wasm: "application/wasm",
  txt: "text/plain; charset=utf-8",
  md: "text/markdown; charset=utf-8",
  xml: "application/xml; charset=utf-8",
  pdf: "application/pdf",
};

const SKIP_DIRS = [
  ".git/", "node_modules/", "vendor/", ".next/cache/", ".vercel/", "__pycache__/"
];

function setStatus(el, message, type = "") {
  el.textContent = message;
  el.className = `status ${type}`.trim();
}

function log(message) {
  const now = new Date().toLocaleTimeString();
  els.logOutput.textContent += `[${now}] ${message}\n`;
  els.logOutput.scrollTop = els.logOutput.scrollHeight;
}

function bytes(n) {
  if (!Number.isFinite(n)) return "-";
  const units = ["B", "KB", "MB", "GB"];
  let size = n;
  let i = 0;
  while (size >= 1024 && i < units.length - 1) {
    size /= 1024;
    i++;
  }
  return `${size.toFixed(size >= 10 || i === 0 ? 0 : 1)} ${units[i]}`;
}

function extOf(path) {
  const base = path.split("/").pop() || "";
  const idx = base.lastIndexOf(".");
  return idx >= 0 ? base.slice(idx + 1).toLowerCase() : "";
}

function mimeOf(path) {
  return MIME_TYPES[extOf(path)] || "application/octet-stream";
}

function normalizeZipPath(path) {
  return String(path)
    .replaceAll("\\", "/")
    .split("/")
    .filter(part => part && part !== "." && part !== "..")
    .join("/");
}

function shouldSkip(path) {
  return SKIP_DIRS.some(dir => path.includes(`/${dir}`) || path.startsWith(dir));
}

function isLikelyText(path, size) {
  if (size > 2 * 1024 * 1024) return false;
  if (TEXT_EXTENSIONS.has(extOf(path))) return true;
  const name = path.split("/").pop()?.toLowerCase() || "";
  return ["dockerfile", "makefile", "license", "readme"].includes(name);
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return btoa(binary);
}

function githubErrorMessage(err) {
  const parts = [];
  if (err.status) parts.push(`HTTP ${err.status}`);
  if (err.message) parts.push(err.message);
  if (Array.isArray(err.data?.errors)) {
    parts.push(err.data.errors.map(e => e.message || e.code || JSON.stringify(e)).join("; "));
  }
  if (err.data?.documentation_url) parts.push(`Docs: ${err.data.documentation_url}`);

  if (err.status === 401) parts.push("Token salah/expired.");
  if (err.status === 403) parts.push("Cek permission token. Minimal butuh Contents: Read and write. Jika upload workflow, token juga butuh permission Workflows.");
  if (err.status === 404) parts.push("Repo/branch tidak ditemukan, atau token tidak punya akses ke repo.");
  if (err.status === 409) parts.push("Konflik branch. Load repo lagi lalu ulangi push.");
  if (err.status === 422) parts.push("Validasi GitHub gagal. Cek nama branch, file terlalu besar, path invalid, atau branch protection.");

  return parts.join(" | ");
}

function revokeObjectUrls() {
  for (const url of state.objectUrls.values()) URL.revokeObjectURL(url);
  state.objectUrls.clear();
}

async function github(path, options = {}) {
  if (!state.token) throw new Error("Belum login GitHub.");
  const res = await fetch(`https://api.github.com${path}`, {
    ...options,
    headers: {
      "Accept": "application/vnd.github+json",
      "Authorization": `Bearer ${state.token}`,
      "X-GitHub-Api-Version": "2022-11-28",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {}),
    },
  });

  const text = await res.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; } catch { data = text; }

  if (!res.ok) {
    const msg = data?.message || text || `${res.status} ${res.statusText}`;
    const err = new Error(msg);
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

async function registerPreviewServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    setStatus(els.runStatus, "Browser tidak support Service Worker. Preview static butuh deploy HTTPS atau localhost.", "err");
    return false;
  }

  if (location.protocol === "file:") {
    setStatus(els.runStatus, "Preview static tidak bisa dari file://. Jalankan via Vercel, localhost server, atau GitHub Pages.", "err");
    return false;
  }

  const reg = await navigator.serviceWorker.register("./sw.js", { scope: "./" });
  if (!navigator.serviceWorker.controller) {
    await new Promise(resolve => {
      navigator.serviceWorker.addEventListener("controllerchange", resolve, { once: true });
      reg.update().catch(() => {});
      setTimeout(resolve, 1200);
    });
  }
  return true;
}

async function login() {
  const token = els.tokenInput.value.trim();
  if (!token) return setStatus(els.authStatus, "Masukkan GitHub token.", "err");

  state.token = token;
  setStatus(els.authStatus, "Mengecek token...");
  try {
    const user = await github("/user");
    state.user = user;
    if (els.rememberToken.checked) {
      localStorage.setItem("zip2gh_token", token);
    } else {
      localStorage.removeItem("zip2gh_token");
    }
    setStatus(els.authStatus, `Login sebagai ${user.login}.`, "ok");
  } catch (err) {
    state.token = "";
    setStatus(els.authStatus, `Login gagal: ${githubErrorMessage(err)}`, "err");
  }
}

async function handleZip(file) {
  if (!file) return;
  state.originalZipFile = file;
  revokeObjectUrls();
  state.files = [];
  state.zip = null;
  state.previewUrl = "";
  state.previewSession = "";
  els.previewFrame.removeAttribute("src");
  els.previewFrame.removeAttribute("srcdoc");
  els.codePreview.textContent = "Pilih file teks untuk preview.";
  setStatus(els.zipStatus, `Membaca ${file.name} (${bytes(file.size)})...`);

  try {
    const zip = await JSZip.loadAsync(file);
    const entries = [];
    for (const [rawPath, entry] of Object.entries(zip.files)) {
      const path = normalizeZipPath(rawPath);
      if (!path || entry.dir || shouldSkip(path)) continue;
      entries.push({ path, entry, size: entry._data?.uncompressedSize || 0 });
    }

    entries.sort((a, b) => a.path.localeCompare(b.path));
    state.zip = zip;
    state.files = entries;

    renderSummary();
    renderFileTree();
    setStatus(els.zipStatus, `ZIP berhasil diekstrak: ${entries.length} file siap dipreview/push.`, "ok");
  } catch (err) {
    setStatus(els.zipStatus, `Gagal membaca ZIP: ${err.message}`, "err");
  }
}

function detectProject(files) {
  const names = new Set(files.map(f => f.path.toLowerCase()));
  const has = (...candidates) => candidates.some(c => names.has(c));
  const any = pattern => files.some(f => pattern.test(f.path.toLowerCase()));

  const detected = [];
  if (has("package.json") || any(/\/package\.json$/)) detected.push("Node/React/Vite/Next");
  if (has("composer.json") || any(/\.php$/)) detected.push("PHP");
  if (has("requirements.txt", "pyproject.toml") || any(/\.py$/)) detected.push("Python");
  if (has("go.mod") || any(/\.go$/)) detected.push("Go");
  if (has("cargo.toml") || any(/\.rs$/)) detected.push("Rust");
  if (any(/\.java$/)) detected.push("Java");
  if (any(/index\.html$/)) detected.push("Static HTML");
  return detected.length ? detected.join(", ") : "Tidak terdeteksi";
}

function renderSummary() {
  const totalBytes = state.files.reduce((sum, f) => sum + f.size, 0);
  const textFiles = state.files.filter(f => isLikelyText(f.path, f.size)).length;
  const largeFiles = state.files.filter(f => f.size > 100 * 1024 * 1024).length;
  const workflowFiles = state.files.filter(f => f.path.startsWith(".github/workflows/")).length;

  els.projectSummary.innerHTML = [
    ["Files", state.files.length],
    ["Size", bytes(totalBytes)],
    ["Text preview", textFiles],
    ["GitHub >100MiB", largeFiles],
  ].map(([label, value]) => `
    <div class="summary-card"><strong>${value}</strong><span>${label}</span></div>
  `).join("");

  if (largeFiles > 0) {
    setStatus(
      els.zipStatus,
      `ZIP terbaca, tetapi ada ${largeFiles} file >100MiB. GitHub biasanya menolak file sebesar itu di repo biasa.`,
      "err"
    );
  } else if (workflowFiles > 0) {
    setStatus(
      els.zipStatus,
      `ZIP terbaca. Catatan: ada ${workflowFiles} file di .github/workflows; token mungkin butuh permission Workflows.`,
      "ok"
    );
  }
}

function renderFileTree() {
  if (!state.files.length) {
    els.fileTree.textContent = "Tidak ada file.";
    els.fileTree.className = "file-tree empty";
    return;
  }

  const folderSet = new Set();
  for (const f of state.files) {
    const parts = f.path.split("/");
    for (let i = 1; i < parts.length; i++) folderSet.add(parts.slice(0, i).join("/"));
  }

  const all = [
    ...[...folderSet].map(path => ({ path, folder: true, size: 0 })),
    ...state.files.map(f => ({ ...f, folder: false })),
  ].sort((a, b) => {
    const depthA = a.path.split("/").length;
    const depthB = b.path.split("/").length;
    if (depthA !== depthB) return depthA - depthB;
    if (a.folder !== b.folder) return a.folder ? -1 : 1;
    return a.path.localeCompare(b.path);
  });

  const rows = all.map(item => {
    const depth = item.path.split("/").length - 1;
    const pad = depth * 14;
    if (item.folder) {
      return `
        <div class="file-row folder" style="padding-left:${pad + 6}px">
          <span>📁 ${item.path.split("/").pop()}</span><small></small>
        </div>
      `;
    }
    return `
      <div class="file-row" data-path="${encodeURIComponent(item.path)}" style="padding-left:${pad + 6}px">
        <span>📄 ${item.path.split("/").pop()}</span><small>${bytes(item.size)}</small>
      </div>
    `;
  });

  els.fileTree.className = "file-tree";
  els.fileTree.innerHTML = rows.join("");
  els.fileTree.querySelectorAll(".file-row[data-path]").forEach(row => {
    row.addEventListener("click", () => previewFile(decodeURIComponent(row.dataset.path)));
  });
}

async function previewFile(path) {
  const file = state.files.find(f => f.path === path);
  if (!file) return;
  els.selectedFileName.textContent = path;
  els.selectedFileMeta.textContent = bytes(file.size);

  if (!isLikelyText(file.path, file.size)) {
    els.codePreview.textContent = "File binary atau terlalu besar untuk preview teks.";
    return;
  }

  try {
    const text = await file.entry.async("text");
    els.codePreview.textContent = text || "(file kosong)";
  } catch (err) {
    els.codePreview.textContent = `Gagal preview: ${err.message}`;
  }
}

function injectPreviewBase(htmlText, session, indexPath) {
  const dir = indexPath.includes("/") ? indexPath.split("/").slice(0, -1).join("/") + "/" : "";
  const baseHref = `/__zip_preview__/${session}/${dir}`;

  if (/<base\s/i.test(htmlText)) {
    return htmlText.replace(/<base\b[^>]*>/i, `<base href="${baseHref}">`);
  }

  if (/<head[^>]*>/i.test(htmlText)) {
    return htmlText.replace(/<head([^>]*)>/i, `<head$1><base href="${baseHref}">`);
  }

  return `<base href="${baseHref}">\n${htmlText}`;
}

async function putPreviewCache(session, indexFile) {
  const cache = await caches.open("zip-preview-v2");
  const oldKeys = await cache.keys();
  await Promise.all(oldKeys.map(key => cache.delete(key)));

  for (const file of state.files) {
    let body;
    let contentType = mimeOf(file.path);

    if (file.path === indexFile.path && /html/i.test(contentType)) {
      const html = await file.entry.async("text");
      body = injectPreviewBase(html, session, indexFile.path);
    } else if (isLikelyText(file.path, file.size)) {
      body = await file.entry.async("text");
    } else {
      body = await file.entry.async("blob");
    }

    const url = new URL(`/__zip_preview__/${session}/${file.path}`, location.origin);
    await cache.put(url.toString(), new Response(body, {
      headers: {
        "Content-Type": contentType,
        "Cache-Control": "no-store",
      },
    }));
  }

  const indexUrl = new URL(`/__zip_preview__/${session}/${indexFile.path}`, location.origin).toString();
  const indexResponse = await cache.match(indexUrl, { ignoreSearch: true });

  if (indexResponse) {
    await cache.put(new URL(`/__zip_preview__/${session}/`, location.origin).toString(), indexResponse.clone());
    await cache.put(new URL(`/__zip_preview__/${session}/index.html`, location.origin).toString(), indexResponse.clone());
  }
}

async function staticPreview() {
  if (!state.files.length) return setStatus(els.runStatus, "Upload ZIP dulu.", "err");

  const ok = await registerPreviewServiceWorker();
  if (!ok) return;

  const indexFile =
    state.files.find(f => f.path.toLowerCase() === "index.html") ||
    state.files.find(f => f.path.toLowerCase().endsWith("/index.html"));

  if (!indexFile) {
    return setStatus(els.runStatus, "Tidak menemukan index.html untuk preview statis.", "err");
  }

  try {
    setStatus(els.runStatus, `Menyiapkan virtual local server untuk ${indexFile.path}...`);
    const session = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    await putPreviewCache(session, indexFile);

    state.previewSession = session;
    state.previewUrl = `/__zip_preview__/${session}/${indexFile.path}`;

    if (navigator.serviceWorker.controller) {
      navigator.serviceWorker.controller.postMessage({ type: "ZIP_PREVIEW_SESSION", session });
    }

    // Small delay so SW receives session message before subresources are requested.
    await new Promise(resolve => setTimeout(resolve, 120));

    els.previewFrame.removeAttribute("srcdoc");
    els.previewFrame.src = state.previewUrl;
    setStatus(els.runStatus, `Preview static berjalan: ${state.previewUrl}`, "ok");
  } catch (err) {
    setStatus(els.runStatus, `Preview gagal: ${err.message}`, "err");
  }
}

function openPreviewTab() {
  if (!state.previewUrl) {
    setStatus(els.runStatus, "Klik Preview Static dulu, lalu buka tab.", "err");
    return;
  }
  window.open(state.previewUrl, "_blank", "noopener,noreferrer");
}

async function runViaRunner() {
  if (!state.originalZipFile) return setStatus(els.runStatus, "Upload ZIP dulu.", "err");
  const base = els.runnerUrlInput.value.trim().replace(/\/$/, "");
  if (!base) return setStatus(els.runStatus, "Isi URL runner dulu.", "err");

  const form = new FormData();
  form.append("zip", state.originalZipFile, state.originalZipFile.name);

  setStatus(els.runStatus, "Mengirim ZIP ke runner...");
  try {
    const res = await fetch(`${base}/run`, { method: "POST", body: form });
    const data = await res.json().catch(() => null);
    if (!res.ok) throw new Error(data?.error || `${res.status} ${res.statusText}`);

    if (data.logs) log(data.logs);
    if (data.url) {
      els.previewFrame.src = data.url;
      state.previewUrl = data.url;
      setStatus(els.runStatus, `Runner aktif: ${data.url}`, "ok");
    } else {
      setStatus(els.runStatus, "Runner selesai, tetapi tidak mengembalikan URL.", "err");
    }
  } catch (err) {
    setStatus(els.runStatus, `Runner gagal: ${err.message}`, "err");
  }
}

async function loadRepos() {
  if (!state.user) return setStatus(els.repoStatus, "Login GitHub dulu.", "err");
  setStatus(els.repoStatus, "Mengambil repository...");
  try {
    const repos = await github("/user/repos?per_page=100&sort=updated&affiliation=owner,collaborator,organization_member");
    els.repoSelect.innerHTML = `<option value="">Pilih repo...</option>` + repos.map(repo =>
      `<option value="${repo.full_name}" data-default-branch="${repo.default_branch || "main"}">${repo.full_name} (${repo.private ? "private" : "public"})</option>`
    ).join("");
    setStatus(els.repoStatus, `${repos.length} repository dimuat.`, "ok");
  } catch (err) {
    setStatus(els.repoStatus, `Gagal load repo: ${githubErrorMessage(err)}`, "err");
  }
}

async function createRepo() {
  if (!state.user) return setStatus(els.repoStatus, "Login GitHub dulu.", "err");
  const name = els.newRepoName.value.trim();
  if (!name) return setStatus(els.repoStatus, "Isi nama repo.", "err");

  setStatus(els.repoStatus, `Membuat repo ${name}...`);
  try {
    const repo = await github("/user/repos", {
      method: "POST",
      body: JSON.stringify({
        name,
        private: els.newRepoPrivate.value === "true",
        auto_init: true,
      }),
    });
    await loadRepos();
    els.repoSelect.value = repo.full_name;
    setStatus(els.repoStatus, `Repo dibuat: ${repo.full_name}.`, "ok");
  } catch (err) {
    setStatus(els.repoStatus, `Gagal create repo: ${githubErrorMessage(err)}`, "err");
  }
}

function selectedRepo() {
  const value = els.repoSelect.value;
  if (!value) return null;
  const [owner, repo] = value.split("/");
  return { owner, repo };
}

function targetPath(path) {
  const folder = normalizeZipPath(els.targetFolderInput.value.trim());
  return folder ? `${folder}/${path}` : path;
}

function historyKey(owner, repo) {
  return `zip2gh_history_${owner}/${repo}`;
}

function readPushHistory(owner, repo) {
  try {
    return JSON.parse(localStorage.getItem(historyKey(owner, repo)) || "[]");
  } catch {
    return [];
  }
}

function writePushHistory(owner, repo, item) {
  const list = readPushHistory(owner, repo);
  list.unshift(item);
  localStorage.setItem(historyKey(owner, repo), JSON.stringify(list.slice(0, 20)));
}

async function getRefOrNull(owner, repo, branch) {
  try {
    return await github(`/repos/${owner}/${repo}/git/ref/heads/${encodeURIComponent(branch)}`);
  } catch (err) {
    if (err.status === 404) return null;
    throw err;
  }
}

async function ensureMainRef(owner, repo) {
  let ref = await getRefOrNull(owner, repo, "main");
  if (ref) return ref;

  log("Branch main belum ada. Mencoba membuat main dari default branch...");
  const repoInfo = await github(`/repos/${owner}/${repo}`);

  if (repoInfo.default_branch && repoInfo.default_branch !== "main") {
    const defaultRef = await getRefOrNull(owner, repo, repoInfo.default_branch);
    if (defaultRef) {
      await github(`/repos/${owner}/${repo}/git/refs`, {
        method: "POST",
        body: JSON.stringify({ ref: "refs/heads/main", sha: defaultRef.object.sha }),
      });
      log(`Branch main dibuat dari ${repoInfo.default_branch}.`);
      return await getRefOrNull(owner, repo, "main");
    }
  }

  log("Repo tampaknya kosong. Membuat initial file agar branch main tersedia...");
  const content = btoa("Initial commit for ZIP to GitHub app.\n");
  try {
    await github(`/repos/${owner}/${repo}/contents/.zip-to-github-init`, {
      method: "PUT",
      body: JSON.stringify({
        message: "Initial commit",
        content,
        branch: "main",
      }),
    });
  } catch (err) {
    log(`Initial commit via Contents API gagal: ${githubErrorMessage(err)}`);
    throw err;
  }

  ref = await getRefOrNull(owner, repo, "main");
  if (!ref) throw new Error("Branch main belum tersedia setelah initial commit.");
  return ref;
}

async function pushToGithub() {
  const repoInfo = selectedRepo();
  if (!repoInfo) return setStatus(els.pushStatus, "Pilih repository dulu.", "err");
  if (!state.user) return setStatus(els.pushStatus, "Login GitHub dulu.", "err");
  if (!state.files.length) return setStatus(els.pushStatus, "Upload ZIP dulu.", "err");

  const huge = state.files.filter(f => f.size > 100 * 1024 * 1024);
  if (huge.length) {
    return setStatus(
      els.pushStatus,
      `Dibatalkan: ${huge.length} file lebih besar dari 100MiB. GitHub akan menolak file ini di repository biasa.`,
      "err"
    );
  }

  const branch = "main";
  const message = els.commitMessageInput.value.trim() || "Upload extracted ZIP files";
  els.pushProgress.value = 0;
  els.logOutput.textContent = "";
  setStatus(els.pushStatus, "Menyiapkan commit...");
  log(`Target: ${repoInfo.owner}/${repoInfo.repo}@${branch}`);

  try {
    const ref = await ensureMainRef(repoInfo.owner, repoInfo.repo);
    const parentCommitSha = ref.object.sha;
    log(`Current main commit: ${parentCommitSha}`);

    const commit = await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/git/commits/${parentCommitSha}`);
    const baseTreeSha = commit.tree.sha;
    log(`Base tree: ${baseTreeSha}`);

    const entries = [];
    for (let i = 0; i < state.files.length; i++) {
      const file = state.files[i];
      const buffer = await file.entry.async("arraybuffer");
      const base64 = arrayBufferToBase64(buffer);

      const blob = await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/git/blobs`, {
        method: "POST",
        body: JSON.stringify({ content: base64, encoding: "base64" }),
      });

      entries.push({
        path: targetPath(file.path),
        mode: "100644",
        type: "blob",
        sha: blob.sha,
      });

      const pct = Math.round(((i + 1) / state.files.length) * 70);
      els.pushProgress.value = pct;
      if (i === 0 || (i + 1) % 10 === 0 || i === state.files.length - 1) {
        log(`Uploaded blobs: ${i + 1}/${state.files.length}`);
      }
    }

    const tree = await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/git/trees`, {
      method: "POST",
      body: JSON.stringify({
        base_tree: baseTreeSha,
        tree: entries,
      }),
    });
    els.pushProgress.value = 82;
    log(`Created tree: ${tree.sha}`);

    const newCommit = await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/git/commits`, {
      method: "POST",
      body: JSON.stringify({
        message,
        tree: tree.sha,
        parents: [parentCommitSha],
      }),
    });
    els.pushProgress.value = 92;
    log(`Created commit: ${newCommit.sha}`);

    await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/git/refs/heads/${branch}`, {
      method: "PATCH",
      body: JSON.stringify({ sha: newCommit.sha, force: false }),
    });

    writePushHistory(repoInfo.owner, repoInfo.repo, {
      repo: `${repoInfo.owner}/${repoInfo.repo}`,
      branch,
      before: parentCommitSha,
      after: newCommit.sha,
      message,
      at: new Date().toISOString(),
      files: state.files.length,
    });

    els.pushProgress.value = 100;
    const folder = normalizeZipPath(els.targetFolderInput.value.trim());
    setStatus(
      els.pushStatus,
      `Berhasil push ${state.files.length} file ke ${repoInfo.owner}/${repoInfo.repo}@main${folder ? ` folder ${folder}` : ""}.`,
      "ok"
    );
    log("Done. Rollback point tersimpan di browser.");
  } catch (err) {
    els.pushProgress.value = 0;
    const msg = githubErrorMessage(err);
    setStatus(els.pushStatus, `Push gagal: ${msg}`, "err");
    log(`ERROR: ${msg}`);
    if (err.data) log(`RAW ERROR: ${JSON.stringify(err.data, null, 2)}`);
  }
}

async function loadCommits() {
  const repoInfo = selectedRepo();
  if (!repoInfo) return setStatus(els.rollbackStatus, "Pilih repository dulu.", "err");
  if (!state.user) return setStatus(els.rollbackStatus, "Login GitHub dulu.", "err");

  setStatus(els.rollbackStatus, "Mengambil recent commits dari main...");
  try {
    const commits = await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/commits?sha=main&per_page=30`);
    els.commitSelect.innerHTML = `<option value="">Pilih commit rollback...</option>` + commits.map(c => {
      const date = new Date(c.commit?.committer?.date || c.commit?.author?.date || Date.now()).toLocaleString();
      const short = c.sha.slice(0, 7);
      const msg = (c.commit?.message || "").split("\n")[0].replace(/[<>&"]/g, s => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;" }[s]));
      return `<option value="${c.sha}">${short} — ${date} — ${msg}</option>`;
    }).join("");
    setStatus(els.rollbackStatus, `${commits.length} commit dimuat.`, "ok");
  } catch (err) {
    setStatus(els.rollbackStatus, `Gagal load commits: ${githubErrorMessage(err)}`, "err");
  }
}

async function rollbackToSha(sha, label) {
  const repoInfo = selectedRepo();
  if (!repoInfo) return setStatus(els.rollbackStatus, "Pilih repository dulu.", "err");
  if (!sha) return setStatus(els.rollbackStatus, "Pilih commit rollback dulu.", "err");

  const ok = confirm(`Rollback branch main ke ${label || sha}?\n\nIni akan memindahkan pointer main. File di main kembali ke kondisi commit tersebut.`);
  if (!ok) return;

  setStatus(els.rollbackStatus, `Rollback main ke ${sha.slice(0, 7)}...`);
  try {
    await github(`/repos/${repoInfo.owner}/${repoInfo.repo}/git/refs/heads/main`, {
      method: "PATCH",
      body: JSON.stringify({ sha, force: true }),
    });
    setStatus(els.rollbackStatus, `Rollback berhasil ke ${sha.slice(0, 7)}.`, "ok");
    await loadCommits();
  } catch (err) {
    setStatus(els.rollbackStatus, `Rollback gagal: ${githubErrorMessage(err)}`, "err");
  }
}

async function rollbackLastPush() {
  const repoInfo = selectedRepo();
  if (!repoInfo) return setStatus(els.rollbackStatus, "Pilih repository dulu.", "err");

  const history = readPushHistory(repoInfo.owner, repoInfo.repo);
  const last = history[0];
  if (!last?.before) {
    return setStatus(els.rollbackStatus, "Tidak ada riwayat push dari app ini di browser ini.", "err");
  }

  await rollbackToSha(last.before, `commit sebelum push terakhir (${last.before.slice(0, 7)})`);
}

async function downloadExtractedZip() {
  if (!state.files.length) return;
  const out = new JSZip();
  for (const file of state.files) {
    const buf = await file.entry.async("arraybuffer");
    out.file(file.path, buf);
  }
  const blob = await out.generateAsync({ type: "blob" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "extracted-clean.zip";
  a.click();
  URL.revokeObjectURL(url);
}

function boot() {
  const saved = localStorage.getItem("zip2gh_token");
  if (saved) {
    els.tokenInput.value = saved;
    els.rememberToken.checked = true;
  }

  els.loginBtn.addEventListener("click", login);
  els.zipInput.addEventListener("change", e => handleZip(e.target.files?.[0]));
  els.expandAllBtn.addEventListener("click", renderFileTree);
  els.downloadExtractedBtn.addEventListener("click", downloadExtractedZip);
  els.staticPreviewBtn.addEventListener("click", staticPreview);
  els.openPreviewBtn.addEventListener("click", openPreviewTab);
  els.runnerBtn.addEventListener("click", runViaRunner);
  els.loadReposBtn.addEventListener("click", loadRepos);
  els.createRepoBtn.addEventListener("click", createRepo);
  els.pushBtn.addEventListener("click", pushToGithub);
  els.loadCommitsBtn.addEventListener("click", loadCommits);
  els.rollbackLastBtn.addEventListener("click", rollbackLastPush);
  els.rollbackSelectedBtn.addEventListener("click", () => {
    const sha = els.commitSelect.value;
    const label = els.commitSelect.selectedOptions?.[0]?.textContent || sha;
    rollbackToSha(sha, label);
  });
  els.repoSelect.addEventListener("change", () => {
    const repo = selectedRepo();
    setStatus(els.repoStatus, repo ? `Dipilih: ${repo.owner}/${repo.repo}, target branch main.` : "Repository belum dipilih.", repo ? "ok" : "");
    els.commitSelect.innerHTML = `<option value="">Load commits dulu...</option>`;
  });
}

boot();


// Expose selected helpers for Repo Terminal extension.
window.ZipGitApi = {
  github,
  selectedRepo,
  ensureMainRef,
  githubErrorMessage,
  isLikelyText,
  bytes,
  writePushHistory,
  readPushHistory,
};
