(() => {
  const ts = {
    repo: null,
    branch: "main",
    cwd: "",
    tree: [],
    treeByPath: new Map(),
    fileCache: new Map(),
    staged: new Map(),
    selectedFile: "",
    baseCommitSha: "",
    baseTreeSha: "",
  };

  const q = id => document.getElementById(id);
  const els = {
    repoSelect: q("repoSelect"),
    loadWorkspaceBtn: q("loadWorkspaceBtn"),
    terminalRepoBadge: q("terminalRepoBadge"),
    terminalPathBadge: q("terminalPathBadge"),
    terminalStageBadge: q("terminalStageBadge"),
    terminalFileSelect: q("terminalFileSelect"),
    terminalOpenSelectedBtn: q("terminalOpenSelectedBtn"),
    terminalOutput: q("terminalOutput"),
    terminalPrompt: q("terminalPrompt"),
    terminalInput: q("terminalInput"),
    terminalRunBtn: q("terminalRunBtn"),
    editorPathInput: q("editorPathInput"),
    editorSaveBtn: q("editorSaveBtn"),
    repoEditor: q("repoEditor"),
    editorNewBtn: q("editorNewBtn"),
    editorReloadBtn: q("editorReloadBtn"),
    editorDeleteBtn: q("editorDeleteBtn"),
    terminalCommitBtn: q("terminalCommitBtn"),
  };

  const TEXT_EXTENSIONS = new Set([
    "txt","md","markdown","html","htm","css","scss","sass","less","js","jsx","ts","tsx","mjs","cjs",
    "json","xml","svg","yml","yaml","toml","ini","env","example","gitignore","dockerfile","dockerignore",
    "py","php","rb","go","rs","java","kt","kts","c","h","cpp","hpp","cs","swift","dart","lua",
    "sh","bash","zsh","ps1","bat","cmd","sql","graphql","gql","vue","svelte","astro"
  ]);

  function api() {
    if (!window.ZipGitApi) throw new Error("GitHub API bridge belum siap. Refresh halaman.");
    return window.ZipGitApi;
  }

  function bytes(n) {
    const fromApi = window.ZipGitApi?.bytes;
    if (fromApi) return fromApi(n);
    if (!Number.isFinite(n)) return "-";
    const units = ["B", "KB", "MB", "GB"];
    let size = n, i = 0;
    while (size >= 1024 && i < units.length - 1) { size /= 1024; i++; }
    return `${size.toFixed(size >= 10 || i === 0 ? 0 : 1)} ${units[i]}`;
  }

  function extOf(path) {
    const base = String(path).split("/").pop() || "";
    const idx = base.lastIndexOf(".");
    return idx >= 0 ? base.slice(idx + 1).toLowerCase() : "";
  }

  function isText(path, size) {
    if (size > 2 * 1024 * 1024) return false;
    if (TEXT_EXTENSIONS.has(extOf(path))) return true;
    const name = String(path).split("/").pop()?.toLowerCase() || "";
    return ["dockerfile", "makefile", "license", "readme"].includes(name);
  }

  function base64ToUtf8(base64) {
    const binary = atob(String(base64 || "").replace(/\n/g, ""));
    const arr = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) arr[i] = binary.charCodeAt(i);
    return new TextDecoder("utf-8", { fatal: false }).decode(arr);
  }

  function utf8ToBase64(text) {
    const arr = new TextEncoder().encode(text || "");
    let out = "";
    for (let i = 0; i < arr.length; i += 0x8000) {
      out += String.fromCharCode(...arr.subarray(i, i + 0x8000));
    }
    return btoa(out);
  }

  function normalizePath(path) {
    const input = String(path || "").trim().replaceAll("\\", "/");
    let parts = [];
    if (!input.startsWith("/")) parts = ts.cwd ? ts.cwd.split("/") : [];
    for (const part of input.split("/")) {
      if (!part || part === ".") continue;
      if (part === "..") parts.pop();
      else parts.push(part);
    }
    return parts.join("/");
  }

  function promptText() {
    return `${ts.cwd ? "/" + ts.cwd : "/"} $`;
  }

  function term(text = "", kind = "line") {
    const div = document.createElement("div");
    div.className = `terminal-${kind}`;
    div.textContent = text;
    els.terminalOutput.appendChild(div);
    els.terminalOutput.scrollTop = els.terminalOutput.scrollHeight;
  }

  function termCommand(command) {
    term(`${promptText()} ${command}`, "command");
  }

  function errorMessage(err) {
    if (window.ZipGitApi?.githubErrorMessage) return window.ZipGitApi.githubErrorMessage(err);
    return err?.message || String(err);
  }

  function updateBadges() {
    els.terminalRepoBadge.textContent = `Repo: ${ts.repo?.fullName || "belum dipilih"}`;
    els.terminalPathBadge.textContent = `Path: /${ts.cwd || ""}`;
    els.terminalStageBadge.textContent = `Stage: ${ts.staged.size}`;
    els.terminalPrompt.textContent = promptText();
  }

  function reset(repo = null) {
    ts.repo = repo;
    ts.cwd = "";
    ts.tree = [];
    ts.treeByPath = new Map();
    ts.fileCache = new Map();
    ts.staged = new Map();
    ts.selectedFile = "";
    ts.baseCommitSha = "";
    ts.baseTreeSha = "";
    els.terminalFileSelect.innerHTML = '<option value="">File repo...</option>';
    els.editorPathInput.value = "";
    els.repoEditor.value = "";
    updateBadges();
  }

  async function getMainBase(repo) {
    const ref = await api().ensureMainRef(repo.owner, repo.repo);
    const parentCommitSha = ref.object.sha;
    const commit = await api().github(`/repos/${repo.owner}/${repo.repo}/git/commits/${parentCommitSha}`);
    return { parentCommitSha, baseTreeSha: commit.tree.sha };
  }

  async function loadWorkspace(verbose = true) {
    const repo = api().selectedRepo();
    if (!repo) { term("Pilih repository dulu.", "error"); return; }
    repo.fullName = `${repo.owner}/${repo.repo}`;
    reset(repo);
    if (verbose) {
      termCommand("pull");
      term(`Loading ${repo.fullName}@main...`, "dim");
    }
    try {
      const base = await getMainBase(repo);
      ts.baseCommitSha = base.parentCommitSha;
      ts.baseTreeSha = base.baseTreeSha;
      const data = await api().github(`/repos/${repo.owner}/${repo.repo}/git/trees/${base.baseTreeSha}?recursive=1`);
      ts.tree = (data.tree || []).filter(x => x.path && x.type).sort((a, b) => a.path.localeCompare(b.path));
      ts.treeByPath = new Map(ts.tree.map(x => [x.path, x]));
      renderFileSelect();
      updateBadges();
      term(`Workspace loaded: ${ts.tree.length} entries. Base ${base.parentCommitSha.slice(0, 7)}.`, "line");
      term("Tips HP: tombol open/save/commit mengurangi kebutuhan mengetik.", "dim");
    } catch (err) {
      term(`Load workspace gagal: ${errorMessage(err)}`, "error");
    }
  }

  function renderFileSelect() {
    const files = ts.tree.filter(x => x.type === "blob").slice(0, 1200);
    els.terminalFileSelect.innerHTML = '<option value="">File repo...</option>' + files.map(x => {
      const mark = ts.staged.has(x.path) ? "*" : "";
      return `<option value="${escapeHtml(x.path)}">${mark}${escapeHtml(x.path)}</option>`;
    }).join("");
  }

  function escapeHtml(text) {
    return String(text ?? "").replace(/[<>&"]/g, c => ({"<":"&lt;", ">":"&gt;", "&":"&amp;", '"':"&quot;"}[c]));
  }

  function isDirectory(path) {
    const clean = normalizePath(path);
    if (!clean) return true;
    const item = ts.treeByPath.get(clean);
    if (item?.type === "tree") return true;
    const prefix = `${clean}/`;
    return ts.tree.some(x => x.path.startsWith(prefix));
  }

  function children(path = ts.cwd) {
    const base = normalizePath(path);
    const prefix = base ? `${base}/` : "";
    const seen = new Map();
    for (const item of ts.tree) {
      if (base && item.path !== base && !item.path.startsWith(prefix)) continue;
      const rest = base ? item.path.slice(prefix.length) : item.path;
      if (!rest) continue;
      const [first, ...more] = rest.split("/");
      const childPath = prefix + first;
      const type = more.length || item.type === "tree" ? "tree" : item.type;
      if (!seen.has(childPath)) seen.set(childPath, { path: childPath, name: first, type, size: item.size });
    }
    for (const [pathKey, staged] of ts.staged) {
      if (staged.delete) continue;
      if (base && pathKey !== base && !pathKey.startsWith(prefix)) continue;
      const rest = base ? pathKey.slice(prefix.length) : pathKey;
      if (!rest) continue;
      const [first, ...more] = rest.split("/");
      const childPath = prefix + first;
      if (!seen.has(childPath)) seen.set(childPath, { path: childPath, name: first, type: more.length ? "tree" : "blob", size: staged.content?.length || 0 });
    }
    return [...seen.values()].sort((a, b) => a.type !== b.type ? (a.type === "tree" ? -1 : 1) : a.name.localeCompare(b.name));
  }

  async function getFileText(path) {
    const clean = normalizePath(path);
    const staged = ts.staged.get(clean);
    if (staged?.delete) throw new Error("File sedang distage untuk delete.");
    if (staged && !staged.delete) return staged.content || "";
    if (ts.fileCache.has(clean)) return ts.fileCache.get(clean);
    const item = ts.treeByPath.get(clean);
    if (!item || item.type !== "blob") throw new Error(`File tidak ditemukan: ${clean}`);
    if (!isText(clean, item.size || 0)) throw new Error("File binary atau terlalu besar untuk editor teks.");
    const blob = await api().github(`/repos/${ts.repo.owner}/${ts.repo.repo}/git/blobs/${item.sha}`);
    const text = base64ToUtf8(blob.content || "");
    ts.fileCache.set(clean, text);
    return text;
  }

  async function openFile(path) {
    const clean = normalizePath(path);
    if (!clean) throw new Error("Path file kosong.");
    const text = await getFileText(clean);
    ts.selectedFile = clean;
    els.editorPathInput.value = clean;
    els.repoEditor.value = text;
    term(`Opened ${clean} (${text.length} chars).`);
  }

  function newFile(path) {
    const clean = normalizePath(path);
    if (!clean) throw new Error("Path file kosong.");
    if (isDirectory(clean)) throw new Error("Path adalah folder. Tulis nama file lengkap.");
    ts.selectedFile = clean;
    els.editorPathInput.value = clean;
    els.repoEditor.value = "";
    term(`New file ready: ${clean}. Edit lalu save.`);
  }

  function stageSave() {
    const clean = normalizePath(els.editorPathInput.value || ts.selectedFile);
    if (!clean) { term("Isi path file di editor terlebih dahulu.", "error"); return; }
    ts.staged.set(clean, { type: "blob", content: els.repoEditor.value, delete: false });
    ts.selectedFile = clean;
    ts.fileCache.set(clean, els.repoEditor.value);
    renderFileSelect();
    updateBadges();
    term(`Staged save: ${clean}`);
  }

  function stageDelete(path = "") {
    const clean = normalizePath(path || els.editorPathInput.value || ts.selectedFile);
    if (!clean) { term("Pilih file untuk delete.", "error"); return; }
    ts.staged.set(clean, { type: "blob", delete: true });
    renderFileSelect();
    updateBadges();
    term(`Staged delete: ${clean}`);
  }

  function stageTouch(path) {
    const clean = normalizePath(path);
    if (!clean) throw new Error("Path file kosong.");
    ts.staged.set(clean, { type: "blob", content: "", delete: false });
    renderFileSelect();
    updateBadges();
    term(`Staged empty file: ${clean}`);
  }

  function stageMkdir(path) {
    const clean = normalizePath(path);
    if (!clean) throw new Error("Path folder kosong.");
    const keep = `${clean}/.gitkeep`;
    ts.staged.set(keep, { type: "blob", content: "", delete: false });
    renderFileSelect();
    updateBadges();
    term(`Staged folder placeholder: ${keep}`);
  }

  function help() {
    term("Commands:", "dim");
    term("  help                         bantuan");
    term("  pwd                          folder aktif");
    term("  ls [folder]                  lihat isi folder");
    term("  cd <folder|..|/>             pindah folder");
    term("  open <file> / cat <file>     buka file ke editor");
    term("  new <file>                   buat file baru di editor");
    term("  save                         stage isi editor");
    term("  touch <file>                 stage file kosong");
    term("  mkdir <folder>               stage folder via .gitkeep");
    term("  rm <file>                    stage delete file");
    term("  status                       lihat staged changes");
    term("  discard [file]               hapus staged file / semua staged");
    term("  commit <message>             commit staged ke main");
    term("  pull                         reload workspace dari main");
    term("  clear                        bersihkan terminal");
  }

  async function run(raw) {
    const command = String(raw || "").trim();
    if (!command) return;
    termCommand(command);
    els.terminalInput.value = "";
    const [cmdRaw] = command.split(/\s+/);
    const cmd = cmdRaw.toLowerCase();
    const rest = command.slice(cmdRaw.length).trim();
    try {
      if (!ts.repo && !["help", "clear"].includes(cmd)) { term("Load Workspace Repo dulu.", "error"); return; }
      if (cmd === "help") return help();
      if (cmd === "clear") { els.terminalOutput.textContent = ""; return; }
      if (cmd === "pull") return loadWorkspace(false);
      if (cmd === "pwd") return term(`/${ts.cwd || ""}`);
      if (cmd === "ls" || cmd === "dir") {
        const target = rest ? normalizePath(rest) : ts.cwd;
        if (!isDirectory(target)) throw new Error("Folder tidak ditemukan.");
        const list = children(target);
        if (!list.length) return term("(kosong)", "dim");
        for (const item of list) {
          const mark = item.type === "tree" ? "📁" : "📄";
          const staged = ts.staged.has(item.path) ? "*" : " ";
          term(`${staged} ${mark} ${item.name}${item.type === "tree" ? "/" : `  ${bytes(item.size || 0)}`}`);
        }
        return;
      }
      if (cmd === "cd") {
        const clean = normalizePath(rest || "/");
        if (!isDirectory(clean)) throw new Error(`Folder tidak ditemukan: ${rest}`);
        ts.cwd = clean;
        updateBadges();
        return term(`/${ts.cwd || ""}`);
      }
      if (["open", "cat", "edit"].includes(cmd)) {
        if (!rest) throw new Error("Tulis path file. Contoh: open README.md");
        return openFile(rest);
      }
      if (cmd === "new") {
        if (!rest) throw new Error("Tulis path file baru. Contoh: new docs/catatan.md");
        return newFile(rest);
      }
      if (cmd === "save") return stageSave();
      if (cmd === "touch") { if (!rest) throw new Error("Tulis path file."); return stageTouch(rest); }
      if (cmd === "mkdir") { if (!rest) throw new Error("Tulis path folder."); return stageMkdir(rest); }
      if (cmd === "rm" || cmd === "delete") { if (!rest) throw new Error("Tulis path file."); return stageDelete(rest); }
      if (cmd === "status") {
        if (!ts.staged.size) return term("No staged changes.", "dim");
        for (const [path, item] of ts.staged) term(`${item.delete ? "DELETE" : ts.treeByPath.has(path) ? "MODIFY" : "ADD"}  ${path}`);
        return;
      }
      if (cmd === "discard") {
        if (rest) { const clean = normalizePath(rest); ts.staged.delete(clean); term(`Discarded staged: ${clean}`); }
        else { ts.staged.clear(); term("Discarded all staged changes."); }
        renderFileSelect(); updateBadges(); return;
      }
      if (cmd === "commit" || cmd === "push") return commitStaged(rest || "Edit repository from web terminal");
      term(`Command tidak dikenal: ${cmd}. Ketik help.`, "error");
    } catch (err) {
      term(`ERROR: ${err.message}`, "error");
    }
  }

  async function commitStaged(message) {
    if (!ts.repo) { term("Load Workspace Repo dulu.", "error"); return; }
    if (!ts.staged.size) { term("Tidak ada staged changes. Jalankan save/new/touch/mkdir/rm dulu.", "error"); return; }
    term(`Committing ${ts.staged.size} staged changes to ${ts.repo.fullName}@main...`, "dim");
    try {
      const base = await getMainBase(ts.repo);
      const treeEntries = [];
      for (const [path, item] of ts.staged) {
        if (item.delete) {
          treeEntries.push({ path, mode: "100644", type: "blob", sha: null });
          continue;
        }
        const blob = await api().github(`/repos/${ts.repo.owner}/${ts.repo.repo}/git/blobs`, {
          method: "POST",
          body: JSON.stringify({ content: utf8ToBase64(item.content || ""), encoding: "base64" }),
        });
        treeEntries.push({ path, mode: "100644", type: "blob", sha: blob.sha });
      }
      const tree = await api().github(`/repos/${ts.repo.owner}/${ts.repo.repo}/git/trees`, {
        method: "POST",
        body: JSON.stringify({ base_tree: base.baseTreeSha, tree: treeEntries }),
      });
      const commit = await api().github(`/repos/${ts.repo.owner}/${ts.repo.repo}/git/commits`, {
        method: "POST",
        body: JSON.stringify({ message, tree: tree.sha, parents: [base.parentCommitSha] }),
      });
      await api().github(`/repos/${ts.repo.owner}/${ts.repo.repo}/git/refs/heads/main`, {
        method: "PATCH",
        body: JSON.stringify({ sha: commit.sha, force: false }),
      });
      if (api().writePushHistory) {
        api().writePushHistory(ts.repo.owner, ts.repo.repo, {
          repo: ts.repo.fullName,
          branch: "main",
          before: base.parentCommitSha,
          after: commit.sha,
          message,
          at: new Date().toISOString(),
          files: ts.staged.size,
          source: "terminal",
        });
      }
      term(`Commit OK: ${commit.sha.slice(0, 7)} — ${message}`);
      ts.staged.clear();
      await loadWorkspace(false);
    } catch (err) {
      term(`Commit gagal: ${errorMessage(err)}`, "error");
      if (err.data) term(JSON.stringify(err.data, null, 2), "error");
    }
  }

  function bind() {
    if (!els.loadWorkspaceBtn) return;
    els.loadWorkspaceBtn.addEventListener("click", () => loadWorkspace(true));
    els.terminalRunBtn.addEventListener("click", () => run(els.terminalInput.value));
    els.terminalInput.addEventListener("keydown", e => { if (e.key === "Enter") run(els.terminalInput.value); });
    document.querySelectorAll("[data-terminal-run]").forEach(btn => btn.addEventListener("click", () => run(btn.dataset.terminalRun)));
    document.querySelectorAll("[data-terminal-insert]").forEach(btn => btn.addEventListener("click", () => {
      const value = btn.dataset.terminalInsert;
      els.terminalInput.value = value;
      els.terminalInput.focus();
      els.terminalInput.setSelectionRange(value.length, value.length);
    }));
    els.terminalOpenSelectedBtn.addEventListener("click", () => {
      if (els.terminalFileSelect.value) run(`open ${els.terminalFileSelect.value}`);
    });
    els.editorSaveBtn.addEventListener("click", stageSave);
    els.editorNewBtn.addEventListener("click", () => {
      const path = prompt("Path file baru, contoh: docs/catatan.md");
      if (path) { try { newFile(path); } catch (err) { term(`ERROR: ${err.message}`, "error"); } }
    });
    els.editorReloadBtn.addEventListener("click", () => {
      const path = els.editorPathInput.value || ts.selectedFile;
      if (path) run(`open ${path}`);
    });
    els.editorDeleteBtn.addEventListener("click", () => stageDelete());
    els.terminalCommitBtn.addEventListener("click", () => {
      const msg = prompt("Commit message", "Edit repository from web terminal");
      if (msg !== null) commitStaged(msg || "Edit repository from web terminal");
    });
    els.repoSelect.addEventListener("change", () => {
      const repo = window.ZipGitApi?.selectedRepo?.();
      if (repo) { repo.fullName = `${repo.owner}/${repo.repo}`; reset(repo); term(`Repo selected: ${repo.fullName}. Klik Load Workspace Repo.`, "dim"); }
      else reset(null);
    });
    updateBadges();
    term("Repo Terminal ready. Pilih repo lalu klik Load Workspace Repo.", "dim");
  }

  bind();
})();
