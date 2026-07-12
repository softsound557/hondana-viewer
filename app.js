const CLIENT_ID = "9a019c14-5029-4055-b5c8-0e42509f6999";
const REDIRECT_URI = "https://softsound557.github.io/hondana-viewer/";
const SCOPES = ["Files.Read"];
const FOLDER_NAME = "本棚エクスポート";
const GRAPH_ROOT = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(FOLDER_NAME)}`;

const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/consumers",
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "localStorage",
  },
});

let books = [];
let currentBook = null; // { title, cover, pageCount, pages }
let currentPageIndex = 0;
const blobUrlCache = new Map(); // key: graph path, value: object URL

const setupView = document.getElementById("setup-view");
const shelfView = document.getElementById("shelf-view");
const readerView = document.getElementById("reader-view");
const shelfGridEl = document.getElementById("shelf-grid");
const countEl = document.getElementById("count");
const statusEl = document.getElementById("status");
const readerImage = document.getElementById("reader-image");
const readerViewport = document.getElementById("reader-viewport");
const readerPageIndicator = document.getElementById("reader-page-indicator");

function showStatus(message) {
  statusEl.textContent = message;
  statusEl.hidden = !message;
}

function showSignIn() {
  setupView.hidden = false;
  shelfView.hidden = true;
  readerView.hidden = true;
}

async function getToken() {
  const account = msalInstance.getActiveAccount();
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch (e) {
    await msalInstance.acquireTokenRedirect({ scopes: SCOPES });
    throw e;
  }
}

function graphPathUrl(...segments) {
  const encoded = segments.map((s) => encodeURIComponent(s)).join("/");
  return `${GRAPH_ROOT}/${encoded}:/content`;
}

async function loadBlobUrl(graphPath) {
  if (blobUrlCache.has(graphPath)) return blobUrlCache.get(graphPath);
  const token = await getToken();
  const res = await fetch(graphPath, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error("fetch failed: " + res.status);
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  blobUrlCache.set(graphPath, url);
  return url;
}

function evictFarBlobUrls(keepPaths) {
  for (const [path, url] of blobUrlCache.entries()) {
    if (!keepPaths.has(path)) {
      URL.revokeObjectURL(url);
      blobUrlCache.delete(path);
    }
  }
}

const coverObserver = new IntersectionObserver((entries, obs) => {
  for (const entry of entries) {
    if (!entry.isIntersecting) continue;
    const el = entry.target;
    obs.unobserve(el);
    const graphPath = el.dataset.graphPath;
    loadBlobUrl(graphPath)
      .then((url) => { el.src = url; })
      .catch(() => {});
  }
});

async function loadData() {
  setupView.hidden = true;
  shelfView.hidden = false;

  try {
    const token = await getToken();
    const res = await fetch(`${GRAPH_ROOT}/data.json:/content`, {
      headers: { Authorization: `Bearer ${token}` },
      cache: "no-store",
    });
    if (!res.ok) throw new Error("status " + res.status);
    books = await res.json();
    showStatus("");
  } catch (e) {
    showStatus("最新データを取得できませんでした(前回表示分を表示しています)");
  }
  renderShelf();
}

function renderShelf() {
  countEl.textContent = `${books.length}冊`;
  shelfGridEl.innerHTML = "";

  if (books.length === 0) {
    shelfGridEl.innerHTML = '<div class="empty">本がまだありません</div>';
    return;
  }

  for (const book of books) {
    const card = document.createElement("a");
    card.href = "#";
    card.className = "book-card";
    card.addEventListener("click", (e) => {
      e.preventDefault();
      openBook(book);
    });

    const cover = document.createElement(book.cover ? "img" : "div");
    cover.className = "book-cover" + (book.cover ? "" : " no-cover");
    if (book.cover) {
      cover.dataset.graphPath = graphPathUrl("books", book.title, book.cover);
      cover.loading = "lazy";
      cover.alt = "";
      coverObserver.observe(cover);
    } else {
      cover.textContent = "表紙なし";
    }

    const titleEl = document.createElement("div");
    titleEl.className = "book-title";
    titleEl.textContent = book.title;

    card.appendChild(cover);
    card.appendChild(titleEl);
    shelfGridEl.appendChild(card);
  }
}

async function openBook(book) {
  try {
    const token = await getToken();
    const res = await fetch(graphPathUrl("books", book.title, "pages.json"), {
      headers: { Authorization: `Bearer ${token}` },
      cache: "no-store",
    });
    if (!res.ok) throw new Error("status " + res.status);
    const data = await res.json();
    currentBook = { ...book, pages: data.pages };
  } catch (e) {
    showStatus("本のページ一覧を取得できませんでした");
    return;
  }

  currentPageIndex = 0;
  shelfView.hidden = true;
  readerView.hidden = false;
  showPage(0);
}

function closeReader() {
  readerView.hidden = true;
  shelfView.hidden = false;
  currentBook = null;
  resetZoom();
}

function pageGraphPath(index) {
  const filename = currentBook.pages[index];
  return graphPathUrl("books", currentBook.title, "pages", filename);
}

async function showPage(index) {
  if (!currentBook) return;
  if (index < 0 || index >= currentBook.pages.length) return;

  currentPageIndex = index;
  resetZoom();
  readerPageIndicator.textContent = `${index + 1} / ${currentBook.pages.length}`;

  const targetPath = pageGraphPath(index);
  try {
    const url = await loadBlobUrl(targetPath);
    if (currentBook && currentPageIndex === index) {
      readerImage.src = url;
    }
  } catch (e) {
    showStatus("ページを取得できませんでした");
  }

  const keep = new Set();
  for (let i = index - 1; i <= index + 1; i++) {
    if (i >= 0 && i < currentBook.pages.length) {
      keep.add(pageGraphPath(i));
      if (i !== index) loadBlobUrl(pageGraphPath(i)).catch(() => {});
    }
  }
  evictFarBlobUrls(keep);
}

function nextPage() {
  if (currentBook && currentPageIndex < currentBook.pages.length - 1) showPage(currentPageIndex + 1);
}

function prevPage() {
  if (currentBook && currentPageIndex > 0) showPage(currentPageIndex - 1);
}

// --- ズーム・パン・タップ操作 ---

let scale = 1;
let tx = 0;
let ty = 0;
const MIN_SCALE = 1;
const MAX_SCALE = 4;
const DOUBLE_TAP_SCALE = 2.5;

const activePointers = new Map();
let gestureMode = null; // "pan" | "pinch" | null
let panStart = null; // { x, y, tx, ty }
let pinchStart = null; // { dist, scale, midX, midY, tx, ty }
let lastTapTime = 0;
let lastTapX = 0;
let lastTapY = 0;

function applyTransform() {
  readerImage.style.transform = `translate(${tx}px, ${ty}px) scale(${scale})`;
}

function resetZoom() {
  scale = 1;
  tx = 0;
  ty = 0;
  applyTransform();
}

function clampPan() {
  const rect = readerViewport.getBoundingClientRect();
  const maxX = (rect.width * (scale - 1)) / 2;
  const maxY = (rect.height * (scale - 1)) / 2;
  tx = Math.max(-maxX, Math.min(maxX, tx));
  ty = Math.max(-maxY, Math.min(maxY, ty));
}

function distance(p1, p2) {
  return Math.hypot(p1.x - p2.x, p1.y - p2.y);
}

function setZoomAt(newScale, clientX, clientY) {
  const rect = readerViewport.getBoundingClientRect();
  const originX = clientX - rect.left - rect.width / 2;
  const originY = clientY - rect.top - rect.height / 2;
  const ratio = newScale / scale;
  tx = originX - (originX - tx) * ratio;
  ty = originY - (originY - ty) * ratio;
  scale = newScale;
  clampPan();
  applyTransform();
}

readerViewport.addEventListener("pointerdown", (e) => {
  readerViewport.setPointerCapture(e.pointerId);
  activePointers.set(e.pointerId, { x: e.clientX, y: e.clientY });

  if (activePointers.size === 1) {
    gestureMode = "pan";
    panStart = { x: e.clientX, y: e.clientY, tx, ty, moved: false };
  } else if (activePointers.size === 2) {
    const pts = [...activePointers.values()];
    gestureMode = "pinch";
    pinchStart = {
      dist: distance(pts[0], pts[1]),
      scale,
      midX: (pts[0].x + pts[1].x) / 2,
      midY: (pts[0].y + pts[1].y) / 2,
      tx,
      ty,
    };
  }
});

readerViewport.addEventListener("pointermove", (e) => {
  if (!activePointers.has(e.pointerId)) return;
  activePointers.set(e.pointerId, { x: e.clientX, y: e.clientY });

  if (gestureMode === "pinch" && activePointers.size === 2) {
    const pts = [...activePointers.values()];
    const newDist = distance(pts[0], pts[1]);
    const newScale = Math.max(MIN_SCALE, Math.min(MAX_SCALE, pinchStart.scale * (newDist / pinchStart.dist)));
    const midX = (pts[0].x + pts[1].x) / 2;
    const midY = (pts[0].y + pts[1].y) / 2;
    setZoomAt(newScale, midX, midY);
  } else if (gestureMode === "pan" && activePointers.size === 1) {
    const dx = e.clientX - panStart.x;
    const dy = e.clientY - panStart.y;
    if (Math.abs(dx) > 6 || Math.abs(dy) > 6) panStart.moved = true;
    if (scale > 1.001) {
      tx = panStart.tx + dx;
      ty = panStart.ty + dy;
      clampPan();
      applyTransform();
    }
  }
});

function handlePointerEnd(e) {
  const wasSinglePan = gestureMode === "pan" && activePointers.size === 1;
  const notMoved = wasSinglePan && panStart && !panStart.moved;
  activePointers.delete(e.pointerId);

  if (activePointers.size === 0) {
    gestureMode = null;

    if (notMoved) {
      const now = Date.now();
      const isDoubleTap =
        now - lastTapTime < 300 && distance({ x: e.clientX, y: e.clientY }, { x: lastTapX, y: lastTapY }) < 40;

      if (isDoubleTap) {
        lastTapTime = 0;
        if (scale > 1.001) {
          resetZoom();
        } else {
          setZoomAt(DOUBLE_TAP_SCALE, e.clientX, e.clientY);
        }
      } else {
        lastTapTime = now;
        lastTapX = e.clientX;
        lastTapY = e.clientY;

        if (scale <= 1.001) {
          const rect = readerViewport.getBoundingClientRect();
          const isRightHalf = e.clientX - rect.left > rect.width / 2;
          if (isRightHalf) nextPage(); else prevPage();
        }
      }
    }
  } else if (activePointers.size === 1) {
    gestureMode = "pan";
    const [remaining] = activePointers.values();
    panStart = { x: remaining.x, y: remaining.y, tx, ty, moved: false };
  }
}

readerViewport.addEventListener("pointerup", handlePointerEnd);
readerViewport.addEventListener("pointercancel", handlePointerEnd);

document.getElementById("reader-back").addEventListener("click", (e) => {
  e.preventDefault();
  closeReader();
});

document.getElementById("settings-link").addEventListener("click", (e) => {
  e.preventDefault();
  msalInstance.logoutRedirect();
});

document.getElementById("signin-button").addEventListener("click", () => {
  msalInstance.loginRedirect({ scopes: SCOPES });
});

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js");
}

async function init() {
  const result = await msalInstance.handleRedirectPromise();
  let account = msalInstance.getAllAccounts()[0];
  if (result && result.account) {
    account = result.account;
  }
  if (!account) {
    showSignIn();
    return;
  }
  msalInstance.setActiveAccount(account);
  await loadData();
}

init();
