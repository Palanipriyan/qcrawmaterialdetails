/* =====================================================
   SETTINGS (change these numbers to speed up / slow down)
   ===================================================== */
const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

const FADE_OUT_MS = reduceMotion ? 0 : 400;    // section1 fades out
const LOADER_MIN_MS = reduceMotion ? 0 : 1600; // minimum time the loader stays on screen
const LOADER_HIDE_MS = reduceMotion ? 0 : 300; // loader fades out
const SPIN_MS = reduceMotion ? 0 : 800;        // icon rotation time on search

/* =====================================================
   ELEMENTS
   ===================================================== */
const section1 = document.getElementById("section1");
const section2 = document.getElementById("section2");
const loader = document.getElementById("loader");
const loaderBar = document.getElementById("loaderBar");
const form = document.getElementById("searchForm");
const idInput = document.getElementById("idInput");
const searchBtn = form.querySelector("button");
const resultBox = document.getElementById("result");
const iconImg = document.querySelector(".icon-size");

let rows = [];          // Google Sheet data is stored here
let isSearching = false;
let isTransitioning = false;

/* =====================================================
   HELPERS
   ===================================================== */
function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
}

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

// Finds a column value ignoring extra spaces and upper/lower case in header names
function getVal(row, name) {
    const key = Object.keys(row).find(function (k) {
        return k.trim().toLowerCase() === name.toLowerCase();
    });
    return key !== undefined ? String(row[key]).trim() : "";
}

function showError(message) {
    resultBox.innerHTML = '<p class="error">' + escapeHtml(message) + "</p>";
}

/* =====================================================
   LOAD DATA FROM GOOGLE SHEETS (fast updates)
   ===================================================== */

// 1. FAST LINK: paste your Sheet ID here (the long text between /d/ and /edit in the sheet's address bar)
const SHEET_ID = "PASTE_YOUR_SHEET_ID_HERE";
const SHEET_GID = "1185246626";   // tab number: the number after gid= in the sheet's address bar

// 2. BACKUP LINK: your "Publish to web" CSV link (used only if the fast link fails or SHEET_ID is not set)
const PUBLISHED_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTJRhr9oBhh9VhVYuPhLHACu0pnXTOFceSW8mYnljuHNqVE7Zy49RfZfQA0OlmIZzHmLeA64wDrU2ed/pub?gid=2061645292&single=true&output=csv";

function getSheetUrls() {
    const urls = [];
    if (SHEET_ID && SHEET_ID.indexOf("PASTE_") !== 0) {
        urls.push(
            "https://docs.google.com/spreadsheets/d/" + SHEET_ID +
            "/gviz/tq?tqx=out:csv&headers=1&gid=" + SHEET_GID +
            "&_=" + Date.now()                      // stops the browser from reusing old data
        );
    }
    if (PUBLISHED_CSV_URL) {
        urls.push(PUBLISHED_CSV_URL + (PUBLISHED_CSV_URL.indexOf("?") > -1 ? "&" : "?") + "_=" + Date.now());
    }
    return urls;
}

async function fetchCsv() {
    const urls = getSheetUrls();
    if (urls.length === 0) {
        throw new Error("Sheet link is not set in index.js.");
    }
    let lastError = null;
    for (const url of urls) {
        try {
            const res = await fetch(url, { cache: "no-store" });
            if (!res.ok) throw new Error("Google Sheet could not be loaded.");
            const text = await res.text();
            if (text.trim().charAt(0) === "<") {
                throw new Error("Sheet is not shared. Set sharing to 'Anyone with the link can view'.");
            }
            return text;
        } catch (err) {
            lastError = err;
        }
    }
    throw lastError;
}

function loadData() {
    return fetchCsv()
        .then(function (text) {
            const workbook = XLSX.read(text, { type: "string", raw: true });   // raw: keeps IDs and weights exactly as text
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            window.excelError = "";
            window.usingCache = "";           // fresh data from the internet
            try {                             // keep a copy for offline use
                localStorage.setItem("qc_rows", JSON.stringify({ time: Date.now(), rows: rows }));
            } catch (e) { /* storage full or blocked: ignore */ }
            console.log("Rows loaded:", rows.length, rows);   // check in browser console (F12)
        })
        .catch(function (err) {
            // No internet or sheet not reachable: use the copy saved from the last successful load
            try {
                const saved = JSON.parse(localStorage.getItem("qc_rows") || "null");
                if (saved && saved.rows && saved.rows.length) {
                    rows = saved.rows;
                    window.excelError = "";
                    window.usingCache = new Date(saved.time).toLocaleString();
                    return;
                }
            } catch (e) { /* ignore */ }

            if (rows.length === 0) {
                window.excelError = navigator.onLine
                    ? "Unable to load data: " + err.message
                    : "You are offline and no saved data is available. Please connect to the internet once to download the data.";
            }
        });
}

const excelReady = loadData();

/* =====================================================
   GET START BUTTON: fade out -> loader -> fade in section2
   (called from index.html: onclick="display('section2')")
   ===================================================== */
async function display(targetId) {
    if (isTransitioning) return;
    isTransitioning = true;

    const target = document.getElementById(targetId);

    // 1. Fade out the current section
    section1.classList.add("fade-out");
    await sleep(FADE_OUT_MS);

    // 2. Show the loader (stays at least LOADER_MIN_MS and waits for Excel to finish loading)
    if (!reduceMotion) {
        loaderBar.style.animationDuration = LOADER_MIN_MS + "ms";
        loader.classList.add("show");
        await Promise.all([sleep(LOADER_MIN_MS), excelReady]);
    }

    // 3. Swap the sections behind the loader
    section1.classList.add("is-hidden");
    target.classList.remove("is-hidden");
    window.scrollTo(0, 0);

    // 4. Hide the loader and fade in section2
    loader.classList.remove("show");
    target.classList.add("page-enter");
    await sleep(LOADER_HIDE_MS);

    if (window.excelError) {
        showError(window.excelError);
    }
    isTransitioning = false;
}

/* =====================================================
   SEARCH: icon rotates, then the result appears
   ===================================================== */
function spinIcon() {
    iconImg.classList.remove("rotate");
    void iconImg.offsetWidth;          // restarts the animation if clicked again
    iconImg.classList.add("rotate");
}

/* =====================================================
   ICONS (inline SVG, no internet or image files needed)
   ===================================================== */
const ICON_CLOCK =
    '<svg class="status-icon" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>';

const ICON_PIN =
    '<svg class="loc-icon" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>';

function renderResult(rawId) {
    const id = rawId.trim().toLowerCase();

    if (!id) {
        showError("Please enter an ID Number.");
        return;
    }

    if (window.excelError) {
        showError(window.excelError);
        return;
    }

    if (rows.length === 0) {
        showError("Data is not loaded. Please check your internet connection and try again.");
        return;
    }

    const row = rows.find(function (r) {
        return getVal(r, "ID Number").toLowerCase() === id;
    });

    if (!row) {
        showError("ID not found. Please check the ID Number and try again.");
        return;
    }

    // Sheet value (PASS / FAIL / HOLD / blank) -> text, color and icon shown on screen
    const status = getVal(row, "Status");
    const s = status.toLowerCase();
    let statusClass = "";      // color: pass = green, fail = red, hold = orange, progress = yellow
    let statusText = status;   // unknown values are shown as they are
    let statusIcon = "";       // clock icon only for "in process"
    if (s === "") {
        statusClass = "progress";
        statusText = "IN PROCESS";
        statusIcon = ICON_CLOCK;
    } else if (s.indexOf("pass") > -1 || s.indexOf("accept") > -1) {
        statusClass = "pass";
        statusText = "ID ACCEPTED";
    } else if (s.indexOf("fail") > -1 || s.indexOf("reject") > -1) {
        statusClass = "fail";
        statusText = "ID REJECTED";
    } else if (s.indexOf("hold") > -1) {
        statusClass = "hold";
        statusText = "ID IN HOLD";
    }

    // Location row is shown only when the Location cell has a value
    const location = getVal(row, "Location");
    const locationRow = location
        ? '<tr><th>Location</th><td><span class="loc">' + ICON_PIN + "<span>" + escapeHtml(location) + "</span></span></td></tr>"
        : "";

    // Shown only when the data comes from the saved offline copy
    const offlineNote = window.usingCache
        ? '<p class="offline-note">Offline mode: showing data last updated on ' + escapeHtml(window.usingCache) + "</p>"
        : "";

    resultBox.innerHTML =
        '<table class="result-table">' +
        "<tr><th>ID Number</th><td>" + escapeHtml(getVal(row, "ID Number")) + "</td></tr>" +
        "<tr><th>Section</th><td>" + escapeHtml(getVal(row, "Section")) + "</td></tr>" +
        "<tr><th>Grade</th><td>" + escapeHtml(getVal(row, "Grade")) + "</td></tr>" +
        "<tr><th>Embossing</th><td>" + escapeHtml(getVal(row, "Embossing")) + "</td></tr>" +
        "<tr><th>Weight in MT</th><td>" + escapeHtml(getVal(row, "Weight in MT")) + "</td></tr>" +
        locationRow +
        '<tr><th>Status</th><td><span class="status ' + statusClass + '">' + statusIcon + "<span>" + escapeHtml(statusText) + "</span></span></td></tr>" +
        "</table>" + offlineNote;
}

form.addEventListener("submit", async function (e) {
    e.preventDefault();               // stops page reload
    if (isSearching) return;
    isSearching = true;
    searchBtn.disabled = true;

    resultBox.innerHTML = "";         // clear the old result
    spinIcon();                       // icon rotates
    await Promise.all([sleep(SPIN_MS), loadData()]);   // icon rotates while the latest data loads

    renderResult(idInput.value);

    searchBtn.disabled = false;
    isSearching = false;
});

/* =====================================================
   OFFLINE SUPPORT: service worker (sw.js) saves the app files on the device
   ===================================================== */
if ("serviceWorker" in navigator) {
    window.addEventListener("load", function () {
        navigator.serviceWorker.register("sw.js").catch(function (err) {
            console.log("Service worker not registered:", err);
        });
    });
}
