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
   LOAD DATA FROM GOOGLE SHEETS
   Paste your "Publish to web" CSV link between the quotes below
   ===================================================== */
const SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTz7y9Tg1pXpdWwWw8cLW3cE581jQtKfMDLcjvKx48W89-oHrAi0YmGvzLMf2xZ30zrZvGQcdYFVo9D/pub?gid=1319714774&single=true&output=csv";

function loadData() {
    return fetch(SHEET_URL, { cache: "no-store" })
        .then(function (res) {
            if (!res.ok) {
                throw new Error("Google Sheet could not be loaded.");
            }
            return res.text();
        })
        .then(function (text) {
            const workbook = XLSX.read(text, { type: "string", raw: true });   // raw: keeps IDs and weights exactly as text
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            window.excelError = "";
            console.log("Rows loaded:", rows.length, rows);   // check in browser console (F12)
        })
        .catch(function (err) {
            // If old data is already loaded, keep using it instead of showing an error
            if (rows.length === 0) {
                window.excelError = "Unable to load data: " + err.message;
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

    const status = getVal(row, "Status");
    const statusClass = status.toLowerCase().replace(/[^a-z]/g, "");   // pass / fail / hold

    resultBox.innerHTML =
        '<table class="result-table">' +
        "<tr><th>ID Number</th><td>" + escapeHtml(getVal(row, "ID Number")) + "</td></tr>" +
        "<tr><th>Section</th><td>" + escapeHtml(getVal(row, "Section")) + "</td></tr>" +
        "<tr><th>Grade</th><td>" + escapeHtml(getVal(row, "Grade")) + "</td></tr>" +
        "<tr><th>Embossing</th><td>" + escapeHtml(getVal(row, "Embossing")) + "</td></tr>" +
        "<tr><th>Weight in MT</th><td>" + escapeHtml(getVal(row, "Weight in MT")) + "</td></tr>" +
        '<tr><th>Status</th><td><span class="status ' + statusClass + '">' + escapeHtml(status) + "</span></td></tr>" +
        "</table>";
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
