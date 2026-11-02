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

// Intro video elements
const intro = document.getElementById("intro");
const introVideo = document.getElementById("introVideo");
const skipBtn = document.getElementById("skipBtn");
const soundBtn = document.getElementById("soundBtn");
const soundOnIcon = document.getElementById("soundOn");
const soundOffIcon = document.getElementById("soundOff");
const tapPlay = document.getElementById("tapPlay");

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

// Image shown when the searched ID is not in the sheet
const NOT_FOUND_IMAGE = "id-not-found.jpg.jpeg";
new Image().src = NOT_FOUND_IMAGE;     // loads the image in advance so it appears instantly

function showNotFound() {
    const wrap = document.createElement("div");
    wrap.className = "not-found";

    const img = document.createElement("img");
    img.className = "not-found-img";
    img.alt = "ID not found. Please check the ID Number and try again.";
    img.onerror = function () {        // if the image file is missing, show text instead
        showError("ID not found. Please check the ID Number and try again.");
    };
    img.src = NOT_FOUND_IMAGE;

    wrap.appendChild(img);
    resultBox.innerHTML = "";
    resultBox.appendChild(wrap);
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
   INTRO VIDEO: plays when the app opens, then shows the
   Get Start page (section1)
   ===================================================== */
const INTRO_END_HOLD_MS = reduceMotion ? 0 : 300;   // short pause on the last frame
const INTRO_STUCK_MS = 15000;                       // if the video never starts, move on
let introDone = false;
let introWatchdog = null;

async function finishIntro() {
    if (introDone) return;
    introDone = true;
    clearTimeout(introWatchdog);

    intro.classList.add("fade-out");
    await sleep(FADE_OUT_MS);

    try { introVideo.pause(); } catch (e) { /* ignore */ }
    intro.classList.add("is-hidden");

    section1.classList.remove("is-hidden");
    section1.classList.add("page-enter");
    window.scrollTo(0, 0);
    // remove the animation class so it cannot block the fade-out used by Get Start
    setTimeout(function () { section1.classList.remove("page-enter"); }, 700);
}

function showSoundButton(muted) {
    soundBtn.classList.remove("is-hidden");
    soundOffIcon.classList.toggle("is-hidden", !muted);
    soundOnIcon.classList.toggle("is-hidden", muted);
    soundBtn.setAttribute("aria-label", muted ? "Turn sound on" : "Turn sound off");
}

async function startIntro() {
    if (reduceMotion) {            // user prefers less motion: go straight to Get Start
        finishIntro();
        return;
    }

    introWatchdog = setTimeout(function () {
        if (introVideo.currentTime === 0 && tapPlay.classList.contains("is-hidden")) {
            finishIntro();         // video could not load: never leave the user stuck
        }
    }, INTRO_STUCK_MS);

    try {
        introVideo.muted = false;          // 1st try: with sound
        await introVideo.play();
    } catch (e1) {
        try {
            introVideo.muted = true;       // 2nd try: browsers always allow muted autoplay
            await introVideo.play();
            showSoundButton(true);
        } catch (e2) {
            tapPlay.classList.remove("is-hidden");   // 3rd: ask the user to tap once
        }
    }
}

introVideo.addEventListener("ended", function () {
    setTimeout(finishIntro, INTRO_END_HOLD_MS);
});
introVideo.addEventListener("error", finishIntro);
skipBtn.addEventListener("click", finishIntro);

soundBtn.addEventListener("click", function () {
    introVideo.muted = !introVideo.muted;
    showSoundButton(introVideo.muted);
});

tapPlay.addEventListener("click", function () {
    tapPlay.classList.add("is-hidden");
    introVideo.muted = false;
    introVideo.play().catch(function () {
        introVideo.muted = true;
        introVideo.play().then(function () { showSoundButton(true); }).catch(finishIntro);
    });
});

startIntro();

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
        showNotFound();
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

    resultBox.innerHTML =
        '<table class="result-table">' +
        "<tr><th>ID Number</th><td>" + escapeHtml(getVal(row, "ID Number")) + "</td></tr>" +
        "<tr><th>Section</th><td>" + escapeHtml(getVal(row, "Section")) + "</td></tr>" +
        "<tr><th>Grade</th><td>" + escapeHtml(getVal(row, "Grade")) + "</td></tr>" +
        "<tr><th>Embossing</th><td>" + escapeHtml(getVal(row, "Embossing")) + "</td></tr>" +
        "<tr><th>Weight in MT</th><td>" + escapeHtml(getVal(row, "Weight in MT")) + "</td></tr>" +
        locationRow +
        '<tr><th>Status</th><td><span class="status ' + statusClass + '">' + statusIcon + "<span>" + escapeHtml(statusText) + "</span></span></td></tr>" +
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
