const CONFIG = {
  APP_VERSION: "2.0.6",
  MAX_REPORTS: 10,
  DRAFT_SAVE_DELAY_MS: 300,
  QUIET_HOURS_START: 22,
  QUIET_HOURS_END: 7,
  SIGNALS: {
    "27": "CIS / peer support required",
    "83": "deceased person",
    "55": "hostile act / serious violence situation",
    "40": "urgent police attendance required",
    "56": "police attendance required, non-urgent"
  }
};

const STORAGE_KEYS = {
  PROFILE: "turnout_profile_v206",
  REPORTS: "turnout_reports_v206",
  DRAFT: "turnout_draft_v206"
};

const MEMBER_FILES = {
  CONN: "CONN.members.json",
  GROV: "GROV.members.json",
  FRES: "FRES.members.json"
};

const KNOWN_PAGED_UNITS = new Set([
  "CONN",
  "GROV",
  "FRES",
  "BARW",
  "P64",
  "P63B",
  "TRQY",
  "STHB1",
  "MTDU",
  "R63",
  "AFPR",
  "MODE"
]);

const DEFAULT_INCIDENT = () => ({
  eventNumber: "F",
  pagerDate: "",
  pagerTime: "",
  brigadeCode: "",
  brigadeRole: "",
  incidentType: "",
  pagerDetails: "",
  actualLocation: "",
  controlName: "",
  firsCode: "",
  brigadesOnScene: [],
  firstAgency: "",
  firstAgencyOther: "",
  weather1: "",
  weather2: "",
  distanceToScene: "",
  comments: "",
  injuryNotes: "",
  hoses: {
    hose64Qty: "0",
    hose38Qty: "0",
    hose25Qty: "0",
    hoseOtherType: ""
  },
  flags: {
    membersBefore: false,
    aar: false,
    hotDebrief: false,
    injuriesManual: false
  },
  signals: []
});

const DEFAULT_RESPONDERS = () => ({
  connewarre: [],
  mtd: [],
  applianceCodes: {
    t1: "",
    t2: "",
    mtd: ""
  }
});

const state = {
  incident: DEFAULT_INCIDENT(),
  responders: DEFAULT_RESPONDERS(),
  agencies: [],
  profile: {
    name: "",
    memberNumber: "",
    contactNumber: "",
    email: "",
    brigade: "Connewarre"
  },
  savedReports: [],
  memberLists: {
    CONN: [],
    GROV: [],
    FRES: []
  },
  ui: {
    currentPage: "incidentPage",
    previewUrl: "",
    restoredDraft: false,
    waitingServiceWorker: null,
    lastDraftSavedAt: "",
    reportGenerated: false,
    hosesOpen: false
  }
};

let draftSaveTimer = null;

document.addEventListener("DOMContentLoaded", init);

async function init() {
  console.log("TURNOUT SHEET VERSION", CONFIG.APP_VERSION, window.location.href);

  setAppVersion();
  populateDistanceDropdown();
  loadProfile();
  loadSavedReports();
  restoreDraftIfPresent();
  normaliseResponderState();
  ensureRows();

  bindStaticEvents();
  bindIncidentEvents();
  bindConnectionEvents();
  registerServiceWorker();

  renderEverything();
  updateConnectionBanner();
  updateDraftMeta("Draft autosave ready.");

  try {
    await loadMemberLists();
    renderResponders();
  } catch (error) {
    console.error("Member list load failed:", error);
  }
}

function el(id) {
  return document.getElementById(id);
}

function text(v) {
  return String(v || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function setAppVersion() {
  const target = el("appVersionText");
  if (target) target.textContent = CONFIG.APP_VERSION;
}

function pagerDateToInputValue(dateText) {
  const value = String(dateText || "").trim();
  const match = value.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!match) return value;
  const [, dd, mm, yyyy] = match;
  return `${yyyy}-${mm}-${dd}`;
}

function inputDateToPagerDate(dateText) {
  const value = String(dateText || "").trim();
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value;
  const [, yyyy, mm, dd] = match;
  return `${dd}-${mm}-${yyyy}`;
}

function populateDistanceDropdown() {
  const select = el("distanceToScene");
  if (!select) return;

  select.innerHTML = `<option value="">Select</option>`;
  for (let i = 0; i <= 20; i += 1) {
    const option = document.createElement("option");
    option.value = `${i} km`;
    option.textContent = `${i} km`;
    select.appendChild(option);
  }

  const extra = document.createElement("option");
  extra.value = "21+ km";
  extra.textContent = "21+ km";
  select.appendChild(extra);
}

async function loadMemberLists() {
  for (const [key, path] of Object.entries(MEMBER_FILES)) {
    try {
      const res = await fetch(path);
      if (!res.ok) throw new Error(path);
      const data = await res.json();
      state.memberLists[key] = Array.isArray(data) ? data : [];
    } catch {
      state.memberLists[key] = [];
    }
  }
}

function loadProfile() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.PROFILE);
    if (raw) state.profile = { ...state.profile, ...JSON.parse(raw) };
  } catch {}
}

function saveProfile() {
  localStorage.setItem(STORAGE_KEYS.PROFILE, JSON.stringify(state.profile));
}

function loadSavedReports() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.REPORTS);
    state.savedReports = raw ? JSON.parse(raw) : [];
  } catch {
    state.savedReports = [];
  }
}

function saveSavedReports() {
  localStorage.setItem(STORAGE_KEYS.REPORTS, JSON.stringify(state.savedReports));
}

function createResponder(group) {
  return {
    id: uid(),
    group,
    brigade: group === "connewarre" ? "CONN" : "",
    name: "",
    phone: "",
    destination: group === "mtd" ? "MTD P/T" : "",
    isDriver: false,
    isCrewLeader: false,
    ba: false,
    injured: false,
    oic: false
  };
}

function resetForNewPagerUpload() {
  state.incident = DEFAULT_INCIDENT();
  state.responders = DEFAULT_RESPONDERS();
  state.agencies = [];
  state.ui.currentPage = "incidentPage";
  state.ui.reportGenerated = false;
  state.ui.hosesOpen = false;

  ensureRows();

  if (el("reportPreview")) el("reportPreview").value = "";
  if (el("validationText")) el("validationText").textContent = "";
  if (el("validationChips")) el("validationChips").innerHTML = "";

  updateReportActionState();
}

function normaliseResponderState() {
  ["connewarre", "mtd"].forEach((groupKey) => {
    state.responders[groupKey] = (state.responders[groupKey] || []).map((person) => {
      const next = { ...createResponder(groupKey), ...person };
      delete next.truckRole;

      if (groupKey === "mtd" && !next.destination) {
        next.destination = "MTD P/T";
      }

      return next;
    });
  });

  if (!state.responders.applianceCodes) {
    state.responders.applianceCodes = { t1: "", t2: "", mtd: "" };
  }
}

function ensureRows() {
  if (!state.responders.connewarre.length) {
    state.responders.connewarre.push(createResponder("connewarre"));
  }
  if (!state.responders.mtd.length) {
    state.responders.mtd.push(createResponder("mtd"));
  }
}

function getAllResponders() {
  return [...state.responders.connewarre, ...state.responders.mtd];
}

function bindStaticEvents() {
  document.querySelectorAll(".tab-btn[data-page]").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.page));
  });

  document.querySelectorAll("[data-next]").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.next));
  });

  document.querySelectorAll("[data-back]").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.back));
  });

  el("openSettingsBtn")?.addEventListener("click", openSettings);
  el("closeSettingsBtn")?.addEventListener("click", closeSettings);
  el("saveProfileBtn")?.addEventListener("click", saveProfileFromUi);

  el("editOicBtn")?.addEventListener("click", () => {
    showPage("respondersPage");
    updateDraftMeta("Review and update OIC in the responders section if needed.");
  });

  el("pagerUpload")?.addEventListener("change", handleImageUpload);
  el("scanBtn")?.addEventListener("click", rerunOcrFromPreview);

  el("toggleHosesBtn")?.addEventListener("click", () => {
    state.ui.hosesOpen = !state.ui.hosesOpen;
    renderHosesPanel();
  });

  el("addSceneBrigadeBtn")?.addEventListener("click", addSceneBrigade);
  el("sceneBrigadeSelect")?.addEventListener("change", () => {
    el("sceneBrigadeOther").classList.toggle("hidden", el("sceneBrigadeSelect").value !== "Other");
  });

  el("firstAgency")?.addEventListener("change", () => {
    el("firstAgencyOther").classList.toggle("hidden", el("firstAgency").value !== "Other");
    state.incident.firstAgency = el("firstAgency").value;
    markReportStale();
    scheduleDraftSave();
    renderEverything();
  });

  el("firstAgencyOther")?.addEventListener("input", () => {
    state.incident.firstAgencyOther = el("firstAgencyOther").value;
    markReportStale();
    scheduleDraftSave();
  });

  el("agencyType")?.addEventListener("change", autoAddAgencyFromSelect);

  el("flagMembersBeforeBtn")?.addEventListener("click", () => toggleFlag("membersBefore", "flagMembersBeforeBtn"));
  el("flagAarBtn")?.addEventListener("click", () => toggleFlag("aar", "flagAarBtn"));
  el("flagHotDebriefBtn")?.addEventListener("click", () => toggleFlag("hotDebrief", "flagHotDebriefBtn"));
  el("flagInjuriesBtn")?.addEventListener("click", toggleManualInjuriesFlag);

  document.querySelectorAll(".signal-btn").forEach((btn) => {
    btn.addEventListener("click", () => toggleSignal(btn.dataset.signalCode));
  });

  el("resetFirsBtn")?.addEventListener("click", () => {
    state.incident.firsCode = "";
    el("firsCode").value = "";
    markReportStale();
    scheduleDraftSave();
    updateReportSummary();
  });

  el("finishBtn")?.addEventListener("click", generateReportFlow);
  el("saveLocalBtn")?.addEventListener("click", () => saveCurrentReport({ dedupe: false, silent: false }));
  el("sendEmailBtn")?.addEventListener("click", sendEmail);
  el("sendSmsBtn")?.addEventListener("click", sendSms);

  el("saveDraftBtnIncident")?.addEventListener("click", forceDraftSave);
  el("saveDraftBtnResponders")?.addEventListener("click", forceDraftSave);
  el("saveDraftBtnSend")?.addEventListener("click", forceDraftSave);

  el("clearDraftBtn")?.addEventListener("click", clearDraft);
  el("reloadAppBtn")?.addEventListener("click", reloadForUpdate);

  el("copyAppLinkBtn")?.addEventListener("click", copyAppLink);
  el("shareAppBtn")?.addEventListener("click", shareAppLink);

  el("t1CodeC1Btn")?.addEventListener("click", () => setApplianceCode("t1", "C1"));
  el("t1CodeC3Btn")?.addEventListener("click", () => setApplianceCode("t1", "C3"));
  el("t1CodeClearBtn")?.addEventListener("click", () => setApplianceCode("t1", ""));

  el("t2CodeC1Btn")?.addEventListener("click", () => setApplianceCode("t2", "C1"));
  el("t2CodeC3Btn")?.addEventListener("click", () => setApplianceCode("t2", "C3"));
  el("t2CodeClearBtn")?.addEventListener("click", () => setApplianceCode("t2", ""));

  el("mtdCodeC1Btn")?.addEventListener("click", () => setApplianceCode("mtd", "C1"));
  el("mtdCodeC3Btn")?.addEventListener("click", () => setApplianceCode("mtd", "C3"));
  el("mtdCodeClearBtn")?.addEventListener("click", () => setApplianceCode("mtd", ""));

  window.addEventListener("beforeunload", () => {
    saveDraft();
  });
}

function bindIncidentEvents() {
  const map = [
    ["eventNumber", "eventNumber"],
    ["pagerDate", "pagerDate"],
    ["pagerTime", "pagerTime"],
    ["brigadeCode", "brigadeCode"],
    ["incidentType", "incidentType"],
    ["pagerDetails", "pagerDetails"],
    ["actualLocation", "actualLocation"],
    ["controlName", "controlName"],
    ["firsCode", "firsCode"],
    ["comments", "comments"],
    ["injuryNotes", "injuryNotes"]
  ];

  map.forEach(([id, key]) => {
    el(id)?.addEventListener("input", (e) => {
      let value = e.target.value;

      if (key === "eventNumber") {
        if (!value.startsWith("F")) value = `F${value.replace(/^F/i, "")}`;
        e.target.value = value;
      }

      if (key === "pagerDate") {
        value = inputDateToPagerDate(value);
      }

      state.incident[key] = value;

      if (key === "brigadeCode") {
        const code = value.trim().toUpperCase();
        state.incident.brigadeRole = code.startsWith("CONN") ? "Primary" : code ? "Support" : "";
        el("brigadeRole").value = state.incident.brigadeRole;
      }

      markReportStale();
      scheduleDraftSave();
      updateReportSummary();
    });
  });

  ["weather1", "weather2", "distanceToScene", "hose64Qty", "hose38Qty", "hose25Qty", "hoseOtherType"].forEach((id) => {
    el(id)?.addEventListener("change", (e) => {
      if (id.startsWith("hose")) {
        state.incident.hoses[id] = e.target.value;
      } else {
        state.incident[id] = e.target.value;
      }

      if (id === "weather1") {
        if (!state.incident.weather1) {
          state.incident.weather2 = "";
          if (el("weather2")) el("weather2").value = "";
        }
        syncWeather2Visibility();
        syncWeather2Options();
      }

      if (id === "weather2") {
        if (!areWeatherOptionsCompatible(state.incident.weather1, state.incident.weather2)) {
          window.alert("That weather combination does not make sense together.");
          state.incident.weather2 = "";
          if (el("weather2")) el("weather2").value = "";
        }
      }

      markReportStale();
      scheduleDraftSave();
      updateReportSummary();
    });
  });
}

function bindConnectionEvents() {
  window.addEventListener("online", updateConnectionBanner);
  window.addEventListener("offline", updateConnectionBanner);
}

function updateConnectionBanner() {
  const banner = el("connectionBanner");
  if (!banner) return;

  if (navigator.onLine) {
    banner.className = "status-banner online centered";
    banner.textContent = "Online";
  } else {
    banner.className = "status-banner offline centered";
    banner.textContent = "Offline";
  }
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;

  window.addEventListener("load", async () => {
    try {
      const reg = await navigator.serviceWorker.register("./service-worker.js");

      if (reg.waiting) {
        state.ui.waitingServiceWorker = reg.waiting;
        showUpdateBanner();
      }

      reg.addEventListener("updatefound", () => {
        const newWorker = reg.installing;
        if (!newWorker) return;

        newWorker.addEventListener("statechange", () => {
          if (newWorker.state === "installed" && navigator.serviceWorker.controller) {
            state.ui.waitingServiceWorker = newWorker;
            showUpdateBanner();
          }
        });
      });

      navigator.serviceWorker.addEventListener("controllerchange", () => {
        window.location.reload();
      });
    } catch (error) {
      console.error("Service worker registration failed:", error);
    }
  });
}

function showUpdateBanner() {
  el("updateBanner")?.classList.remove("hidden");
}

function reloadForUpdate() {
  if (state.ui.waitingServiceWorker) {
    state.ui.waitingServiceWorker.postMessage({ type: "SKIP_WAITING" });
    return;
  }
  window.location.reload();
}

function getWeatherConflicts() {
  return {
    Hot: ["Cold"],
    Cold: ["Hot"],
    Calm: ["Windy", "Very Windy", "Stormy"],
    Windy: ["Calm"],
    "Very Windy": ["Calm"],
    Sunny: ["Overcast", "Cloudy", "Rain", "Showers", "Stormy"],
    Overcast: ["Sunny"],
    Cloudy: ["Sunny"],
    Rain: ["Sunny"],
    Showers: ["Sunny"],
    Stormy: ["Sunny", "Calm"]
  };
}

function areWeatherOptionsCompatible(weather1, weather2) {
  if (!weather1 || !weather2) return true;
  if (weather1 === weather2) return false;

  const conflicts = getWeatherConflicts();
  const blockedByOne = conflicts[weather1] || [];
  const blockedByTwo = conflicts[weather2] || [];

  return !blockedByOne.includes(weather2) && !blockedByTwo.includes(weather1);
}

function syncWeather2Visibility() {
  const wrap = el("weather2Wrap");
  if (!wrap) return;
  wrap.classList.toggle("hidden", !state.incident.weather1);
}

function syncWeather2Options() {
  const select = el("weather2");
  if (!select) return;

  const weather1 = state.incident.weather1;

  Array.from(select.options).forEach((option) => {
    const value = option.value;

    if (!value) {
      option.hidden = false;
      option.disabled = false;
      return;
    }

    const allowed = areWeatherOptionsCompatible(weather1, value);
    option.hidden = weather1 ? !allowed : false;
    option.disabled = weather1 ? !allowed : false;
  });

  if (!areWeatherOptionsCompatible(state.incident.weather1, state.incident.weather2)) {
    state.incident.weather2 = "";
    select.value = "";
  }
}

function markReportStale() {
  state.ui.reportGenerated = false;
  updateReportActionState();
}

function updateReportActionState() {
  const enabled = state.ui.reportGenerated;
  ["sendSmsBtn", "sendEmailBtn", "saveLocalBtn"].forEach((id) => {
    const btn = el(id);
    if (btn) btn.disabled = !enabled;
  });
}

function renderHosesPanel() {
  const panel = el("hosesPanel");
  if (!panel) return;
  panel.classList.toggle("hidden", !state.ui.hosesOpen);
}

async function handleImageUpload(e) {
  const file = e.target.files?.[0];
  if (!file) return;

  resetForNewPagerUpload();

  const dataUrl = await fileToDataUrl(file);
  state.ui.previewUrl = dataUrl;

  el("pagerPreview").src = dataUrl;
  el("pagerPreview").classList.remove("hidden");
  el("pagerPreviewEmpty").classList.add("hidden");
  el("scanStatus").textContent = `Loaded ${file.name}. Reading pager...`;

  renderEverything();
  scheduleDraftSave();

  await runPagerOcr(dataUrl);
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function rerunOcrFromPreview() {
  if (!state.ui.previewUrl) {
    el("scanStatus").textContent = "Upload a pager screenshot first.";
    return;
  }

  await runPagerOcr(state.ui.previewUrl);
}

async function runPagerOcr(imageSrc) {
  if (!window.Tesseract) {
    el("scanStatus").textContent = "OCR library did not load.";
    return;
  }

  try {
    el("scanStatus").textContent = "Reading pager screenshot...";

    const processedImage = await preprocessImageForOcr(imageSrc);

    const result = await Tesseract.recognize(processedImage, "eng", {
      logger: (msg) => {
        if (msg.status === "recognizing text" && typeof msg.progress === "number") {
          el("scanStatus").textContent = `Reading pager screenshot... ${Math.round(msg.progress * 100)}%`;
        }
      }
    });

    const rawText = result?.data?.text || "";
    const normalizedText = normalizeOcrText(rawText);
    const blocks = extractEmergencyBlocks(normalizedText);

    if (!blocks.length) {
      state.incident.pagerDetails = normalizedText || rawText || "";
      renderEverything();
      el("scanStatus").textContent = "OCR read text, but could not confidently isolate the emergency block. Pager text has been placed in Pager Details.";
      return;
    }

    const mergedEvents = mergeEmergencyBlocks(blocks);
    const firstEventNumber = blocks.map((b) => b.eventNumber).find(Boolean);
    const selectedEvent =
      mergedEvents.find((e) => e.eventNumber === firstEventNumber) ||
      mergedEvents[0];

    if (!selectedEvent) {
      state.incident.pagerDetails = normalizedText || rawText || "";
      renderEverything();
      el("scanStatus").textContent = "Pager text found, but no event could be parsed. Raw text has been placed in Pager Details.";
      return;
    }

    populateIncidentFromOcr(selectedEvent);
    renderEverything();

    const missingBits = [];
    if (!state.incident.pagerDate) missingBits.push("date");
    if (!state.incident.brigadesOnScene.length) missingBits.push("brigades");

    if (missingBits.length) {
      el("scanStatus").textContent =
        `OCR complete with minor gaps. Loaded ${selectedEvent.eventNumber || "event"} from ${selectedEvent.blockCount} pager message${selectedEvent.blockCount === 1 ? "" : "s"}. Check ${missingBits.join(" and ")}.`;
    } else {
      el("scanStatus").textContent =
        `OCR complete. Loaded ${selectedEvent.eventNumber || "event"} from ${selectedEvent.blockCount} pager message${selectedEvent.blockCount === 1 ? "" : "s"}.`;
    }
  } catch (error) {
    console.error(error);
    el("scanStatus").textContent = "OCR failed before parsing completed. Upload again or correct fields manually.";
  }
}

async function preprocessImageForOcr(imageSrc) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const scale = 2;
      const canvas = document.createElement("canvas");
      canvas.width = img.width * scale;
      canvas.height = img.height * scale;

      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const data = imageData.data;

      for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];
        const gray = Math.round((r + g + b) / 3);
        const boosted = gray > 140 ? 255 : gray < 70 ? 0 : gray;

        data[i] = boosted;
        data[i + 1] = boosted;
        data[i + 2] = boosted;
      }

      ctx.putImageData(imageData, 0, 0);
      resolve(canvas.toDataURL("image/png"));
    };

    img.onerror = reject;
    img.src = imageSrc;
  });
}

function normalizeOcrText(textValue) {
  return String(textValue || "")
    .replace(/\r/g, "")
    .replace(/[|]/g, "1")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[—–]/g, "-")
    .replace(/EMERGENCV/g, "EMERGENCY")
    .replace(/EMERGENC Y/g, "EMERGENCY")
    .replace(/ALARCI/g, "ALARC1")
    .replace(/STRUCI/g, "STRUC1")
    .replace(/INCII/g, "INCI1")
    .replace(/RESCCI/g, "RESCC1")
.replace(/2&8\s*[\\T]/g, "MT DUNEED ALL")
.replace(/2&8\s*T/g, "MT DUNEED ALL")
.replace(/2&8\s*VT/g, "MT DUNEED ALL")
    .replace(/MT\s*DUNEED\s*AIL/g, "MT DUNEED ALL")
    .replace(/MT\s*DUNEED\s*ALLL/g, "MT DUNEED ALL")
    .replace(/[^\S\n]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .toUpperCase()
    .trim();
}

function extractEmergencyBlocks(normalizedText) {
  const rawLines = normalizedText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const lines = [];

  rawLines.forEach((line) => {
    const splitEmergency = line.replace(/(EMERGENCY)(\d{2}:\d{2}:\d{2})/g, "$1\n$2");
    splitEmergency.split("\n").forEach((part) => {
      const clean = part.trim();
      if (clean) lines.push(clean);
    });
  });

  const blocks = [];
  const isEmergencyLike = (line) =>
    /EMERGEN/.test(line) || (/ALERT /.test(line) && /F\d{6,}/.test(normalizedText));

  for (let i = 0; i < lines.length; i += 1) {
    if (!isEmergencyLike(lines[i])) continue;

    const blockLines = [];
    let foundAlert = false;

    for (let j = i; j < lines.length; j += 1) {
      const line = lines[j];

      if (j > i && /NON-EMERGENCY|ADMIN/.test(line)) break;
      if (j > i && /EMERGEN/.test(line) && blockLines.length >= 4) break;

      if (/RESPOND|REVIEW/.test(line)) continue;
      if (/SINCE ALERT/.test(line)) continue;

      blockLines.push(line);

      if (line.includes("ALERT ")) foundAlert = true;
      if (/F\d{6,}/.test(line) && foundAlert) break;
      if (blockLines.length >= 10 && foundAlert) break;
    }

    const parsed = parseEmergencyBlock(blockLines);
    if (parsed.eventNumber || parsed.alertLine || parsed.rawText.includes("ALERT ")) {
      blocks.push(parsed);
    }
  }

  return blocks;
}

function parseEmergencyBlock(lines) {
  const blockText = lines.join("\n");
  const flatText = lines.join(" ");

  const dateTimeMatch = flatText.match(/(\d{2}:\d{2}:\d{2})\s+(\d{2}-\d{2}-\d{4})/);
  const pagerTime = dateTimeMatch?.[1] || "";
  const pagerDate = dateTimeMatch?.[2] || "";

  const eventMatch = flatText.match(/(F\d{6,})\b/g);
  const eventNumber = eventMatch ? eventMatch[eventMatch.length - 1] : "";

  const alertStartIndex = lines.findIndex((line) => line.includes("ALERT "));
  const alertAndBody = alertStartIndex >= 0 ? lines.slice(alertStartIndex).join(" ") : flatText;

  const alertLineMatch = alertAndBody.match(/ALERT\s+([A-Z0-9]+)\s+([A-Z]{4})C([13])\b/);
  const brigadeCode = alertLineMatch?.[1] || "";
  const incidentType = alertLineMatch?.[2] || "";

  const actualLocation = extractActualLocation(alertAndBody);
  const units = extractPagedUnits(flatText, eventNumber);

  return {
    rawText: blockText,
    lines,
    pagerDate,
    pagerTime,
    eventNumber,
    brigadeCode,
    incidentType,
    actualLocation,
    units,
    alertLine: alertLineMatch?.[0] || ""
  };
}

function extractActualLocation(textValue) {
  const noEvent = textValue.replace(/F\d{6,}.*/g, " ");
  const locationMatch = noEvent.match(
    /\b(\d+\s+[A-Z0-9 ]+?(?:ST|RD|AV|AVE|DR|BVD|BLVD|HWY|CT|CRT|CRES|PL|WAY|LANE|LN)\s+[A-Z ]+?)(?=\s*(?:\/|\/\/|\sJ[A-Z]|\s+M\s+\d+))/i
  );

  if (locationMatch?.[1]) {
    return cleanLocation(locationMatch[1]);
  }

  const roughLineMatch = noEvent.match(/\b\d+\s+[A-Z0-9 ].+?(?=\s+M\s+\d+|F\d{6,}|$)/i);
  return cleanLocation(roughLineMatch?.[0] || "");
}

function cleanLocation(value) {
  let cleaned = String(value || "")
    .replace(/\s+/g, " ")
    .trim();

  cleaned = cleaned
    .replace(/\s\/\/\s*/g, " / ")
    .replace(/\s\/\s*/g, " / ")
    .replace(/\sJ(?=[A-Z]{2,})/g, " /");

  const separatorMatch = cleaned.match(/\s\/|\/\/|\sJ(?=[A-Z]{2,})/);
  if (separatorMatch) {
    cleaned = cleaned.slice(0, separatorMatch.index).trim();
  }

  return cleaned;
}

function extractPagedUnits(textValue, eventNumber) {
  let working = String(textValue || "");

  if (eventNumber) {
    working = working.replace(eventNumber, " ");
  }

 function extractPagedUnits(textValue, eventNumber) {
  let working = String(textValue || "");

  if (eventNumber) {
    working = working.replace(eventNumber, " ");
  }

  const tokens = working
    .split(/\s+/)
    .map((token) => token.replace(/[^A-Z0-9]/g, ""))
    .filter(Boolean);

  const ignore = new Set([
    "EMERGENCY",
    "ALERT",
    "SINCE",
    "ALERTS",
    "SINCEALERT",
    "RESPOND",
    "REVIEW",
    "OTHER",
    "ATTENDING",
    "NONEMERGENCY",
    "ADMIN",
    "FP",
    "F",
    "AFPRALL",
    "CONNEWARRE",
    "BRIGADE",
    "ALL",
    "MT",
    "DUNEED"
  ]);

  const units = [];

  tokens.forEach((token) => {
    if (ignore.has(token)) return;
    if (/^\d+$/.test(token)) return;
    if (/^F\d{6,}$/.test(token)) return;
    if (/^[A-Z]{4}C[13]$/.test(token)) return;
    if (/^ALARC[13]$/.test(token)) return;
    if (/^STRUC[13]$/.test(token)) return;
    if (/^INCI[13]$/.test(token)) return;
    if (/^RESCC[13]$/.test(token)) return;

    let normalized = token;

    if (token.startsWith("C") && KNOWN_PAGED_UNITS.has(token.slice(1))) {
      normalized = token.slice(1);
    }

    if (!KNOWN_PAGED_UNITS.has(normalized)) return;
    if (!units.includes(normalized)) units.push(normalized);
  });

  return units;
}

  const units = [];

  tokens.forEach((token) => {
    if (ignore.has(token)) return;
    if (/^\d+$/.test(token)) return;
    if (/^\d{2}:\d{2}:\d{2}$/.test(token)) return;
    if (/^\d{2}-\d{2}-\d{4}$/.test(token)) return;
    if (/^F\d{6,}$/.test(token)) return;
    if (/^[A-Z]{4}C[13]$/.test(token)) return;
    if (/^ALARC[13]$/.test(token)) return;
    if (/^STRUC[13]$/.test(token)) return;
    if (/^INCI[13]$/.test(token)) return;
    if (/^RESCC[13]$/.test(token)) return;

    let normalized = token;

    if (/^C[A-Z0-9]{4,6}$/.test(token)) {
      normalized = token.slice(1);
    }

    const isOperationalCode =
      /^(?:CONN|GROV|FRES|BARW|TRQY|MTDU|P63B|P64|P\d+[A-Z]?|R\d+|STHB\d+|SES\d*|[A-Z]{4,6})$/.test(normalized);

    if (!isOperationalCode) return;
    if (!units.includes(normalized)) units.push(normalized);
  });

  return units;
}

function mergeEmergencyBlocks(blocks) {
  const byEvent = new Map();

  blocks.forEach((block) => {
    const key = block.eventNumber || `NO_EVENT_${block.pagerDate}_${block.pagerTime}`;

    if (!byEvent.has(key)) {
      byEvent.set(key, {
        eventNumber: block.eventNumber,
        pagerDate: block.pagerDate,
        pagerTime: block.pagerTime,
        brigadeCode: block.brigadeCode,
        incidentType: block.incidentType,
        actualLocation: block.actualLocation,
        units: [...block.units],
        earliestRawText: block.rawText,
        blockCount: 1
      });
      return;
    }

    const existing = byEvent.get(key);

    if (isEarlierPagerTime(block.pagerDate, block.pagerTime, existing.pagerDate, existing.pagerTime)) {
      existing.pagerDate = block.pagerDate || existing.pagerDate;
      existing.pagerTime = block.pagerTime || existing.pagerTime;
      existing.earliestRawText = block.rawText || existing.earliestRawText;
    }

    existing.brigadeCode = existing.brigadeCode || block.brigadeCode;
    existing.incidentType = existing.incidentType || block.incidentType;
    existing.actualLocation = existing.actualLocation || block.actualLocation;

    block.units.forEach((unit) => {
      if (!existing.units.includes(unit)) existing.units.push(unit);
    });

    existing.blockCount += 1;
  });

  return Array.from(byEvent.values()).map((event) => ({
    ...event,
    pagerDetails: cleanPagerDetails(event.earliestRawText || "")
  }));
}

function cleanPagerDetails(rawText) {
  const lines = String(rawText || "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const startIndex = lines.findIndex((line) => line.includes("EMERGENCY"));
  const usefulLines = startIndex >= 0 ? lines.slice(startIndex) : lines;

  return usefulLines
    .filter((line) => !line.includes("RESPOND"))
    .filter((line) => !line.includes("REVIEW"))
    .filter((line) => !line.includes("SINCE ALERT"))
    .filter((line) => !line.includes("NON-EMERGENCY"))
    .filter((line) => !line.includes("ADMIN"))
    .join("\n")
    .trim();
}

function isEarlierPagerTime(dateA, timeA, dateB, timeB) {
  const a = parsePagerDateTime(dateA, timeA);
  const b = parsePagerDateTime(dateB, timeB);

  if (!a && !b) return false;
  if (a && !b) return true;
  if (!a && b) return false;

  return a.getTime() < b.getTime();
}

function parsePagerDateTime(dateText, timeText) {
  if (!dateText || !timeText) return null;

  const parts = dateText.split("-");
  if (parts.length !== 3) return null;

  const [dd, mm, yyyy] = parts;
  const iso = `${yyyy}-${mm}-${dd}T${timeText}`;
  const date = new Date(iso);
  return Number.isNaN(date.getTime()) ? null : date;
}

function populateIncidentFromOcr(parsed) {
  state.incident.eventNumber = parsed.eventNumber || state.incident.eventNumber || "F";
  state.incident.pagerDate = parsed.pagerDate || state.incident.pagerDate;
  state.incident.pagerTime = parsed.pagerTime || state.incident.pagerTime;
  state.incident.brigadeCode = parsed.brigadeCode || state.incident.brigadeCode;
  state.incident.brigadeRole = state.incident.brigadeCode.startsWith("CONN")
    ? "Primary"
    : state.incident.brigadeCode
      ? "Support"
      : "";
  state.incident.incidentType = parsed.incidentType || state.incident.incidentType;
  state.incident.pagerDetails = parsed.pagerDetails || state.incident.pagerDetails;
  state.incident.actualLocation = parsed.actualLocation || state.incident.actualLocation;
  state.incident.brigadesOnScene = parsed.units?.length ? parsed.units : state.incident.brigadesOnScene;

  if (el("pagerDate")) el("pagerDate").value = pagerDateToInputValue(state.incident.pagerDate);
  if (el("pagerTime")) el("pagerTime").value = state.incident.pagerTime;
  if (el("brigadeCode")) el("brigadeCode").value = state.incident.brigadeCode;
  if (el("brigadeRole")) el("brigadeRole").value = state.incident.brigadeRole;
  if (el("incidentType")) el("incidentType").value = state.incident.incidentType;
  if (el("pagerDetails")) el("pagerDetails").value = state.incident.pagerDetails;
  if (el("actualLocation")) el("actualLocation").value = state.incident.actualLocation;
  if (el("eventNumber")) el("eventNumber").value = state.incident.eventNumber;

  markReportStale();
  scheduleDraftSave();
}

function addSceneBrigade() {
  let value = el("sceneBrigadeSelect").value;
  if (!value) return;

  if (value === "Other") value = el("sceneBrigadeOther").value.trim().toUpperCase();
  if (!value) return;

  if (!state.incident.brigadesOnScene.includes(value)) {
    state.incident.brigadesOnScene.push(value);
  }

  el("sceneBrigadeSelect").value = "";
  el("sceneBrigadeOther").value = "";
  el("sceneBrigadeOther").classList.add("hidden");

  markReportStale();
  renderSceneBrigades();
  scheduleDraftSave();
  updateReportSummary();
}

function renderSceneBrigades() {
  const wrap = el("sceneBrigadeChips");
  if (!wrap) return;

  wrap.innerHTML = "";

  state.incident.brigadesOnScene.forEach((code) => {
    const chip = document.createElement("div");
    chip.className = "scene-chip";
    chip.innerHTML = `<span>${text(code)}</span><button type="button" aria-label="Remove ${text(code)}">×</button>`;
    chip.querySelector("button").addEventListener("click", () => {
      state.incident.brigadesOnScene = state.incident.brigadesOnScene.filter((x) => x !== code);
      markReportStale();
      renderSceneBrigades();
      scheduleDraftSave();
      updateReportSummary();
    });
    wrap.appendChild(chip);
  });
}

function autoAddAgencyFromSelect() {
  const select = el("agencyType");
  const type = select.value;
  if (!type) return;

  state.agencies.push({
    id: uid(),
    type,
    otherName: "",
    officerName: "",
    contactNumber: "",
    station: "",
    badgeNumber: "",
    comments: ""
  });

  select.value = "";
  markReportStale();
  renderAgencies();
  scheduleDraftSave();
  updateReportSummary();
}

function renderAgencies() {
  const wrap = el("agencyBlocks");
  if (!wrap) return;

  wrap.innerHTML = "";

  state.agencies.forEach((agency) => {
    const block = document.createElement("div");
    block.className = "agency-block";
    block.innerHTML = `
      <div class="row between">
        <strong>${text(agency.type)}</strong>
        <button class="tiny-btn" type="button">Remove</button>
      </div>
      <div class="grid">
        ${agency.type === "Other" ? `<label>Other Agency Name<input data-field="otherName" type="text" value="${text(agency.otherName)}"></label>` : ""}
        <label>Officer Name<input data-field="officerName" type="text" value="${text(agency.officerName)}"></label>
        <label>Contact Number<input data-field="contactNumber" type="text" inputmode="tel" value="${text(agency.contactNumber)}"></label>
        <label>Station<input data-field="station" type="text" value="${text(agency.station)}"></label>
        <label>Badge Number<input data-field="badgeNumber" type="text" inputmode="numeric" value="${text(agency.badgeNumber)}"></label>
        <label class="full">Comments<textarea data-field="comments" rows="2">${text(agency.comments)}</textarea></label>
      </div>
    `;

    block.querySelector("button").addEventListener("click", () => {
      const confirmed = window.confirm(`Remove agency entry for ${agency.type}?`);
      if (!confirmed) return;
      state.agencies = state.agencies.filter((a) => a.id !== agency.id);
      markReportStale();
      renderAgencies();
      scheduleDraftSave();
      updateReportSummary();
    });

    block.querySelectorAll("[data-field]").forEach((field) => {
      field.addEventListener("input", (e) => {
        agency[e.target.dataset.field] = e.target.value;
        markReportStale();
        scheduleDraftSave();
        updateReportSummary();
      });
    });

    wrap.appendChild(block);
  });
}

function toggleFlag(key, buttonId) {
  state.incident.flags[key] = !state.incident.flags[key];
  el(buttonId)?.classList.toggle("active", state.incident.flags[key]);
  markReportStale();
  scheduleDraftSave();
  updateReportSummary();
}

function hasResponderInjury() {
  return getAllResponders().some((r) => r.injured);
}

function isInjuriesFlagActive() {
  return state.incident.flags.injuriesManual || hasResponderInjury();
}

function toggleManualInjuriesFlag() {
  state.incident.flags.injuriesManual = !state.incident.flags.injuriesManual;
  syncInjuriesFlagButton();
  markReportStale();
  scheduleDraftSave();
  updateReportSummary();
}

function syncInjuriesFlagButton() {
  el("flagInjuriesBtn")?.classList.toggle("active", isInjuriesFlagActive());
}

function toggleSignal(code) {
  const list = state.incident.signals;
  if (list.includes(code)) {
    state.incident.signals = list.filter((x) => x !== code);
  } else {
    state.incident.signals.push(code);
  }

  syncSignalButtons();
  markReportStale();
  scheduleDraftSave();
  updateReportSummary();
}

function syncSignalButtons() {
  document.querySelectorAll(".signal-btn").forEach((btn) => {
    btn.classList.toggle("active", state.incident.signals.includes(btn.dataset.signalCode));
  });
}

function resolveResponderMemberDetails(groupKey, person) {
  const enteredName = String(person.name || "").trim().toUpperCase();

  if (!enteredName) {
    person.brigade = groupKey === "connewarre" ? "CONN" : "";
    person.phone = "";
    return;
  }

  if (groupKey === "connewarre") {
    person.brigade = "CONN";
    const found = state.memberLists.CONN.find((m) => {
      const name = (typeof m === "string" ? m : m.name || "").trim().toUpperCase();
      return name === enteredName;
    });
    person.phone = found && typeof found !== "string" ? found.phone || "" : "";
    return;
  }

  const found = findMemberAcrossBrigades(person.name);
  person.brigade = found?.brigade || "";
  person.phone = found?.phone || "";
}

function updateResponderCardDisplay(card, person, groupKey) {
  const badge = card.querySelector(".badge");
  if (badge) {
    badge.textContent = person.brigade || (groupKey === "connewarre" ? "CONN" : "");
  }

  card.querySelectorAll(".responder-destinations .chip-btn").forEach((btn) => {
    btn.disabled = !person.name.trim();
  });

  card.querySelectorAll(".responder-flags .chip-btn").forEach((btn) => {
    btn.disabled = !person.name.trim();
  });
}

function updateOicBanner() {
  const oic = getAllResponders().find((r) => r.oic && r.name.trim());

  if (!oic) {
    el("oicBanner").textContent = "APPOINT OIC";
    el("oicBanner").classList.add("missing");
    return;
  }

  el("oicBanner").textContent = `OIC: ${oic.name}${oic.phone ? " – " + oic.phone : ""}`;
  el("oicBanner").classList.remove("missing");
}

function renderResponders() {
  renderResponderGroup("connewarre", "connewarreList", ["T1", "T2", "Station", "Direct"]);
  renderResponderGroup("mtd", "mtdList", ["MTD P/T", "Station", "Direct"]);
  updateOicBanner();
  syncInjuriesFlagButton();
  updateReportSummary();
}

function renderResponderGroup(groupKey, containerId, destinations) {
  const wrap = el(containerId);
  if (!wrap) return;

  wrap.innerHTML = "";

  state.responders[groupKey].forEach((person, index) => {
    const card = document.createElement("div");
    card.className = "responder-card";
    card.dataset.id = person.id;
    card.dataset.group = groupKey;

    const listId = `list_${groupKey}_${person.id}`;
    const options = groupKey === "connewarre"
      ? state.memberLists.CONN
      : [...state.memberLists.CONN, ...state.memberLists.GROV, ...state.memberLists.FRES];

    const optionHtml = options.map((m) => {
      const name = typeof m === "string" ? m : m.name;
      return `<option value="${text(name)}"></option>`;
    }).join("");

    const showTruckRole = person.destination === "T1" || person.destination === "T2" || person.destination === "MTD P/T";

    card.innerHTML = `
      <div class="responder-card-top">
        <div>
          <input type="text" list="${listId}" value="${text(person.name)}" placeholder="Name" autocomplete="off" />
          <datalist id="${listId}">${optionHtml}</datalist>
        </div>
        <div class="badge">${text(person.brigade || (groupKey === "connewarre" ? "CONN" : ""))}</div>
        <button class="tiny-btn" type="button">Remove</button>
      </div>

      <div class="responder-stage">
        <div class="stage-label">Response</div>
        <div class="chips responder-destinations"></div>
      </div>

      ${showTruckRole ? `
        <div class="responder-stage">
          <div class="stage-label">Truck role</div>
          <div class="chips responder-roles"></div>
        </div>
      ` : ""}

      <div class="responder-stage">
        <div class="stage-label">Flags</div>
        <div class="chips responder-flags"></div>
      </div>
    `;

    const nameInput = card.querySelector("input");
    nameInput.addEventListener("input", (e) => {
      person.name = e.target.value;
      resolveResponderMemberDetails(groupKey, person);

      if (groupKey === "mtd" && !person.destination) {
        person.destination = "MTD P/T";
      }

      updateResponderCardDisplay(card, person, groupKey);
      updateOicBanner();
      updateReportSummary();
      markReportStale();
      scheduleDraftSave();
    });

    card.querySelector(".tiny-btn").addEventListener("click", () => {
      const confirmed = window.confirm(`Remove responder "${person.name || "Unnamed member"}"?`);
      if (!confirmed) return;

      state.responders[groupKey] = state.responders[groupKey].filter((x) => x.id !== person.id);
      if (!state.responders[groupKey].length) state.responders[groupKey].push(createResponder(groupKey));
      markReportStale();
      renderResponders();
      scheduleDraftSave();
    });

    const destWrap = card.querySelector(".responder-destinations");
    destinations.forEach((dest) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `chip-btn ${person.destination === dest ? "active" : ""}`;
      btn.textContent = dest;
      btn.disabled = !person.name.trim();

      btn.addEventListener("click", () => {
        person.destination = dest;

        if (!(dest === "T1" || dest === "T2" || dest === "MTD P/T")) {
          person.isDriver = false;
          person.isCrewLeader = false;
        }

        markReportStale();
        renderResponders();
        scheduleDraftSave();
      });

      destWrap.appendChild(btn);
    });

    if (showTruckRole) {
      const roleWrap = card.querySelector(".responder-roles");

      const driverBtn = document.createElement("button");
      driverBtn.type = "button";
      driverBtn.className = `chip-btn ${person.isDriver ? "active" : ""}`;
      driverBtn.textContent = "Driver";
      driverBtn.addEventListener("click", () => {
        if (!person.isDriver && isRoleTaken(person.destination, "driver", person.id)) {
          window.alert("A driver has already been appointed for that appliance.");
          return;
        }

        person.isDriver = !person.isDriver;
        markReportStale();
        renderResponders();
        scheduleDraftSave();
      });
      roleWrap.appendChild(driverBtn);

      const clBtn = document.createElement("button");
      clBtn.type = "button";
      clBtn.className = `chip-btn ${person.isCrewLeader ? "active" : ""}`;
      clBtn.textContent = "CL";
      clBtn.addEventListener("click", () => {
        if (!person.isCrewLeader && isRoleTaken(person.destination, "crewLeader", person.id)) {
          window.alert("A crew leader has already been appointed for that appliance.");
          return;
        }

        person.isCrewLeader = !person.isCrewLeader;
        markReportStale();
        renderResponders();
        scheduleDraftSave();
      });
      roleWrap.appendChild(clBtn);
    }

    const flagsWrap = card.querySelector(".responder-flags");
    [
      ["BA", "ba"],
      ["Injured", "injured"],
      ["OIC", "oic"]
    ].forEach(([label, key]) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `chip-btn ${person[key] ? "active" : ""}`;
      btn.textContent = label;
      btn.disabled = !person.name.trim();

      btn.addEventListener("click", () => {
        if (!person.name.trim()) return;

        if (key === "oic") {
          const currentOic = getAllResponders().find((r) => r.oic && r.id !== person.id && r.name.trim());

          if (!person.oic && currentOic) {
            const confirmed = window.confirm(`OIC is currently ${currentOic.name}. Replace with ${person.name}?`);
            if (!confirmed) return;
          }

          const willBeOic = !person.oic;
          clearAllOic();
          person.oic = willBeOic;
          markReportStale();
          renderResponders();
          scheduleDraftSave();
          return;
        }

        person[key] = !person[key];
        if (key === "injured") syncInjuriesFlagButton();
        btn.classList.toggle("active", person[key]);
        updateOicBanner();
        updateReportSummary();
        markReportStale();
        scheduleDraftSave();
      });

      flagsWrap.appendChild(btn);
    });

    wrap.appendChild(card);
    updateResponderCardDisplay(card, person, groupKey);

    const list = state.responders[groupKey];
    const isLast = index === list.length - 1;
    if (isLast && person.name.trim() && person.destination) {
      const addBtn = document.createElement("button");
      addBtn.type = "button";
      addBtn.className = "secondary-btn";
      addBtn.textContent = "Add Member";
      addBtn.addEventListener("click", () => {
        state.responders[groupKey].push(createResponder(groupKey));
        markReportStale();
        renderResponders();
        scheduleDraftSave();
      });
      wrap.appendChild(addBtn);
    }
  });
}

function findMemberAcrossBrigades(name) {
  const target = String(name || "").trim().toUpperCase();
  for (const key of ["CONN", "GROV", "FRES"]) {
    const found = state.memberLists[key].find((m) => {
      const n = (typeof m === "string" ? m : m.name).toUpperCase();
      return n === target;
    });
    if (found) {
      return {
        brigade: key,
        phone: typeof found === "string" ? "" : found.phone || ""
      };
    }
  }
  return null;
}

function isRoleTaken(destination, roleKey, currentId) {
  return getAllResponders().some((r) => {
    if (r.id === currentId) return false;
    if (r.destination !== destination) return false;
    if (roleKey === "driver") return Boolean(r.isDriver);
    if (roleKey === "crewLeader") return Boolean(r.isCrewLeader);
    return false;
  });
}

function clearAllOic() {
  getAllResponders().forEach((r) => {
    r.oic = false;
  });
}

function formatResponderName(r) {
  if (!r.name.trim()) return "";
  if (r.brigade === "CONN" || !r.brigade) return r.name.trim();
  return `${r.name.trim()} (${r.brigade})`;
}

function formatResponderFlags(r) {
  const parts = [];
  if (r.isDriver) parts.push("Driver");
  if (r.isCrewLeader) parts.push("CL");
  if (r.oic) parts.push("OIC");
  if (r.ba) parts.push("BA");
  return parts.join(" ");
}

function setApplianceCode(applianceKey, value) {
  state.responders.applianceCodes[applianceKey] = value;
  syncApplianceCodeButtons();
  renderResponders();
  markReportStale();
  scheduleDraftSave();
  updateReportSummary();
}

function syncApplianceCodeButtons() {
  const configs = [
    ["t1", "t1CodeC1Btn", "C1"],
    ["t1", "t1CodeC3Btn", "C3"],
    ["t2", "t2CodeC1Btn", "C1"],
    ["t2", "t2CodeC3Btn", "C3"],
    ["mtd", "mtdCodeC1Btn", "C1"],
    ["mtd", "mtdCodeC3Btn", "C3"]
  ];

  configs.forEach(([key, id, code]) => {
    const button = el(id);
    if (!button) return;
    if (state.responders.applianceCodes[key] === code) {
      button.classList.add("active");
    } else {
      button.classList.remove("active");
    }
  });

  const clearConfigs = [
    ["t1", "t1CodeClearBtn"],
    ["t2", "t2CodeClearBtn"],
    ["mtd", "mtdCodeClearBtn"]
  ];

  clearConfigs.forEach(([key, id]) => {
    const button = el(id);
    if (!button) return;
    if (!state.responders.applianceCodes[key]) {
      button.classList.add("active");
    } else {
      button.classList.remove("active");
    }
  });
}

function buildPagerDateTime() {
  const date = state.incident.pagerDate || "";
  const time = state.incident.pagerTime || "";
  return [date, time].filter(Boolean).join(" ");
}

function buildApplianceCodeLine() {
  const parts = [];

  if (state.responders.applianceCodes.t1) {
    parts.push(`Conn T1 - Code ${state.responders.applianceCodes.t1.replace("C", "")}`);
  }

  if (state.responders.applianceCodes.t2) {
    parts.push(`Conn T2 - Code ${state.responders.applianceCodes.t2.replace("C", "")}`);
  }

  if (state.responders.applianceCodes.mtd) {
    parts.push(`MTD P/T - Code ${state.responders.applianceCodes.mtd.replace("C", "")}`);
  }

  return parts.join(", ");
}

function buildWeatherLine() {
  const a = state.incident.weather1;
  const b = state.incident.weather2;
  if (a && b) return `${a} and ${b}`;
  return a || b || "";
}

function buildHosesLine() {
  const parts = [];
  const qty64 = Number(state.incident.hoses.hose64Qty || 0);
  const qty38 = Number(state.incident.hoses.hose38Qty || 0);
  const qty25 = Number(state.incident.hoses.hose25Qty || 0);
  const other = state.incident.hoses.hoseOtherType || "";

  if (qty64 > 0) parts.push(`${qty64}x64`);
  if (qty38 > 0) parts.push(`${qty38}x38`);
  if (qty25 > 0) parts.push(`${qty25}x25`);
  if (other) parts.push(other);

  return parts.join(", ");
}

function buildControlLine() {
  const name = (state.incident.controlName || "").trim();
  const role = (state.incident.brigadeRole || "").trim();
  if (!name && !role) return "";

  const controlText = name
    ? `${name}${name.toUpperCase().endsWith("CONTROL") ? "" : " Control"}`
    : "Control";

  return role ? `${controlText} | ${role}` : controlText;
}

function getSupportTargetFromBrigadeCode() {
  const code = String(state.incident.brigadeCode || "").trim().toUpperCase();

  const targets = ["GROV", "FRES", "BARW", "TRQY", "MTDU"];
  return targets.find((target) => code.startsWith(target)) || "";
}

function buildOperationalStatusLine() {
  if (state.incident.brigadeRole === "Primary") return "Primary";

  if (state.incident.brigadeRole === "Support") {
    const target = getSupportTargetFromBrigadeCode();
    return target ? `Support to ${target}` : "Support";
  }

  return "";
}
  const name = (state.incident.controlName || "").trim();
  const role = (state.incident.brigadeRole || "").trim();
  if (!name && !role) return "";

  const controlText = name
    ? `${name}${name.toUpperCase().endsWith("CONTROL") ? "" : " Control"}`
    : "Control";

  return role ? `${controlText} | ${role}` : controlText;
}

function buildSignalsLine() {
  if (!state.incident.signals.length) return "";
  return state.incident.signals.map((code) => `${code} = ${CONFIG.SIGNALS[code]}`).join("; ");
}

function buildRespondersSection(lines) {
  const groups = [
    ["CONN T1", "T1"],
    ["CONN T2", "T2"],
    ["MTD P/T", "MTD P/T"],
    ["Direct Responders", "Direct"],
    ["Station Responders", "Station"]
  ];

  const all = getAllResponders();
  const hasAny = groups.some(([, code]) => all.some((r) => r.destination === code && r.name.trim()));
  if (!hasAny) return;

  lines.push("");
  lines.push("Responders");

  groups.forEach(([title, code]) => {
    const crew = all.filter((r) => r.destination === code && r.name.trim());
    if (!crew.length) return;

    lines.push(`${title}:`);
    crew.forEach((r) => {
      const base = formatResponderName(r);
      const flags = formatResponderFlags(r);
      lines.push(`- ${base}${flags ? ` ${flags}` : ""}`);
    });
  });
}

function buildReport() {
  const lines = [];

  pushLine(lines, "Event Number", state.incident.eventNumber);
  if (lines.length) lines.push("");

 const operationalStatus = buildOperationalStatusLine();
if (operationalStatus) {
  lines.push(operationalStatus);
  lines.push("");
}
  
  pushLine(lines, "Pager Date / Time", buildPagerDateTime());
  pushLine(lines, "Appliances", buildApplianceCodeLine());

  if (state.incident.pagerDetails.trim()) {
    lines.push("");
    lines.push("Pager Details");
    lines.push(state.incident.pagerDetails.trim());
  }

  pushLine(lines, "Actual Location", state.incident.actualLocation);
  pushLine(lines, "Control Name", buildControlLine());

  const firstAgency = state.incident.firstAgency === "Other"
    ? state.incident.firstAgencyOther
    : state.incident.firstAgency;

  pushLine(lines, "First Agency On Scene", firstAgency);
  pushLine(lines, "Brigades On Scene", state.incident.brigadesOnScene.join(", "));
  pushLine(lines, "Weather", buildWeatherLine());
  pushLine(lines, "Distance to Scene", state.incident.distanceToScene);
  pushLine(lines, "Hoses Used", buildHosesLine());

  if (state.agencies.length) {
    lines.push("");
    lines.push("Agency Details");

    state.agencies.forEach((a, i) => {
      lines.push(`${i + 1}. ${a.type === "Other" ? (a.otherName || "Other") : a.type}`);
      pushLine(lines, "Officer", a.officerName);
      pushLine(lines, "Contact", a.contactNumber);
      pushLine(lines, "Station", a.station);
      pushLine(lines, "Badge", a.badgeNumber);
      pushLine(lines, "Comments", a.comments);
    });
  }

  pushLine(lines, "FIRS Code", state.incident.firsCode);

  buildRespondersSection(lines);

  if (state.incident.comments.trim()) {
    lines.push("");
    lines.push("Comments");
    lines.push(state.incident.comments.trim());
  }

  const flags = [];

  if (state.incident.flags.membersBefore) flags.push("Members direct before 1st appliance");
  if (state.incident.flags.aar) flags.push("AAR required");
  if (state.incident.flags.hotDebrief) flags.push("Hot debrief conducted");
  if (isInjuriesFlagActive()) flags.push("Injuries");

  const signalsLine = buildSignalsLine();

  if (flags.length || state.incident.injuryNotes.trim() || signalsLine) {
    lines.push("");
    lines.push("Incident Flags");

    flags.forEach((f) => lines.push(`- ${f}`));

    if (state.incident.injuryNotes.trim()) {
      lines.push(`Injury Notes: ${state.incident.injuryNotes.trim()}`);
    }

    if (signalsLine) {
      lines.push(`Signal: ${signalsLine}`);
    }
  }

  lines.push("");
  lines.push("Created by");

  pushLine(lines, "Name", state.profile.name);
  pushLine(lines, "Brigade", state.profile.brigade);
  pushLine(lines, "CFA Member Number", state.profile.memberNumber);
  pushLine(lines, "Contact Number", state.profile.contactNumber);

  return lines.join("\n");
}

const firstAgency = state.incident.firstAgency === "Other"
    ? state.incident.firstAgencyOther
    : state.incident.firstAgency;
  pushLine(lines, "First Agency On Scene", firstAgency);
  pushLine(lines, "Brigades On Scene", state.incident.brigadesOnScene.join(", "));
  pushLine(lines, "Weather", buildWeatherLine());
  pushLine(lines, "Distance to Scene", state.incident.distanceToScene);
  pushLine(lines, "Hoses Used", buildHosesLine());

  if (state.agencies.length) {
    lines.push("");
    lines.push("Agency Details");
    state.agencies.forEach((a, i) => {
      lines.push(`${i + 1}. ${a.type === "Other" ? (a.otherName || "Other") : a.type}`);
      pushLine(lines, "Officer", a.officerName);
      pushLine(lines, "Contact", a.contactNumber);
      pushLine(lines, "Station", a.station);
      pushLine(lines, "Badge", a.badgeNumber);
      pushLine(lines, "Comments", a.comments);
    });
  }

  pushLine(lines, "FIRS Code", state.incident.firsCode);
  buildRespondersSection(lines);

  if (state.incident.comments.trim()) {
    lines.push("");
    lines.push("Comments");
    lines.push(state.incident.comments.trim());
  }

  const flags = [];
  if (state.incident.flags.membersBefore) flags.push("Members direct before 1st appliance");
  if (state.incident.flags.aar) flags.push("AAR required");
  if (state.incident.flags.hotDebrief) flags.push("Hot debrief conducted");
  if (isInjuriesFlagActive()) flags.push("Injuries");

  const signalsLine = buildSignalsLine();

  if (flags.length || state.incident.injuryNotes.trim() || signalsLine) {
    lines.push("");
    lines.push("Incident Flags");
    flags.forEach((f) => lines.push(`- ${f}`));
    if (state.incident.injuryNotes.trim()) lines.push(`Injury Notes: ${state.incident.injuryNotes.trim()}`);
    if (signalsLine) lines.push(`Signal: ${signalsLine}`);
  }

  lines.push("");
  lines.push("Created by");
  pushLine(lines, "Name", state.profile.name);
  pushLine(lines, "Brigade", state.profile.brigade);
  pushLine(lines, "CFA Member Number", state.profile.memberNumber);
  pushLine(lines, "Contact Number", state.profile.contactNumber);

  return lines.join("\n");
}

function pushLine(lines, label, value) {
  if (String(value || "").trim()) lines.push(`${label}: ${value}`);
}

function getValidationIssues() {
  const issues = [];
  const all = getAllResponders();
  const hasResponder = all.some((r) => r.name.trim());
  const hasOic = all.some((r) => r.oic && r.name.trim());
  const needsOic = state.incident.brigadeCode.trim().toUpperCase().startsWith("CONN");

  if (!state.incident.eventNumber.trim() || state.incident.eventNumber.trim() === "F") issues.push("Event Number");
  if (!buildPagerDateTime().trim()) issues.push("Pager Date / Time");
  if (!state.incident.actualLocation.trim()) issues.push("Actual Location");
  if (!hasResponder) issues.push("Responder");
  if (needsOic && !hasOic) issues.push("OIC");

  return issues;
}

function renderValidationIssues(issues) {
  const wrap = el("validationChips");
  if (!wrap) return;
  wrap.innerHTML = issues.map((item) => `<div class="validation-chip">${text(item)} missing</div>`).join("");
}

function isQuietHours() {
  const hour = new Date().getHours();
  return hour >= CONFIG.QUIET_HOURS_START || hour < CONFIG.QUIET_HOURS_END;
}

function generateReportFlow() {
  const issues = getValidationIssues();
  renderValidationIssues(issues);

  if (issues.length) {
    el("validationText").textContent = "Complete the missing fields below before generating the report.";
    state.ui.reportGenerated = false;
    updateReportActionState();
    return;
  }

  if (isQuietHours()) {
    el("validationText").textContent = "Report generated. It’s a bit late to send paperwork. Consider saving and sending during normal hours unless urgent.";
  } else {
    el("validationText").textContent = "Report generated.";
  }

  const report = buildReport();
  el("reportPreview").value = report;
  saveCurrentReport({ dedupe: true, silent: true });
  state.ui.reportGenerated = true;
  updateReportActionState();
  updateReportSummary();
}

function buildReportTitle() {
  return `${state.incident.eventNumber || "NO_EVENT"} – ${state.incident.incidentType || "UNKNOWN"} – ${state.incident.actualLocation || "NO LOCATION"}`;
}

function saveCurrentReport(options = {}) {
  const { dedupe = false, silent = false } = options;
  const title = buildReportTitle();
  const body = buildReport();

  if (dedupe) {
    const existing = state.savedReports.find((r) => r.title === title && r.body === body);
    if (existing) {
      if (!silent) updateDraftMeta("Matching report already saved locally.");
      return;
    }
  }

  state.savedReports.unshift({
    id: uid(),
    title,
    body,
    createdAt: new Date().toISOString()
  });

  state.savedReports = state.savedReports.slice(0, CONFIG.MAX_REPORTS);
  saveSavedReports();
  renderSavedReports();

  if (!silent) updateDraftMeta("Report saved locally.");
}

function renderSavedReports() {
  const wrap = el("savedReports");
  if (!wrap) return;

  wrap.innerHTML = "";

  if (!state.savedReports.length) {
    wrap.innerHTML = `<div class="saved-item">No saved reports yet.</div>`;
    return;
  }

  state.savedReports.forEach((r) => {
    const item = document.createElement("div");
    item.className = "saved-item";
    item.innerHTML = `
      <strong>${text(r.title)}</strong>
      <div class="help-text">${new Date(r.createdAt).toLocaleString("en-AU")}</div>
      <div class="saved-actions">
        <button class="tiny-btn" type="button">Load Preview</button>
        <button class="tiny-btn" type="button">Delete</button>
      </div>
    `;

    const [loadBtn, delBtn] = item.querySelectorAll("button");

    loadBtn.addEventListener("click", () => {
      el("reportPreview").value = r.body;
      showPage("sendPage");
      el("validationText").textContent = "Saved report loaded for review.";
      state.ui.reportGenerated = true;
      updateReportActionState();
    });

    delBtn.addEventListener("click", () => {
      const confirmed = window.confirm(`Delete saved report "${r.title}"?`);
      if (!confirmed) return;
      state.savedReports = state.savedReports.filter((x) => x.id !== r.id);
      saveSavedReports();
      renderSavedReports();
    });

    wrap.appendChild(item);
  });
}

function openSettings() {
  el("profileName").value = state.profile.name;
  el("profileMemberNumber").value = state.profile.memberNumber;
  el("profileContactNumber").value = state.profile.contactNumber;
  el("profileEmail").value = state.profile.email;
  el("profileBrigade").value = state.profile.brigade;
  el("settingsModal").classList.remove("hidden");
}

function closeSettings() {
  el("settingsModal").classList.add("hidden");
}

function saveProfileFromUi() {
  state.profile.name = el("profileName").value.trim();
  state.profile.memberNumber = el("profileMemberNumber").value.trim();
  state.profile.contactNumber = el("profileContactNumber").value.trim();
  state.profile.email = el("profileEmail").value.trim();
  state.profile.brigade = el("profileBrigade").value.trim() || "Connewarre";
  saveProfile();
  updateReportSummary();
  closeSettings();
}

function sendEmail() {
  const subject = buildReportTitle();
  const body = buildReport();
  window.location.href = `mailto:${encodeURIComponent(state.profile.email || "")}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

async function sendSms() {
  const body = buildReport();
  try {
    await navigator.clipboard.writeText(body);
    updateDraftMeta("Report copied to clipboard for SMS.");
  } catch {}
  window.location.href = `sms:?body=${encodeURIComponent(body)}`;
}

function showPage(pageId) {
  state.ui.currentPage = pageId;
  document.querySelectorAll(".page").forEach((p) => p.classList.toggle("active", p.id === pageId));
  document.querySelectorAll(".tab-btn[data-page]").forEach((b) => b.classList.toggle("active", b.dataset.page === pageId));

  if (pageId === "sendPage") {
    el("reportPreview").value = buildReport();
    renderValidationIssues(getValidationIssues());
    updateReportSummary();
  }

  scheduleDraftSave();
}

function renderIncidentFields() {
  el("eventNumber").value = state.incident.eventNumber || "F";
  el("pagerDate").value = pagerDateToInputValue(state.incident.pagerDate);
  el("pagerTime").value = state.incident.pagerTime;
  el("brigadeCode").value = state.incident.brigadeCode;
  el("brigadeRole").value = state.incident.brigadeRole;
  el("incidentType").value = state.incident.incidentType;
  el("pagerDetails").value = state.incident.pagerDetails;
  el("actualLocation").value = state.incident.actualLocation;
  el("controlName").value = state.incident.controlName;
  el("firsCode").value = state.incident.firsCode;
  el("comments").value = state.incident.comments;
  el("injuryNotes").value = state.incident.injuryNotes;
  el("firstAgency").value = state.incident.firstAgency;
  el("firstAgencyOther").value = state.incident.firstAgencyOther;
  el("firstAgencyOther").classList.toggle("hidden", state.incident.firstAgency !== "Other");

  el("weather1").value = state.incident.weather1;
  el("weather2").value = state.incident.weather2;
  syncWeather2Visibility();
  syncWeather2Options();

  el("distanceToScene").value = state.incident.distanceToScene;
  el("hose64Qty").value = state.incident.hoses.hose64Qty;
  el("hose38Qty").value = state.incident.hoses.hose38Qty;
  el("hose25Qty").value = state.incident.hoses.hose25Qty;
  el("hoseOtherType").value = state.incident.hoses.hoseOtherType;

  el("flagMembersBeforeBtn").classList.toggle("active", state.incident.flags.membersBefore);
  el("flagAarBtn").classList.toggle("active", state.incident.flags.aar);
  el("flagHotDebriefBtn")?.classList.toggle("active", state.incident.flags.hotDebrief);
  syncInjuriesFlagButton();
  syncSignalButtons();
  syncApplianceCodeButtons();

  renderHosesPanel();

  if (state.ui.previewUrl) {
    el("pagerPreview").src = state.ui.previewUrl;
    el("pagerPreview").classList.remove("hidden");
    el("pagerPreviewEmpty").classList.add("hidden");
  } else {
    el("pagerPreview").src = "";
    el("pagerPreview").classList.add("hidden");
    el("pagerPreviewEmpty").classList.remove("hidden");
  }
}

function renderEverything() {
  renderIncidentFields();
  renderSceneBrigades();
  renderAgencies();
  renderResponders();
  renderSavedReports();
  updateConnectionBanner();
  updateReportSummary();
  renderValidationIssues(getValidationIssues());
  updateReportActionState();
  showPage(state.ui.currentPage);
  renderDraftRestoreBanner();
}

function updateReportSummary() {
  const wrap = el("reportSummary");
  if (!wrap) return;

  const all = getAllResponders();
  const oic = all.find((r) => r.oic && r.name.trim());
  const applianceCount = ["T1", "T2", "MTD P/T"].filter((code) => all.some((r) => r.destination === code && r.name.trim())).length;
  const directCount = all.filter((r) => r.destination === "Direct" && r.name.trim()).length;
  const stationCount = all.filter((r) => r.destination === "Station" && r.name.trim()).length;
  const flagsCount = [
    state.incident.flags.membersBefore,
    state.incident.flags.aar,
    state.incident.flags.hotDebrief,
    isInjuriesFlagActive(),
    state.incident.signals.length > 0
  ].filter(Boolean).length;

  const items = [
    ["Event", state.incident.eventNumber || "Not set"],
    ["Type", state.incident.incidentType || "Not set"],
    ["OIC", oic?.name || "Not set"],
    ["Responders", all.filter((r) => r.name.trim()).length],
    ["Appliances", applianceCount],
    ["Direct", directCount],
    ["Station", stationCount],
    ["Agencies", state.agencies.length],
    ["Flags", flagsCount]
  ];

  wrap.innerHTML = items
    .map(([label, value]) => `<div class="summary-item">${text(label)}: ${text(value)}</div>`)
    .join("");
}

function getDraftPayload() {
  return {
    incident: clone(state.incident),
    responders: clone(state.responders),
    agencies: clone(state.agencies),
    ui: {
      currentPage: state.ui.currentPage,
      previewUrl: state.ui.previewUrl,
      hosesOpen: state.ui.hosesOpen
    },
    savedAt: new Date().toISOString()
  };
}

function isMeaningfulDraft(payload) {
  if (!payload) return false;

  const { incident, responders, agencies, ui } = payload;
  return Boolean(
    incident?.eventNumber ||
    incident?.pagerDate ||
    incident?.pagerTime ||
    incident?.brigadeCode ||
    incident?.incidentType ||
    incident?.actualLocation ||
    incident?.comments ||
    incident?.brigadesOnScene?.length ||
    agencies?.length ||
    responders?.connewarre?.some((r) => r.name?.trim()) ||
    responders?.mtd?.some((r) => r.name?.trim()) ||
    ui?.previewUrl
  );
}

function saveDraft() {
  const payload = getDraftPayload();
  localStorage.setItem(STORAGE_KEYS.DRAFT, JSON.stringify(payload));
  state.ui.lastDraftSavedAt = payload.savedAt;
  updateDraftMeta(
    `Draft autosaved ${new Date(payload.savedAt).toLocaleTimeString("en-AU", {
      hour: "2-digit",
      minute: "2-digit"
    })}.`
  );
}

function scheduleDraftSave() {
  clearTimeout(draftSaveTimer);
  draftSaveTimer = setTimeout(saveDraft, CONFIG.DRAFT_SAVE_DELAY_MS);
}

function forceDraftSave() {
  clearTimeout(draftSaveTimer);
  saveDraft();
  const stamp = new Date().toLocaleTimeString("en-AU", {
    hour: "2-digit",
    minute: "2-digit"
  });
  updateDraftMeta(`Draft saved manually at ${stamp}.`);
}

function restoreDraftIfPresent() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DRAFT);
    if (!raw) return;

    const payload = JSON.parse(raw);
    if (!isMeaningfulDraft(payload)) return;

    state.incident = { ...DEFAULT_INCIDENT(), ...payload.incident };
    state.responders = {
      connewarre: Array.isArray(payload.responders?.connewarre) ? payload.responders.connewarre : [],
      mtd: Array.isArray(payload.responders?.mtd) ? payload.responders.mtd : [],
      applianceCodes: {
        t1: payload.responders?.applianceCodes?.t1 || "",
        t2: payload.responders?.applianceCodes?.t2 || "",
        mtd: payload.responders?.applianceCodes?.mtd || ""
      }
    };
    state.agencies = Array.isArray(payload.agencies) ? payload.agencies : [];
    state.ui.currentPage = "incidentPage";
    state.ui.previewUrl = payload.ui?.previewUrl || "";
    state.ui.hosesOpen = Boolean(payload.ui?.hosesOpen);
    state.ui.restoredDraft = true;
    state.ui.lastDraftSavedAt = payload.savedAt || "";
  } catch {}
}

function renderDraftRestoreBanner() {
  const banner = el("draftRestoreBanner");
  const textEl = el("draftRestoreText");
  if (!banner || !textEl) return;

  if (!state.ui.restoredDraft) {
    banner.classList.add("hidden");
    return;
  }

  const stamp = state.ui.lastDraftSavedAt
    ? new Date(state.ui.lastDraftSavedAt).toLocaleString("en-AU")
    : "this device";

  textEl.textContent = `Draft restored from ${stamp}.`;
  banner.classList.remove("hidden");
}

function clearDraft() {
  const confirmed = window.confirm("Clear the current local draft from this device?");
  if (!confirmed) return;

  localStorage.removeItem(STORAGE_KEYS.DRAFT);
  state.ui.restoredDraft = false;
  state.ui.lastDraftSavedAt = "";
  el("draftRestoreBanner")?.classList.add("hidden");
  updateDraftMeta("Local draft cleared.");
}

function updateDraftMeta(message) {
  const target = el("draftMeta");
  if (target) target.textContent = message;
}

async function copyAppLink() {
  const link = `${window.location.origin}${window.location.pathname}`;
  try {
    await navigator.clipboard.writeText(link);
    updateDraftMeta("App link copied.");
  } catch {
    window.prompt("Copy this app link:", link);
  }
}

async function shareAppLink() {
  const link = `${window.location.origin}${window.location.pathname}`;
  try {
    if (navigator.share) {
      await navigator.share({
        title: "Connewarre Fire Brigade Turnout Sheet",
        text: "Open the turnout sheet",
        url: link
      });
      updateDraftMeta("App link shared.");
      return;
    }
    await copyAppLink();
  } catch {}
}
