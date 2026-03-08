/* =========================================================
   TURNOUT SHEET - CFA DIGITAL TURNOUT SHEET
   Stable full replacement app.js
   ========================================================= */

const CONFIG = {
  MAX_REPORTS: 10,
  MAX_VEHICLES: 5,
  QUIET_HOURS: { start: 22, end: 7 },
  MEMBER_FILES: [
    { key: "CONN", url: "CONN.members.json" },
    { key: "GROV", url: "GROV.members.json" },
    { key: "FRES", url: "FRES.members.json" }
  ],
  BRIGADE_SCENE_OPTIONS: ["CONN", "GROV", "FRES", "P64", "P32", "P31", "P33", "P35", "P41"],
  AGENCY_TYPES: ["Police", "Ambulance", "PowerCor", "Gas", "SES", "Local Government", "Other"],
  INCIDENT_TYPES: ["ALAR", "STRU", "NONS", "INCI", "G&SC"],
  CODE_LEVELS: ["C1", "C3"],
  VEHICLE_MAKES: [
    "Toyota", "Ford", "Holden", "Mazda", "Hyundai", "Kia", "Mitsubishi", "Nissan",
    "Subaru", "Volkswagen", "BMW", "Mercedes", "Audi", "Isuzu", "LDV", "Tesla",
    "BYD", "MG", "GWM", "Volvo", "Skoda", "Jeep", "Suzuki", "Lexus", "Ram", "Other"
  ],
  STATES: ["VIC", "NSW", "SA", "TAS", "ACT", "QLD", "WA", "NT"],
  OCCUPANTS: ["0", "1", "2", "3", "4+", "Unknown"],
  TRI_STATE: ["Yes", "No", "Unknown"],
  VEHICLE_STABILITY: ["Upright", "On side", "On roof", "Unknown"],
  VEHICLE_PROPULSION: [
    "ICE (Internal Combustion Engine – Petrol/Diesel)",
    "LPG",
    "Hybrid Electric Vehicle (HEV)",
    "Plug-in Hybrid Electric Vehicle (PHEV)",
    "Battery Electric Vehicle (BEV)",
    "Hydrogen Fuel Cell (FCEV)",
    "Heavy Vehicle / Diesel Truck",
    "Unknown",
    "Other"
  ],
  FIRST_AGENCIES: ["CFA", "FRV", "Police", "Ambulance", "SES", "PowerCor", "Gas", "Other"],
  RESPONSE_DESTINATIONS_CONNEWARRE: ["T1", "T2", "Station", "Direct"],
  RESPONSE_DESTINATIONS_MTD: ["MTD P/T", "Station", "Direct"],
  TRUCK_ROLES: ["Driver", "CL"],
  MVA_KEYWORDS: [
    "INCI", "MVA", "MVC", "MVI", "VEHICLE", "COLLISION", "CRASH",
    "ROLLOVER", "TRAPPED", "ROAD RESCUE", "CAR INTO", "TRUCK INTO", "VEHICLE INTO"
  ]
};

const STORAGE_KEYS = {
  PROFILE: "turnout_profile",
  SAVED_REPORTS: "turnout_saved_reports",
  DRAFT: "turnout_draft_state"
};

const state = {
  incident: {
    eventNumber: "",
    pagerDate: "",
    pagerTime: "",
    brigadeCode: "",
    brigadeRole: "",
    incidentType: "",
    codeLevel: "",
    address: "",
    firsCode: "",
    brigadesOnScene: [],
    firstAgency: "",
    firstAgencyOther: "",
    flags: {
      membersBefore: false,
      aarRequired: false,
      hotDebrief: false
    },
    notes: "",
    pagerRawText: ""
  },
  responders: {
    connewarre: [],
    mtd: []
  },
  agencies: [],
  vehicles: [],
  structure: {
    required: false,
    propertyType: "",
    levels: "",
    occupancy: "",
    construction: "",
    fireAreaOrigin: "",
    fireAreaExtent: "",
    fireAreaComments: "",
    fireFuelLoad: "",
    fireBehaviour: "",
    fireMaterialComments: "",
    detectionType: "",
    alarmActivated: "",
    detectionComments: "",
    suppressionType: "",
    suppressionWorked: "",
    suppressionComments: "",
    portableExtinguisher: "",
    portableOther: "",
    portableComments: ""
  },
  hoseUse: {
    "64": "",
    "38": "",
    "25": "",
    "Live Reel": ""
  },
  profile: {
    name: "",
    memberNumber: "",
    contactNumber: "",
    email: "",
    brigade: "Connewarre"
  },
  memberLists: {
    CONN: [],
    GROV: [],
    FRES: []
  },
  savedReports: [],
  ui: {
    currentPage: "incidentPage",
    pagerImageFile: null,
    pagerPreviewUrl: "",
    pagerCropPreviewUrl: "",
    reportPreview: ""
  }
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  try {
    await loadMemberLists();
    loadProfile();
    loadSavedReports();
    restoreDraft();

    ensureMinimumResponderRows();
    bindBaseEvents();
    syncAllFieldsToUi();
    renderResponders();
    renderAgencies();
    renderVehicles();
    renderBrigadesOnScene();
    renderSavedReports();
    updateHeaderOicStatus();
    showPage(state.ui.currentPage || "incidentPage");
    registerServiceWorker();
  } catch (error) {
    console.error("Init failed", error);
    setStatus("ocrStatus", "App failed to initialise. Refresh and try again.");
  }
}

/* =========================================================
   BASIC HELPERS
   ========================================================= */

function el(id) {
  return document.getElementById(id);
}

function setStatus(id, text) {
  const node = el(id);
  if (node) node.textContent = text;
}

function normaliseSpacing(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function cryptoRandomId() {
  return `id_${Math.random().toString(36).slice(2, 11)}`;
}

function lineIfValue(label, value) {
  if (value === null || value === undefined || String(value).trim() === "") return "";
  return `${label}: ${value}`;
}

function formatDateForReport(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  if (!year || !month || !day) return value;
  return `${day}/${month}/${year}`;
}

function formatBrigadeName(brigadeKey) {
  if (brigadeKey === "CONN") return "Connewarre";
  if (brigadeKey === "GROV") return "Grovedale";
  if (brigadeKey === "FRES") return "Freshwater Creek";
  return brigadeKey || "";
}

function appendBrigadeIfNeeded(responder) {
  return responder.brigade && responder.brigade !== "CONN"
    ? `${responder.name} (${formatBrigadeName(responder.brigade)})`
    : responder.name;
}

function isPrimaryConnJob() {
  return (state.incident.brigadeCode || "").toUpperCase().startsWith("CONN");
}

function getAllResponders() {
  return [...state.responders.connewarre, ...state.responders.mtd];
}

function findResponderById(id) {
  return getAllResponders().find((r) => r.id === id);
}

/* =========================================================
   STORAGE
   ========================================================= */

function loadProfile() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.PROFILE);
    if (!raw) return;
    state.profile = { ...state.profile, ...JSON.parse(raw) };
  } catch (error) {
    console.warn("Profile load failed", error);
  }
}

function saveProfile() {
  localStorage.setItem(STORAGE_KEYS.PROFILE, JSON.stringify(state.profile));
}

function loadSavedReports() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.SAVED_REPORTS);
    state.savedReports = raw ? JSON.parse(raw) : [];
  } catch (error) {
    state.savedReports = [];
  }
}

function saveSavedReports() {
  localStorage.setItem(STORAGE_KEYS.SAVED_REPORTS, JSON.stringify(state.savedReports));
}

function persistDraft() {
  localStorage.setItem(STORAGE_KEYS.DRAFT, JSON.stringify(state));
}

function restoreDraft() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DRAFT);
    if (!raw) return;
    const parsed = JSON.parse(raw);

    if (parsed.incident) {
      state.incident = { ...state.incident, ...parsed.incident };
      if (!Array.isArray(state.incident.brigadesOnScene)) state.incident.brigadesOnScene = [];
    }
    if (parsed.responders) {
      state.responders = { ...state.responders, ...parsed.responders };
    }
    if (parsed.agencies) state.agencies = parsed.agencies;
    if (parsed.vehicles) state.vehicles = parsed.vehicles;
    if (parsed.structure) state.structure = { ...state.structure, ...parsed.structure };
    if (parsed.hoseUse) state.hoseUse = { ...state.hoseUse, ...parsed.hoseUse };
    if (parsed.profile) state.profile = { ...state.profile, ...parsed.profile };
    if (parsed.savedReports) state.savedReports = parsed.savedReports;
    if (parsed.ui) {
      state.ui.currentPage = parsed.ui.currentPage || "incidentPage";
      state.ui.reportPreview = parsed.ui.reportPreview || "";
    }
  } catch (error) {
    console.warn("Draft restore failed", error);
  }
}

/* =========================================================
   MEMBER LISTS
   ========================================================= */

async function loadMemberLists() {
  for (const file of CONFIG.MEMBER_FILES) {
    try {
      const res = await fetch(file.url);
      if (!res.ok) throw new Error(`Failed to load ${file.url}`);
      const data = await res.json();
      state.memberLists[file.key] = Array.isArray(data) ? data : [];
    } catch (error) {
      console.warn(`Could not load ${file.url}`, error);
      state.memberLists[file.key] = [];
    }
  }
}

function getMemberRecord(brigadeKey, memberName) {
  if (!brigadeKey || !memberName) return null;
  const members = state.memberLists[brigadeKey] || [];
  const target = memberName.trim().toUpperCase();

  return members.find((member) => {
    const name = typeof member === "string" ? member : (member.name || "");
    return name.trim().toUpperCase() === target;
  }) || null;
}

function getMemberPhone(brigadeKey, memberName) {
  const record = getMemberRecord(brigadeKey, memberName);
  if (!record || typeof record === "string") return "";
  return record.phone || "";
}

function getMemberOptionsHtml(brigadeKey, currentValue = "") {
  const options = (state.memberLists[brigadeKey] || []).map((member) => {
    const name = typeof member === "string" ? member : (member.name || "");
    return `<option value="${escapeHtml(name)}"></option>`;
  });

  if (currentValue && !options.some((opt) => opt.includes(`value="${escapeHtml(currentValue)}"`))) {
    options.unshift(`<option value="${escapeHtml(currentValue)}"></option>`);
  }

  return options.join("");
}

function getCombinedMtdMemberOptionsHtml(currentValue = "") {
  const names = [];
  ["CONN", "GROV", "FRES"].forEach((key) => {
    (state.memberLists[key] || []).forEach((member) => {
      const name = typeof member === "string" ? member : (member.name || "");
      if (name) names.push(name);
    });
  });

  const unique = [...new Set(names)];
  const html = unique.map((name) => `<option value="${escapeHtml(name)}"></option>`);

  if (currentValue && !unique.includes(currentValue)) {
    html.unshift(`<option value="${escapeHtml(currentValue)}"></option>`);
  }

  return html.join("");
}

function inferMtdMemberRecord(memberName) {
  const target = String(memberName || "").trim().toUpperCase();
  if (!target) return null;

  for (const key of ["CONN", "GROV", "FRES"]) {
    const found = (state.memberLists[key] || []).find((member) => {
      const name = typeof member === "string" ? member : (member.name || "");
      return name.trim().toUpperCase() === target;
    });

    if (found) {
      return {
        brigade: key,
        phone: typeof found === "string" ? "" : (found.phone || "")
      };
    }
  }

  return null;
}

/* =========================================================
   BASE EVENTS
   ========================================================= */

function bindBaseEvents() {
  document.querySelectorAll(".nav-btn[data-page]").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.page));
  });

  document.querySelectorAll(".next-btn").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.next));
  });

  document.querySelectorAll(".back-btn").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.back));
  });

  document.querySelectorAll(".stack-toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = el(btn.dataset.target);
      if (target) target.classList.toggle("open");
    });
  });

  document.querySelectorAll(".accordion-toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = el(btn.dataset.target);
      if (target) target.classList.toggle("open");
    });
  });

  if (el("settingsBtn")) el("settingsBtn").addEventListener("click", openSettingsModal);
  if (el("closeSettingsBtn")) el("closeSettingsBtn").addEventListener("click", closeSettingsModal);
  if (el("saveProfileBtn")) el("saveProfileBtn").addEventListener("click", saveProfileFromUi);

  if (el("pagerUpload")) el("pagerUpload").addEventListener("change", handleImageUpload);
  if (el("runOcrBtn")) el("runOcrBtn").addEventListener("click", runOcrPipeline);

  bindIncidentFieldEvents();
  bindSceneBrigadeEvents();
  bindStructureEvents();
  bindVehicleEvents();
  bindAgencyEvents();
  bindSendEvents();
}

function bindIncidentFieldEvents() {
  const textBindings = [
    ["eventNumber", "eventNumber"],
    ["pagerDate", "pagerDate"],
    ["pagerTime", "pagerTime"],
    ["brigadeCode", "brigadeCode"],
    ["incidentType", "incidentType"],
    ["codeLevel", "codeLevel"],
    ["address", "address"],
    ["firsCode", "firsCode"],
    ["notes", "notes"]
  ];

  textBindings.forEach(([id, key]) => {
    const node = el(id);
    if (!node) return;
    node.addEventListener("input", (e) => {
      state.incident[key] = e.target.value;
      if (key === "brigadeCode") updateBrigadeRole();
      if (key === "incidentType") updateStructureVisibilityByIncidentType();
      if (key === "firsCode") updateFirsLabel();
      persistDraft();
    });
  });

  if (el("resetFirsBtn")) {
    el("resetFirsBtn").addEventListener("click", () => {
      state.incident.firsCode = "";
      el("firsCode").value = "";
      updateFirsLabel();
      persistDraft();
    });
  }

  if (el("firstAgencySelect")) {
    el("firstAgencySelect").addEventListener("change", (e) => {
      state.incident.firstAgency = e.target.value;
      el("firstAgencyOther").classList.toggle("hidden", e.target.value !== "Other");
      persistDraft();
    });
  }

  if (el("firstAgencyOther")) {
    el("firstAgencyOther").addEventListener("input", (e) => {
      state.incident.firstAgencyOther = e.target.value;
      persistDraft();
    });
  }

  if (el("flagMembersBeforeBtn")) {
    el("flagMembersBeforeBtn").addEventListener("click", () => {
      state.incident.flags.membersBefore = !state.incident.flags.membersBefore;
      renderIncidentFlagButtons();
      persistDraft();
    });
  }

  if (el("flagAarRequiredBtn")) {
    el("flagAarRequiredBtn").addEventListener("click", () => {
      state.incident.flags.aarRequired = !state.incident.flags.aarRequired;
      renderIncidentFlagButtons();
      persistDraft();
    });
  }

  if (el("flagHotDebriefBtn")) {
    el("flagHotDebriefBtn").addEventListener("click", () => {
      state.incident.flags.hotDebrief = !state.incident.flags.hotDebrief;
      renderIncidentFlagButtons();
      persistDraft();
    });
  }
}

function bindSceneBrigadeEvents() {
  if (el("brigadeOnSceneSelect")) {
    el("brigadeOnSceneSelect").addEventListener("change", (e) => {
      el("brigadeOnSceneOther").classList.toggle("hidden", e.target.value !== "Other");
    });
  }

  if (el("addBrigadeOnSceneBtn")) {
    el("addBrigadeOnSceneBtn").addEventListener("click", addBrigadeOnSceneFromUi);
  }
}

function bindStructureEvents() {
  document.querySelectorAll('input[name="structureRequired"]').forEach((radio) => {
    radio.addEventListener("change", () => {
      const checked = document.querySelector('input[name="structureRequired"]:checked');
      state.structure.required = checked?.value === "yes";
      el("structureFormWrap").classList.toggle("hidden", !state.structure.required);
      persistDraft();
    });
  });

  [
    "structurePropertyType", "structureLevels", "structureOccupancy", "structureConstruction",
    "fireAreaOrigin", "fireAreaExtent", "fireAreaComments",
    "fireFuelLoad", "fireBehaviour", "fireMaterialComments",
    "detectionType", "alarmActivated", "detectionComments",
    "suppressionType", "suppressionWorked", "suppressionComments",
    "portableExtinguisher", "portableOther", "portableComments"
  ].forEach((id) => {
    const node = el(id);
    if (!node) return;
    node.addEventListener("input", collectStructureData);
  });

  ["hose64", "hose38", "hose25", "hoseLiveReel"].forEach((id) => {
    const node = el(id);
    if (!node) return;
    node.addEventListener("input", collectHoseUse);
  });
}

function bindVehicleEvents() {
  if (el("vehicleCount")) {
    el("vehicleCount").addEventListener("change", (e) => {
      setVehicleCount(Number(e.target.value));
    });
  }
}

function bindAgencyEvents() {
  if (el("addAgencyBtn")) {
    el("addAgencyBtn").addEventListener("click", () => {
      const type = el("agencyType").value;
      if (!type) return;
      addAgencyBlock(type);
      el("agencyType").value = "";
    });
  }
}

function bindSendEvents() {
  if (el("finishBtn")) el("finishBtn").addEventListener("click", handleFinish);
  if (el("sendSmsBtn")) el("sendSmsBtn").addEventListener("click", sendSms);
  if (el("sendEmailBtn")) el("sendEmailBtn").addEventListener("click", sendEmail);
  if (el("saveLocalBtn")) el("saveLocalBtn").addEventListener("click", saveReportLocally);
}

/* =========================================================
   OCR / IMAGE PREVIEW
   ========================================================= */

function handleImageUpload(e) {
  const file = e.target.files?.[0] || null;
  state.ui.pagerImageFile = file;

  if (!file) {
    state.ui.pagerPreviewUrl = "";
    state.ui.pagerCropPreviewUrl = "";
    renderPagerPreviews();
    setStatus("ocrStatus", "No screenshot selected.");
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    state.ui.pagerPreviewUrl = reader.result;
    state.ui.pagerCropPreviewUrl = "";
    renderPagerPreviews();
    setStatus("ocrStatus", `Loaded ${file.name}. Scanning...`);
    runOcrPipeline();
  };
  reader.onerror = () => {
    setStatus("ocrStatus", "Could not read screenshot file.");
  };
  reader.readAsDataURL(file);
}

async function runOcrPipeline() {
  if (!state.ui.pagerImageFile) {
    setStatus("ocrStatus", "Please upload a pager screenshot first.");
    return;
  }

  if (typeof Tesseract === "undefined") {
    setStatus("ocrStatus", "OCR library not loaded.");
    return;
  }

  try {
    setStatus("ocrStatus", "Scanning screenshot...");

    const imageBitmap = await createImageBitmap(state.ui.pagerImageFile);
    const rects = findPagerRectangles(imageBitmap);
    const selectedRect = rects.length ? selectBestPagerRectangle(rects) : null;

    const cropCanvas = cropPagerRectangle(imageBitmap, selectedRect);
    state.ui.pagerCropPreviewUrl = cropCanvas.toDataURL("image/png");
    renderPagerPreviews();

    let text = await runPagerOCR(cropCanvas);
    let parsed = parsePagerMessage(text);

    if (!isUsefulPagerParse(parsed)) {
      const fullCanvas = cropPagerRectangle(imageBitmap, null);
      text = await runPagerOCR(fullCanvas);
      const fullParsed = parsePagerMessage(text);
      if (scorePagerParse(fullParsed) >= scorePagerParse(parsed)) {
        parsed = fullParsed;
        state.ui.pagerCropPreviewUrl = fullCanvas.toDataURL("image/png");
        renderPagerPreviews();
      }
    }

    if (!isUsefulPagerParse(parsed)) {
      setStatus("ocrStatus", "Scan ran, but the pager details were too weak to populate safely.");
      return;
    }

    populateIncidentFields(parsed);
    state.incident.pagerRawText = parsed.rawText || "";
    setStatus("ocrStatus", "Pager details extracted.");
    persistDraft();
  } catch (error) {
    console.error("OCR failed", error);
    setStatus("ocrStatus", "Scan failed. Try a clearer screenshot.");
  }
}

function renderPagerPreviews() {
  const preview = el("pagerPreviewImage");
  const previewEmpty = el("pagerPreviewEmpty");
  const crop = el("pagerCropPreviewImage");
  const cropEmpty = el("pagerCropPreviewEmpty");

  if (preview) {
    if (state.ui.pagerPreviewUrl) {
      preview.src = state.ui.pagerPreviewUrl;
      preview.classList.remove("hidden");
      if (previewEmpty) previewEmpty.classList.add("hidden");
    } else {
      preview.src = "";
      preview.classList.add("hidden");
      if (previewEmpty) previewEmpty.classList.remove("hidden");
    }
  }

  if (crop) {
    if (state.ui.pagerCropPreviewUrl) {
      crop.src = state.ui.pagerCropPreviewUrl;
      crop.classList.remove("hidden");
      if (cropEmpty) cropEmpty.classList.add("hidden");
    } else {
      crop.src = "";
      crop.classList.add("hidden");
      if (cropEmpty) cropEmpty.classList.remove("hidden");
    }
  }
}

function findPagerRectangles(imageBitmap) {
  const canvas = document.createElement("canvas");
  canvas.width = imageBitmap.width;
  canvas.height = imageBitmap.height;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(imageBitmap, 0, 0);

  const { width, height } = canvas;
  const imageData = ctx.getImageData(0, 0, width, height);
  const data = imageData.data;

  const rowScores = [];
  for (let y = 0; y < height; y++) {
    let darkPixels = 0;
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
      if (avg < 190) darkPixels++;
    }
    rowScores.push(darkPixels / width);
  }

  const bands = [];
  let start = null;

  for (let i = 0; i < rowScores.length; i++) {
    if (rowScores[i] > 0.10 && start === null) start = i;
    if ((rowScores[i] <= 0.10 || i === rowScores.length - 1) && start !== null) {
      const end = i;
      const bandHeight = end - start;
      if (bandHeight > 100) {
        bands.push({ x: 0, y: start, width, height: bandHeight });
      }
      start = null;
    }
  }

  if (!bands.length) return [{ x: 0, y: 0, width, height }];
  return bands.slice(0, 5);
}

function selectBestPagerRectangle(rectangles) {
  return rectangles.sort((a, b) => a.y - b.y)[0];
}

function cropPagerRectangle(imageBitmap, rect) {
  const canvas = document.createElement("canvas");
  const x = rect?.x || 0;
  const y = rect?.y || 0;
  const width = rect?.width || imageBitmap.width;
  const height = rect?.height || imageBitmap.height;

  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d");
  ctx.drawImage(imageBitmap, x, y, width, height, 0, 0, width, height);
  return canvas;
}

async function runPagerOCR(canvas) {
  const result = await Tesseract.recognize(canvas, "eng", { logger: () => {} });
  return result?.data?.text || "";
}

function parsePagerMessage(rawText) {
  const text = String(rawText || "").replace(/\r/g, "\n").trim();
  const upper = text.toUpperCase();
  const compact = upper.replace(/\n/g, " ");

  const eventNumber = (compact.match(/\bF\d{9}\b/) || [])[0] || "";

  const dateMatch = compact.match(/\b(\d{2})[\/\-](\d{2})[\/\-](\d{2,4})\b/);
  let pagerDate = "";
  if (dateMatch) {
    const day = dateMatch[1];
    const month = dateMatch[2];
    let year = dateMatch[3];
    if (year.length === 2) year = `20${year}`;
    pagerDate = `${year}-${month}-${day}`;
  }

  const timeMatch = compact.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
  const pagerTime = timeMatch ? `${timeMatch[1].padStart(2, "0")}:${timeMatch[2]}` : "";

  const brigadeCode = (compact.match(/\b(CONN|GROV|FRES)\d\b/) || [])[0] || "";
  const incidentClassMatch = compact.match(/\b(ALAR|STRU|NONS|INCI|G&SC)(C1|C3)?\b/);
  const incidentType = incidentClassMatch?.[1] || "";
  const codeLevel = incidentClassMatch?.[2] || "";

  const address = extractPagerAddress(compact);
  const brigadesOnScene = extractKnownBrigades(compact);

  return {
    rawText: text,
    upper: compact,
    eventNumber,
    pagerDate,
    pagerTime,
    brigadeCode,
    incidentType,
    codeLevel,
    address,
    brigadesOnScene
  };
}

function extractPagerAddress(text) {
  const withSlash = text.match(/(\d{1,5}\s+[A-Z0-9.' -]+?\s(?:ST|RD|AV))\s*\//i);
  if (withSlash) return normaliseSpacing(withSlash[1]);

  const fallback = text.match(/(\d{1,5}\s+[A-Z0-9.' -]+?\s(?:ST|RD|AV))\b/i);
  if (fallback) return normaliseSpacing(fallback[1]);

  return "";
}

function extractKnownBrigades(text) {
  const found = [];
  CONFIG.BRIGADE_SCENE_OPTIONS.forEach((code) => {
    if (text.includes(code) && !found.includes(code)) found.push(code);
  });
  return found;
}

function scorePagerParse(parsed) {
  let score = 0;
  if (parsed.eventNumber) score += 3;
  if (parsed.address) score += 4;
  if (parsed.brigadeCode) score += 2;
  if (parsed.incidentType) score += 2;
  if (parsed.pagerTime) score += 1;
  return score;
}

function isUsefulPagerParse(parsed) {
  return scorePagerParse(parsed) >= 4;
}

function populateIncidentFields(parsed) {
  state.incident.eventNumber = parsed.eventNumber || state.incident.eventNumber;
  state.incident.pagerDate = parsed.pagerDate || state.incident.pagerDate;
  state.incident.pagerTime = parsed.pagerTime || state.incident.pagerTime;
  state.incident.brigadeCode = parsed.brigadeCode || state.incident.brigadeCode;
  state.incident.incidentType = parsed.incidentType || state.incident.incidentType;
  state.incident.codeLevel = parsed.codeLevel || state.incident.codeLevel;
  state.incident.address = parsed.address || state.incident.address;

  if (Array.isArray(parsed.brigadesOnScene) && parsed.brigadesOnScene.length) {
    state.incident.brigadesOnScene = [...new Set([...state.incident.brigadesOnScene, ...parsed.brigadesOnScene])];
  }

  syncIncidentFieldsToUi();
  renderBrigadesOnScene();
  updateBrigadeRole();
  updateStructureVisibilityByIncidentType();
  evaluateMvaAutoTrigger(parsed.upper || "");
}

/* =========================================================
   INCIDENT RENDER
   ========================================================= */

function addBrigadeOnSceneFromUi() {
  const select = el("brigadeOnSceneSelect");
  const other = el("brigadeOnSceneOther");
  let value = select.value;

  if (!value) return;

  if (value === "Other") {
    value = normaliseSpacing(other.value).toUpperCase();
    if (!value) return;
  }

  if (!state.incident.brigadesOnScene.includes(value)) {
    state.incident.brigadesOnScene.push(value);
  }

  select.value = "";
  other.value = "";
  other.classList.add("hidden");
  renderBrigadesOnScene();
  persistDraft();
}

function removeBrigadeOnScene(code) {
  state.incident.brigadesOnScene = state.incident.brigadesOnScene.filter((b) => b !== code);
  renderBrigadesOnScene();
  persistDraft();
}

function renderBrigadesOnScene() {
  const container = el("brigadesOnSceneChips");
  if (!container) return;
  container.innerHTML = "";

  state.incident.brigadesOnScene.forEach((code) => {
    const chip = document.createElement("div");
    chip.className = "scene-chip";
    chip.innerHTML = `<span>${escapeHtml(code)}</span><button type="button">×</button>`;
    chip.querySelector("button").addEventListener("click", () => removeBrigadeOnScene(code));
    container.appendChild(chip);
  });
}

function updateBrigadeRole() {
  const code = (state.incident.brigadeCode || "").toUpperCase().trim();
  state.incident.brigadeRole = code.startsWith("CONN") ? "Primary" : code ? "Support" : "";
  if (el("brigadeRole")) el("brigadeRole").value = state.incident.brigadeRole;
  persistDraft();
}

function updateFirsLabel() {
  if (el("firsLabel")) {
    el("firsLabel").textContent = state.incident.firsCode ? "FIRS Code" : "Paste FIRS";
  }
}

function updateStructureVisibilityByIncidentType() {
  const isStructure = state.incident.incidentType === "STRU";
  if (!isStructure) {
    state.structure.required = false;
    const noRadio = document.querySelector('input[name="structureRequired"][value="no"]');
    if (noRadio) noRadio.checked = true;
    if (el("structureFormWrap")) el("structureFormWrap").classList.add("hidden");
  }
  persistDraft();
}

function renderIncidentFlagButtons() {
  el("flagMembersBeforeBtn")?.classList.toggle("mini-chip-active", state.incident.flags.membersBefore);
  el("flagAarRequiredBtn")?.classList.toggle("mini-chip-active", state.incident.flags.aarRequired);
  el("flagHotDebriefBtn")?.classList.toggle("mini-chip-active", state.incident.flags.hotDebrief);
}

function evaluateMvaAutoTrigger(text) {
  const hasMatch = CONFIG.MVA_KEYWORDS.some((keyword) => text.includes(keyword));
  if (hasMatch) {
    el("vehicleSection")?.classList.add("open");
    if (state.vehicles.length === 0) setVehicleCount(1);
  }
}

/* =========================================================
   RESPONDER RENDER
   ========================================================= */

function renderResponders() {
  renderConnewarreResponders();
  renderMtdResponders();
  renderOtherResponderLists();
  updateHeaderOicStatus();
}

function renderConnewarreResponders() {
  const container = el("connewarreResponders");
  if (!container) return;

  container.innerHTML = `<div class="responder-list-wrap"></div>`;
  const list = container.firstElementChild;

  state.responders.connewarre.forEach((responder, index) => {
    list.appendChild(buildResponderRow(responder, index));
  });

  bindResponderEvents(container);
}

function renderMtdResponders() {
  const container = el("mtdResponders");
  if (!container) return;

  container.innerHTML = `<div class="responder-list-wrap"></div>`;
  const list = container.firstElementChild;

  state.responders.mtd.forEach((responder, index) => {
    list.appendChild(buildResponderRow(responder, index));
  });

  bindResponderEvents(container);
}

function renderOtherResponderLists() {
  const directContainer = el("directRespondersContainer");
  const stationContainer = el("stationRespondersContainer");
  if (!directContainer || !stationContainer) return;

  const direct = getAllResponders().filter((r) => r.name.trim() && r.destination === "Direct");
  const station = getAllResponders().filter((r) => r.name.trim() && r.destination === "Station");

  directContainer.innerHTML = direct.length
    ? direct.map((r) => `<div class="simple-note">${escapeHtml(appendBrigadeIfNeeded(r))}${escapeHtml(buildResponderSuffix(r))}</div>`).join("")
    : `<div class="simple-note">No direct responders selected.</div>`;

  stationContainer.innerHTML = station.length
    ? station.map((r) => `<div class="simple-note">${escapeHtml(appendBrigadeIfNeeded(r))}${escapeHtml(buildResponderSuffix(r))}</div>`).join("")
    : `<div class="simple-note">No station responders selected.</div>`;
}

function buildResponderRow(responder, index) {
  const wrapper = document.createElement("div");
  const isConn = responder.group === "connewarre";
  const badge = responder.group === "mtd" && responder.brigade
    ? `<div class="responder-badge">${formatBrigadeName(responder.brigade)}</div>`
    : "";

  const destinations = isConn
    ? CONFIG.RESPONSE_DESTINATIONS_CONNEWARRE
    : CONFIG.RESPONSE_DESTINATIONS_MTD;

  const destinationButtons = destinations.map((dest) => {
    const active = responder.destination === dest;
    return `
      <button
        type="button"
        class="mini-chip ${active ? "mini-chip-active" : ""}"
        data-action="destination"
        data-responder-id="${responder.id}"
        data-value="${dest}"
        ${!responder.name.trim() ? "disabled" : ""}
      >${dest}</button>
    `;
  }).join("");

  const showTruckRoles = responder.destination === "T1" || responder.destination === "T2";
  const roleButtons = showTruckRoles
    ? CONFIG.TRUCK_ROLES.map((role) => {
        const disabled = isTruckRoleUnavailable(responder, responder.destination, role);
        const active = responder.truckRole === role;
        return `
          <button
            type="button"
            class="mini-chip ${active ? "mini-chip-active" : ""}"
            data-action="truck-role"
            data-responder-id="${responder.id}"
            data-value="${role}"
            ${disabled && !active ? "disabled" : ""}
          >${role}</button>
        `;
      }).join("")
    : "";

  const flagsVisible = !!responder.destination;
  const flagButtons = flagsVisible ? `
    <button type="button" class="mini-chip ${responder.ba ? "mini-chip-active" : ""}" data-action="toggle-flag" data-flag="ba" data-responder-id="${responder.id}">BA</button>
    <button type="button" class="mini-chip ${responder.injury ? "mini-chip-active" : ""}" data-action="toggle-flag" data-flag="injury" data-responder-id="${responder.id}">Injured</button>
    <button type="button" class="mini-chip ${responder.oic ? "mini-chip-active" : ""}" data-action="toggle-flag" data-flag="oic" data-responder-id="${responder.id}">OIC</button>
  ` : "";

  const datalistHtml = isConn
    ? getMemberOptionsHtml("CONN", responder.name)
    : getCombinedMtdMemberOptionsHtml(responder.name);

  const showAddMember = shouldShowAddMemberButton(responder, index);

  wrapper.innerHTML = `
    <div class="responder-row" data-responder-id="${responder.id}">
      <div class="responder-row-top">
        <div class="responder-name-wrap">
          <input
            type="text"
            list="member-list-${responder.id}"
            placeholder="Name"
            data-action="name-input"
            data-responder-id="${responder.id}"
            value="${escapeHtml(responder.name)}"
          />
          <datalist id="member-list-${responder.id}">
            ${datalistHtml}
          </datalist>
        </div>

        ${badge}

        <button
          type="button"
          class="tiny-btn"
          data-action="remove-member"
          data-group="${responder.group}"
          data-responder-id="${responder.id}"
        >Remove</button>
      </div>

      <div class="responder-stage">
        <div class="stage-label">Response</div>
        <div class="chip-row">${destinationButtons}</div>
      </div>

      ${showTruckRoles ? `
        <div class="responder-stage">
          <div class="stage-label">Truck role</div>
          <div class="chip-row">${roleButtons}<div class="role-hint">Leave blank = Crew</div></div>
        </div>
      ` : ""}

      ${flagsVisible ? `
        <div class="responder-stage">
          <div class="stage-label">Flags</div>
          <div class="chip-row">${flagButtons}</div>
        </div>
      ` : ""}

      ${showAddMember ? `<button type="button" class="secondary-btn compact-add-btn" data-action="add-member" data-group="${responder.group}">Add Member</button>` : ""}
    </div>
  `;

  return wrapper.firstElementChild;
}

function bindResponderEvents(container) {
  container.querySelectorAll('[data-action="name-input"]').forEach((input) => {
    input.addEventListener("input", handleResponderNameTyping);
    input.addEventListener("change", handleResponderNameCommit);
    input.addEventListener("blur", handleResponderNameCommit);
  });

  container.querySelectorAll('[data-action="destination"]').forEach((btn) => {
    btn.addEventListener("click", handleDestinationSelect);
  });

  container.querySelectorAll('[data-action="truck-role"]').forEach((btn) => {
    btn.addEventListener("click", handleTruckRoleSelect);
  });

  container.querySelectorAll('[data-action="toggle-flag"]').forEach((btn) => {
    btn.addEventListener("click", handleResponderFlagToggle);
  });

  container.querySelectorAll('[data-action="add-member"]').forEach((btn) => {
    btn.addEventListener("click", handleAddMember);
  });

  container.querySelectorAll('[data-action="remove-member"]').forEach((btn) => {
    btn.addEventListener("click", handleRemoveMember);
  });
}

function handleResponderNameTyping(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;
  responder.name = e.target.value;
}

function handleResponderNameCommit(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  responder.name = e.target.value.trim();

  if (responder.group === "connewarre") {
    responder.brigade = "CONN";
    responder.phone = getMemberPhone("CONN", responder.name) || "";
  } else {
    const inferred = inferMtdMemberRecord(responder.name);
    responder.brigade = inferred?.brigade || responder.brigade || "";
    responder.phone = inferred?.phone || responder.phone || "";
  }

  persistDraft();
  renderResponders();
}

function handleDestinationSelect(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  responder.destination = e.target.dataset.value;
  if (!(responder.destination === "T1" || responder.destination === "T2")) {
    responder.truckRole = "";
  }

  persistDraft();
  renderResponders();
}

function handleTruckRoleSelect(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  const role = e.target.dataset.value;
  if (isTruckRoleUnavailable(responder, responder.destination, role)) return;

  responder.truckRole = responder.truckRole === role ? "" : role;
  persistDraft();
  renderResponders();
}

function handleResponderFlagToggle(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  const flag = e.target.dataset.flag;

  if (flag === "oic") {
    if (!responder.oic) {
      clearExistingOic();
      responder.oic = true;
      responder.phone = responder.phone || getResponderPhoneFallback(responder);
    } else {
      responder.oic = false;
    }
  } else {
    responder[flag] = !responder[flag];
  }

  persistDraft();
  renderResponders();
}

function handleAddMember(e) {
  const group = e.target.dataset.group;
  state.responders[group].push(createResponder(group));
  persistDraft();
  renderResponders();
}

function handleRemoveMember(e) {
  const group = e.target.dataset.group;
  const responderId = e.target.dataset.responderId;
  state.responders[group] = state.responders[group].filter((r) => r.id !== responderId);

  if (!state.responders[group].length) {
    state.responders[group].push(createResponder(group));
  }

  persistDraft();
  renderResponders();
}

function shouldShowAddMemberButton(responder, index) {
  const list = state.responders[responder.group];
  const isLast = index === list.length - 1;
  return isLast && responder.name.trim() && responder.destination;
}

function isTruckRoleUnavailable(responder, truck, role) {
  return getAllResponders().some((r) =>
    r.id !== responder.id &&
    r.destination === truck &&
    r.truckRole === role
  );
}

function clearExistingOic() {
  getAllResponders().forEach((r) => {
    r.oic = false;
  });
}

function getResponderPhoneFallback(responder) {
  if (responder.brigade) {
    return getMemberPhone(responder.brigade, responder.name) || responder.phone || "";
  }
  return responder.phone || "";
}

function updateHeaderOicStatus() {
  const node = el("oicStatus");
  if (!node) return;

  const oic = getAllResponders().find((r) => r.oic && r.name.trim());

  if (!oic) {
    node.textContent = "APPOINT OIC";
    node.classList.add("oic-missing");
    return;
  }

  oic.phone = oic.phone || getResponderPhoneFallback(oic);
  node.textContent = `OIC: ${oic.name}${oic.phone ? " – " + oic.phone : ""}`;
  node.classList.remove("oic-missing");
}

/* =========================================================
   AGENCIES
   ========================================================= */

function addAgencyBlock(type) {
  state.agencies.push({
    id: cryptoRandomId(),
    type,
    otherName: "",
    officerName: "",
    contactNumber: "",
    station: "",
    badgeNumber: "",
    comments: ""
  });
  renderAgencies();
  persistDraft();
}

function renderAgencies() {
  const container = el("agencyBlocks");
  if (!container) return;

  container.innerHTML = "";

  state.agencies.forEach((agency) => {
    const block = document.createElement("div");
    block.className = "agency-block";
    block.innerHTML = `
      <div class="responder-card-top">
        <div class="responder-title">${escapeHtml(agency.type)}</div>
        <button class="tiny-btn" type="button" data-remove-agency="${agency.id}">Remove</button>
      </div>

      <div class="form-grid">
        ${agency.type === "Other" ? `
          <label>
            Other Agency Name
            <input type="text" data-agency-id="${agency.id}" data-field="otherName" value="${escapeHtml(agency.otherName)}" />
          </label>
        ` : ""}

        <label>
          Officer Name
          <input type="text" data-agency-id="${agency.id}" data-field="officerName" value="${escapeHtml(agency.officerName)}" />
        </label>

        <label>
          Contact Number
          <input type="text" data-agency-id="${agency.id}" data-field="contactNumber" value="${escapeHtml(agency.contactNumber)}" />
        </label>

        <label>
          Station
          <input type="text" data-agency-id="${agency.id}" data-field="station" value="${escapeHtml(agency.station)}" />
        </label>

        <label>
          Badge Number
          <input type="text" data-agency-id="${agency.id}" data-field="badgeNumber" value="${escapeHtml(agency.badgeNumber)}" />
        </label>

        <label class="full-width">
          Add Comments
          <textarea rows="2" data-agency-id="${agency.id}" data-field="comments">${escapeHtml(agency.comments)}</textarea>
        </label>
      </div>
    `;
    container.appendChild(block);
  });

  container.querySelectorAll("[data-agency-id]").forEach((field) => {
    field.addEventListener("input", (e) => {
      const agency = state.agencies.find((a) => a.id === e.target.dataset.agencyId);
      if (!agency) return;
      agency[e.target.dataset.field] = e.target.value;
      persistDraft();
    });
  });

  container.querySelectorAll("[data-remove-agency]").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.agencies = state.agencies.filter((a) => a.id !== btn.dataset.removeAgency);
      renderAgencies();
      persistDraft();
    });
  });
}

/* =========================================================
   VEHICLES
   ========================================================= */

function renderVehicles() {
  renderVehicleBlocks();
}

/* =========================================================
   SAVED REPORTS
   ========================================================= */

function saveReportSummary() {
  const title = buildReportTitle();
  const body = state.ui.reportPreview || generateReport();

  state.savedReports.unshift({
    id: cryptoRandomId(),
    title,
    body,
    createdAt: new Date().toISOString()
  });

  state.savedReports = state.savedReports.slice(0, CONFIG.MAX_REPORTS);
  saveSavedReports();
  renderSavedReports();
}

function deleteSavedReport(id) {
  state.savedReports = state.savedReports.filter((r) => r.id !== id);
  saveSavedReports();
  renderSavedReports();
}

function renderSavedReports() {
  const container = el("savedReportsList");
  if (!container) return;

  container.innerHTML = "";

  if (!state.savedReports.length) {
    container.innerHTML = `<div class="saved-report-item">No saved reports yet.</div>`;
    return;
  }

  state.savedReports.forEach((report) => {
    const item = document.createElement("div");
    item.className = "saved-report-item";
    item.innerHTML = `
      <div><strong>${escapeHtml(report.title)}</strong></div>
      <div>${new Date(report.createdAt).toLocaleString("en-AU")}</div>
      <div class="saved-report-actions">
        <button class="tiny-btn" type="button" data-load-report="${report.id}">Load Preview</button>
        <button class="tiny-btn" type="button" data-delete-report="${report.id}">Delete</button>
      </div>
    `;
    container.appendChild(item);
  });

  container.querySelectorAll("[data-load-report]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const report = state.savedReports.find((r) => r.id === btn.dataset.loadReport);
      if (!report) return;
      el("reportPreview").value = report.body;
    });
  });

  container.querySelectorAll("[data-delete-report]").forEach((btn) => {
    btn.addEventListener("click", () => deleteSavedReport(btn.dataset.deleteReport));
  });
}

/* =========================================================
   PROFILE MODAL
   ========================================================= */

function openSettingsModal() {
  if (!el("settingsModal")) return;
  el("profileName").value = state.profile.name;
  el("profileMemberNumber").value = state.profile.memberNumber;
  el("profileContactNumber").value = state.profile.contactNumber;
  el("profileEmail").value = state.profile.email;
  el("profileBrigade").value = state.profile.brigade;
  el("settingsModal").classList.remove("hidden");
}

function closeSettingsModal() {
  el("settingsModal")?.classList.add("hidden");
}

function saveProfileFromUi() {
  state.profile.name = el("profileName").value.trim();
  state.profile.memberNumber = el("profileMemberNumber").value.trim();
  state.profile.contactNumber = el("profileContactNumber").value.trim();
  state.profile.email = el("profileEmail").value.trim();
  state.profile.brigade = el("profileBrigade").value.trim() || "Connewarre";
  saveProfile();
  closeSettingsModal();
}

/* =========================================================
   REPORTS / VALIDATION
   ========================================================= */

function validateResponders() {
  return getAllResponders().some((r) => r.name.trim());
}

function validateReportReady() {
  const messages = [];
  if (!validateResponders()) messages.push("At least one responder must be entered.");
  if (isPrimaryConnJob() && !getAllResponders().some((r) => r.oic && r.name.trim())) {
    messages.push("OIC must be selected for primary Connewarre jobs.");
  }
  return { valid: messages.length === 0, messages };
}

function generateReport() {
  collectStructureData();
  collectHoseUse();

  const sections = [
    buildJobSection(),
    buildOicSection(),
    buildApplianceSections(),
    buildResponderSections(),
    formatStructureReport(),
    buildVehicleSection(),
    buildFirstAgencySection(),
    formatAgencySection(),
    buildHoseSection(),
    buildNotesSection(),
    buildIncidentFlagsSection(),
    buildSignoff()
  ].filter(Boolean);

  const report = sections.join("\n\n").trim();
  state.ui.reportPreview = report;
  if (el("reportPreview")) el("reportPreview").value = report;
  return report;
}

function buildJobSection() {
  return [
    "Job Details",
    lineIfValue("Event Number", state.incident.eventNumber),
    lineIfValue("Pager Date", formatDateForReport(state.incident.pagerDate)),
    lineIfValue("Pager Time", state.incident.pagerTime),
    lineIfValue("Brigade Code", state.incident.brigadeCode),
    lineIfValue("Primary / Support", state.incident.brigadeRole),
    lineIfValue("Incident Type / Class", state.incident.incidentType),
    lineIfValue("Code Level", state.incident.codeLevel),
    lineIfValue("Address", state.incident.address),
    lineIfValue("FIRS Code", state.incident.firsCode),
    lineIfValue("Brigades on Scene", state.incident.brigadesOnScene.join(", "))
  ].filter(Boolean).join("\n");
}

function buildOicSection() {
  const oic = getAllResponders().find((r) => r.oic && r.name.trim());
  if (!oic) return "";
  return [
    "OIC",
    lineIfValue("Name", appendBrigadeIfNeeded(oic)),
    lineIfValue("Phone", oic.phone)
  ].filter(Boolean).join("\n");
}

function buildApplianceSections() {
  const grouped = {};

  getAllResponders()
    .filter((r) => r.name.trim() && ["T1", "T2", "MTD P/T"].includes(r.destination))
    .forEach((r) => {
      if (!grouped[r.destination]) grouped[r.destination] = [];
      grouped[r.destination].push(r);
    });

  const ordered = ["T1", "T2", "MTD P/T"];
  const sectionBlocks = [];

  ordered.forEach((key) => {
    if (!grouped[key]?.length) return;
    const title = key === "T1" ? "Conn T1" : key === "T2" ? "Conn T2" : "MTD P/T";
    const lines = [title];
    grouped[key].forEach((r) => {
      lines.push(`- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)} | ${r.truckRole || "Crew"}`);
    });
    sectionBlocks.push(lines.join("\n"));
  });

  return sectionBlocks.length ? ["Appliances", ...sectionBlocks].join("\n") : "";
}

function buildResponderSections() {
  const direct = getAllResponders()
    .filter((r) => r.name.trim() && r.destination === "Direct")
    .map((r) => `- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)}`);

  const station = getAllResponders()
    .filter((r) => r.name.trim() && r.destination === "Station")
    .map((r) => `- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)}`);

  const parts = [];
  if (direct.length) parts.push(["Direct Responders", ...direct].join("\n"));
  if (station.length) parts.push(["Station Responders", ...station].join("\n"));
  return parts.join("\n\n");
}

function buildVehicleSection() {
  const vehicles = state.vehicles.filter((v) => v.make || v.model || v.rego || v.contactName);
  if (!vehicles.length) return "";

  const lines = ["Vehicle Section"];
  vehicles.forEach((v, index) => {
    lines.push(`Vehicle ${index + 1}`);
    lines.push(...[
      lineIfValue("Make", v.make === "Other" ? v.makeOther : v.make),
      lineIfValue("Model", v.model),
      lineIfValue("Rego", v.rego),
      lineIfValue("State", v.state),
      lineIfValue("Driver / Contact", v.contactName),
      lineIfValue("Contact Number", v.contactNumber),
      lineIfValue("Occupants", v.occupants),
      lineIfValue("Trapped", v.trapped),
      lineIfValue("Airbags Deployed", v.airbagsDeployed),
      lineIfValue("Vehicle Stability", v.stability),
      lineIfValue("Vehicle Type / Propulsion", v.propulsion === "Other" ? v.propulsionOther : v.propulsion)
    ].filter(Boolean));
  });

  return lines.join("\n");
}

function buildFirstAgencySection() {
  if (!state.incident.firstAgency) return "";
  const value = state.incident.firstAgency === "Other" ? state.incident.firstAgencyOther : state.incident.firstAgency;
  if (!value) return "";
  return ["First Agency On Scene", value].join("\n");
}

function buildHoseSection() {
  const used = Object.entries(state.hoseUse).filter(([, qty]) => String(qty).trim() !== "");
  if (!used.length) return "";
  return ["Hose Use", ...used.map(([type, qty]) => `${type}: ${qty}`)].join("\n");
}

function buildNotesSection() {
  return state.incident.notes.trim() ? ["Notes", state.incident.notes.trim()].join("\n") : "";
}

function buildIncidentFlagsSection() {
  const flags = [];
  if (state.incident.flags.membersBefore) flags.push("Members before 1st appliance");
  if (state.incident.flags.aarRequired) flags.push("AAR required");
  if (state.incident.flags.hotDebrief) flags.push("Hot debrief conducted");
  return flags.length ? ["Incident Flags", ...flags.map((f) => `- ${f}`)].join("\n") : "";
}

function buildSignoff() {
  return [
    "Sign-off",
    lineIfValue("Name", state.profile.name),
    lineIfValue("Brigade", state.profile.brigade),
    lineIfValue("CFA Member Number", state.profile.memberNumber),
    lineIfValue("Contact Number", state.profile.contactNumber)
  ].filter(Boolean).join("\n");
}

function formatAgencySection() {
  if (!state.agencies.length) return "";
  const lines = ["Agencies"];
  state.agencies.forEach((agency, index) => {
    const title = agency.type === "Other" ? (agency.otherName || "Other") : agency.type;
    lines.push(`${index + 1}. ${title}`);
    lines.push(...[
      lineIfValue("Officer", agency.officerName),
      lineIfValue("Contact", agency.contactNumber),
      lineIfValue("Station", agency.station),
      lineIfValue("Badge", agency.badgeNumber),
      lineIfValue("Comments", agency.comments)
    ].filter(Boolean));
  });
  return lines.join("\n");
}

/* =========================================================
   SEND / PAGE UI
   ========================================================= */

function handleFinish() {
  const validation = validateReportReady();
  if (!validation.valid) {
    setStatus("validationStatus", validation.messages.join(" "));
    el("finishActions")?.classList.add("hidden");
    return;
  }

  setStatus("validationStatus", "Report ready.");
  generateReport();
  updateQuietHoursWarning();
  el("finishActions")?.classList.remove("hidden");
}

async function sendSms() {
  const report = generateReport();
  try {
    await navigator.clipboard.writeText(report);
  } catch (error) {
    console.warn("Clipboard copy failed", error);
  }
  window.location.href = `sms:?body=${encodeURIComponent(report)}`;
}

function sendEmail() {
  const report = generateReport();
  const subject = buildReportTitle();
  const email = state.profile.email || "";
  window.location.href = `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(report)}`;
}

function saveReportLocally() {
  generateReport();
  saveReportSummary();
}

function showPage(pageId) {
  state.ui.currentPage = pageId;
  document.querySelectorAll(".page").forEach((page) => {
    page.classList.toggle("active", page.id === pageId);
  });

  document.querySelectorAll(".nav-btn[data-page]").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.page === pageId);
  });

  if (pageId === "sendPage") {
    const validation = validateReportReady();
    setStatus("validationStatus", validation.valid ? "Ready to finish." : validation.messages.join(" "));
    generateReport();
    updateQuietHoursWarning();
  }

  persistDraft();
}

function updateQuietHoursWarning() {
  const warning = el("quietHoursWarning");
  if (!warning) return;
  const hour = new Date().getHours();
  const quiet = hour >= CONFIG.QUIET_HOURS.start || hour < CONFIG.QUIET_HOURS.end;
  warning.classList.toggle("hidden", !quiet);
}

/* =========================================================
   SYNC UI
   ========================================================= */

function syncIncidentFieldsToUi() {
  if (el("eventNumber")) el("eventNumber").value = state.incident.eventNumber;
  if (el("pagerDate")) el("pagerDate").value = state.incident.pagerDate;
  if (el("pagerTime")) el("pagerTime").value = state.incident.pagerTime;
  if (el("brigadeCode")) el("brigadeCode").value = state.incident.brigadeCode;
  if (el("brigadeRole")) el("brigadeRole").value = state.incident.brigadeRole;
  if (el("incidentType")) el("incidentType").value = state.incident.incidentType;
  if (el("codeLevel")) el("codeLevel").value = state.incident.codeLevel;
  if (el("address")) el("address").value = state.incident.address;
  if (el("firsCode")) el("firsCode").value = state.incident.firsCode;
  if (el("notes")) el("notes").value = state.incident.notes;
  if (el("firstAgencySelect")) el("firstAgencySelect").value = state.incident.firstAgency;
  if (el("firstAgencyOther")) {
    el("firstAgencyOther").value = state.incident.firstAgencyOther;
    el("firstAgencyOther").classList.toggle("hidden", state.incident.firstAgency !== "Other");
  }

  updateFirsLabel();
  renderIncidentFlagButtons();
  renderPagerPreviews();
}

function syncAllFieldsToUi() {
  syncIncidentFieldsToUi();

  const structureRadio = document.querySelector(`input[name="structureRequired"][value="${state.structure.required ? "yes" : "no"}"]`);
  if (structureRadio) structureRadio.checked = true;
  if (el("structureFormWrap")) el("structureFormWrap").classList.toggle("hidden", !state.structure.required);

  [
    "structurePropertyType", "structureLevels", "structureOccupancy", "structureConstruction",
    "fireAreaOrigin", "fireAreaExtent", "fireAreaComments",
    "fireFuelLoad", "fireBehaviour", "fireMaterialComments",
    "detectionType", "alarmActivated", "detectionComments",
    "suppressionType", "suppressionWorked", "suppressionComments",
    "portableExtinguisher", "portableOther", "portableComments"
  ].forEach((id) => {
    if (el(id)) {
      const key = id
        .replace("structure", "")
        .replace("fire", "fire")
        .replace("portable", "portable");
    }
  });

  if (el("structurePropertyType")) el("structurePropertyType").value = state.structure.propertyType;
  if (el("structureLevels")) el("structureLevels").value = state.structure.levels;
  if (el("structureOccupancy")) el("structureOccupancy").value = state.structure.occupancy;
  if (el("structureConstruction")) el("structureConstruction").value = state.structure.construction;
  if (el("fireAreaOrigin")) el("fireAreaOrigin").value = state.structure.fireAreaOrigin;
  if (el("fireAreaExtent")) el("fireAreaExtent").value = state.structure.fireAreaExtent;
  if (el("fireAreaComments")) el("fireAreaComments").value = state.structure.fireAreaComments;
  if (el("fireFuelLoad")) el("fireFuelLoad").value = state.structure.fireFuelLoad;
  if (el("fireBehaviour")) el("fireBehaviour").value = state.structure.fireBehaviour;
  if (el("fireMaterialComments")) el("fireMaterialComments").value = state.structure.fireMaterialComments;
  if (el("detectionType")) el("detectionType").value = state.structure.detectionType;
  if (el("alarmActivated")) el("alarmActivated").value = state.structure.alarmActivated;
  if (el("detectionComments")) el("detectionComments").value = state.structure.detectionComments;
  if (el("suppressionType")) el("suppressionType").value = state.structure.suppressionType;
  if (el("suppressionWorked")) el("suppressionWorked").value = state.structure.suppressionWorked;
  if (el("suppressionComments")) el("suppressionComments").value = state.structure.suppressionComments;
  if (el("portableExtinguisher")) el("portableExtinguisher").value = state.structure.portableExtinguisher;
  if (el("portableOther")) el("portableOther").value = state.structure.portableOther;
  if (el("portableComments")) el("portableComments").value = state.structure.portableComments;

  if (el("hose64")) el("hose64").value = state.hoseUse["64"];
  if (el("hose38")) el("hose38").value = state.hoseUse["38"];
  if (el("hose25")) el("hose25").value = state.hoseUse["25"];
  if (el("hoseLiveReel")) el("hoseLiveReel").value = state.hoseUse["Live Reel"];

  renderBrigadesOnScene();
}

/* =========================================================
   SERVICE WORKER
   ========================================================= */

function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch((error) => {
      console.warn("SW registration failed", error);
    });
  }
}
