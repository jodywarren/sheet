/* =========================================================
   TURNOUT SHEET - CFA DIGITAL TURNOUT SHEET
   Full replacement file
   ========================================================= */

/* =========================================================
   1. CONFIGURATION BLOCK
   ========================================================= */
const CONFIG = {
  APP_NAME: "Turnout Sheet",
  MAX_REPORTS: 10,
  MAX_VEHICLES: 5,
  QUIET_HOURS: {
    start: 22,
    end: 7,
  },
  MEMBER_FILES: [
    { key: "CONN", url: "CONN.members.json", label: "Connewarre" },
    { key: "GROV", url: "GROV.members.json", label: "Grovedale" },
    { key: "FRES", url: "FRES.members.json", label: "Freshwater Creek" },
  ],
  BRIGADE_SCENE_OPTIONS: ["CONN", "GROV", "FRES", "P64", "P32", "P31", "P33", "P35", "P41"],
  VEHICLE_MAKES: [
    "Toyota", "Ford", "Holden", "Mazda", "Hyundai", "Kia", "Mitsubishi", "Nissan",
    "Subaru", "Volkswagen", "BMW", "Mercedes", "Audi", "Isuzu", "LDV", "Tesla",
    "BYD", "MG", "GWM", "Volvo", "Skoda", "Jeep", "Suzuki", "Lexus", "Ram", "Other"
  ],
  STATES: ["VIC", "NSW", "SA", "TAS", "ACT", "QLD", "WA", "NT"],
  INCIDENT_TYPES: ["ALAR", "STRU", "NONS", "INCI", "G&SC"],
  CODE_LEVELS: ["C1", "C3"],
  FIRST_AGENCIES: ["CFA", "FRV", "Police", "Ambulance", "SES", "PowerCor", "Gas", "Other"],
  AGENCY_TYPES: ["Police", "Ambulance", "PowerCor", "Gas", "SES", "Local Government", "Other"],
  MVA_KEYWORDS: [
    "INCI", "MVA", "MVC", "MVI", "VEHICLE", "COLLISION", "CRASH",
    "ROLLOVER", "TRAPPED", "ROAD RESCUE", "CAR INTO", "TRUCK INTO", "VEHICLE INTO"
  ],
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
  OCCUPANTS: ["0", "1", "2", "3", "4+", "Unknown"],
  TRI_STATE: ["Yes", "No", "Unknown"],
  VEHICLE_STABILITY: ["Upright", "On side", "On roof", "Unknown"],
  RESPONSE_DESTINATIONS_CONNEWARRE: ["T1", "T2", "Station", "Direct"],
  RESPONSE_DESTINATIONS_MTD: ["MTD P/T", "Station", "Direct"],
  TRUCK_ROLES: ["Driver", "CL"],
};

const STORAGE_KEYS = {
  PROFILE: "turnout_profile",
  SAVED_REPORTS: "turnout_saved_reports",
  DRAFT_STATE: "turnout_draft_state"
};

/* =========================================================
   2. APPLICATION STATE
   ========================================================= */
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
  vehicles: [],
  agencies: [],
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
  savedReports: [],
  memberLists: {
    CONN: [],
    GROV: [],
    FRES: []
  },
  ui: {
    currentPage: "incidentPage",
    oicResponderId: "",
    oicName: "",
    oicPhone: "",
    pagerImageFile: null,
    reportPreview: ""
  }
};

/* =========================================================
   3. INITIALISATION
   ========================================================= */
document.addEventListener("DOMContentLoaded", init);

async function init() {
  await loadMemberLists();
  loadProfile();
  loadSavedReports();
  restoreDraftState();
  ensureMinimumResponderRows();
  bindEventListeners();
  renderResponders();
  renderVehicleBlocks();
  renderAgencyBlocks();
  renderSavedReports();
  renderBrigadesOnScene();
  syncAllFieldsToUi();
  updateHeaderOicStatus();
  showPage(state.ui.currentPage || "incidentPage");
  registerServiceWorker();
}

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

function bindEventListeners() {
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
    btn.addEventListener("click", () => toggleSection(btn.dataset.target));
  });

  document.querySelectorAll(".accordion-toggle").forEach((btn) => {
    btn.addEventListener("click", () => toggleAccordion(btn.dataset.target));
  });

  document.getElementById("settingsBtn").addEventListener("click", openSettingsModal);
  document.getElementById("closeSettingsBtn").addEventListener("click", closeSettingsModal);
  document.getElementById("saveProfileBtn").addEventListener("click", saveProfileFromUi);

  document.getElementById("pagerUpload").addEventListener("change", handleImageUpload);
  document.getElementById("runOcrBtn").addEventListener("click", runOcrPipeline);

  bindIncidentListeners();
  bindStructureListeners();
  bindVehicleListeners();
  bindAgencyListeners();
  bindSendListeners();
  bindSceneBrigadesListeners();
}

function bindIncidentListeners() {
  const bindings = [
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

  bindings.forEach(([id, key]) => {
    document.getElementById(id).addEventListener("input", (e) => {
      state.incident[key] = e.target.value;
      if (key === "brigadeCode") updateBrigadeRole();
      if (key === "incidentType") updateStructureVisibilityByIncidentType();
      if (key === "firsCode") updateFirsLabel();
      persistDraftState();
    });
  });

  document.getElementById("resetFirsBtn").addEventListener("click", () => {
    state.incident.firsCode = "";
    document.getElementById("firsCode").value = "";
    updateFirsLabel();
    persistDraftState();
  });

  document.getElementById("firstAgencySelect").addEventListener("change", (e) => {
    state.incident.firstAgency = e.target.value;
    const other = document.getElementById("firstAgencyOther");
    other.classList.toggle("hidden", e.target.value !== "Other");
    persistDraftState();
  });

  document.getElementById("firstAgencyOther").addEventListener("input", (e) => {
    state.incident.firstAgencyOther = e.target.value;
    persistDraftState();
  });

  document.getElementById("flagMembersBeforeBtn").addEventListener("click", () => {
    state.incident.flags.membersBefore = !state.incident.flags.membersBefore;
    renderIncidentFlagButtons();
    persistDraftState();
  });

  document.getElementById("flagAarRequiredBtn").addEventListener("click", () => {
    state.incident.flags.aarRequired = !state.incident.flags.aarRequired;
    renderIncidentFlagButtons();
    persistDraftState();
  });

  document.getElementById("flagHotDebriefBtn").addEventListener("click", () => {
    state.incident.flags.hotDebrief = !state.incident.flags.hotDebrief;
    renderIncidentFlagButtons();
    persistDraftState();
  });
}

function bindSceneBrigadesListeners() {
  document.getElementById("brigadeOnSceneSelect").addEventListener("change", (e) => {
    const otherInput = document.getElementById("brigadeOnSceneOther");
    otherInput.classList.toggle("hidden", e.target.value !== "Other");
  });

  document.getElementById("addBrigadeOnSceneBtn").addEventListener("click", addBrigadeOnSceneFromUi);
}

function addBrigadeOnSceneFromUi() {
  const select = document.getElementById("brigadeOnSceneSelect");
  const other = document.getElementById("brigadeOnSceneOther");
  let value = select.value;

  if (!value) return;

  if (value === "Other") {
    value = other.value.trim().toUpperCase();
    if (!value) return;
  }

  if (!state.incident.brigadesOnScene.includes(value)) {
    state.incident.brigadesOnScene.push(value);
  }

  select.value = "";
  other.value = "";
  other.classList.add("hidden");
  renderBrigadesOnScene();
  persistDraftState();
}

function removeBrigadeOnScene(code) {
  state.incident.brigadesOnScene = state.incident.brigadesOnScene.filter((b) => b !== code);
  renderBrigadesOnScene();
  persistDraftState();
}

function renderBrigadesOnScene() {
  const container = document.getElementById("brigadesOnSceneChips");
  container.innerHTML = "";

  state.incident.brigadesOnScene.forEach((code) => {
    const chip = document.createElement("div");
    chip.className = "scene-chip";
    chip.innerHTML = `<span>${escapeHtml(code)}</span><button type="button" aria-label="Remove ${escapeHtml(code)}">×</button>`;
    chip.querySelector("button").addEventListener("click", () => removeBrigadeOnScene(code));
    container.appendChild(chip);
  });
}

function bindStructureListeners() {
  document.querySelectorAll('input[name="structureRequired"]').forEach((radio) => {
    radio.addEventListener("change", () => {
      state.structure.required = document.querySelector('input[name="structureRequired"]:checked').value === "yes";
      document.getElementById("structureFormWrap").classList.toggle("hidden", !state.structure.required);
      persistDraftState();
    });
  });

  const structureFields = [
    "structurePropertyType", "structureLevels", "structureOccupancy", "structureConstruction",
    "fireAreaOrigin", "fireAreaExtent", "fireAreaComments",
    "fireFuelLoad", "fireBehaviour", "fireMaterialComments",
    "detectionType", "alarmActivated", "detectionComments",
    "suppressionType", "suppressionWorked", "suppressionComments",
    "portableExtinguisher", "portableOther", "portableComments"
  ];

  structureFields.forEach((id) => {
    document.getElementById(id).addEventListener("input", collectStructureData);
  });

  document.getElementById("hose64").addEventListener("input", collectHoseUse);
  document.getElementById("hose38").addEventListener("input", collectHoseUse);
  document.getElementById("hose25").addEventListener("input", collectHoseUse);
  document.getElementById("hoseLiveReel").addEventListener("input", collectHoseUse);
}

function bindVehicleListeners() {
  document.getElementById("vehicleCount").addEventListener("change", (e) => {
    setVehicleCount(Number(e.target.value));
  });
}

function bindAgencyListeners() {
  document.getElementById("addAgencyBtn").addEventListener("click", () => {
    const type = document.getElementById("agencyType").value;
    if (!type) return;
    addAgencyBlock(type);
    document.getElementById("agencyType").value = "";
  });
}

function bindSendListeners() {
  document.getElementById("finishBtn").addEventListener("click", handleFinish);
  document.getElementById("sendSmsBtn").addEventListener("click", sendSms);
  document.getElementById("sendEmailBtn").addEventListener("click", sendEmail);
  document.getElementById("saveLocalBtn").addEventListener("click", saveReportLocally);
}

/* =========================================================
   4. OCR MODULE
   ========================================================= */
function handleImageUpload(e) {
  const file = e.target.files?.[0] || null;
  state.ui.pagerImageFile = file;
  setOcrStatus(file ? `Selected: ${file.name}` : "");
  if (file) {
    runOcrPipeline();
  }
}

async function runOcrPipeline() {
  if (!state.ui.pagerImageFile) {
    setOcrStatus("Please upload a pager screenshot first.");
    return;
  }

  setOcrStatus("Scanning screenshot...");
  try {
    const imageBitmap = await createImageBitmap(state.ui.pagerImageFile);
    const rectangles = findPagerRectangles(imageBitmap);
    const selected = rectangles.length ? selectBestPagerRectangle(rectangles) : null;
    const cropCanvas = cropPagerRectangle(imageBitmap, selected);
    const rawText = await runPagerOCR(cropCanvas);
    const parsed = parsePagerMessage(rawText);

    if (!validatePagerText(parsed)) {
      setOcrStatus("Scan ran, but the pager details were not strong enough to safely populate fields.");
      return;
    }

    populateIncidentFields(parsed);
    state.incident.pagerRawText = rawText;
    setOcrStatus("Pager details extracted.");
    persistDraftState();
  } catch (error) {
    console.error(error);
    setOcrStatus("Scan failed. Try a cleaner screenshot or enter details manually.");
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
      const index = (y * width + x) * 4;
      const avg = (data[index] + data[index + 1] + data[index + 2]) / 3;
      if (avg < 180) darkPixels++;
    }
    rowScores.push(darkPixels / width);
  }

  const candidateBands = [];
  let start = null;
  for (let i = 0; i < rowScores.length; i++) {
    if (rowScores[i] > 0.12 && start === null) start = i;
    if ((rowScores[i] <= 0.12 || i === rowScores.length - 1) && start !== null) {
      const end = i;
      if (end - start > 120) {
        candidateBands.push({ x: 0, y: start, width, height: end - start });
      }
      start = null;
    }
  }

  if (!candidateBands.length) return [{ x: 0, y: 0, width, height }];
  return candidateBands.slice(0, 5);
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
  const text = (rawText || "").replace(/\r/g, "\n").trim();
  const upper = text.toUpperCase();
  const compact = upper.replace(/\n/g, " ");

  const eventNumber = (upper.match(/\bF\d{9}\b/) || [])[0] || "";

  const dateMatch = upper.match(/\b(\d{2})[\/\-](\d{2})[\/\-](\d{2,4})\b/);
  let pagerDate = "";
  if (dateMatch) {
    const day = dateMatch[1];
    const month = dateMatch[2];
    let year = dateMatch[3];
    if (year.length === 2) year = `20${year}`;
    pagerDate = `${year}-${month}-${day}`;
  }

  const timeMatch = upper.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
  const pagerTime = timeMatch ? `${timeMatch[1].padStart(2, "0")}:${timeMatch[2]}` : "";

  const brigadeCode = (upper.match(/\b(CONN|GROV|FRES)\d\b/) || [])[0] || "";
  const incidentClassMatch = upper.match(/\b(ALAR|STRU|NONS|INCI|G&SC)(C1|C3)?\b/);
  const incidentType = incidentClassMatch?.[1] || "";
  const codeLevel = incidentClassMatch?.[2] || "";

  const address = extractPagerAddress(compact);
  const brigadesOnScene = extractKnownBrigades(compact);

  return {
    rawText: text,
    upper,
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
  const slashMatch = text.match(/(\d{1,5}\s+[A-Z0-9\s.'-]+?\s(?:ST|RD|AV))\s*\//i);
  if (slashMatch) return normaliseSpacing(slashMatch[1]);

  const fallback = text.match(/(\d{1,5}\s+[A-Z0-9\s.'-]+?\s(?:ST|RD|AV))\b/i);
  if (fallback) return normaliseSpacing(fallback[1]);

  return "";
}

function extractKnownBrigades(text) {
  const found = [];
  CONFIG.BRIGADE_SCENE_OPTIONS.forEach((code) => {
    if (text.includes(code) && !found.includes(code)) {
      found.push(code);
    }
  });
  return found;
}

function validatePagerText(parsed) {
  return !!(parsed.eventNumber || parsed.address || parsed.brigadeCode);
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
   5. INCIDENT MODULE
   ========================================================= */
function updateBrigadeRole() {
  const code = (state.incident.brigadeCode || "").toUpperCase().trim();
  state.incident.brigadeRole = code.startsWith("CONN") ? "Primary" : code ? "Support" : "";
  document.getElementById("brigadeRole").value = state.incident.brigadeRole;
  persistDraftState();
}

function updateFirsLabel() {
  document.getElementById("firsLabel").textContent = state.incident.firsCode ? "FIRS Code" : "Paste FIRS";
}

function evaluateMvaAutoTrigger(text) {
  const hasMatch = CONFIG.MVA_KEYWORDS.some((keyword) => text.includes(keyword));
  if (hasMatch) {
    document.getElementById("vehicleSection").classList.add("open");
    if (state.vehicles.length === 0) setVehicleCount(1);
  }
}

function updateStructureVisibilityByIncidentType() {
  const isStructure = state.incident.incidentType === "STRU";
  if (!isStructure) {
    state.structure.required = false;
    document.querySelector('input[name="structureRequired"][value="no"]').checked = true;
    document.getElementById("structureFormWrap").classList.add("hidden");
  }
  persistDraftState();
}

function renderIncidentFlagButtons() {
  document.getElementById("flagMembersBeforeBtn").classList.toggle("mini-chip-active", state.incident.flags.membersBefore);
  document.getElementById("flagAarRequiredBtn").classList.toggle("mini-chip-active", state.incident.flags.aarRequired);
  document.getElementById("flagHotDebriefBtn").classList.toggle("mini-chip-active", state.incident.flags.hotDebrief);
}

/* =========================================================
   6. VEHICLE MODULE
   ========================================================= */
function setVehicleCount(count) {
  const safeCount = Math.max(0, Math.min(CONFIG.MAX_VEHICLES, count));
  while (state.vehicles.length < safeCount) state.vehicles.push(createBlankVehicle(state.vehicles.length + 1));
  while (state.vehicles.length > safeCount) state.vehicles.pop();
  renderVehicleBlocks();
  persistDraftState();
}

function createBlankVehicle(index) {
  return {
    id: cryptoRandomId(),
    index,
    make: "",
    makeOther: "",
    model: "",
    rego: "",
    state: "VIC",
    contactName: "",
    contactNumber: "",
    occupants: "Unknown",
    trapped: "Unknown",
    airbagsDeployed: "Unknown",
    stability: "Unknown",
    propulsion: "Unknown",
    propulsionOther: ""
  };
}

function renderVehicleBlocks() {
  const container = document.getElementById("vehicleBlocks");
  container.innerHTML = "";

  state.vehicles.forEach((vehicle, i) => {
    const block = document.createElement("div");
    block.className = "vehicle-block";
    block.innerHTML = `
      <h4>Vehicle ${i + 1}</h4>
      <div class="form-grid">
        <label>
          Vehicle Make
          <input type="text" list="vehicleMakeList" data-vehicle-id="${vehicle.id}" data-field="make" value="${escapeHtml(vehicle.make)}" placeholder="Search make" />
        </label>

        <label class="${vehicle.make === "Other" ? "" : "hidden"}" data-other-make-wrap="${vehicle.id}">
          Other Make
          <input type="text" data-vehicle-id="${vehicle.id}" data-field="makeOther" value="${escapeHtml(vehicle.makeOther)}" />
        </label>

        <label>Vehicle Model<input type="text" data-vehicle-id="${vehicle.id}" data-field="model" value="${escapeHtml(vehicle.model)}" /></label>
        <label>Rego<input type="text" data-vehicle-id="${vehicle.id}" data-field="rego" value="${escapeHtml(vehicle.rego)}" /></label>

        <label>
          State
          <select data-vehicle-id="${vehicle.id}" data-field="state">
            ${CONFIG.STATES.map((s) => `<option value="${s}" ${vehicle.state === s ? "selected" : ""}>${s}</option>`).join("")}
          </select>
        </label>

        <label>Driver / Contact Name<input type="text" data-vehicle-id="${vehicle.id}" data-field="contactName" value="${escapeHtml(vehicle.contactName)}" /></label>
        <label>Driver / Contact Number<input type="text" data-vehicle-id="${vehicle.id}" data-field="contactNumber" value="${escapeHtml(vehicle.contactNumber)}" /></label>

        <label>
          Occupants
          <select data-vehicle-id="${vehicle.id}" data-field="occupants">
            ${CONFIG.OCCUPANTS.map((v) => `<option value="${v}" ${vehicle.occupants === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label>
          Trapped
          <select data-vehicle-id="${vehicle.id}" data-field="trapped">
            ${CONFIG.TRI_STATE.map((v) => `<option value="${v}" ${vehicle.trapped === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label>
          Airbags Deployed
          <select data-vehicle-id="${vehicle.id}" data-field="airbagsDeployed">
            ${CONFIG.TRI_STATE.map((v) => `<option value="${v}" ${vehicle.airbagsDeployed === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label>
          Vehicle Stability
          <select data-vehicle-id="${vehicle.id}" data-field="stability">
            ${CONFIG.VEHICLE_STABILITY.map((v) => `<option value="${v}" ${vehicle.stability === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label>
          Vehicle Type / Propulsion
          <select data-vehicle-id="${vehicle.id}" data-field="propulsion">
            ${CONFIG.VEHICLE_PROPULSION.map((v) => `<option value="${escapeHtml(v)}" ${vehicle.propulsion === v ? "selected" : ""}>${escapeHtml(v)}</option>`).join("")}
          </select>
        </label>

        <label class="${vehicle.propulsion === "Other" ? "" : "hidden"}" data-other-propulsion-wrap="${vehicle.id}">
          Other Propulsion
          <input type="text" data-vehicle-id="${vehicle.id}" data-field="propulsionOther" value="${escapeHtml(vehicle.propulsionOther)}" />
        </label>
      </div>
    `;
    container.appendChild(block);
  });

  renderVehicleMakesDatalist();
  bindVehicleBlockInputs();
  document.getElementById("vehicleCount").value = String(state.vehicles.length);
}

function renderVehicleMakesDatalist() {
  let list = document.getElementById("vehicleMakeList");
  if (!list) {
    list = document.createElement("datalist");
    list.id = "vehicleMakeList";
    document.body.appendChild(list);
  }
  list.innerHTML = CONFIG.VEHICLE_MAKES.map((make) => `<option value="${make}"></option>`).join("");
}

function bindVehicleBlockInputs() {
  document.querySelectorAll("[data-vehicle-id]").forEach((el) => {
    el.addEventListener("input", updateVehicleField);
    el.addEventListener("change", updateVehicleField);
  });
}

function updateVehicleField(e) {
  const vehicle = state.vehicles.find((v) => v.id === e.target.dataset.vehicleId);
  if (!vehicle) return;
  vehicle[e.target.dataset.field] = e.target.value;

  const otherMakeWrap = document.querySelector(`[data-other-make-wrap="${vehicle.id}"]`);
  const otherPropulsionWrap = document.querySelector(`[data-other-propulsion-wrap="${vehicle.id}"]`);
  if (otherMakeWrap) otherMakeWrap.classList.toggle("hidden", vehicle.make !== "Other");
  if (otherPropulsionWrap) otherPropulsionWrap.classList.toggle("hidden", vehicle.propulsion !== "Other");

  persistDraftState();
}

/* =========================================================
   7. STRUCTURE FIRE MODULE
   ========================================================= */
function collectStructureData() {
  state.structure.propertyType = document.getElementById("structurePropertyType").value;
  state.structure.levels = document.getElementById("structureLevels").value;
  state.structure.occupancy = document.getElementById("structureOccupancy").value;
  state.structure.construction = document.getElementById("structureConstruction").value;
  state.structure.fireAreaOrigin = document.getElementById("fireAreaOrigin").value;
  state.structure.fireAreaExtent = document.getElementById("fireAreaExtent").value;
  state.structure.fireAreaComments = document.getElementById("fireAreaComments").value;
  state.structure.fireFuelLoad = document.getElementById("fireFuelLoad").value;
  state.structure.fireBehaviour = document.getElementById("fireBehaviour").value;
  state.structure.fireMaterialComments = document.getElementById("fireMaterialComments").value;
  state.structure.detectionType = document.getElementById("detectionType").value;
  state.structure.alarmActivated = document.getElementById("alarmActivated").value;
  state.structure.detectionComments = document.getElementById("detectionComments").value;
  state.structure.suppressionType = document.getElementById("suppressionType").value;
  state.structure.suppressionWorked = document.getElementById("suppressionWorked").value;
  state.structure.suppressionComments = document.getElementById("suppressionComments").value;
  state.structure.portableExtinguisher = document.getElementById("portableExtinguisher").value;
  state.structure.portableOther = document.getElementById("portableOther").value;
  state.structure.portableComments = document.getElementById("portableComments").value;
  persistDraftState();
}

function formatStructureReport() {
  if (!(state.incident.incidentType === "STRU" && state.structure.required)) return "";

  const s = state.structure;
  return [
    "Structure Fire",
    lineIfValue("Property Type", s.propertyType),
    lineIfValue("Building Height / Levels", s.levels),
    lineIfValue("Occupancy", s.occupancy),
    lineIfValue("Construction Type", s.construction),
    lineIfValue("Area of Origin", s.fireAreaOrigin),
    lineIfValue("Extent of Fire", s.fireAreaExtent),
    lineIfValue("Fire Area Comments", s.fireAreaComments),
    lineIfValue("Main Fuel Load", s.fireFuelLoad),
    lineIfValue("Smoke / Fire Behaviour", s.fireBehaviour),
    lineIfValue("Material Comments", s.fireMaterialComments),
    lineIfValue("Detection Type", s.detectionType),
    lineIfValue("Alarm Activated", s.alarmActivated),
    lineIfValue("Detection Comments", s.detectionComments),
    lineIfValue("Suppression Type", s.suppressionType),
    lineIfValue("Suppression Worked", s.suppressionWorked),
    lineIfValue("Suppression Comments", s.suppressionComments),
    lineIfValue("Extinguisher Use", s.portableExtinguisher),
    lineIfValue("Portable Other", s.portableOther),
    lineIfValue("Portable Equipment Comments", s.portableComments)
  ].filter(Boolean).join("\n");
}

function collectHoseUse() {
  state.hoseUse["64"] = document.getElementById("hose64").value;
  state.hoseUse["38"] = document.getElementById("hose38").value;
  state.hoseUse["25"] = document.getElementById("hose25").value;
  state.hoseUse["Live Reel"] = document.getElementById("hoseLiveReel").value;
  persistDraftState();
}

/* =========================================================
   8. RESPONDER MODULE
   ========================================================= */
function ensureMinimumResponderRows() {
  if (!Array.isArray(state.responders.connewarre) || !state.responders.connewarre.length) {
    state.responders.connewarre = [createResponder("connewarre")];
  }
  if (!Array.isArray(state.responders.mtd) || !state.responders.mtd.length) {
    state.responders.mtd = [createResponder("mtd")];
  }
}

function createResponder(group) {
  return {
    id: cryptoRandomId(),
    group,
    brigade: group === "connewarre" ? "CONN" : "",
    name: "",
    phone: "",
    destination: "",
    truckRole: "",
    ba: false,
    injury: false,
    oic: false
  };
}

function renderResponders() {
  renderConnewarreResponders();
  renderMtdResponders();
  renderOtherResponderLists();
  updateHeaderOicStatus();
}

function renderConnewarreResponders() {
  const container = document.getElementById("connewarreResponders");
  container.innerHTML = `<div class="responder-list-wrap"></div>`;
  const list = container.firstElementChild;

  state.responders.connewarre.forEach((responder, index) => {
    list.appendChild(buildResponderRow(responder, index));
  });

  bindResponderEvents(container);
}

function renderMtdResponders() {
  const container = document.getElementById("mtdResponders");
  container.innerHTML = `<div class="responder-list-wrap"></div>`;
  const list = container.firstElementChild;

  state.responders.mtd.forEach((responder, index) => {
    list.appendChild(buildResponderRow(responder, index));
  });

  bindResponderEvents(container);
}

function renderOtherResponderLists() {
  const directContainer = document.getElementById("directRespondersContainer");
  const stationContainer = document.getElementById("stationRespondersContainer");

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
  wrapper.innerHTML = buildResponderRowHtml(responder, index);
  return wrapper.firstElementChild;
}

function buildResponderRowHtml(responder, index) {
  const isConn = responder.group === "connewarre";
  const isMtd = responder.group === "mtd";
  const nameListId = `member-list-${responder.id}`;
  const badge = isMtd && responder.brigade ? `<div class="responder-badge">${formatBrigadeName(responder.brigade)}</div>` : "";
  const destinations = isConn ? CONFIG.RESPONSE_DESTINATIONS_CONNEWARRE : CONFIG.RESPONSE_DESTINATIONS_MTD;
  const hasName = !!responder.name.trim();

  const destinationButtons = destinations.map((dest) => {
    const disabled = !hasName;
    const active = responder.destination === dest;
    return `
      <button
        type="button"
        class="mini-chip ${active ? "mini-chip-active" : ""}"
        data-action="destination"
        data-responder-id="${responder.id}"
        data-value="${dest}"
        ${disabled ? "disabled" : ""}
      >${dest}</button>
    `;
  }).join("");

  const showTruckRoles = responder.destination === "T1" || responder.destination === "T2";
  const roleButtons = showTruckRoles ? CONFIG.TRUCK_ROLES.map((role) => {
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
  }).join("") : "";

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

  return `
    <div class="responder-row" data-responder-id="${responder.id}">
      <div class="responder-row-top">
        <div class="responder-name-wrap">
          <input
            type="text"
            list="${nameListId}"
            placeholder="Name"
            data-action="name-input"
            data-responder-id="${responder.id}"
            value="${escapeHtml(responder.name)}"
          />
          <datalist id="${nameListId}">
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

      ${showAddMember ? `
        <button type="button" class="secondary-btn compact-add-btn" data-action="add-member" data-group="${responder.group}">Add Member</button>
      ` : ""}
    </div>
  `;
}

function bindResponderEvents(container) {
  container.querySelectorAll('[data-action="name-input"]').forEach((input) => {
    input.addEventListener("input", handleResponderNameTyping);
    input.addEventListener("change", handleResponderNameCommit);
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

  persistDraftState();
  renderResponders();
}

function handleDestinationSelect(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  responder.destination = e.target.dataset.value;
  if (!(responder.destination === "T1" || responder.destination === "T2")) {
    responder.truckRole = "";
  }

  persistDraftState();
  renderResponders();
}

function handleTruckRoleSelect(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  const role = e.target.dataset.value;
  if (isTruckRoleUnavailable(responder, responder.destination, role)) return;
  responder.truckRole = responder.truckRole === role ? "" : role;

  persistDraftState();
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

  persistDraftState();
  renderResponders();
}

function handleAddMember(e) {
  const group = e.target.dataset.group;
  state.responders[group].push(createResponder(group));
  persistDraftState();
  renderResponders();
}

function handleRemoveMember(e) {
  const group = e.target.dataset.group;
  const responderId = e.target.dataset.responderId;
  state.responders[group] = state.responders[group].filter((r) => r.id !== responderId);

  if (!state.responders[group].length) {
    state.responders[group].push(createResponder(group));
  }

  persistDraftState();
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

function updateHeaderOicStatus() {
  const el = document.getElementById("oicStatus");
  const oic = getAllResponders().find((r) => r.oic && r.name.trim());

  if (!oic) {
    el.textContent = "APPOINT OIC";
    el.classList.add("oic-missing");
    state.ui.oicResponderId = "";
    state.ui.oicName = "";
    state.ui.oicPhone = "";
    return;
  }

  oic.phone = oic.phone || getResponderPhoneFallback(oic) || "";
  state.ui.oicResponderId = oic.id;
  state.ui.oicName = oic.name;
  state.ui.oicPhone = oic.phone;
  el.textContent = `OIC: ${oic.name}${oic.phone ? " – " + oic.phone : ""}`;
  el.classList.remove("oic-missing");
}

function getAllResponders() {
  return [
    ...state.responders.connewarre,
    ...state.responders.mtd
  ];
}

function findResponderById(id) {
  return getAllResponders().find((r) => r.id === id);
}

function getResponderPhoneFallback(responder) {
  if (responder.brigade) {
    return getMemberPhone(responder.brigade, responder.name) || responder.phone || "";
  }
  return responder.phone || "";
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
  const all = [
    ...((state.memberLists.CONN || []).map((m) => ({ ...m, brigade: "CONN" }))),
    ...((state.memberLists.GROV || []).map((m) => ({ ...m, brigade: "GROV" }))),
    ...((state.memberLists.FRES || []).map((m) => ({ ...m, brigade: "FRES" })))
  ];

  const names = all.map((member) => {
    const name = typeof member === "string" ? member : member.name || "";
    return `<option value="${escapeHtml(name)}"></option>`;
  });

  if (currentValue && !names.some((opt) => opt.includes(`value="${escapeHtml(currentValue)}"`))) {
    names.unshift(`<option value="${escapeHtml(currentValue)}"></option>`);
  }

  return [...new Set(names)].join("");
}

function inferMtdMemberRecord(memberName) {
  const name = (memberName || "").trim().toUpperCase();
  if (!name) return null;

  const sources = [
    { brigade: "CONN", list: state.memberLists.CONN || [] },
    { brigade: "GROV", list: state.memberLists.GROV || [] },
    { brigade: "FRES", list: state.memberLists.FRES || [] }
  ];

  for (const source of sources) {
    const found = source.list.find((member) => {
      const memberNameValue = typeof member === "string" ? member : member.name || "";
      return memberNameValue.trim().toUpperCase() === name;
    });

    if (found) {
      return {
        brigade: source.brigade,
        phone: typeof found === "string" ? "" : (found.phone || "")
      };
    }
  }

  return null;
}

/* =========================================================
   9. AGENCY MODULE
   ========================================================= */
function addAgencyBlock(type) {
  state.agencies.push({
    id: cryptoRandomId(),
    type,
    officerName: "",
    contactNumber: "",
    station: "",
    badgeNumber: "",
    comments: "",
    otherName: ""
  });
  renderAgencyBlocks();
  persistDraftState();
}

function renderAgencyBlocks() {
  const container = document.getElementById("agencyBlocks");
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

  container.querySelectorAll("[data-agency-id]").forEach((el) => {
    el.addEventListener("input", updateAgencyField);
  });

  container.querySelectorAll("[data-remove-agency]").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.agencies = state.agencies.filter((agency) => agency.id !== btn.dataset.removeAgency);
      renderAgencyBlocks();
      persistDraftState();
    });
  });
}

function updateAgencyField(e) {
  const agency = state.agencies.find((a) => a.id === e.target.dataset.agencyId);
  if (!agency) return;
  agency[e.target.dataset.field] = e.target.value;
  persistDraftState();
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
   10. PROFILE MODULE
   ========================================================= */
function loadProfile() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.PROFILE);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    state.profile = { ...state.profile, ...parsed };
  } catch (error) {
    console.warn("Profile load failed", error);
  }
}

function saveProfileFromUi() {
  state.profile.name = document.getElementById("profileName").value.trim();
  state.profile.memberNumber = document.getElementById("profileMemberNumber").value.trim();
  state.profile.contactNumber = document.getElementById("profileContactNumber").value.trim();
  state.profile.email = document.getElementById("profileEmail").value.trim();
  state.profile.brigade = document.getElementById("profileBrigade").value.trim() || "Connewarre";

  localStorage.setItem(STORAGE_KEYS.PROFILE, JSON.stringify(state.profile));
  closeSettingsModal();
}

function openSettingsModal() {
  document.getElementById("profileName").value = state.profile.name;
  document.getElementById("profileMemberNumber").value = state.profile.memberNumber;
  document.getElementById("profileContactNumber").value = state.profile.contactNumber;
  document.getElementById("profileEmail").value = state.profile.email;
  document.getElementById("profileBrigade").value = state.profile.brigade;
  document.getElementById("settingsModal").classList.remove("hidden");
}

function closeSettingsModal() {
  document.getElementById("settingsModal").classList.add("hidden");
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

function saveReportSummary() {
  const subject = buildReportTitle();
  const body = state.ui.reportPreview || generateReport();
  const summary = {
    id: cryptoRandomId(),
    title: subject,
    body,
    createdAt: new Date().toISOString()
  };

  state.savedReports.unshift(summary);
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
  const container = document.getElementById("savedReportsList");
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
      document.getElementById("reportPreview").value = report.body;
    });
  });

  container.querySelectorAll("[data-delete-report]").forEach((btn) => {
    btn.addEventListener("click", () => deleteSavedReport(btn.dataset.deleteReport));
  });
}

/* =========================================================
   11. VALIDATION MODULE
   ========================================================= */
function validateResponders() {
  return getAllResponders().some((r) => r.name.trim());
}

function isPrimaryConnJob() {
  return (state.incident.brigadeCode || "").toUpperCase().startsWith("CONN");
}

function validateReportReady() {
  const hasResponder = validateResponders();
  const hasOic = !!getAllResponders().find((r) => r.oic && r.name.trim());

  const messages = [];
  if (!hasResponder) messages.push("At least one responder must be entered.");
  if (isPrimaryConnJob() && !hasOic) messages.push("OIC must be selected for primary Connewarre jobs.");

  return {
    valid: messages.length === 0,
    messages
  };
}

/* =========================================================
   12. REPORT GENERATOR MODULE
   ========================================================= */
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
  document.getElementById("reportPreview").value = report;
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
  const oic = getAllResponders().find((r) => r.oic);
  if (!oic || !oic.name) return "";
  return [
    "OIC",
    lineIfValue("Name", appendBrigadeIfNeeded(oic)),
    lineIfValue("Phone", oic.phone)
  ].filter(Boolean).join("\n");
}

function buildApplianceSections() {
  const grouped = {};

  getAllResponders()
    .filter((r) => r.name.trim() && (r.destination === "T1" || r.destination === "T2" || r.destination === "MTD P/T"))
    .forEach((r) => {
      const key = r.destination;
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(r);
    });

  const order = ["T1", "T2", "MTD P/T"];
  const sections = [];

  order.forEach((key) => {
    if (!grouped[key]?.length) return;
    const title = key === "T1" ? "Conn T1" : key === "T2" ? "Conn T2" : "MTD P/T";
    const lines = [title];

    grouped[key].forEach((r) => {
      const role = r.truckRole || "Crew";
      lines.push(`- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)} | ${role}`);
    });

    sections.push(lines.join("\n"));
  });

  if (!sections.length) return "";
  return ["Appliances", ...sections].join("\n");
}

function buildResponderSections() {
  const directLines = getAllResponders()
    .filter((r) => r.name.trim() && r.destination === "Direct")
    .map((r) => `- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)}`);

  const stationLines = getAllResponders()
    .filter((r) => r.name.trim() && r.destination === "Station")
    .map((r) => `- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)}`);

  const sections = [];
  if (directLines.length) sections.push(["Direct Responders", ...directLines].join("\n"));
  if (stationLines.length) sections.push(["Station Responders", ...stationLines].join("\n"));

  return sections.join("\n\n");
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
  const name = state.incident.firstAgency === "Other" ? state.incident.firstAgencyOther : state.incident.firstAgency;
  if (!name) return "";
  return ["First Agency On Scene", name].join("\n");
}

function buildHoseSection() {
  const used = Object.entries(state.hoseUse).filter(([, qty]) => String(qty).trim() !== "");
  if (!used.length) return "";
  return ["Hose Use", ...used.map(([type, qty]) => `${type}: ${qty}`)].join("\n");
}

function buildNotesSection() {
  if (!state.incident.notes.trim()) return "";
  return ["Notes", state.incident.notes.trim()].join("\n");
}

function buildIncidentFlagsSection() {
  const flags = [];
  if (state.incident.flags.membersBefore) flags.push("Members before 1st appliance");
  if (state.incident.flags.aarRequired) flags.push("AAR required");
  if (state.incident.flags.hotDebrief) flags.push("Hot debrief conducted");
  if (!flags.length) return "";
  return ["Incident Flags", ...flags.map((f) => `- ${f}`)].join("\n");
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

function buildResponderSuffix(responder) {
  const suffixes = [];
  if (responder.oic) suffixes.push("OIC");
  if (responder.ba) suffixes.push("BA");
  if (responder.injury) suffixes.push("Injured");
  return suffixes.length ? ` – ${suffixes.join(", ")}` : "";
}

/* =========================================================
   13. SEND MODULE
   ========================================================= */
function handleFinish() {
  const validation = validateReportReady();
  const validationStatus = document.getElementById("validationStatus");

  if (!validation.valid) {
    validationStatus.textContent = validation.messages.join(" ");
    document.getElementById("finishActions").classList.add("hidden");
    return;
  }

  validationStatus.textContent = "Report ready.";
  generateReport();
  updateQuietHoursWarning();
  document.getElementById("finishActions").classList.remove("hidden");
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
  const mailto = `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(report)}`;
  window.location.href = mailto;
}

function saveReportLocally() {
  generateReport();
  saveReportSummary();
}

/* =========================================================
   14. UI HELPERS
   ========================================================= */
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
    document.getElementById("validationStatus").textContent = validation.valid
      ? "Ready to finish."
      : validation.messages.join(" ");
    generateReport();
    updateQuietHoursWarning();
  }

  persistDraftState();
}

function toggleSection(targetId) {
  document.getElementById(targetId).classList.toggle("open");
}

function toggleAccordion(targetId) {
  document.getElementById(targetId).classList.toggle("open");
}

function updateQuietHoursWarning() {
  const warning = document.getElementById("quietHoursWarning");
  const now = new Date();
  const hour = now.getHours();
  const quiet = hour >= CONFIG.QUIET_HOURS.start || hour < CONFIG.QUIET_HOURS.end;
  warning.classList.toggle("hidden", !quiet);
}

function syncIncidentFieldsToUi() {
  document.getElementById("eventNumber").value = state.incident.eventNumber;
  document.getElementById("pagerDate").value = state.incident.pagerDate;
  document.getElementById("pagerTime").value = state.incident.pagerTime;
  document.getElementById("brigadeCode").value = state.incident.brigadeCode;
  document.getElementById("brigadeRole").value = state.incident.brigadeRole;
  document.getElementById("incidentType").value = state.incident.incidentType;
  document.getElementById("codeLevel").value = state.incident.codeLevel;
  document.getElementById("address").value = state.incident.address;
  document.getElementById("firsCode").value = state.incident.firsCode;
  document.getElementById("notes").value = state.incident.notes;
  document.getElementById("firstAgencySelect").value = state.incident.firstAgency;
  document.getElementById("firstAgencyOther").value = state.incident.firstAgencyOther;
  document.getElementById("firstAgencyOther").classList.toggle("hidden", state.incident.firstAgency !== "Other");
  updateFirsLabel();
  renderIncidentFlagButtons();
}

function syncAllFieldsToUi() {
  syncIncidentFieldsToUi();

  document.querySelector(`input[name="structureRequired"][value="${state.structure.required ? "yes" : "no"}"]`).checked = true;
  document.getElementById("structureFormWrap").classList.toggle("hidden", !state.structure.required);

  document.getElementById("structurePropertyType").value = state.structure.propertyType;
  document.getElementById("structureLevels").value = state.structure.levels;
  document.getElementById("structureOccupancy").value = state.structure.occupancy;
  document.getElementById("structureConstruction").value = state.structure.construction;
  document.getElementById("fireAreaOrigin").value = state.structure.fireAreaOrigin;
  document.getElementById("fireAreaExtent").value = state.structure.fireAreaExtent;
  document.getElementById("fireAreaComments").value = state.structure.fireAreaComments;
  document.getElementById("fireFuelLoad").value = state.structure.fireFuelLoad;
  document.getElementById("fireBehaviour").value = state.structure.fireBehaviour;
  document.getElementById("fireMaterialComments").value = state.structure.fireMaterialComments;
  document.getElementById("detectionType").value = state.structure.detectionType;
  document.getElementById("alarmActivated").value = state.structure.alarmActivated;
  document.getElementById("detectionComments").value = state.structure.detectionComments;
  document.getElementById("suppressionType").value = state.structure.suppressionType;
  document.getElementById("suppressionWorked").value = state.structure.suppressionWorked;
  document.getElementById("suppressionComments").value = state.structure.suppressionComments;
  document.getElementById("portableExtinguisher").value = state.structure.portableExtinguisher;
  document.getElementById("portableOther").value = state.structure.portableOther;
  document.getElementById("portableComments").value = state.structure.portableComments;

  document.getElementById("hose64").value = state.hoseUse["64"];
  document.getElementById("hose38").value = state.hoseUse["38"];
  document.getElementById("hose25").value = state.hoseUse["25"];
  document.getElementById("hoseLiveReel").value = state.hoseUse["Live Reel"];

  renderBrigadesOnScene();
  renderResponders();
}

function persistDraftState() {
  localStorage.setItem(STORAGE_KEYS.DRAFT_STATE, JSON.stringify(state));
}

function restoreDraftState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DRAFT_STATE);
    if (!raw) return;
    const parsed = JSON.parse(raw);

    if (parsed.incident) {
      state.incident = { ...state.incident, ...parsed.incident };
      if (!Array.isArray(state.incident.brigadesOnScene)) {
        state.incident.brigadesOnScene = [];
      }
    }
    if (parsed.responders) {
      state.responders = { ...state.responders, ...parsed.responders };
    }
    if (parsed.vehicles) state.vehicles = parsed.vehicles;
    if (parsed.agencies) state.agencies = parsed.agencies;
    if (parsed.structure) state.structure = { ...state.structure, ...parsed.structure };
    if (parsed.hoseUse) state.hoseUse = { ...state.hoseUse, ...parsed.hoseUse };
    if (parsed.ui) state.ui = { ...state.ui, currentPage: parsed.ui.currentPage || "incidentPage" };
  } catch (error) {
    console.warn("Draft restore failed", error);
  }
}

/* =========================================================
   UTILITIES
   ========================================================= */
function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch((error) => {
      console.warn("SW registration failed", error);
    });
  }
}

function setOcrStatus(message) {
  document.getElementById("ocrStatus").textContent = message;
}

function buildReportTitle() {
  const event = state.incident.eventNumber || "NO_EVENT";
  const type = state.incident.incidentType || "UNKNOWN";
  const address = (state.incident.address || "").trim() || "NO ADDRESS";
  return `${event} – ${type} – ${address}`;
}

function formatDateForReport(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  return `${day}/${month}/${year}`;
}

function appendBrigadeIfNeeded(responder) {
  const brigadeName = formatBrigadeName(responder.brigade || "CONN");
  return responder.brigade && responder.brigade !== "CONN"
    ? `${responder.name} (${brigadeName})`
    : responder.name;
}

function formatBrigadeName(brigadeKey) {
  if (brigadeKey === "CONN") return "Connewarre";
  if (brigadeKey === "GROV") return "Grovedale";
  if (brigadeKey === "FRES") return "Freshwater Creek";
  return brigadeKey || "";
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
  if (!record) return "";
  if (typeof record === "string") return "";
  return record.phone || "";
}

function lineIfValue(label, value) {
  if (value === null || value === undefined || value === "") return "";
  return `${label}: ${value}`;
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
