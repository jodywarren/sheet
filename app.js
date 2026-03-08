/* =========================================================
   TURNOUT SHEET - CFA DIGITAL TURNOUT SHEET
   Vanilla JS Progressive Web App
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
  RESPONSE_TYPES: ["Appliance", "Direct", "Station"],
  TASKS: ["Crew Leader", "Driver", "Crew", "Other"],
  SEAT_LAYOUTS: {
    "Conn T1": ["Crew Leader", "Driver", "Crew 1", "Crew 2", "Crew 3"],
    "Conn T2": ["Crew Leader", "Driver", "Crew 1", "Crew 2", "Crew 3"],
    "Other": ["Crew Leader", "Driver", "Crew 1", "Crew 2", "Crew 3"],
    "MTD P/T": ["Crew Leader", "Driver", "Crew 1", "Crew 2"]
  }
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
    brigadesOnScene: "",
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
    appliances: {
      "Conn T1": [],
      "Conn T2": [],
      "Other": [],
      "MTD P/T": []
    },
    direct: [],
    station: []
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
    "64": 0,
    "38": 0,
    "25": 0,
    "Live Reel": 0
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
  bindEventListeners();
  initialiseResponderLayouts();
  renderVehicleBlocks();
  renderAgencyBlocks();
  renderSavedReports();
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
    btn.addEventListener("click", () => toggleSection(btn.dataset.target, btn));
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

  document.querySelectorAll(".add-responder-btn").forEach((btn) => {
    btn.addEventListener("click", () => addSimpleResponder(btn.dataset.list));
  });
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
    ["brigadesOnScene", "brigadesOnScene"],
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
    document.getElementById("firstAgencyOtherWrap").classList.toggle("hidden", e.target.value !== "Other");
    persistDraftState();
  });

  document.getElementById("firstAgencyOther").addEventListener("input", (e) => {
    state.incident.firstAgencyOther = e.target.value;
    persistDraftState();
  });

  document.getElementById("flagMembersBefore").addEventListener("change", (e) => {
    state.incident.flags.membersBefore = e.target.checked;
    persistDraftState();
  });

  document.getElementById("flagAarRequired").addEventListener("change", (e) => {
    state.incident.flags.aarRequired = e.target.checked;
    persistDraftState();
  });

  document.getElementById("flagHotDebrief").addEventListener("change", (e) => {
    state.incident.flags.hotDebrief = e.target.checked;
    persistDraftState();
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
}

async function runOcrPipeline() {
  if (!state.ui.pagerImageFile) {
    setOcrStatus("Please upload a pager screenshot first.");
    return;
  }

  setOcrStatus("Analysing screenshot...");
  try {
    const imageBitmap = await createImageBitmap(state.ui.pagerImageFile);
    const rectangles = findPagerRectangles(imageBitmap);
    const selected = rectangles.length ? selectBestPagerRectangle(rectangles) : null;
    const cropCanvas = cropPagerRectangle(imageBitmap, selected);
    const rawText = await runPagerOCR(cropCanvas);
    const parsed = parsePagerMessage(rawText);

    if (!validatePagerText(parsed)) {
      setOcrStatus("OCR ran, but the rectangle did not contain enough pager content to safely populate fields.");
      return;
    }

    populateIncidentFields(parsed);
    state.incident.pagerRawText = rawText;
    setOcrStatus("Pager details extracted.");
    persistDraftState();
  } catch (error) {
    console.error(error);
    setOcrStatus("OCR failed. Try a cleaner screenshot or enter details manually.");
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
      const r = data[index];
      const g = data[index + 1];
      const b = data[index + 2];
      const avg = (r + g + b) / 3;
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

  if (!candidateBands.length) {
    return [{ x: 0, y: 0, width, height }];
  }

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
  const result = await Tesseract.recognize(canvas, "eng", {
    logger: () => {}
  });
  return result?.data?.text || "";
}

function parsePagerMessage(rawText) {
  const text = (rawText || "").replace(/\r/g, "").trim();
  const upper = text.toUpperCase();
  const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);

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

  const brigadeLines = lines.filter((line) => /(CONN|GROV|FRES)\d/.test(line.toUpperCase()));
  const brigadesOnScene = brigadeLines.length ? brigadeLines.join(", ") : "";

  const address = cleanAddress(extractLikelyAddress(lines));

  return {
    rawText: text,
    upper,
    lines,
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

function extractLikelyAddress(lines) {
  const joined = lines.join("\n").toUpperCase();
  const addressPattern = /(\d{1,5}\s+[A-Z0-9\s.'-]+(?:ROAD|RD|STREET|ST|AVENUE|AVE|DRIVE|DR|COURT|CT|CRESCENT|CRES|HIGHWAY|HWY|LANE|LN|PLACE|PL|WAY|BOULEVARD|BLVD)[A-Z0-9\s\/-]*)(?:\n([A-Z\s]+))?/;
  const match = joined.match(addressPattern);
  if (!match) return "";
  return [match[1], match[2] || ""].filter(Boolean).join("\n");
}

function validatePagerText(parsed) {
  const hasEvent = !!parsed.eventNumber;
  const hasBrigade = !!parsed.brigadeCode;
  const hasIncident = !!parsed.incidentType;
  const hasAddress = !!parsed.address;
  return hasEvent && hasBrigade && hasIncident && hasAddress;
}

function populateIncidentFields(parsed) {
  state.incident.eventNumber = parsed.eventNumber || state.incident.eventNumber;
  state.incident.pagerDate = parsed.pagerDate || state.incident.pagerDate;
  state.incident.pagerTime = parsed.pagerTime || state.incident.pagerTime;
  state.incident.brigadeCode = parsed.brigadeCode || state.incident.brigadeCode;
  state.incident.incidentType = parsed.incidentType || state.incident.incidentType;
  state.incident.codeLevel = parsed.codeLevel || state.incident.codeLevel;
  state.incident.address = parsed.address || state.incident.address;
  state.incident.brigadesOnScene = parsed.brigadesOnScene || state.incident.brigadesOnScene;

  syncIncidentFieldsToUi();
  updateBrigadeRole();
  updateStructureVisibilityByIncidentType();
  evaluateMvaAutoTrigger(parsed.upper || "");
}

function setOcrStatus(message) {
  document.getElementById("ocrStatus").textContent = message;
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

function cleanAddress(address) {
  if (!address) return "";
  const lines = address.split("\n").map((line) => line.trim()).filter(Boolean);
  const cleaned = lines.map((line, index) => {
    if (index === 0) {
      let value = line;
      value = value.split("//")[0];
      value = value.split("/")[0];
      return value.trim();
    }
    return line;
  });
  return cleaned.join("\n");
}

function updateFirsLabel() {
  const label = document.getElementById("firsLabel");
  label.textContent = state.incident.firsCode ? "FIRS Code" : "Paste FIRS";
}

function evaluateMvaAutoTrigger(text) {
  const hasMatch = CONFIG.MVA_KEYWORDS.some((keyword) => text.includes(keyword));
  if (hasMatch) {
    const vehicleSection = document.getElementById("vehicleSection");
    const toggle = document.querySelector('[data-target="vehicleSection"]');
    vehicleSection.classList.add("open");
    if (toggle) toggle.classList.add("open");
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

function collectHoseUse() {
  state.hoseUse["64"] = Number(document.getElementById("hose64").value || 0);
  state.hoseUse["38"] = Number(document.getElementById("hose38").value || 0);
  state.hoseUse["25"] = Number(document.getElementById("hose25").value || 0);
  state.hoseUse["Live Reel"] = Number(document.getElementById("hoseLiveReel").value || 0);
  persistDraftState();
}

/* =========================================================
   6. VEHICLE MODULE
   ========================================================= */
function setVehicleCount(count) {
  const safeCount = Math.max(0, Math.min(CONFIG.MAX_VEHICLES, count));
  while (state.vehicles.length < safeCount) {
    state.vehicles.push(createBlankVehicle(state.vehicles.length + 1));
  }
  while (state.vehicles.length > safeCount) {
    state.vehicles.pop();
  }
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
  const vehicleId = e.target.dataset.vehicleId;
  const field = e.target.dataset.field;
  const vehicle = state.vehicles.find((v) => v.id === vehicleId);
  if (!vehicle) return;

  vehicle[field] = e.target.value;

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
  const lines = [
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
  ].filter(Boolean);

  return lines.join("\n");
}

/* =========================================================
   8. RESPONDER MODULE
   ========================================================= */
function initialiseResponderLayouts() {
  renderSeatLayout("Conn T1", "connT1SeatContainer", "CONN");
  renderSeatLayout("Conn T2", "connT2SeatContainer", "CONN");
  renderSeatLayout("Other", "connOtherSeatContainer", "CONN");
  renderSeatLayout("MTD P/T", "mtdPtSeatContainer", "MTD");
  renderSimpleResponders();
}

function renderSeatLayout(applianceName, containerId, mode) {
  const container = document.getElementById(containerId);
  const seats = CONFIG.SEAT_LAYOUTS[applianceName];
  const existing = state.responders.appliances[applianceName];

  while (existing.length < seats.length) {
    existing.push(createBlankResponder(applianceName, seats[existing.length], mode === "MTD" ? "Appliance" : "Appliance"));
  }

  container.innerHTML = "";

  existing.forEach((responder, index) => {
    responder.task = seats[index];
    const card = document.createElement("div");
    card.className = "responder-card";
    card.innerHTML = buildResponderCardHtml(responder, mode, true);
    container.appendChild(card);
  });

  bindResponderCardInputs(container);
}

function buildResponderCardHtml(responder, mode = "CONN", seatLocked = false) {
  const brigadeOptions = mode === "MTD"
    ? `<option value="">Select Brigade</option>
       <option value="CONN" ${responder.brigade === "CONN" ? "selected" : ""}>Connewarre</option>
       <option value="GROV" ${responder.brigade === "GROV" ? "selected" : ""}>Grovedale</option>
       <option value="FRES" ${responder.brigade === "FRES" ? "selected" : ""}>Freshwater Creek</option>`
    : `<option value="CONN" selected>Connewarre</option>`;

  const nameOptions = getMemberOptionsHtml(responder.brigade || (mode === "MTD" ? "" : "CONN"), responder.name);

  return `
    <div class="responder-card-top">
      <div class="responder-title">${escapeHtml(responder.appliance)} - ${escapeHtml(responder.task)}</div>
      <button class="tiny-btn remove-responder-btn ${seatLocked ? "hidden" : ""}" type="button" data-responder-id="${responder.id}">Remove</button>
    </div>

    <div class="responder-grid">
      <label>
        Brigade
        <select data-responder-id="${responder.id}" data-field="brigade" class="responder-brigade">
          ${brigadeOptions}
        </select>
      </label>

      <label>
        Name
        <input type="text" list="memberList-${responder.id}" data-responder-id="${responder.id}" data-field="name" value="${escapeHtml(responder.name)}" />
        <datalist id="memberList-${responder.id}">
          ${nameOptions}
        </datalist>
      </label>

      <label>
        Contact Number
        <input type="text" data-responder-id="${responder.id}" data-field="phone" value="${escapeHtml(responder.phone)}" />
      </label>

      <label>
        Task
        <select data-responder-id="${responder.id}" data-field="task" ${seatLocked ? "disabled" : ""}>
          ${CONFIG.TASKS.map((task) => `<option value="${task}" ${responder.task === task ? "selected" : ""}>${task}</option>`).join("")}
        </select>
      </label>

      <label>
        Response Type
        <select data-responder-id="${responder.id}" data-field="responseType">
          ${CONFIG.RESPONSE_TYPES.map((rt) => `<option value="${rt}" ${responder.responseType === rt ? "selected" : ""}>${rt}</option>`).join("")}
        </select>
      </label>

      <label>
        Notes
        <input type="text" data-responder-id="${responder.id}" data-field="notes" value="${escapeHtml(responder.notes)}" />
      </label>
    </div>

    <div class="responder-flags">
      <label><input type="checkbox" data-responder-id="${responder.id}" data-flag="ba" ${responder.ba ? "checked" : ""} /> BA</label>
      <label><input type="checkbox" data-responder-id="${responder.id}" data-flag="injury" ${responder.injury ? "checked" : ""} /> Injury</label>
      <label><input type="checkbox" data-responder-id="${responder.id}" data-flag="oic" ${responder.oic ? "checked" : ""} /> OIC</label>
    </div>
  `;
}

function createBlankResponder(appliance = "", task = "", responseType = "Appliance") {
  return {
    id: cryptoRandomId(),
    brigade: appliance === "MTD P/T" ? "" : "CONN",
    name: "",
    phone: "",
    appliance,
    task,
    ba: false,
    injury: false,
    oic: false,
    responseType,
    notes: ""
  };
}

function bindResponderCardInputs(container) {
  container.querySelectorAll("[data-responder-id]").forEach((el) => {
    if (el.matches("[data-field]")) {
      el.addEventListener("input", handleResponderFieldChange);
      el.addEventListener("change", handleResponderFieldChange);
    }
    if (el.matches("[data-flag]")) {
      el.addEventListener("change", handleResponderFlagChange);
    }
    if (el.classList.contains("remove-responder-btn")) {
      el.addEventListener("click", removeResponderById);
    }
  });
}

function handleResponderFieldChange(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  const field = e.target.dataset.field;
  responder[field] = e.target.value;

  if (field === "brigade") {
    const list = e.target.closest(".responder-card").querySelector("datalist");
    list.innerHTML = getMemberOptionsHtml(responder.brigade, responder.name);
  }

  persistDraftState();
  updateHeaderOicStatus();
}

function handleResponderFlagChange(e) {
  const responder = findResponderById(e.target.dataset.responderId);
  if (!responder) return;

  const flag = e.target.dataset.flag;

  if (flag === "oic" && e.target.checked) {
    clearExistingOic();
  }

  responder[flag] = e.target.checked;

  if (flag === "oic") {
    state.ui.oicResponderId = responder.oic ? responder.id : "";
    state.ui.oicName = responder.oic ? responder.name : "";
    state.ui.oicPhone = responder.oic ? responder.phone : "";
  }

  if (flag === "oic" && !e.target.checked) {
    state.ui.oicResponderId = "";
    state.ui.oicName = "";
    state.ui.oicPhone = "";
  }

  updateHeaderOicStatus();
  persistDraftState();
}

function clearExistingOic() {
  getAllResponders().forEach((responder) => {
    responder.oic = false;
  });
  document.querySelectorAll('[data-flag="oic"]').forEach((box) => {
    box.checked = false;
  });
}

function updateHeaderOicStatus() {
  const oic = getAllResponders().find((r) => r.oic);
  const el = document.getElementById("oicStatus");

  if (!oic || !oic.name) {
    el.textContent = "APPOINT OIC";
    el.classList.add("oic-missing");
    return;
  }

  state.ui.oicResponderId = oic.id;
  state.ui.oicName = oic.name;
  state.ui.oicPhone = oic.phone || "";
  el.textContent = `OIC: ${oic.name}${oic.phone ? " – " + oic.phone : ""}`;
  el.classList.remove("oic-missing");
}

function getAllResponders() {
  return [
    ...state.responders.appliances["Conn T1"],
    ...state.responders.appliances["Conn T2"],
    ...state.responders.appliances["Other"],
    ...state.responders.appliances["MTD P/T"],
    ...state.responders.direct,
    ...state.responders.station
  ];
}

function findResponderById(id) {
  return getAllResponders().find((r) => r.id === id);
}

function addSimpleResponder(listName) {
  const responseType = listName === "direct" ? "Direct" : "Station";
  state.responders[listName].push(createBlankResponder("", "Other", responseType));
  renderSimpleResponders();
  persistDraftState();
}

function renderSimpleResponders() {
  renderSimpleResponderList("direct", "directRespondersContainer");
  renderSimpleResponderList("station", "stationRespondersContainer");
}

function renderSimpleResponderList(listName, containerId) {
  const container = document.getElementById(containerId);
  container.innerHTML = "";

  state.responders[listName].forEach((responder) => {
    const card = document.createElement("div");
    card.className = "responder-card";
    card.innerHTML = buildSimpleResponderHtml(responder);
    container.appendChild(card);
  });

  bindResponderCardInputs(container);
}

function buildSimpleResponderHtml(responder) {
  return `
    <div class="responder-card-top">
      <div class="responder-title">${escapeHtml(responder.responseType)} Responder</div>
      <button class="tiny-btn remove-responder-btn" type="button" data-responder-id="${responder.id}">Remove</button>
    </div>

    <div class="responder-grid">
      <label>
        Brigade
        <select data-responder-id="${responder.id}" data-field="brigade" class="responder-brigade">
          <option value="">Select Brigade</option>
          <option value="CONN" ${responder.brigade === "CONN" ? "selected" : ""}>Connewarre</option>
          <option value="GROV" ${responder.brigade === "GROV" ? "selected" : ""}>Grovedale</option>
          <option value="FRES" ${responder.brigade === "FRES" ? "selected" : ""}>Freshwater Creek</option>
        </select>
      </label>

      <label>
        Name
        <input type="text" list="memberList-${responder.id}" data-responder-id="${responder.id}" data-field="name" value="${escapeHtml(responder.name)}" />
        <datalist id="memberList-${responder.id}">
          ${getMemberOptionsHtml(responder.brigade, responder.name)}
        </datalist>
      </label>

      <label>
        Contact Number
        <input type="text" data-responder-id="${responder.id}" data-field="phone" value="${escapeHtml(responder.phone)}" />
      </label>

      <label>
        Task
        <select data-responder-id="${responder.id}" data-field="task">
          ${CONFIG.TASKS.map((task) => `<option value="${task}" ${responder.task === task ? "selected" : ""}>${task}</option>`).join("")}
        </select>
      </label>

      <label>
        Response Type
        <select data-responder-id="${responder.id}" data-field="responseType">
          ${CONFIG.RESPONSE_TYPES.map((rt) => `<option value="${rt}" ${responder.responseType === rt ? "selected" : ""}>${rt}</option>`).join("")}
        </select>
      </label>

      <label>
        Notes
        <input type="text" data-responder-id="${responder.id}" data-field="notes" value="${escapeHtml(responder.notes)}" />
      </label>
    </div>

    <div class="responder-flags">
      <label><input type="checkbox" data-responder-id="${responder.id}" data-flag="ba" ${responder.ba ? "checked" : ""} /> BA</label>
      <label><input type="checkbox" data-responder-id="${responder.id}" data-flag="injury" ${responder.injury ? "checked" : ""} /> Injury</label>
      <label><input type="checkbox" data-responder-id="${responder.id}" data-flag="oic" ${responder.oic ? "checked" : ""} /> OIC</label>
    </div>
  `;
}

function removeResponderById(e) {
  const id = e.target.dataset.responderId;
  state.responders.direct = state.responders.direct.filter((r) => r.id !== id);
  state.responders.station = state.responders.station.filter((r) => r.id !== id);
  renderSimpleResponders();
  updateHeaderOicStatus();
  persistDraftState();
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
      <button class="tiny-btn" type="button" data-load-report="${report.id}">Load Preview</button>
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
}

/* =========================================================
   11. VALIDATION MODULE
   ========================================================= */
function validateIncident() {
  return true;
}

function validateResponders() {
  const responders = getAllResponders().filter((r) => r.name.trim());
  return responders.length > 0;
}

function validateReportReady() {
  const hasOic = !!getAllResponders().find((r) => r.oic && r.name.trim());
  const hasResponder = validateResponders();

  const messages = [];
  if (!hasOic) messages.push("OIC must be selected.");
  if (!hasResponder) messages.push("At least one responder must be entered.");

  return {
    valid: hasOic && hasResponder,
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
    formatAgencySection(),
    buildFirstAgencySection(),
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
    lineIfValue("Address", state.incident.address.replace(/\n/g, ", ")),
    lineIfValue("FIRS Code", state.incident.firsCode),
    lineIfValue("Brigades on Scene", state.incident.brigadesOnScene)
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
  const lines = ["Appliances"];
  let hasAny = false;

  ["Conn T1", "Conn T2", "Other", "MTD P/T"].forEach((appliance) => {
    const responders = state.responders.appliances[appliance].filter((r) => r.name.trim());
    if (!responders.length) return;
    hasAny = true;
    lines.push(`${appliance}`);
    responders.forEach((r) => {
      lines.push(`- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)} | ${r.task}`);
    });
  });

  return hasAny ? lines.join("\n") : "";
}

function buildResponderSections() {
  const directLines = state.responders.direct.filter((r) => r.name.trim()).map((r) => `- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)}`);
  const stationLines = state.responders.station.filter((r) => r.name.trim()).map((r) => `- ${appendBrigadeIfNeeded(r)}${buildResponderSuffix(r)}`);

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
  const name = state.incident.firstAgency === "Other"
    ? state.incident.firstAgencyOther
    : state.incident.firstAgency;
  if (!name) return "";
  return ["First Agency On Scene", name].join("\n");
}

function buildHoseSection() {
  const used = Object.entries(state.hoseUse).filter(([, qty]) => Number(qty) > 0);
  if (!used.length) return "";
  return [
    "Hose Use",
    ...used.map(([type, qty]) => `${type}: ${qty}`)
  ].join("\n");
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

function updateQuietHoursWarning() {
  const warning = document.getElementById("quietHoursWarning");
  const now = new Date();
  const hour = now.getHours();
  const quiet = hour >= CONFIG.QUIET_HOURS.start || hour < CONFIG.QUIET_HOURS.end;
  warning.classList.toggle("hidden", !quiet);
}

/* =========================================================
   14. UI RENDER HELPERS
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
  const target = document.getElementById(targetId);
  target.classList.toggle("open");
}

function toggleAccordion(targetId) {
  const panel = document.getElementById(targetId);
  panel.classList.toggle("open");
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
  document.getElementById("brigadesOnScene").value = state.incident.brigadesOnScene;
  document.getElementById("notes").value = state.incident.notes;
  document.getElementById("firstAgencySelect").value = state.incident.firstAgency;
  document.getElementById("firstAgencyOther").value = state.incident.firstAgencyOther;
  document.getElementById("flagMembersBefore").checked = state.incident.flags.membersBefore;
  document.getElementById("flagAarRequired").checked = state.incident.flags.aarRequired;
  document.getElementById("flagHotDebrief").checked = state.incident.flags.hotDebrief;
  document.getElementById("firstAgencyOtherWrap").classList.toggle("hidden", state.incident.firstAgency !== "Other");
  updateFirsLabel();
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

  initialiseResponderLayouts();
}

function persistDraftState() {
  localStorage.setItem(STORAGE_KEYS.DRAFT_STATE, JSON.stringify(state));
}

function restoreDraftState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DRAFT_STATE);
    if (!raw) return;
    const parsed = JSON.parse(raw);

    if (parsed.incident) state.incident = { ...state.incident, ...parsed.incident };
    if (parsed.responders) state.responders = parsed.responders;
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

function buildReportTitle() {
  const event = state.incident.eventNumber || "NO_EVENT";
  const type = state.incident.incidentType || "UNKNOWN";
  const address = (state.incident.address || "").replace(/\n/g, " ").trim() || "NO ADDRESS";
  return `${event} – ${type} – ${address}`;
}

function formatDateForReport(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  return `${day}/${month}/${year}`;
}

function appendBrigadeIfNeeded(responder) {
  const brigadeName = responder.brigade === "GROV"
    ? "Grovedale"
    : responder.brigade === "FRES"
    ? "Freshwater Creek"
    : "Connewarre";

  return responder.brigade && responder.brigade !== "CONN"
    ? `${responder.name} (${brigadeName})`
    : responder.name;
}

function lineIfValue(label, value) {
  if (value === null || value === undefined || value === "") return "";
  return `${label}: ${value}`;
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
