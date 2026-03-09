const CONFIG = {
  APP_VERSION: "2.0.3",
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
  PROFILE: "turnout_profile_v203",
  REPORTS: "turnout_reports_v203",
  DRAFT: "turnout_draft_v203"
};

const MEMBER_FILES = {
  CONN: "CONN.members.json",
  GROV: "GROV.members.json",
  FRES: "FRES.members.json"
};

const state = {
  incident: {
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
      injuriesManual: false
    },
    signals: []
  },
  responders: {
    connewarre: [],
    mtd: [],
    applianceCodes: {
      t1: "",
      t2: "",
      mtd: ""
    }
  },
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
  setAppVersion();
  populateDistanceDropdown();
  await loadMemberLists();
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
    destination: "",
    isDriver: false,
    isCrewLeader: false,
    ba: false,
    injured: false,
    oic: false
  };
}

function normaliseResponderState() {
  ["connewarre", "mtd"].forEach((groupKey) => {
    state.responders[groupKey] = (state.responders[groupKey] || []).map((person) => {
      const next = { ...createResponder(groupKey), ...person };
      delete next.truckRole;
      return next;
    });
  });

  if (!state.responders.applianceCodes) {
    state.responders.applianceCodes = { t1: "", t2: "", mtd: "" };
  }
}

function ensureRows() {
  if (!state.responders.connewarre.length) state.responders.connewarre.push(createResponder("connewarre"));
  if (!state.responders.mtd.length) state.responders.mtd.push(createResponder("mtd"));
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
  el("scanBtn")?.addEventListener("click", () => {
    el("scanStatus").textContent = "OCR is intentionally disabled in this version. The report structure is ready for it.";
  });

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
      markReportStale();
      scheduleDraftSave();
      updateReportSummary();
    });
  });

  ["t1Code", "t2Code", "mtdCode"].forEach((id) => {
    el(id)?.addEventListener("change", (e) => {
      if (id === "t1Code") state.responders.applianceCodes.t1 = e.target.value;
      if (id === "t2Code") state.responders.applianceCodes.t2 = e.target.value;
      if (id === "mtdCode") state.responders.applianceCodes.mtd = e.target.value;
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

function handleImageUpload(e) {
  const file = e.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    state.ui.previewUrl = reader.result;
    el("pagerPreview").src = reader.result;
    el("pagerPreview").classList.remove("hidden");
    el("pagerPreviewEmpty").classList.add("hidden");
    el("scanStatus").textContent = `Loaded ${file.name}`;
    markReportStale();
    scheduleDraftSave();
  };
  reader.readAsDataURL(file);
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
  el(buttonId).classList.toggle("active", state.incident.flags[key]);
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
        if (key === "injured") {
          syncInjuriesFlagButton();
        }
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

function buildPagerDateTime() {
  const date = state.incident.pagerDate || "";
  const time = state.incident.pagerTime || "";
  return [date, time].filter(Boolean).join(" ");
}

function buildApplianceCodeLine() {
  const parts = [];
  if (state.responders.applianceCodes.t1) parts.push(`T1 ${state.responders.applianceCodes.t1}`);
  if (state.responders.applianceCodes.t2) parts.push(`T2 ${state.responders.applianceCodes.t2}`);
  if (state.responders.applianceCodes.mtd) parts.push(`MTD P/T ${state.responders.applianceCodes.mtd}`);
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
    ? `${name}${name.toLowerCase().endsWith("control") ? "" : " Control"}`
    : "Control";

  return role ? `${controlText} | ${role}` : controlText;
}

function buildSignalsLine() {
  if (!state.incident.signals.length) return "";
  return state.incident.signals
    .map((code) => `${code} = ${CONFIG.SIGNALS[code]}`)
    .join("; ");
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
  pushLine(lines, "Pager Date / Time", buildPagerDateTime());
  pushLine(lines, "Appliance Codes", buildApplianceCodeLine());

  if (state.incident.pagerDetails.trim()) {
    lines.push("");
    lines.push("Pager Details");
    lines.push(state.incident.pagerDetails.trim());
  }

  pushLine(lines, "Actual Location", state.incident.actualLocation);
  pushLine(lines, "Control", buildControlLine());

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

  if (!silent) {
    updateDraftMeta("Report saved locally.");
  }
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
  el("pagerDate").value = state.incident.pagerDate;
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
  el("distanceToScene").value = state.incident.distanceToScene;
  el("hose64Qty").value = state.incident.hoses.hose64Qty;
  el("hose38Qty").value = state.incident.hoses.hose38Qty;
  el("hose25Qty").value = state.incident.hoses.hose25Qty;
  el("hoseOtherType").value = state.incident.hoses.hoseOtherType;

  el("flagMembersBeforeBtn").classList.toggle("active", state.incident.flags.membersBefore);
  el("flagAarBtn").classList.toggle("active", state.incident.flags.aar);
  syncInjuriesFlagButton();
  syncSignalButtons();

  el("t1Code").value = state.responders.applianceCodes.t1;
  el("t2Code").value = state.responders.applianceCodes.t2;
  el("mtdCode").value = state.responders.applianceCodes.mtd;

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
  updateDraftMeta(`Draft autosaved ${new Date(payload.savedAt).toLocaleTimeString("en-AU", { hour: "2-digit", minute: "2-digit" })}.`);
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

    state.incident = { ...state.incident, ...payload.incident };
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
    state.ui.currentPage = payload.ui?.currentPage || "incidentPage";
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
