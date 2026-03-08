const CONFIG = {
  APP_VERSION: "2.0.1",
  MAX_REPORTS: 10,
  DRAFT_SAVE_DELAY_MS: 300,
  QUIET_HOURS_START: 22,
  QUIET_HOURS_END: 7
};

const STORAGE_KEYS = {
  PROFILE: "turnout_profile_v201",
  REPORTS: "turnout_reports_v201",
  DRAFT: "turnout_draft_v201"
};

const MEMBER_FILES = {
  CONN: "CONN.members.json",
  GROV: "GROV.members.json",
  FRES: "FRES.members.json"
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
    notes: "",
    flags: {
      membersBefore: false,
      aar: false,
      hotDebrief: false
    }
  },
  responders: {
    connewarre: [],
    mtd: []
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
    lastDraftSavedAt: ""
  }
};

let draftSaveTimer = null;

document.addEventListener("DOMContentLoaded", init);

async function init() {
  setAppVersion();
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

      if (typeof person.truckRole === "string") {
        if (person.truckRole === "Driver") next.isDriver = true;
        if (person.truckRole === "CL") next.isCrewLeader = true;
      }

      if (person.truckRole === "Driver/CL") {
        next.isDriver = true;
        next.isCrewLeader = true;
      }

      delete next.truckRole;
      return next;
    });
  });
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
    el("scanStatus").textContent = "OCR is intentionally disabled in this hardened build. Reliability first.";
  });

  el("addSceneBrigadeBtn")?.addEventListener("click", addSceneBrigade);
  el("sceneBrigadeSelect")?.addEventListener("change", () => {
    el("sceneBrigadeOther").classList.toggle("hidden", el("sceneBrigadeSelect").value !== "Other");
  });

  el("firstAgency")?.addEventListener("change", () => {
    el("firstAgencyOther").classList.toggle("hidden", el("firstAgency").value !== "Other");
    state.incident.firstAgency = el("firstAgency").value;
    scheduleDraftSave();
    renderEverything();
  });

  el("firstAgencyOther")?.addEventListener("input", () => {
    state.incident.firstAgencyOther = el("firstAgencyOther").value;
    scheduleDraftSave();
  });

  el("agencyType")?.addEventListener("change", autoAddAgencyFromSelect);

  el("flagMembersBeforeBtn")?.addEventListener("click", () => toggleFlag("membersBefore", "flagMembersBeforeBtn"));
  el("flagAarBtn")?.addEventListener("click", () => toggleFlag("aar", "flagAarBtn"));
  el("flagHotDebriefBtn")?.addEventListener("click", () => toggleFlag("hotDebrief", "flagHotDebriefBtn"));

  el("resetFirsBtn")?.addEventListener("click", () => {
    state.incident.firsCode = "";
    el("firsCode").value = "";
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
    ["codeLevel", "codeLevel"],
    ["address", "address"],
    ["firsCode", "firsCode"],
    ["notes", "notes"]
  ];

  map.forEach(([id, key]) => {
    el(id)?.addEventListener("input", (e) => {
      state.incident[key] = e.target.value;

      if (key === "brigadeCode") {
        const code = e.target.value.trim().toUpperCase();
        state.incident.brigadeRole = code.startsWith("CONN") ? "Primary" : code ? "Support" : "";
        el("brigadeRole").value = state.incident.brigadeRole;
      }

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
    banner.className = "status-banner online";
    banner.textContent = "Online";
  } else {
    banner.className = "status-banner offline";
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
        <label>Contact Number<input data-field="contactNumber" type="text" value="${text(agency.contactNumber)}"></label>
        <label>Station<input data-field="station" type="text" value="${text(agency.station)}"></label>
        <label>Badge Number<input data-field="badgeNumber" type="text" value="${text(agency.badgeNumber)}"></label>
        <label class="full">Comments<textarea data-field="comments" rows="2">${text(agency.comments)}</textarea></label>
      </div>
    `;

    block.querySelector("button").addEventListener("click", () => {
      const confirmed = window.confirm(`Remove agency entry for ${agency.type}?`);
      if (!confirmed) return;
      state.agencies = state.agencies.filter((a) => a.id !== agency.id);
      renderAgencies();
      scheduleDraftSave();
      updateReportSummary();
    });

    block.querySelectorAll("[data-field]").forEach((field) => {
      field.addEventListener("input", (e) => {
        agency[e.target.dataset.field] = e.target.value;
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
  scheduleDraftSave();
  updateReportSummary();
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
      scheduleDraftSave();
    });

    card.querySelector(".tiny-btn").addEventListener("click", () => {
      const confirmed = window.confirm(`Remove responder "${person.name || "Unnamed member"}"?`);
      if (!confirmed) return;

      state.responders[groupKey] = state.responders[groupKey].filter((x) => x.id !== person.id);
      if (!state.responders[groupKey].length) state.responders[groupKey].push(createResponder(groupKey));
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
          renderResponders();
          scheduleDraftSave();
          return;
        }

        person[key] = !person[key];
        btn.classList.toggle("active", person[key]);
        updateOicBanner();
        updateReportSummary();
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

function buildResponderRoleLabel(r) {
  const roles = [];
  if (r.isDriver) roles.push("Driver");
  if (r.isCrewLeader) roles.push("CL");
  return roles.length ? roles.join("/") : "Crew";
}

function buildReport() {
  const all = getAllResponders();
  const oic = all.find((r) => r.oic && r.name.trim());

  const lines = [];
  lines.push("Job Details");
  pushLine(lines, "Event Number", state.incident.eventNumber);
  pushLine(lines, "Pager Date", state.incident.pagerDate);
  pushLine(lines, "Pager Time", state.incident.pagerTime);
  pushLine(lines, "Brigade Code", state.incident.brigadeCode);
  pushLine(lines, "Primary / Support", state.incident.brigadeRole);
  pushLine(lines, "Incident Type / Class", state.incident.incidentType);
  pushLine(lines, "Code Level", state.incident.codeLevel);
  pushLine(lines, "Address", state.incident.address);
  pushLine(lines, "FIRS Code", state.incident.firsCode);
  pushLine(lines, "Brigades on Scene", state.incident.brigadesOnScene.join(", "));

  if (oic) {
    lines.push("");
    lines.push("OIC");
    pushLine(lines, "Name", oic.name);
    pushLine(lines, "Phone", oic.phone);
  }

  const appliances = [
    ["Conn T1", "T1"],
    ["Conn T2", "T2"],
    ["MTD P/T", "MTD P/T"]
  ];

  const applianceSections = [];
  appliances.forEach(([title, code]) => {
    const crew = all.filter((r) => r.destination === code && r.name.trim());
    if (crew.length) {
      applianceSections.push(title);
      crew.forEach((r) => {
        applianceSections.push(`- ${r.name}${responderSuffix(r)} | ${buildResponderRoleLabel(r)}`);
      });
    }
  });

  if (applianceSections.length) {
    lines.push("");
    lines.push("Appliances");
    lines.push(...applianceSections);
  }

  const direct = all.filter((r) => r.destination === "Direct" && r.name.trim());
  if (direct.length) {
    lines.push("");
    lines.push("Direct Responders");
    direct.forEach((r) => lines.push(`- ${r.name}${responderSuffix(r)}`));
  }

  const station = all.filter((r) => r.destination === "Station" && r.name.trim());
  if (station.length) {
    lines.push("");
    lines.push("Station Responders");
    station.forEach((r) => lines.push(`- ${r.name}${responderSuffix(r)}`));
  }

  if (state.agencies.length) {
    lines.push("");
    lines.push("Agencies");
    state.agencies.forEach((a, i) => {
      lines.push(`${i + 1}. ${a.type === "Other" ? (a.otherName || "Other") : a.type}`);
      pushLine(lines, "Officer", a.officerName);
      pushLine(lines, "Contact", a.contactNumber);
      pushLine(lines, "Station", a.station);
      pushLine(lines, "Badge", a.badgeNumber);
      pushLine(lines, "Comments", a.comments);
    });
  }

  const firstAgency = state.incident.firstAgency === "Other" ? state.incident.firstAgencyOther : state.incident.firstAgency;
  if (firstAgency) {
    lines.push("");
    lines.push("First Agency On Scene");
    lines.push(firstAgency);
  }

  if (state.incident.notes.trim()) {
    lines.push("");
    lines.push("Notes");
    lines.push(state.incident.notes.trim());
  }

  const flags = [];
  if (state.incident.flags.membersBefore) flags.push("Members before 1st appliance");
  if (state.incident.flags.aar) flags.push("AAR required");
  if (state.incident.flags.hotDebrief) flags.push("Hot debrief conducted");

  if (flags.length) {
    lines.push("");
    lines.push("Incident Flags");
    flags.forEach((f) => lines.push(`- ${f}`));
  }

  lines.push("");
  lines.push("Sign-off");
  pushLine(lines, "Name", state.profile.name);
  pushLine(lines, "Brigade", state.profile.brigade);
  pushLine(lines, "CFA Member Number", state.profile.memberNumber);
  pushLine(lines, "Contact Number", state.profile.contactNumber);

  return lines.join("\n");
}

function responderSuffix(r) {
  const flags = [];
  if (r.oic) flags.push("OIC");
  if (r.ba) flags.push("BA");
  if (r.injured) flags.push("Injured");
  return flags.length ? " – " + flags.join(", ") : "";
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

  if (!state.incident.eventNumber.trim()) issues.push("Event Number");
  if (!state.incident.incidentType.trim()) issues.push("Incident Type");
  if (!state.incident.address.trim()) issues.push("Address");
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
    el("finishActions").classList.add("hidden");
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
  el("finishActions").classList.remove("hidden");
  updateReportSummary();
}

function buildReportTitle() {
  return `${state.incident.eventNumber || "NO_EVENT"} – ${state.incident.incidentType || "UNKNOWN"} – ${state.incident.address || "NO ADDRESS"}`;
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
  el("eventNumber").value = state.incident.eventNumber;
  el("pagerDate").value = state.incident.pagerDate;
  el("pagerTime").value = state.incident.pagerTime;
  el("brigadeCode").value = state.incident.brigadeCode;
  el("brigadeRole").value = state.incident.brigadeRole;
  el("incidentType").value = state.incident.incidentType;
  el("codeLevel").value = state.incident.codeLevel;
  el("address").value = state.incident.address;
  el("firsCode").value = state.incident.firsCode;
  el("notes").value = state.incident.notes;
  el("firstAgency").value = state.incident.firstAgency;
  el("firstAgencyOther").value = state.incident.firstAgencyOther;
  el("firstAgencyOther").classList.toggle("hidden", state.incident.firstAgency !== "Other");

  el("flagMembersBeforeBtn").classList.toggle("active", state.incident.flags.membersBefore);
  el("flagAarBtn").classList.toggle("active", state.incident.flags.aar);
  el("flagHotDebriefBtn").classList.toggle("active", state.incident.flags.hotDebrief);

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
  const flagsCount = Object.values(state.incident.flags).filter(Boolean).length;

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
      previewUrl: state.ui.previewUrl
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
    incident?.address ||
    incident?.notes ||
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
      mtd: Array.isArray(payload.responders?.mtd) ? payload.responders.mtd : []
    };
    state.agencies = Array.isArray(payload.agencies) ? payload.agencies : [];
    state.ui.currentPage = payload.ui?.currentPage || "incidentPage";
    state.ui.previewUrl = payload.ui?.previewUrl || "";
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
