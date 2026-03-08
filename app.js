const STORAGE_KEYS = {
  PROFILE: "turnout_profile_v1",
  REPORTS: "turnout_reports_v1"
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
    previewUrl: ""
  }
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  await loadMemberLists();
  loadProfile();
  loadSavedReports();

  ensureRows();
  bindStaticEvents();
  bindIncidentEvents();
  renderEverything();
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
    truckRole: "",
    ba: false,
    injured: false,
    oic: false
  };
}

function ensureRows() {
  if (!state.responders.connewarre.length) state.responders.connewarre.push(createResponder("connewarre"));
  if (!state.responders.mtd.length) state.responders.mtd.push(createResponder("mtd"));
}

function getAllResponders() {
  return [...state.responders.connewarre, ...state.responders.mtd];
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

  el("openSettingsBtn").addEventListener("click", openSettings);
  el("closeSettingsBtn").addEventListener("click", closeSettings);
  el("saveProfileBtn").addEventListener("click", saveProfileFromUi);

  el("pagerUpload").addEventListener("change", handleImageUpload);
  el("scanBtn").addEventListener("click", () => {
    el("scanStatus").textContent = "OCR is intentionally disabled in this clean rebuild. Base app first.";
  });

  el("addSceneBrigadeBtn").addEventListener("click", addSceneBrigade);
  el("sceneBrigadeSelect").addEventListener("change", () => {
    el("sceneBrigadeOther").classList.toggle("hidden", el("sceneBrigadeSelect").value !== "Other");
  });

  el("firstAgency").addEventListener("change", () => {
    el("firstAgencyOther").classList.toggle("hidden", el("firstAgency").value !== "Other");
    state.incident.firstAgency = el("firstAgency").value;
    renderEverything();
  });

  el("firstAgencyOther").addEventListener("input", () => {
    state.incident.firstAgencyOther = el("firstAgencyOther").value;
  });

  el("addAgencyBtn").addEventListener("click", addAgency);

  el("flagMembersBeforeBtn").addEventListener("click", () => toggleFlag("membersBefore", "flagMembersBeforeBtn"));
  el("flagAarBtn").addEventListener("click", () => toggleFlag("aar", "flagAarBtn"));
  el("flagHotDebriefBtn").addEventListener("click", () => toggleFlag("hotDebrief", "flagHotDebriefBtn"));

  el("resetFirsBtn").addEventListener("click", () => {
    state.incident.firsCode = "";
    el("firsCode").value = "";
  });

  el("finishBtn").addEventListener("click", finishReport);
  el("saveLocalBtn").addEventListener("click", saveCurrentReport);
  el("sendEmailBtn").addEventListener("click", sendEmail);
  el("sendSmsBtn").addEventListener("click", sendSms);
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
    el(id).addEventListener("input", (e) => {
      state.incident[key] = e.target.value;
      if (key === "brigadeCode") {
        const code = e.target.value.trim().toUpperCase();
        state.incident.brigadeRole = code.startsWith("CONN") ? "Primary" : code ? "Support" : "";
        el("brigadeRole").value = state.incident.brigadeRole;
      }
    });
  });
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
}

function renderSceneBrigades() {
  const wrap = el("sceneBrigadeChips");
  wrap.innerHTML = "";

  state.incident.brigadesOnScene.forEach((code) => {
    const chip = document.createElement("div");
    chip.className = "scene-chip";
    chip.innerHTML = `<span>${text(code)}</span><button type="button">×</button>`;
    chip.querySelector("button").addEventListener("click", () => {
      state.incident.brigadesOnScene = state.incident.brigadesOnScene.filter((x) => x !== code);
      renderSceneBrigades();
    });
    wrap.appendChild(chip);
  });
}

function addAgency() {
  const type = el("agencyType").value;
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

  el("agencyType").value = "";
  renderAgencies();
}

function renderAgencies() {
  const wrap = el("agencyBlocks");
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
      state.agencies = state.agencies.filter((a) => a.id !== agency.id);
      renderAgencies();
    });

    block.querySelectorAll("[data-field]").forEach((field) => {
      field.addEventListener("input", (e) => {
        agency[e.target.dataset.field] = e.target.value;
      });
    });

    wrap.appendChild(block);
  });
}

function toggleFlag(key, buttonId) {
  state.incident.flags[key] = !state.incident.flags[key];
  el(buttonId).classList.toggle("active", state.incident.flags[key]);
}

function renderResponders() {
  renderResponderGroup("connewarre", "connewarreList", ["T1", "T2", "Station", "Direct"]);
  renderResponderGroup("mtd", "mtdList", ["MTD P/T", "Station", "Direct"]);
  updateOicBanner();
}

function renderResponderGroup(groupKey, containerId, destinations) {
  const wrap = el(containerId);
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

    const showTruckRole = person.destination === "T1" || person.destination === "T2";

    card.innerHTML = `
      <div class="responder-card-top">
        <div>
          <input type="text" list="${listId}" value="${text(person.name)}" placeholder="Name" />
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

      ${person.destination ? `
        <div class="responder-stage">
          <div class="stage-label">Flags</div>
          <div class="chips responder-flags"></div>
        </div>
      ` : ""}
    `;

    const nameInput = card.querySelector("input");
    nameInput.addEventListener("input", (e) => {
      person.name = e.target.value;
      resolveResponderMemberDetails(groupKey, person);
      updateResponderCardDisplay(card, person, groupKey);
      updateOicBanner();
    });

    card.querySelector(".tiny-btn").addEventListener("click", () => {
      state.responders[groupKey] = state.responders[groupKey].filter((x) => x.id !== person.id);
      if (!state.responders[groupKey].length) state.responders[groupKey].push(createResponder(groupKey));
      renderResponders();
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
        if (dest !== "T1" && dest !== "T2") person.truckRole = "";
        renderResponders();
      });
      destWrap.appendChild(btn);
    });

    if (showTruckRole) {
      const roleWrap = card.querySelector(".responder-roles");
      ["Driver", "CL"].forEach((role) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = `chip-btn ${person.truckRole === role ? "active" : ""}`;
        btn.textContent = role;
        btn.addEventListener("click", () => {
          if (person.truckRole === role) {
            person.truckRole = "";
          } else if (!roleTaken(person.destination, role, person.id)) {
            person.truckRole = role;
          }
          renderResponders();
        });
        roleWrap.appendChild(btn);
      });
    }

    if (person.destination) {
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

        btn.addEventListener("click", () => {
          if (key === "oic") {
            const willBeOic = !person.oic;
            clearAllOic();
            person.oic = willBeOic;
            renderResponders();
            return;
          }

          person[key] = !person[key];
          btn.classList.toggle("active", person[key]);
          updateOicBanner();
        });

        flagsWrap.appendChild(btn);
      });
    }

    wrap.appendChild(card);

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

function roleTaken(destination, role, currentId) {
  return [...state.responders.connewarre, ...state.responders.mtd].some((r) => {
    return r.id !== currentId && r.destination === destination && r.truckRole === role;
  });
}

function clearAllOic() {
  [...state.responders.connewarre, ...state.responders.mtd].forEach((r) => {
    r.oic = false;
  });
}

function buildReport() {
  const all = [...state.responders.connewarre, ...state.responders.mtd];
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
        applianceSections.push(`- ${r.name}${responderSuffix(r)} | ${r.truckRole || "Crew"}`);
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

function finishReport() {
  const all = [...state.responders.connewarre, ...state.responders.mtd];
  const hasResponder = all.some((r) => r.name.trim());
  const hasOic = all.some((r) => r.oic && r.name.trim());

  const needsOic = state.incident.brigadeCode.trim().toUpperCase().startsWith("CONN");

  if (!hasResponder) {
    el("validationText").textContent = "At least one responder must be entered.";
    el("finishActions").classList.add("hidden");
    return;
  }

  if (needsOic && !hasOic) {
    el("validationText").textContent = "OIC must be selected for primary Connewarre jobs.";
    el("finishActions").classList.add("hidden");
    return;
  }

  el("validationText").textContent = "Report ready.";
  el("reportPreview").value = buildReport();
  el("finishActions").classList.remove("hidden");
}

function saveCurrentReport() {
  const title = `${state.incident.eventNumber || "NO_EVENT"} – ${state.incident.incidentType || "UNKNOWN"} – ${state.incident.address || "NO ADDRESS"}`;
  const body = buildReport();

  state.savedReports.unshift({
    id: uid(),
    title,
    body,
    createdAt: new Date().toISOString()
  });

  state.savedReports = state.savedReports.slice(0, 10);
  saveSavedReports();
  renderSavedReports();
}

function renderSavedReports() {
  const wrap = el("savedReports");
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
    });

    delBtn.addEventListener("click", () => {
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
  closeSettings();
}

function sendEmail() {
  const subject = `${state.incident.eventNumber || "NO_EVENT"} – ${state.incident.incidentType || "UNKNOWN"} – ${state.incident.address || "NO ADDRESS"}`;
  const body = buildReport();
  window.location.href = `mailto:${encodeURIComponent(state.profile.email || "")}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

async function sendSms() {
  const body = buildReport();
  try {
    await navigator.clipboard.writeText(body);
  } catch {}
  window.location.href = `sms:?body=${encodeURIComponent(body)}`;
}

function showPage(pageId) {
  state.ui.currentPage = pageId;
  document.querySelectorAll(".page").forEach((p) => p.classList.toggle("active", p.id === pageId));
  document.querySelectorAll(".tab-btn[data-page]").forEach((b) => b.classList.toggle("active", b.dataset.page === pageId));

  if (pageId === "sendPage") {
    el("reportPreview").value = buildReport();
  }
}

function renderEverything() {
  el("brigadeRole").value = state.incident.brigadeRole;
  renderSceneBrigades();
  renderAgencies();
  renderResponders();
  renderSavedReports();
  showPage(state.ui.currentPage);
}
