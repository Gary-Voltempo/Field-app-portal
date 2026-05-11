"use strict";

const CLIENT_ID = "28753d72-1ee9-450b-9946-43247f9b0a24";
const TENANT = "voltempo.com";
const ENGINEER_EMAIL = "commissioning@voltempo.com";
const APPROVER_1_EMAIL = "gary.grant@voltempo.com";
const APPROVER_2_EMAIL = "duble_g@me.com";
const AUTHORISED_ACCOUNTS = [
  ENGINEER_EMAIL,
  APPROVER_1_EMAIL,
  APPROVER_2_EMAIL,
];
const ROLE_LABELS = {
  engineer: "Engineer",
  approver1: "Approver 1",
  approver2: "Approver 2",
};
const APP_ROOT_FOLDER = "Voltempo Field Application";
const SHARED_ROOT_LINK =
  "https://evcpowertechcom-my.sharepoint.com/:f:/r/personal/commissioning_voltempo_com/Documents/Voltempo%20Field%20Application?csf=1&web=1&e=5alMCS";
const BACKUPS_FOLDER = `${APP_ROOT_FOLDER}/app/backups`;
const REPORTS_FILE = "voltempo_commissioning_reports.json";
const LEGACY_REPORTS_FILE = "commissioning_backup.json";
const REPORTS_FILE_PATH = `${BACKUPS_FOLDER}/${REPORTS_FILE}`;
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Files.ReadWrite",
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Mail.Send",
];
const MSAL_SCRIPT_URLS = [
  "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js",
  "https://unpkg.com/@azure/msal-browser@2.38.3/lib/msal-browser.min.js",
];
const PDFJS_SCRIPT_URLS = [
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js",
  "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js",
];

const SELECT_OPTIONS = [
  "-- Select --",
  "Pass",
  "Work Required",
  "N/A",
  "Yes",
  "No",
];

const menuToggle = document.getElementById("menuToggle");
const topbarMenu = document.getElementById("topbarMenu");
const topbar = document.querySelector(".topbar");

function closeTopbarMenu() {
  if (!menuToggle || !topbarMenu) return;
  topbarMenu.classList.remove("open");
  menuToggle.setAttribute("aria-expanded", "false");
}

if (menuToggle && topbarMenu) {
  menuToggle.addEventListener("click", () => {
    const isOpen = topbarMenu.classList.toggle("open");
    menuToggle.setAttribute("aria-expanded", String(isOpen));
  });

  topbarMenu.addEventListener("click", (event) => {
    if (event.target.closest("button")) {
      closeTopbarMenu();
    }
  });

  document.addEventListener("click", (event) => {
    if (
      !topbarMenu.contains(event.target) &&
      !menuToggle.contains(event.target)
    ) {
      closeTopbarMenu();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeTopbarMenu();
    }
  });
}

const STATUS_OPTIONS = [
  "Draft",
  "Submitted",
  "In Review",
  "Changes Requested",
  "Reviewed",
  "Final Approval",
  "Final Approved",
  "Archived",
];
const STATUS_FILTERS = ["All", ...STATUS_OPTIONS];
const WORKFLOW_ACTIONS = {
  Draft: [
    {
      label: "Submit Report",
      status: "Submitted",
      action: "Submitted for review",
      notify: "review",
      roles: ["engineer"],
    },
  ],
  "In Progress": [
    {
      label: "Submit Report",
      status: "Submitted",
      action: "Submitted for review",
      notify: "review",
      roles: ["engineer"],
    },
  ],
  Completed: [
    {
      label: "Submit Report",
      status: "Submitted",
      action: "Submitted for review",
      notify: "review",
      roles: ["engineer"],
    },
  ],
  Submitted: [
    {
      label: "Start Review",
      status: "In Review",
      action: "Review started",
      roles: ["approver1"],
    },
  ],
  "In Review": [
    {
      label: "Request Changes",
      status: "Changes Requested",
      action: "Changes requested",
      notify: "changes",
      roles: ["approver1"],
    },
    {
      label: "Approve Review",
      status: "Reviewed",
      action: "Review approved",
      notify: "final",
      roles: ["approver1"],
    },
  ],
  "Changes Requested": [
    {
      label: "Resubmit",
      status: "Submitted",
      action: "Resubmitted for review",
      notify: "review",
      roles: ["engineer"],
    },
  ],
  Reviewed: [
    {
      label: "Send For Final Approval",
      status: "Final Approval",
      action: "Sent for final approval",
      notify: "final",
      roles: ["approver1"],
    },
  ],
  "Final Approval": [
    {
      label: "Final Approve",
      status: "Final Approved",
      action: "Final approval granted",
      roles: ["approver2"],
    },
  ],
  "Final Approved": [
    {
      label: "Archive",
      status: "Archived",
      action: "Archived",
      roles: ["approver2"],
    },
  ],
};

const FIELD_SECTIONS = [
  {
    title: "Site Information",
    fields: [
      { key: "siteName", label: "Site Name" },
      { key: "siteId", label: "Site ID" },
    ],
  },
  {
    title: "Systems",
    fields: [
      {
        key: "chargingDistribution",
        label: "Charging Distribution",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "chargingSatelites",
        label: "Charging Satellites",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "coolingSystem",
        label: "Cooling System",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "communication",
        label: "Communication",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preChecksSystemsCertificateName",
        label: "Systems Certificate / PDF",
      },
      {
        key: "preChecksSystemsCertificatePath",
        label: "Systems Certificate Path",
        multiline: true,
      },
    ],
  },
  {
    title: "Incoming Supply",
    fields: [
      { key: "supplyCable", label: "Supply Cable", multiline: true },
      { key: "supplyType", label: "Supply Type" },
      { key: "nominalVoltage", label: "Nominal Voltage" },
      { key: "earthingArrangement", label: "Earthing Arrangement" },
      { key: "maxCurrent", label: "Maximum Available Current" },
      { key: "circuitProtective", label: "Circuit Protective Device" },
    ],
  },
  {
    title: "Transformer Details",
    fields: [
      { key: "transformerSerial", label: "Serial Number" },
      { key: "transformerManufacturer", label: "Manufacturer" },
      { key: "transformerYear", label: "Year" },
      { key: "transformerRatedPower", label: "Rated Power" },
      { key: "transformerPrimaryVoltage", label: "Primary Voltage" },
      { key: "transformerSecondaryVoltage", label: "Secondary Voltage" },
      { key: "transformerFrequency", label: "Frequency" },
    ],
  },
  {
    title: "Pre-Checks",
    fields: [
      {
        key: "manufacturingTestCertificate",
        label: "Manufacturing Test Certificate",
      },
      {
        key: "manufacturingTestCertificatePath",
        label: "Certificate Path",
        multiline: true,
      },
      {
        key: "preCheckAC",
        label: "AC Supply Available",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preCheckDC",
        label: "DC Supply Available",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preCheckPermits",
        label: "Permits Complete",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preCheckRAMS",
        label: "RAMS Complete",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preCheckACB",
        label: "ACB Checked",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preCheckACSupply",
        label: "AC Supply Checked",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "preCheckSafety",
        label: "Safety Signs Present",
        type: "select",
        options: ["", "Yes", "No"],
      },
      {
        key: "ev_charger_test_report_name",
        label: "EV Charger Test Report",
      },
      {
        key: "ev_charger_test_report_path",
        label: "EV Charger Test Report Path",
        multiline: true,
      },
    ],
  },
  {
    title: "Sign Off",
    fields: [
      { key: "summaryFullName", label: "Full Name" },
      { key: "summaryDate", label: "Date" },
      { key: "summaryTime", label: "Time" },
    ],
  },
];

const INSPECTION_SECTIONS = [
  {
    title: "Foundation Checks",
    rows: [
      {
        key: "foundation_ducts",
        label: "Are ducts installed correctly with no damage to cabling",
      },
      {
        key: "foundation_cables",
        label: "Are all DC Cables installed to correct standard",
      },
      {
        key: "foundation_system",
        label: "System fixed down adequately",
      },
      {
        key: "foundation_concrete",
        label: "No signs of damage to concrete base",
      },
    ],
  },
  {
    title: "Charge Pod Visual Inspections",
    rows: [
      {
        key: "charge_pod_inspect_damage",
        label: "Inspect pod and charge points for damage",
      },
      {
        key: "charge_pod_inspect_seals",
        label: "Inspect door seals",
      },
    ],
  },
  {
    title: "Pod Visual Inspections",
    rows: [
      {
        key: "pod_section_1_witness_marks",
        label: "Check witness marks on AC incoming bus bars for movement",
      },
      {
        key: "pod_section_1_inspect_lv",
        label:
          "Inspect LV cabinet for secure terminations, seated connectors, inserted fuses, and visible damage",
      },
      {
        key: "pod_section_2_fan_thermostats",
        label: "Ensure fan thermostats are set to 25C",
      },
      {
        key: "pod_section_2_network_compartment",
        label: "Inspect network compartment connectors are seated correctly",
      },
      {
        key: "pod_section_2_heater_control",
        label: "Ensure heater Control is set to 75%",
      },
      {
        key: "pod_section_5_pc_vsecc",
        label: "Inspect PC & vSECC connectors are fully seated",
      },
      {
        key: "pod_section_5_pc_connections",
        label: "Ensure PC connections are correct and seated correctly",
      },
    ],
  },
  {
    title: "Charge Satellites",
    rows: [1, 2, 3, 4, 5, 6].map((number) => ({
      key: `charge_point_${number}`,
      label: `Charge Satellite ${number}`,
    })),
  },
  {
    title: "Chiller Unit Visual Inspection",
    rows: [
      {
        key: "chiller_unit_exterior_panels",
        label: "Is there any damage to the exterior panels of the Chiller Unit",
      },
      {
        key: "chiller_unit_door_seals",
        label: "Ensure all door seals are intact and free from damage",
      },
      {
        key: "chiller_unit_pipes_secure",
        label: "Ensure pipes are secure and free from damage with no leaks",
      },
      {
        key: "chiller_unit_coolant_levels",
        label: "Check coolant levels",
      },
    ],
  },
  {
    title: "Live Checks",
    rows: [
      {
        key: "live_checks_mcb6_mcb7",
        fallbackKey: "Visual_checks_mcb6_mcb7",
        label: "Check voltage at MCB6 & MCB7 are 400V AC",
      },
      {
        key: "live_checks_mcb1",
        fallbackKey: "Visual_checks_mcb1",
        label: "Check voltage at MCB1 is 230V AC",
      },
      {
        key: "live_checks_mcb1_mcb3",
        fallbackKey: "Visual_checks_mcb1_mcb3",
        label: "Turn ON MCB1-3 & RCD1; verify PSU3 & PSU4 ON",
      },
      {
        key: "live_checks_mcb4",
        fallbackKey: "Visual_checks_mcb4",
        label: "Turn ON MCB4 check voltage at transformer TA1 is 110v AC",
      },
      {
        key: "live_checks_mcb6_mcb7_psu",
        fallbackKey: "Visual_checks_mcb6_mcb7_psu",
        label: "Turn ON MCB6 & MCB7; confirm PSU1 & PSU2 ON; check 24V DC",
      },
      {
        key: "live_checks_power_modules",
        fallbackKey: "Visual_checks_power_modules",
        label: "Energise power modules, ensure each power module in in SLP mode",
      },
      {
        key: "live_checks_thermostats",
        fallbackKey: "Visual_checks_thermostats",
        label: "Turn thermostats to 0C and confirm fans run. Reset thermostats to 25C",
      },
      {
        key: "live_checks_heater_controls",
        fallbackKey: "Visual_checks_heater_controls",
        label: "Ensure heater controls is set to 75%",
      },
      {
        key: "live_checks_control_boards",
        fallbackKey: "Visual_checks_control_boards",
        label: "Ensure control boards are powered on",
      },
      {
        key: "live_checks_vseccs",
        fallbackKey: "Visual_checks_vseccs",
        label: "Ensure vSECCs and PC are powered on",
      },
    ],
  },
];

const state = {
  msalClient: null,
  account: null,
  reportsFile: null,
  reports: [],
  selectedId: null,
  currentRaw: null,
  emptyTitle: "No report selected",
  emptyMessage: "",
  isDirty: false,
  isBusy: false,
  statusFilter: "All",
  targetReportId: "",
  renderedAssets: [],
  viewerAssetIds: [],
  viewerIndex: 0,
  sharedRoot: null,
  pdfViewer: {
    document: null,
    page: 1,
    scale: 1,
    session: 0,
  },
  pendingAttachmentTarget: null,
};

const dom = {};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  bindDom();
  bindEvents();
  bindTopbarScroll();
  state.targetReportId = reportIdFromUrl();
  setBusy(true, "Starting");

  try {
    await ensureMsalClient();
    const response = await state.msalClient.handleRedirectPromise();
    const account = response?.account || state.msalClient.getAllAccounts()[0];
    if (account) {
      await activateAccount(account);
      await syncReports();
    } else {
      renderSignedOut();
    }
  } catch (error) {
    showToast(error.message || String(error), "error");
    renderAuthUnavailable(error.message || String(error));
  } finally {
    setBusy(false);
  }
}

async function ensureMsalClient() {
  if (state.msalClient) return state.msalClient;

  await loadMsalLibrary();

  state.msalClient = new window.msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT}`,
      redirectUri: authRedirectUri(),
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  });

  return state.msalClient;
}

function authRedirectUri() {
  const { origin, pathname } = window.location;

  if (pathname.endsWith("/index.html")) {
    return `${origin}${pathname.slice(0, -"index.html".length)}`;
  }

  if (!pathname.endsWith("/")) {
    return `${origin}${pathname}/`;
  }

  return `${origin}${pathname}`;
}

async function loadMsalLibrary() {
  if (window.msal?.PublicClientApplication) return;

  for (const url of MSAL_SCRIPT_URLS) {
    try {
      await loadScript(url);
      if (window.msal?.PublicClientApplication) return;
    } catch (error) {
      // Try the next CDN.
    }
  }

  throw new Error(
    "Microsoft sign-in library could not be loaded. Check internet access or allow the MSAL CDN.",
  );
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", reject, { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

function bindDom() {
  dom.cloudState = document.getElementById("cloudState");
  dom.signInButton = document.getElementById("signInButton");
  dom.submitReportButton = document.getElementById("submitReportButton");
  dom.syncButton = document.getElementById("syncButton");
  dom.importButton = document.getElementById("importButton");
  dom.uploadButton = document.getElementById("uploadButton");
  dom.saveButton = document.getElementById("saveButton");
  dom.exportButton = document.getElementById("exportButton");
  dom.signOutButton = document.getElementById("signOutButton");
  dom.fileInput = document.getElementById("fileInput");
  dom.attachmentFileInput = document.getElementById("attachmentFileInput");
  dom.searchInput = document.getElementById("searchInput");
  dom.reportCount = document.getElementById("reportCount");
  dom.dirtyState = document.getElementById("dirtyState");
  dom.reportList = document.getElementById("reportList");
  dom.emptyState = document.getElementById("emptyState");
  dom.reportEditor = document.getElementById("reportEditor");
  dom.toast = document.getElementById("toast");
  dom.assetViewer = document.getElementById("assetViewer");
  dom.assetViewerTitle = document.getElementById("assetViewerTitle");
  dom.assetViewerMeta = document.getElementById("assetViewerMeta");
  dom.assetViewerStage = document.getElementById("assetViewerStage");
  dom.assetViewerDownload = document.getElementById("assetViewerDownload");
}

function bindEvents() {
  dom.signInButton.addEventListener("click", signIn);
  dom.submitReportButton.addEventListener("click", submitCurrentReport);
  dom.signOutButton.addEventListener("click", signOut);
  dom.syncButton.addEventListener("click", syncReports);
  dom.importButton.addEventListener("click", () => dom.fileInput.click());
  dom.uploadButton.addEventListener("click", uploadLoadedReportsFile);
  dom.fileInput.addEventListener("change", importReportsFileFromDisk);
  dom.attachmentFileInput.addEventListener("change", handleAttachmentFileSelected);
  dom.saveButton.addEventListener("click", saveCurrentReport);
  dom.exportButton.addEventListener("click", exportCurrentReportPdf);
  dom.searchInput.addEventListener("input", renderReportList);
  dom.reportList.addEventListener("click", handleReportListClick);
  dom.reportEditor.addEventListener("input", handleEditorInput);
  dom.reportEditor.addEventListener("change", handleEditorInput);
  dom.reportEditor.addEventListener("click", handleEditorClick);
  dom.assetViewer.addEventListener("click", handleAssetViewerClick);
  document.addEventListener("keydown", handleGlobalKeydown);
}

function bindTopbarScroll() {
  const editorPane = document.querySelector(".editor-pane");
  const reportList = document.querySelector(".report-list");

  window.addEventListener("scroll", updateTopbarScrollState, { passive: true });
  editorPane?.addEventListener("scroll", updateTopbarScrollState, {
    passive: true,
  });
  reportList?.addEventListener("scroll", updateTopbarScrollState, {
    passive: true,
  });
  updateTopbarScrollState();
}

function updateTopbarScrollState() {
  if (!topbar) return;
  const editorPane = document.querySelector(".editor-pane");
  const reportList = document.querySelector(".report-list");
  const isScrolled =
    window.scrollY > 10 ||
    (editorPane?.scrollTop || 0) > 10 ||
    (reportList?.scrollTop || 0) > 10;

  topbar.classList.toggle("scrolled", isScrolled);
}

function setControlLabel(element, label) {
  if (!element) return;
  const labelNode = element.querySelector(".button-label, .state-label");
  if (labelNode) {
    labelNode.textContent = label;
  } else {
    element.textContent = label;
  }
  element.setAttribute("aria-label", label);
}

async function signIn() {
  if (state.account) return;
  setBusy(true, "Signing in");

  try {
    await ensureMsalClient();
    await state.msalClient.loginRedirect({
      scopes: GRAPH_SCOPES,
      prompt: "select_account",
    });
  } catch (error) {
    showToast(error.message || String(error), "error");
    setBusy(false);
  }
}

async function activateAccount(account) {
  if (!account || !isRequiredAccount(account)) {
    throw new Error("Please sign in with an authorised approval workflow account.");
  }

  state.account = account;
  state.msalClient.setActiveAccount(account);
  renderSignedIn();
}

async function signOut() {
  const account = state.account;
  state.account = null;
  state.reportsFile = null;
  state.reports = [];
  state.selectedId = null;
  state.currentRaw = null;
  state.isDirty = false;
  renderSignedOut();

  if (account) {
    await state.msalClient.logoutRedirect({ account });
  }
}

function renderSignedIn() {
  setControlLabel(dom.cloudState, "Signed in");
  dom.cloudState.className = "state-chip online";
  setControlLabel(dom.signInButton, accountDisplayName());
  dom.signInButton.className = "button subtle account-button";
  dom.signInButton.hidden = false;
  dom.signInButton.disabled = false;
  dom.signOutButton.hidden = false;
  dom.submitReportButton.disabled = !canSubmitSelectedReport();
  dom.syncButton.disabled = false;
  dom.importButton.disabled = false;
  dom.uploadButton.disabled = !state.reportsFile;
  dom.saveButton.disabled = !state.selectedId || !state.isDirty;
  dom.exportButton.disabled = !state.selectedId;
}

function renderSignedOut() {
  setControlLabel(dom.cloudState, "Offline");
  dom.cloudState.className = "state-chip";
  setControlLabel(dom.signInButton, "Sign In");
  dom.signInButton.className = "button primary";
  dom.signInButton.disabled = false;
  dom.signInButton.hidden = false;
  dom.signOutButton.hidden = true;
  dom.submitReportButton.disabled = true;
  dom.syncButton.disabled = true;
  dom.importButton.disabled = false;
  dom.uploadButton.disabled = true;
  dom.saveButton.disabled = true;
  dom.exportButton.disabled = true;
  dom.reportList.innerHTML = "";
  dom.reportCount.textContent = "0 reports";
  dom.dirtyState.textContent = "";
  state.emptyTitle = "Sign in to load reports";
  state.emptyMessage = "";
  dom.reportEditor.hidden = true;
  dom.emptyState.hidden = false;
}

function accountDisplayName() {
  const name =
    state.account?.name ||
    state.account?.username ||
    state.account?.idTokenClaims?.preferred_username ||
    "Signed in";
  const role = ROLE_LABELS[currentUserRole()];
  return role ? `${name} (${role})` : name;
}

function renderAuthUnavailable(message) {
  renderSignedOut();
  dom.signInButton.disabled = false;
  setControlLabel(dom.signInButton, "Retry Sign In");
  state.emptyTitle = "Microsoft sign-in is not ready";
  state.emptyMessage = message;
  renderEditor();
}

async function getAccessToken() {
  const request = {
    scopes: GRAPH_SCOPES,
    account: state.account,
  };

  try {
    const response = await state.msalClient.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    await state.msalClient.acquireTokenRedirect(request);
    throw new Error("Redirecting to Microsoft sign-in for permission approval.");
  }
}

async function syncReports() {
  if (!state.account) {
    showToast("Sign in first.", "warn");
    return;
  }

  if (state.isDirty) {
    const shouldContinue = window.confirm(
      "You have unsaved edits. Syncing will reload the cloud copy.",
    );
    if (!shouldContinue) return;
  }

  setBusy(true, "Syncing");

  try {
    const token = await getAccessToken();
    await ensureSharedRoot(token);
    const file =
      (await downloadJsonFileRecord(token, REPORTS_FILE_PATH)) ||
      (await downloadJsonFileRecord(token, REPORTS_FILE)) ||
      (await downloadJsonFileRecord(token, LEGACY_REPORTS_FILE));
    const data = file?.data || createEmptyReportsFile();
    const loadedFileName = file ? displayFileName(file.fileName) : "";

    if (file) {
      await hydrateReportAttachmentUrls(token, data);
    }

    loadReportsFile(data);
    state.emptyTitle = "No reports found";
    state.emptyMessage = file
      ? `${loadedFileName} exists, but it does not contain any saved reports yet.`
      : "No OneDrive reports file exists yet. Open the mobile app, then upload your saved reports.";
    renderReportList();
    selectFirstReportIfNeeded();
    if (!file) {
      showToast("No OneDrive reports file found yet.", "warn");
    } else if (!state.reports.length) {
      showToast(`${loadedFileName} contains no saved reports yet.`, "warn");
    } else {
      showToast(
        `Loaded ${state.reports.length} reports from ${loadedFileName}.`,
        "success",
      );
    }
  } catch (error) {
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

async function importReportsFileFromDisk(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;

  if (state.isDirty) {
    const shouldContinue = window.confirm(
      "You have unsaved edits. Importing a file will replace the loaded data.",
    );
    if (!shouldContinue) return;
  }

  try {
    const text = await file.text();
    const data = JSON.parse(text);
    loadReportsFile(data);
    state.emptyTitle = "No reports found";
    state.emptyMessage = `${file.name} imported, but it does not contain any saved reports.`;
    renderReportList();
    selectFirstReportIfNeeded();
    showToast(
      `Imported ${state.reports.length} reports from ${file.name}.`,
      state.reports.length ? "success" : "warn",
    );
  } catch (error) {
    showToast(`Import failed: ${error.message || String(error)}`, "error");
  }
}

async function uploadLoadedReportsFile() {
  if (!state.account) {
    showToast("Sign in first.", "warn");
    return;
  }

  if (!state.reportsFile) {
    state.reportsFile = createEmptyReportsFile();
  }

  setBusy(true, "Uploading");

  try {
    if (state.isDirty) {
      applyCurrentReportEditsToState();
    }

    const token = await getAccessToken();
    await uploadReportsFile(token);

    state.isDirty = false;
    updateDirtyState();
    renderReportList();
    renderEditor();
    showToast(
      `Uploaded ${state.reports.length} reports to OneDrive.`,
      "success",
    );
  } catch (error) {
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

function displayFileName(filePath) {
  return String(filePath || "").split("/").pop() || REPORTS_FILE;
}

async function downloadJsonFileRecord(token, fileName) {
  const data = await downloadJsonFile(token, fileName);
  return data ? { fileName, data } : null;
}

async function downloadJsonFile(token, fileName) {
  await ensureSharedRoot(token);
  const response = await fetch(
    graphDriveFileContentUrl(fileName),
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  );

  if (response.status === 404) return null;

  if (!response.ok) {
    throw new Error(await graphError(response, `Could not download ${fileName}`));
  }

  return response.json();
}

async function hydrateReportAttachmentUrls(token, reportsFile) {
  const attachments = collectAttachmentObjects(reportsFile);
  const linkedAttachments = attachments.filter(
    (attachment) => attachment.driveItemId || attachment.storagePath,
  );
  const searchCache = new Map();

  await Promise.all(
    linkedAttachments.slice(0, 160).map(async (attachment) => {
      try {
        const driveItem = await fetchDriveItemForAttachment(token, attachment);
        if (!driveItem) return;

        attachment.driveItemId = driveItem.id || attachment.driveItemId || "";
        attachment.downloadUrl =
          driveItem["@microsoft.graph.downloadUrl"] || attachment.downloadUrl || "";
        attachment.webUrl = driveItem.webUrl || attachment.webUrl || "";
      } catch (_error) {
        // Existing URLs and file tiles still give the user useful context.
      }
    }),
  );

  await hydratePathOnlyReportAssets(token, reportsFile, searchCache);
}

async function fetchDriveItemForAttachment(token, attachment) {
  const url = attachment.driveItemId
    ? graphDriveItemByIdUrl(attachment.driveItemId)
    : graphDriveItemUrl(normalizeAttachmentStoragePath(attachment));

  if (!url) return null;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) return null;
  return response.json();
}

function collectAttachmentObjects(value, found = []) {
  if (Array.isArray(value)) {
    value.forEach((item) => collectAttachmentObjects(item, found));
    return found;
  }

  if (!value || typeof value !== "object") return found;

  if (
    value.driveItemId ||
    value.webUrl ||
    value.downloadUrl ||
    value.storagePath ||
    value.localPath
  ) {
    found.push(value);
  }

  Object.values(value).forEach((item) => collectAttachmentObjects(item, found));
  return found;
}

async function hydratePathOnlyReportAssets(token, reportsFile, searchCache) {
  const reports = Array.isArray(reportsFile?.savedReports)
    ? reportsFile.savedReports
    : [];

  for (const report of reports) {
    const rawData = report?.rawData;
    if (!rawData || typeof rawData !== "object") continue;
    await hydrateRawDataPathAssets(token, rawData, searchCache);
  }
}

async function hydrateRawDataPathAssets(token, rawData, searchCache) {
  for (const [key, value] of Object.entries(rawData)) {
    if (typeof value === "string" && value.trim()) {
      if (/_image$/i.test(key) || isAssetPathField(key)) {
        await hydratePathAssetField(token, rawData, key, value, searchCache);
      }
      continue;
    }

    if (Array.isArray(value) && key === "additional_observations") {
      for (const item of value) {
        if (!item || typeof item !== "object") continue;
        const paths = Array.isArray(item.imagePaths) ? item.imagePaths : [];
        if (!paths.length || Array.isArray(item.imageAttachments)) continue;
        item.imageAttachments = [];
        for (const path of paths) {
          const attachment = await attachmentForPathOnlyValue(
            token,
            "additional_observations",
            path,
            searchCache,
          );
          if (attachment) item.imageAttachments.push(attachment);
        }
      }
    }
  }
}

async function hydratePathAssetField(token, rawData, key, pathValue, searchCache) {
  const attachmentKey = attachmentKeyForSource(key);
  const currentAttachments = Array.isArray(rawData[attachmentKey])
    ? rawData[attachmentKey]
    : [];
  if (currentAttachments.some((asset) => asset?.downloadUrl || asset?.webUrl)) {
    return;
  }

  const attachment = await attachmentForPathOnlyValue(
    token,
    key,
    pathValue,
    searchCache,
  );
  if (attachment) rawData[attachmentKey] = [attachment];
}

async function attachmentForPathOnlyValue(token, sourceKey, pathValue, searchCache) {
  const name = assetLabel(pathValue);
  if (!name || !/\.[a-z0-9]{2,5}$/i.test(name)) return null;

  const driveItem = await findDriveItemByName(token, name, searchCache);
  if (!driveItem) return null;

  return {
    id: `att_${driveItem.id || Date.now()}`,
    fieldKey: sourceKey,
    kind: assetKind({
      name,
      mimeType: driveItem.file?.mimeType || "",
    }),
    category: driveItem.file?.mimeType === "application/pdf" ? "Certificates" : "Images",
    name: driveItem.name || name,
    mimeType: driveItem.file?.mimeType || mimeTypeForFileName(name),
    sizeBytes: driveItem.size || 0,
    localPath: pathValue,
    storagePath: driveItem.parentReference?.path || "",
    driveItemId: driveItem.id || "",
    webUrl: driveItem.webUrl || "",
    downloadUrl: driveItem["@microsoft.graph.downloadUrl"] || "",
    syncStatus: "uploaded",
    createdAt: new Date().toISOString(),
  };
}

async function findDriveItemByName(token, fileName, searchCache) {
  const variants = fileNameSearchVariants(fileName);
  const cacheKey = variants.join("|").toLowerCase();
  if (searchCache.has(cacheKey)) return searchCache.get(cacheKey);

  try {
    const results = [];
    for (const variant of variants) {
      const escaped = variant.replace(/'/g, "''");
      const response = await fetch(
        graphDriveSearchUrl(escaped),
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        },
      );

      if (!response.ok) continue;

      const result = await response.json();
      if (Array.isArray(result.value)) {
        results.push(...result.value);
      }
    }

    const normalizedVariants = variants.map(normalizeFileSearchName);
    const match =
      results.find((item) =>
        normalizedVariants.includes(normalizeFileSearchName(item.name || "")),
      ) ||
      results.find((item) =>
        normalizedVariants.some((variant) =>
          normalizeFileSearchName(item.name || "").includes(variant),
        ),
      );
    searchCache.set(cacheKey, match || null);
    return match || null;
  } catch (_error) {
    searchCache.set(cacheKey, null);
    return null;
  }
}

function fileNameSearchVariants(fileName) {
  const clean = assetLabel(fileName);
  const withoutTimestamp = clean.replace(/^\d{10,17}[_-]+/, "");
  const withoutAttachmentId = withoutTimestamp.replace(/^att_\d{10,17}[_-]+/, "");
  const withoutFieldPrefix = withoutAttachmentId.replace(
    /^[a-z0-9_ ]+(?:image|report|certificate|pdf)[_-]+\d{10,17}[_-]+/i,
    "",
  );

  return [...new Set([clean, withoutTimestamp, withoutAttachmentId, withoutFieldPrefix])]
    .map((value) => value.trim())
    .filter(Boolean);
}

function normalizeFileSearchName(value) {
  return valueToString(value)
    .toLowerCase()
    .replace(/^\d{10,17}[_-]+/, "")
    .replace(/^att_\d{10,17}[_-]+/, "")
    .replace(/\s+/g, " ")
    .trim();
}

async function uploadReportsFile(token) {
  await ensureSharedRoot(token);
  state.reportsFile.updatedAt = new Date().toISOString();
  state.reportsFile.savedReports = state.reports;
  await ensureDriveFolderPath(token, BACKUPS_FOLDER);

  const response = await fetch(
    graphDriveFileContentUrl(REPORTS_FILE_PATH),
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(state.reportsFile, null, 2),
    },
  );

  if (!response.ok) {
    throw new Error(await graphError(response, "Could not save reports"));
  }
}

async function ensureSharedRoot(token) {
  if (state.sharedRoot) return state.sharedRoot;
  if (!SHARED_ROOT_LINK) return null;

  const response = await fetch(`${GRAPH_ROOT}/shares/${shareIdForUrl(SHARED_ROOT_LINK)}/driveItem`, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error(await graphError(response, "Could not open shared Voltempo Field Application folder"));
  }

  const item = await response.json();
  state.sharedRoot = {
    driveId: item.parentReference?.driveId || item.remoteItem?.parentReference?.driveId || "",
    itemId: item.id || item.remoteItem?.id || "",
    name: item.name || APP_ROOT_FOLDER,
  };

  if (!state.sharedRoot.driveId || !state.sharedRoot.itemId) {
    throw new Error("The shared OneDrive folder did not return a usable drive ID.");
  }

  return state.sharedRoot;
}

function shareIdForUrl(url) {
  return `u!${base64UrlEncode(url)}`;
}

function base64UrlEncode(value) {
  return btoa(value)
    .replace(/=+$/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

function graphRootFileContentUrl(fileName) {
  return graphDriveFileContentUrl(fileName);
}

function graphDriveFileContentUrl(filePath) {
  return `${graphDrivePathUrl(filePath)}:/content`;
}

function graphDriveItemUrl(filePath) {
  return graphDrivePathUrl(filePath);
}

function graphDriveItemByIdUrl(itemId) {
  if (state.sharedRoot?.driveId) {
    return `${GRAPH_ROOT}/drives/${encodeURIComponent(state.sharedRoot.driveId)}/items/${encodeURIComponent(itemId)}`;
  }
  return `${GRAPH_ROOT}/me/drive/items/${encodeURIComponent(itemId)}`;
}

function graphDriveItemContentByIdUrl(itemId) {
  return `${graphDriveItemByIdUrl(itemId)}/content`;
}

function graphDrivePathUrl(filePath) {
  const relativePath = driveRelativePath(filePath);
  if (state.sharedRoot?.driveId && state.sharedRoot?.itemId) {
    const root = `${GRAPH_ROOT}/drives/${encodeURIComponent(state.sharedRoot.driveId)}/items/${encodeURIComponent(state.sharedRoot.itemId)}`;
    return relativePath ? `${root}:/${encodeDrivePath(relativePath)}` : root;
  }
  return `${GRAPH_ROOT}/me/drive/root:/${encodeDrivePath(relativePath)}`;
}

function graphDriveChildrenUrl(folderPath) {
  const relativePath = driveRelativePath(folderPath);
  if (state.sharedRoot?.driveId && state.sharedRoot?.itemId) {
    const root = `${GRAPH_ROOT}/drives/${encodeURIComponent(state.sharedRoot.driveId)}/items/${encodeURIComponent(state.sharedRoot.itemId)}`;
    return relativePath ? `${root}:/${encodeDrivePath(relativePath)}:/children` : `${root}/children`;
  }
  return relativePath
    ? `${GRAPH_ROOT}/me/drive/root:/${encodeDrivePath(relativePath)}:/children`
    : `${GRAPH_ROOT}/me/drive/root/children`;
}

function graphDriveSearchUrl(query) {
  const encoded = encodeURIComponent(query);
  if (state.sharedRoot?.driveId && state.sharedRoot?.itemId) {
    return `${GRAPH_ROOT}/drives/${encodeURIComponent(state.sharedRoot.driveId)}/items/${encodeURIComponent(state.sharedRoot.itemId)}/search(q='${encoded}')?$top=25`;
  }
  return `${GRAPH_ROOT}/me/drive/root/search(q='${encoded}')?$top=25`;
}

function driveRelativePath(filePath) {
  let path = valueToString(filePath).replace(/^\/+/, "");
  const rootPrefix = `${APP_ROOT_FOLDER}/`;
  if (path === APP_ROOT_FOLDER) return "";
  if (path.startsWith(rootPrefix)) {
    path = path.slice(rootPrefix.length);
  }
  return path;
}

async function ensureDriveFolderPath(token, folderPath) {
  const parts = folderPath.split("/").filter(Boolean);
  let currentPath = "";

  for (const part of parts) {
    const nextPath = currentPath ? `${currentPath}/${part}` : part;
    if (await drivePathExists(token, nextPath)) {
      currentPath = nextPath;
      continue;
    }
    await createDriveFolder(token, currentPath, part);
    currentPath = nextPath;
  }
}

async function drivePathExists(token, folderPath) {
  const response = await fetch(graphDriveItemUrl(folderPath), {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (response.status === 404) return false;
  if (!response.ok) {
    throw new Error(await graphError(response, `Could not check ${folderPath}`));
  }
  return true;
}

async function createDriveFolder(token, parentPath, folderName) {
  const response = await fetch(graphDriveChildrenUrl(parentPath), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    }),
  });

  if (!response.ok) {
    throw new Error(await graphError(response, `Could not create ${folderName}`));
  }
}

function encodeDrivePath(filePath) {
  return filePath
    .split("/")
    .filter(Boolean)
    .map(encodeURIComponent)
    .join("/");
}

function createEmptyReportsFile() {
  return {
    schemaVersion: 2,
    updatedAt: new Date().toISOString(),
    lastSaved: null,
    draft: null,
    savedReports: [],
  };
}

function loadReportsFile(data) {
  const file = data && typeof data === "object" ? data : createEmptyReportsFile();
  const reports = Array.isArray(file.savedReports)
    ? file.savedReports.filter((item) => item && typeof item === "object")
    : [];

  state.reportsFile = {
    schemaVersion: file.schemaVersion || 2,
    updatedAt: file.updatedAt || new Date().toISOString(),
    lastSaved: file.lastSaved || null,
    draft: file.draft || null,
    savedReports: reports,
  };

  state.reports = reports.map(normalizeReport).sort(sortReports);
  state.reportsFile.savedReports = state.reports;
  state.isDirty = false;
  updateDirtyState();
}

function normalizeReport(report) {
  const rawData =
    report.rawData && typeof report.rawData === "object" ? report.rawData : {};
  const now = new Date().toISOString();
  const id = report.id || `RPT-${Date.now()}`;

  return {
    ...report,
    id,
    title: report.title || buildReportTitle(rawData),
    type: report.type || "Commissioning",
    status: report.status || "Draft",
    date: report.date || formatDate(new Date()),
    savedDate: report.savedDate || formatDateTime(new Date()),
    createdAt: report.createdAt || report.updatedAt || now,
    updatedAt: report.updatedAt || report.createdAt || now,
    location: report.location || rawData.siteName || rawData.siteId || "Unknown site",
    size: report.size || reportSize(rawData),
    tags: Array.isArray(report.tags) ? report.tags : ["saved"],
    notes: report.notes || "",
    workflowHistory: Array.isArray(report.workflowHistory)
      ? report.workflowHistory
      : [],
    isSaved: report.isSaved !== false,
    siteId: report.siteId || rawData.siteId || "",
    rawData: {
      ...rawData,
      editingReportId: rawData.editingReportId || id,
    },
  };
}

function sortReports(a, b) {
  return toDateValue(b.updatedAt || b.createdAt || b.savedDate) -
    toDateValue(a.updatedAt || a.createdAt || a.savedDate);
}

function toDateValue(value) {
  const date = value ? Date.parse(value) : 0;
  return Number.isFinite(date) ? date : 0;
}

function selectFirstReportIfNeeded() {
  if (!state.reports.length) {
    state.selectedId = null;
    state.currentRaw = null;
    renderEditor();
    return;
  }

  const targetId = state.targetReportId;
  if (targetId && state.reports.some((report) => report.id === targetId)) {
    state.statusFilter = "All";
    selectReport(targetId);
    state.targetReportId = "";
  } else if (targetId) {
    showToast("The report from the email link was not found in the loaded reports.", "warn");
    state.targetReportId = "";
  } else if (!state.selectedId || !state.reports.some((r) => r.id === state.selectedId)) {
    selectReport(state.reports[0].id);
  } else {
    selectReport(state.selectedId);
  }
}

function handleReportListClick(event) {
  const actionTarget = event.target.closest("[data-action]");
  if (actionTarget?.dataset.action === "status-filter") {
    state.statusFilter = actionTarget.dataset.status || "All";
    renderReportList();
    return;
  }

  const item = event.target.closest("[data-report-id]");
  if (!item) return;

  if (state.isDirty) {
    const shouldContinue = window.confirm(
      "You have unsaved edits. Change report without saving?",
    );
    if (!shouldContinue) return;
  }

  selectReport(item.dataset.reportId);
}

function selectReport(reportId) {
  const report = state.reports.find((item) => item.id === reportId);
  if (!report) return;

  state.selectedId = reportId;
  state.currentRaw = cloneJson(report.rawData || {});
  state.isDirty = false;
  updateReportUrl(reportId);
  renderReportList();
  renderEditor();
  updateDirtyState();
}

function renderReportList() {
  const query = dom.searchInput.value.trim().toLowerCase();
  const reports = state.reports.filter((report) => {
    const status = report.status || "Draft";
    const matchesStatus =
      state.statusFilter === "All" || status === state.statusFilter;
    if (!matchesStatus) return false;
    if (!query) return true;
    return [
      report.title,
      report.location,
      report.id,
      report.notes,
      report.siteId,
    ]
      .join(" ")
      .toLowerCase()
      .includes(query);
  });

  dom.reportCount.textContent = `${reports.length} ${
    reports.length === 1 ? "report" : "reports"
  }`;

  if (!reports.length) {
    dom.reportList.innerHTML = `
      ${renderApprovalQueueSummary()}
      ${renderStatusFilters()}
      <div class="empty-list">No reports found</div>
    `;
    return;
  }

  dom.reportList.innerHTML = `
    ${renderApprovalQueueSummary()}
    ${renderStatusFilters()}
    ${reports
      .map((report) => {
      const selected = report.id === state.selectedId ? " selected" : "";
      const statusClass = statusClassName(report.status);
      return `
        <button class="report-item${selected}" type="button" data-report-id="${escapeAttr(
          report.id,
        )}">
          <span class="report-title">${escapeHtml(report.title)}</span>
          <span class="report-meta">
            <span class="status-pill ${statusClass}">${escapeHtml(
              report.status,
            )}</span>
            <span>${escapeHtml(report.location || "Unknown site")}</span>
            <span>${escapeHtml(report.savedDate || "")}</span>
          </span>
        </button>
      `;
      })
      .join("")}
  `;
}

function renderApprovalQueueSummary() {
  const queueItems = [
    { label: "To Review", status: "Submitted" },
    { label: "In Review", status: "In Review" },
    { label: "Changes", status: "Changes Requested" },
    { label: "Final Sign-Off", status: "Final Approval" },
  ];

  return `
    <div class="approval-queue" aria-label="Approval queue">
      ${queueItems
        .map((item) => {
          const count = state.reports.filter(
            (report) => (report.status || "Draft") === item.status,
          ).length;
          const active = state.statusFilter === item.status ? " active" : "";
          return `
            <button class="queue-card${active}" type="button" data-action="status-filter" data-status="${escapeAttr(item.status)}">
              <span>${escapeHtml(item.label)}</span>
              <strong>${count}</strong>
            </button>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderStatusFilters() {
  return `
    <div class="status-filter-row">
      ${STATUS_FILTERS.map((status) => {
        const active = state.statusFilter === status ? " active" : "";
        return `<button class="status-filter${active}" type="button" data-action="status-filter" data-status="${escapeAttr(status)}">${escapeHtml(status)}</button>`;
      }).join("")}
    </div>
  `;
}

function renderEditor() {
  const report = selectedReport();

  if (!report || !state.currentRaw) {
    dom.reportEditor.hidden = true;
    dom.emptyState.hidden = false;
    dom.emptyState.innerHTML = `
      <h2>${escapeHtml(state.emptyTitle || "No report selected")}</h2>
      ${
        state.emptyMessage
          ? `<p>${escapeHtml(state.emptyMessage)}</p>`
          : ""
      }
    `;
    dom.saveButton.disabled = true;
    dom.exportButton.disabled = true;
    return;
  }

  dom.emptyState.hidden = true;
  dom.reportEditor.hidden = false;
  dom.saveButton.disabled = !state.isDirty;
  dom.exportButton.disabled = false;
  state.renderedAssets = [];
  state.viewerAssetIds = [];

  dom.reportEditor.innerHTML = `
    <div class="editor-header">
      <div>
        <h2>${escapeHtml(report.title)}</h2>
        <p>${escapeHtml(report.id)} - ${escapeHtml(
          report.savedDate || "",
        )}</p>
      </div>
      <span class="status-pill ${statusClassName(report.status)}">${escapeHtml(
        report.status,
      )}</span>
    </div>

    ${renderReportMetaSection(report)}
    ${renderWorkflowSection(report)}
    ${FIELD_SECTIONS.map(renderFieldSection).join("")}
    ${renderTestEquipmentSection()}
    ${renderInspectionSections()}
    ${renderObservationsSection()}
    ${renderReportAttachmentsSection()}
    ${renderBottomSubmitSection(report)}
  `;
}

function renderReportMetaSection(report) {
  return `
    <section class="section">
      <div class="section-header">
        <h3>Report Record</h3>
      </div>
      <div class="section-body field-grid">
        <div class="field">
          <label for="reportStatus">Status</label>
          <select id="reportStatus" data-report-property="status">
            ${STATUS_OPTIONS.map((option) =>
              optionHtml(option, report.status),
            ).join("")}
          </select>
        </div>
        <div class="field full">
          <label for="reportNotes">Notes</label>
          <textarea id="reportNotes" data-report-property="notes">${escapeHtml(
            report.notes || "",
          )}</textarea>
        </div>
      </div>
    </section>
  `;
}

function renderBottomSubmitSection(report) {
  const disabled = canSubmitReport(report) ? "" : " disabled";
  return `
    <section class="submit-footer">
      <button class="button success" type="button" data-action="submit-report"${disabled}>
        Submit Report
      </button>
    </section>
  `;
}

function renderWorkflowSection(report) {
  const audit = Array.isArray(report.workflowHistory)
    ? report.workflowHistory
    : [];
  const actions = WORKFLOW_ACTIONS[report.status || "Draft"] || [];
  const notifyStage = workflowNotifyStage(report.status || "Draft");
  const role = currentUserRole();
  const availableActions = actions.filter(canPerformWorkflowAction);

  return `
    <section class="section workflow-section">
      <div class="section-header">
        <h3>Submitted Report Status</h3>
        <span class="status-pill ${statusClassName(report.status)}">${escapeHtml(
          report.status || "Draft",
        )}</span>
      </div>
      <div class="section-body workflow-body">
        <div class="workflow-actions">
          ${
            availableActions.length
              ? availableActions.map(renderWorkflowAction).join("")
              : `<span class="empty-asset">No ${ROLE_LABELS[role] || "user"} actions available for this status</span>`
          }
          <button class="small-button" type="button" data-action="notify-approval" data-notify-stage="${escapeAttr(notifyStage)}">
            Email Current Approver
          </button>
        </div>
        <div class="field full workflow-note">
          <label for="workflowNote">Approval Note</label>
          <textarea id="workflowNote" data-report-property="workflowNote" placeholder="Add review notes, requested changes, or final approval comments">${escapeHtml(
            report.workflowNote || "",
          )}</textarea>
        </div>
        <div class="workflow-grid">
          ${renderWorkflowFact("Your Role", ROLE_LABELS[role] || "No approval role")}
          ${renderWorkflowFact("Engineer", ENGINEER_EMAIL)}
          ${renderWorkflowFact("Approver 1", APPROVER_1_EMAIL)}
          ${renderWorkflowFact("Approver 2", APPROVER_2_EMAIL)}
          ${renderWorkflowFact("Last Action", audit[0]?.action || "None yet")}
          ${renderWorkflowFact("Last Updated", audit[0]?.timestamp ? formatDateTime(audit[0].timestamp) : "Not started")}
        </div>
        <div class="audit-list">
          <strong>Audit History</strong>
          ${
            audit.length
              ? audit.map(renderAuditEntry).join("")
              : `<span class="empty-asset">No approval actions recorded yet</span>`
          }
        </div>
      </div>
    </section>
  `;
}

function renderWorkflowAction(action) {
  return `
    <button class="small-button success" type="button" data-action="workflow-transition" data-status="${escapeAttr(
      action.status,
    )}" data-workflow-action="${escapeAttr(action.action)}">
      ${escapeHtml(action.label)}
    </button>
  `;
}

function renderWorkflowFact(label, value) {
  return `
    <div class="workflow-fact">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(valueToString(value) || "Not set")}</strong>
    </div>
  `;
}

function renderAuditEntry(entry) {
  return `
    <div class="audit-entry">
      <span>${escapeHtml(formatDateTime(entry.timestamp))}</span>
      <strong>${escapeHtml(entry.action || "Status updated")}</strong>
      <span>${escapeHtml(entry.by || "Unknown user")}</span>
      ${
        entry.note
          ? `<p>${escapeHtml(entry.note)}</p>`
          : ""
      }
    </div>
  `;
}

function renderFieldSection(section) {
  return `
    <section class="section">
      <div class="section-header">
        <h3>${escapeHtml(section.title)}</h3>
      </div>
      <div class="section-body field-grid">
        ${section.fields.map(renderField).join("")}
      </div>
    </section>
  `;
}

function renderField(field) {
  const value = valueToString(state.currentRaw[field.key]);
  const full = field.multiline ? " full" : "";

  if (isAssetPathField(field.key)) {
    return renderAssetField(field);
  }

  if (field.type === "select") {
    return `
      <div class="field${full}">
        <label for="field-${field.key}">${escapeHtml(field.label)}</label>
        <select id="field-${field.key}" data-field="${escapeAttr(field.key)}">
          ${(field.options || SELECT_OPTIONS).map((option) =>
            optionHtml(option, value),
          ).join("")}
        </select>
      </div>
    `;
  }

  if (field.multiline) {
    return `
      <div class="field${full}">
        <label for="field-${field.key}">${escapeHtml(field.label)}</label>
        <textarea id="field-${field.key}" data-field="${escapeAttr(
          field.key,
        )}">${escapeHtml(value)}</textarea>
      </div>
    `;
  }

  return `
    <div class="field${full}">
      <label for="field-${field.key}">${escapeHtml(field.label)}</label>
      <input id="field-${field.key}" data-field="${escapeAttr(
        field.key,
      )}" value="${escapeAttr(value)}" />
    </div>
  `;
}

function isAssetPathField(key) {
  return /(?:certificate|report).*path$/i.test(key) || /(?:Path|_path)$/.test(key);
}

function renderAssetField(field) {
  const value = assetValueForField(field.key);
  return `
    <div class="field full">
      <div class="attachment-field-header">
        <span class="field-title">${escapeHtml(field.label)}</span>
        <button class="small-button" type="button" data-action="upload-asset" data-source-key="${escapeAttr(field.key)}">
          Upload
        </button>
      </div>
      ${value ? renderAssetList(value, "compact", field.key) : `<div class="empty-asset">No file attached</div>`}
    </div>
  `;
}

function assetValueForField(key) {
  const attachmentKey = key.replace(/(?:Path|_path)$/i, "_attachments");
  return state.currentRaw[attachmentKey] || state.currentRaw[key] || "";
}

function attachmentKeyForSource(sourceKey) {
  if (!sourceKey) return "";
  if (sourceKey.includes(".")) return sourceKey;
  if (sourceKey.endsWith("_attachments")) return sourceKey;
  if (/_image$/i.test(sourceKey)) return `${sourceKey}_attachments`;
  if (/(?:Path|_path)$/i.test(sourceKey)) {
    return sourceKey.replace(/(?:Path|_path)$/i, "_attachments");
  }
  return `${sourceKey}_attachments`;
}

function renderTestEquipmentSection() {
  const equipment = ensureArray("testEquipment", {
    description: "",
    serialNumber: "",
    calibration: "",
  });

  return `
    <section class="section">
      <div class="section-header">
        <h3>Record of Test Equipment</h3>
        <button class="small-button" type="button" data-action="add-equipment">
          Add Equipment
        </button>
      </div>
      <div class="section-body array-list">
        ${equipment.map(renderEquipmentRow).join("")}
      </div>
    </section>
  `;
}

function renderEquipmentRow(item, index) {
  return `
    <div class="array-row">
      <div class="array-row-header">
        <strong>Equipment ${index + 1}</strong>
        <button class="small-button danger" type="button" data-action="remove-equipment" data-index="${index}">
          Remove
        </button>
      </div>
      <div class="field-grid">
        <div class="field">
          <label>Description</label>
          <input data-array="testEquipment" data-index="${index}" data-key="description" value="${escapeAttr(
            item.description || "",
          )}" />
        </div>
        <div class="field">
          <label>Serial Number</label>
          <input data-array="testEquipment" data-index="${index}" data-key="serialNumber" value="${escapeAttr(
            item.serialNumber || "",
          )}" />
        </div>
        <div class="field">
          <label>Calibration</label>
          <input data-array="testEquipment" data-index="${index}" data-key="calibration" value="${escapeAttr(
            item.calibration || "",
          )}" />
        </div>
      </div>
    </div>
  `;
}

function renderInspectionSections() {
  const modeledKeys = new Set();
  const sections = INSPECTION_SECTIONS.map((section) => {
    section.rows.forEach((row) => modeledKeys.add(`${row.key}_selection`));
    return renderInspectionSection(section.title, section.rows);
  });

  const extraRows = Object.keys(state.currentRaw || {})
    .filter((key) => key.endsWith("_selection"))
    .filter((key) => !modeledKeys.has(key))
    .sort();

  if (extraRows.length) {
    sections.push(
      renderInspectionSection(
        "Other Inspection Checks",
        extraRows.map((key) => ({
          key: key.replace(/_selection$/, ""),
          label: labelFromKey(key.replace(/_selection$/, "")),
        })),
      ),
    );
  }

  return sections.join("");
}

function renderInspectionSection(title, rows) {
  if (!rows.length) return "";

  return `
    <section class="section">
      <div class="section-header">
        <h3>${escapeHtml(title)}</h3>
      </div>
      <div class="section-body inspection-table">
        ${rows.map(renderInspectionRow).join("")}
      </div>
    </section>
  `;
}

function renderInspectionRow(row) {
  const fallbackKey = row.fallbackKey || "";
  const selectionKey = `${row.key}_selection`;
  const commentKey = `${row.key}_comment`;
  const imageKey = `${row.key}_image`;
  const value = valueToString(
    formValue(selectionKey, fallbackKey ? `${fallbackKey}_selection` : ""),
  );
  const comment = valueToString(
    formValue(commentKey, fallbackKey ? `${fallbackKey}_comment` : ""),
  );
  const imageValue =
    formValue(`${imageKey}_attachments`, fallbackKey ? `${fallbackKey}_image_attachments` : "") ||
    formValue(imageKey, fallbackKey ? `${fallbackKey}_image` : "");

  return `
    <div class="inspection-row">
      <div>
        <div class="attachment-field-header">
          <div class="inspection-name">${escapeHtml(row.label)}</div>
          <button class="small-button" type="button" data-action="upload-asset" data-source-key="${escapeAttr(imageKey)}">
            Upload Image
          </button>
        </div>
        ${renderAssetList(imageValue, "compact", imageKey)}
      </div>
      <div class="field">
        <span class="field-title">Result</span>
        <select data-field="${escapeAttr(selectionKey)}">
          ${SELECT_OPTIONS.map((option) => optionHtml(option, value)).join("")}
        </select>
      </div>
      <div class="field">
        <span class="field-title">Comment</span>
        <textarea data-field="${escapeAttr(commentKey)}">${escapeHtml(
          comment,
        )}</textarea>
      </div>
    </div>
  `;
}

function renderObservationsSection() {
  const observations = ensureArray("additional_observations", {
    observation: "",
    imagePaths: [],
  });

  return `
    <section class="section">
      <div class="section-header">
        <h3>Additional Observations</h3>
        <button class="small-button" type="button" data-action="add-observation">
          Add Observation
        </button>
      </div>
      <div class="section-body array-list">
        ${observations.map(renderObservationRow).join("")}
      </div>
    </section>
  `;
}

function renderObservationRow(item, index) {
  const imagePaths = Array.isArray(item.imagePaths) ? item.imagePaths : [];
  const imageAttachments = Array.isArray(item.imageAttachments)
    ? item.imageAttachments
    : [];
  const assets = imageAttachments.length ? imageAttachments : imagePaths;
  return `
    <div class="array-row">
      <div class="array-row-header">
        <strong>Observation ${index + 1}</strong>
        <button class="small-button danger" type="button" data-action="remove-observation" data-index="${index}">
          Remove
        </button>
      </div>
      <div class="field full">
        <label>Observation</label>
        <textarea data-array="additional_observations" data-index="${index}" data-key="observation">${escapeHtml(
          item.observation || "",
        )}</textarea>
      </div>
      <div class="attachment-field-header">
        <span class="field-title">Images</span>
        <button class="small-button" type="button" data-action="upload-observation-asset" data-index="${index}">
          Upload Image
        </button>
      </div>
      ${
        assets.length
          ? renderAssetList(assets, "compact", `additional_observations.${index}`)
          : ""
      }
    </div>
  `;
}

function renderAssetList(value, mode = "", sourceKey = "") {
  const assets = normalizeAssetList(value).map((asset, index) =>
    enrichAsset(asset, sourceKey, index),
  );
  if (!assets.length) return "";

  return `<div class="asset-list ${mode === "compact" ? "compact" : ""}">
    ${assets.map((asset) => renderAssetTile(asset, mode)).join("")}
  </div>`;
}

function enrichAsset(asset, sourceKey, index = 0) {
  const nextSourceKey = asset.sourceKey || sourceKey;
  return {
    ...asset,
    sourceKey: nextSourceKey,
    attachmentKey: asset.attachmentKey || attachmentKeyForSource(nextSourceKey),
    assetIndex: Number.isInteger(asset.assetIndex) ? asset.assetIndex : index,
  };
}

function renderAssetTile(asset, mode = "") {
  const assetId = registerRenderedAsset(asset);
  const label = assetLabel(asset);
  const kind = assetKind(asset);
  const previewUrl = kind === "image" ? assetPreviewUrl(asset) : "";
  const status = assetStatus(asset);
  const body = previewUrl
    ? `<img src="${escapeAttr(previewUrl)}" alt="${escapeAttr(label)}" loading="lazy" onerror="this.closest('.asset-tile').classList.add('preview-failed')" />`
    : renderFilePreview(kind);
  const title = `
    <span class="asset-name">${escapeHtml(label)}</span>
    ${status && status !== "uploaded" ? `<span class="sync-note">${escapeHtml(status)}</span>` : ""}
  `;
  const content = `
    <span class="asset-preview">${body}</span>
    <span class="asset-caption">${title}</span>
    <span class="asset-actions">
      ${
        kind === "image" || kind === "pdf"
          ? `<button class="small-button" type="button" data-action="preview-asset" data-asset-id="${assetId}">View</button>`
          : ""
      }
      <button class="small-button" type="button" data-action="replace-asset" data-asset-id="${assetId}">Replace</button>
      <button class="small-button danger" type="button" data-action="remove-asset" data-asset-id="${assetId}">Remove</button>
    </span>
  `;
  const className = `asset-tile ${kind} ${mode === "compact" ? "compact" : ""}`;

  return `<article class="${className}" title="${escapeAttr(assetTitle(asset))}">${content}</article>`;
}

function registerRenderedAsset(asset) {
  const assetId = String(state.renderedAssets.length);
  state.renderedAssets.push(asset);
  if (["image", "pdf"].includes(assetKind(asset))) {
    state.viewerAssetIds.push(assetId);
  }
  return assetId;
}

function renderFilePreview(kind) {
  const label = kind === "pdf" ? "PDF" : kind === "image" ? "IMG" : "FILE";
  const iconPath =
    kind === "pdf"
      ? "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8ZM14 2v6h6M8 16h8M8 12h2"
      : "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8ZM14 2v6h6";

  return `<span class="file-preview ${kind === "pdf" ? "pdf" : ""}">
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="${iconPath}" /></svg>
    <span>${label}</span>
  </span>`;
}

function normalizeAssetList(value) {
  if (Array.isArray(value)) {
    return value.map(normalizeAsset).filter(Boolean);
  }
  const asset = normalizeAsset(value);
  return asset ? [asset] : [];
}

function normalizeAsset(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return {
      name: value.name || value.fileName || value.localPath || value.storagePath || "",
      kind: value.kind || "",
      mimeType: value.mimeType || "",
      localPath: value.localPath || "",
      storagePath: value.storagePath || "",
      driveItemId: value.driveItemId || "",
      webUrl: value.webUrl || "",
      downloadUrl: value.downloadUrl || "",
      syncStatus: value.syncStatus || "",
      syncError: value.syncError || "",
      sourceKey: value.sourceKey || value.fieldKey || "",
    };
  }

  const text = valueToString(value);
  return text ? { name: text, localPath: text } : null;
}

function assetLabel(value) {
  const text =
    typeof value === "object"
      ? value.name || value.localPath || value.storagePath || value.webUrl || ""
      : valueToString(value);
  const parts = text.split(/[\\/]/).filter(Boolean);
  return parts[parts.length - 1] || text;
}

function assetKind(asset) {
  const declaredKind = (asset.kind || "").toLowerCase();
  if (declaredKind) return declaredKind;
  const mimeType = (asset.mimeType || "").toLowerCase();
  if (mimeType === "application/pdf") return "pdf";
  if (mimeType.startsWith("image/")) return "image";
  const label = assetLabel(asset).toLowerCase();
  if (label.endsWith(".pdf")) return "pdf";
  if (/\.(jpe?g|png|webp|gif|heic|heif)$/i.test(label)) return "image";
  return "document";
}

function assetOpenUrl(asset) {
  return asset.downloadUrl || assetLocalBrowserUrl(asset) || asset.webUrl || "";
}

function assetPreviewUrl(asset) {
  return asset.downloadUrl || assetLocalBrowserUrl(asset) || "";
}

function assetBrowserUrl(asset) {
  return asset.webUrl || assetOpenUrl(asset);
}

function assetCanLoadFromGraph(asset) {
  return Boolean(asset.driveItemId || normalizeAttachmentStoragePath(asset));
}

function normalizeAttachmentStoragePath(asset) {
  let storagePath = valueToString(asset?.storagePath).trim();
  if (!storagePath) return "";

  storagePath = storagePath
    .replace(/^\/?drive\/root:\//i, "")
    .replace(/^\/?drives\/[^/]+\/root:\//i, "")
    .replace(/:$/, "");

  const fileName = assetLabel(asset);
  const pathFileName = assetLabel(storagePath);
  if (fileName && pathFileName && pathFileName !== fileName && !/\.[a-z0-9]{2,5}$/i.test(pathFileName)) {
    storagePath = `${storagePath.replace(/\/$/, "")}/${fileName}`;
  }

  return storagePath;
}

function assetLocalBrowserUrl(asset) {
  const localPath = asset.localPath || "";
  if (!localPath || /^[a-z][a-z0-9+.-]*:/i.test(localPath)) return "";
  if (localPath.startsWith("/")) return "";
  return encodeURI(localPath);
}

function assetStatus(asset) {
  return asset.syncStatus || "";
}

function assetTitle(asset) {
  return asset.storagePath || asset.localPath || asset.webUrl || assetLabel(asset);
}

function renderReportAttachmentsSection() {
  const attachments = collectReportAssets();
  const images = attachments.filter((asset) => assetKind(asset) === "image");
  const pdfs = attachments.filter((asset) => assetKind(asset) === "pdf");
  const documents = attachments.filter((asset) => !["image", "pdf"].includes(assetKind(asset)));

  return `
    <section class="section">
      <div class="section-header">
        <h3>Report Attachments</h3>
        <button class="small-button" type="button" data-action="upload-asset" data-source-key="report">
          Upload Attachment
        </button>
      </div>
      <div class="section-body report-attachments">
        ${attachments.length ? "" : `<div class="empty-asset">No report attachments yet</div>`}
        ${images.length ? renderImageGallery(images) : ""}
        ${renderPdfPanel(pdfs)}
        ${documents.length ? `<div class="attachment-strip">${documents.map((asset) => renderAssetTile(asset)).join("")}</div>` : ""}
      </div>
    </section>
  `;
}

function collectReportAssets() {
  const assets = [];
  const seen = new Set();
  const addAssets = (sourceKey, value) => {
    normalizeAssetList(value).forEach((asset, index) => {
      const enriched = { ...asset, sourceKey: asset.sourceKey || sourceKey };
      enriched.attachmentKey = attachmentKeyForSource(enriched.sourceKey);
      enriched.assetIndex = index;
      const key = [
        enriched.driveItemId,
        enriched.downloadUrl,
        enriched.webUrl,
        enriched.storagePath,
        enriched.localPath,
        assetLabel(enriched),
        sourceKey,
      ]
        .filter(Boolean)
        .join("|");
      if (seen.has(key)) return;
      seen.add(key);
      assets.push(enriched);
    });
  };

  Object.entries(state.currentRaw || {}).forEach(([key, value]) => {
    if (key.endsWith("_attachments") && Array.isArray(value)) {
      addAssets(key.replace(/_attachments$/, ""), value);
      return;
    }

    if (/_image$/i.test(key) || isAssetPathField(key)) {
      addAssets(key, value);
    }
  });

  ensureArray("additional_observations").forEach((item, index) => {
    if (!item || typeof item !== "object") return;
    addAssets(
      `additional observation ${index + 1}`,
      Array.isArray(item.imageAttachments) && item.imageAttachments.length
        ? item.imageAttachments
        : item.imagePaths,
    );
  });

  return assets;
}

function renderImageGallery(images) {
  if (!images.length) {
    return `<div class="empty-asset">No image attachments</div>`;
  }

  return `
    <div class="attachment-view">
      <div class="attachment-view-header">
        <strong>Images</strong>
        <div class="gallery-controls">
          <button class="small-button" type="button" data-action="gallery-prev">Previous</button>
          <button class="small-button" type="button" data-action="gallery-next">Next</button>
        </div>
      </div>
      <div class="image-gallery" data-gallery>
        ${images.map((asset) => renderGalleryImage(asset)).join("")}
      </div>
    </div>
  `;
}

function renderGalleryImage(asset) {
  const assetId = registerRenderedAsset(asset);
  const previewUrl = assetPreviewUrl(asset);
  const body = previewUrl
    ? `<img src="${escapeAttr(previewUrl)}" alt="${escapeAttr(assetLabel(asset))}" loading="lazy" onerror="this.closest('.gallery-slide').classList.add('preview-failed')" />`
    : renderFilePreview("image");

  return `
    <figure class="gallery-slide">
      <button class="gallery-open" type="button" data-action="preview-asset" data-asset-id="${assetId}">
        ${body}
      </button>
      <figcaption>
        <strong>${escapeHtml(formatAttachmentSource(asset.sourceKey || "Image"))}</strong>
        <span>${escapeHtml(assetLabel(asset))}</span>
        <span class="asset-actions">
          <button class="small-button" type="button" data-action="replace-asset" data-asset-id="${assetId}">Replace</button>
          <button class="small-button danger" type="button" data-action="remove-asset" data-asset-id="${assetId}">Remove</button>
        </span>
      </figcaption>
    </figure>
  `;
}

function renderPdfPanel(pdfs) {
  if (!pdfs.length) return "";

  return `
    <div class="attachment-view">
      <div class="attachment-view-header">
        <strong>PDFs</strong>
      </div>
      <div class="pdf-list">
        ${pdfs.map(renderPdfCard).join("")}
      </div>
    </div>
  `;
}

function renderPdfCard(asset) {
  const assetId = registerRenderedAsset(asset);
  return `
    <article class="pdf-card">
      <div class="pdf-preview" data-action="preview-asset" data-asset-id="${assetId}" role="button" tabindex="0">
        ${renderFilePreview("pdf")}
      </div>
      <div class="pdf-meta">
        <strong>${escapeHtml(assetLabel(asset))}</strong>
        <span>${escapeHtml(formatAttachmentSource(asset.sourceKey || "PDF"))}</span>
        <span class="asset-actions">
          <button class="small-button" type="button" data-action="preview-asset" data-asset-id="${assetId}">View</button>
          <button class="small-button" type="button" data-action="replace-asset" data-asset-id="${assetId}">Replace</button>
          <button class="small-button danger" type="button" data-action="remove-asset" data-asset-id="${assetId}">Remove</button>
        </span>
      </div>
    </article>
  `;
}

function formatAttachmentSource(key) {
  return valueToString(key)
    .replace(/_attachments$/i, "")
    .replace(/_image$/i, "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function handleEditorInput(event) {
  const target = event.target;
  const report = selectedReport();
  if (!report || !state.currentRaw || !target) return;

  if (target.dataset.field) {
    state.currentRaw[target.dataset.field] = target.value;
    markDirty();
    return;
  }

  if (target.dataset.reportProperty) {
    report[target.dataset.reportProperty] = target.value;
    if (target.dataset.reportProperty === "status") {
      addWorkflowHistory(report, `Status manually set to ${target.value}`);
    }
    markDirty();
    renderReportList();
    return;
  }

  if (target.dataset.array && target.dataset.index && target.dataset.key) {
    const array = ensureArray(target.dataset.array);
    const index = Number(target.dataset.index);
    if (!array[index]) array[index] = {};
    array[index][target.dataset.key] = target.value;
    markDirty();
  }
}

async function updateWorkflowStatus(nextStatus, actionLabel) {
  const report = selectedReport();
  if (!report || !nextStatus) return;

  const workflowAction = workflowActionFor(report.status || "Draft", nextStatus, actionLabel);
  if (!workflowAction || !canPerformWorkflowAction(workflowAction)) {
    showToast("Your account cannot perform that approval step.", "error");
    return;
  }

  setBusy(true, "Updating");

  try {
    const note = valueToString(report.workflowNote).trim();
    report.status = nextStatus;
    addWorkflowHistory(report, actionLabel || `Status changed to ${nextStatus}`, note);
    report.workflowNote = "";
    const rawData = applyCurrentReportEditsToState();
    const token = await getAccessToken();
    await uploadReportsFile(token);
    if (workflowAction?.notify) {
      await sendApprovalEmail(workflowAction.notify, token);
    }
    state.isDirty = false;
    state.currentRaw = cloneJson(rawData);
    updateDirtyState();
    renderReportList();
    renderEditor();
    showToast("Workflow updated and saved.", "success");
  } catch (error) {
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

async function submitCurrentReport() {
  const report = selectedReport();
  if (!report || !state.currentRaw) {
    showToast("Select a report first.", "warn");
    return;
  }

  if (!canSubmitReport(report)) {
    showToast(`This report is already ${report.status || "submitted"}.`, "warn");
    return;
  }

  setBusy(true, "Submitting");

  try {
    const note = valueToString(report.workflowNote).trim();
    report.status = "Submitted";
    addWorkflowHistory(report, "Submitted for review", note);
    report.workflowNote = "";
    const rawData = applyCurrentReportEditsToState();
    const token = await getAccessToken();
    await uploadReportsFile(token);
    await sendApprovalEmail("review", token);
    state.isDirty = false;
    state.currentRaw = cloneJson(rawData);
    updateDirtyState();
    renderReportList();
    renderEditor();
    showToast("Report submitted and Gary has been notified.", "success");
  } catch (error) {
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

function canSubmitSelectedReport() {
  return Boolean(state.account && canSubmitReport(selectedReport()));
}

function canSubmitReport(report) {
  if (!report) return false;
  if (currentUserRole() !== "engineer") return false;
  return ["Draft", "In Progress", "Completed", "Changes Requested"].includes(
    report.status || "Draft",
  );
}

function workflowActionFor(currentStatus, nextStatus, actionLabel) {
  const actions = WORKFLOW_ACTIONS[currentStatus] || [];
  return actions.find(
    (action) => action.status === nextStatus && action.action === actionLabel,
  );
}

function canPerformWorkflowAction(action) {
  if (!action) return false;
  const allowedRoles = Array.isArray(action.roles) ? action.roles : [];
  if (!allowedRoles.length) return true;
  return allowedRoles.includes(currentUserRole());
}

function addWorkflowHistory(report, action, note = "") {
  if (!Array.isArray(report.workflowHistory)) {
    report.workflowHistory = [];
  }

  report.workflowHistory.unshift({
    action,
    status: report.status || "Draft",
    by: state.account?.username || state.account?.name || "Web portal",
    timestamp: new Date().toISOString(),
    note,
  });
}

async function sendApprovalEmail(stage, token) {
  const report = selectedReport();
  const recipient = approvalEmailForStage(stage);
  const subject =
    stage === "final"
      ? `Final approval required: ${report?.title || "Commissioning report"}`
      : stage === "changes"
        ? `Changes requested: ${report?.title || "Commissioning report"}`
        : `Report review required: ${report?.title || "Commissioning report"}`;
  const body = approvalEmailHtml(report, stage);
  const accessToken = token || (await getAccessToken());
  const response = await fetch(`${GRAPH_ROOT}/me/sendMail`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        subject,
        body: {
          contentType: "HTML",
          content: body,
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipient,
            },
          },
        ],
      },
      saveToSentItems: true,
    }),
  });

  if (!response.ok) {
    throw new Error(await graphError(response, "Could not send approval email"));
  }

  return recipient;
}

function approvalEmailHtml(report, stage) {
  const heading =
    stage === "final"
      ? "Final approval required"
      : stage === "changes"
        ? "Changes requested"
        : "Report review required";

  const latestNote = latestWorkflowNote(report);

  return `
    <p>${escapeHtml(heading)}</p>
    <table>
      <tr><td><strong>Report</strong></td><td>${escapeHtml(report?.title || "")}</td></tr>
      <tr><td><strong>ID</strong></td><td>${escapeHtml(report?.id || "")}</td></tr>
      <tr><td><strong>Status</strong></td><td>${escapeHtml(report?.status || "Draft")}</td></tr>
      <tr><td><strong>Site</strong></td><td>${escapeHtml(report?.location || "")}</td></tr>
      ${
        latestNote
          ? `<tr><td><strong>Latest note</strong></td><td>${escapeHtml(latestNote)}</td></tr>`
          : ""
      }
    </table>
    <p><a href="${escapeAttr(reportDeepLink(report))}">Open this report in the Voltempo Field Application portal</a></p>
  `;
}

function reportDeepLink(report) {
  const url = new URL(authRedirectUri());
  if (report?.id) {
    url.searchParams.set("report", report.id);
  }
  return url.toString();
}

function reportIdFromUrl() {
  const params = new URLSearchParams(window.location.search);
  return valueToString(params.get("report")).trim();
}

function updateReportUrl(reportId) {
  if (!window.history?.replaceState || !reportId) return;
  const url = new URL(window.location.href);
  url.searchParams.set("report", reportId);
  window.history.replaceState({}, "", url.toString());
}

function workflowNotifyStage(status) {
  if (status === "Reviewed" || status === "Final Approval") return "final";
  if (status === "Changes Requested") return "changes";
  return "review";
}

function approvalEmailForStage(stage) {
  if (stage === "final") return APPROVER_2_EMAIL;
  if (stage === "changes") return ENGINEER_EMAIL;
  return APPROVER_1_EMAIL;
}

function latestWorkflowNote(report) {
  const audit = Array.isArray(report?.workflowHistory)
    ? report.workflowHistory
    : [];
  const entry = audit.find((item) => valueToString(item.note).trim());
  return valueToString(entry?.note).trim();
}

function handleEditorClick(event) {
  const target = event.target.closest("[data-action]");
  if (!target) return;

  const action = target.dataset.action;
  const index = Number(target.dataset.index);

  if (action === "add-equipment") {
    ensureArray("testEquipment").push({
      description: "",
      serialNumber: "",
      calibration: "",
    });
    markDirty();
    renderEditor();
  }

  if (action === "remove-equipment" && Number.isInteger(index)) {
    ensureArray("testEquipment").splice(index, 1);
    markDirty();
    renderEditor();
  }

  if (action === "status-filter") {
    state.statusFilter = target.dataset.status || "All";
    renderReportList();
  }

  if (action === "submit-report") {
    submitCurrentReport();
  }

  if (action === "workflow-transition") {
    updateWorkflowStatus(target.dataset.status || "", target.dataset.workflowAction || "");
  }

  if (action === "notify-approval") {
    sendApprovalEmail(target.dataset.notifyStage || "review")
      .then((recipient) => showToast(`Email sent to ${recipient}.`, "success"))
      .catch((error) => showToast(error.message || String(error), "error"));
  }

  if (action === "add-observation") {
    ensureArray("additional_observations").push({
      observation: "",
      imagePaths: [],
    });
    markDirty();
    renderEditor();
  }

  if (action === "remove-observation" && Number.isInteger(index)) {
    ensureArray("additional_observations").splice(index, 1);
    markDirty();
    renderEditor();
  }

  if (action === "upload-asset") {
    openAttachmentPicker({
      sourceKey: target.dataset.sourceKey,
    });
  }

  if (action === "upload-observation-asset" && Number.isInteger(index)) {
    openAttachmentPicker({
      sourceKey: `additional_observations.${index}`,
    });
  }

  if (action === "replace-asset") {
    const asset = assetById(target.dataset.assetId);
    if (!asset) return;
    openAttachmentPicker({
      sourceKey: asset.sourceKey,
      replaceAssetId: target.dataset.assetId,
    });
  }

  if (action === "remove-asset") {
    removeAssetById(target.dataset.assetId);
  }

  if (action === "preview-asset") {
    openAssetViewer(target.dataset.assetId);
  }

  if (action === "gallery-prev" || action === "gallery-next") {
    const gallery = target.closest(".attachment-view")?.querySelector("[data-gallery]");
    if (!gallery) return;
    const direction = action === "gallery-next" ? 1 : -1;
    gallery.scrollBy({
      left: direction * Math.max(280, gallery.clientWidth * 0.78),
      behavior: "smooth",
    });
  }
}

function openAttachmentPicker(target) {
  if (!state.account) {
    showToast("Sign in before uploading attachments.", "warn");
    return;
  }
  if (!selectedReport() || !state.currentRaw) {
    showToast("Select a report first.", "warn");
    return;
  }

  state.pendingAttachmentTarget = target;
  dom.attachmentFileInput.value = "";
  dom.attachmentFileInput.click();
}

async function handleAttachmentFileSelected(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  const target = state.pendingAttachmentTarget;
  state.pendingAttachmentTarget = null;
  if (!file || !target) return;

  setBusy(true, "Uploading");

  try {
    const attachment = await uploadAttachmentFile(file, target.sourceKey);
    addOrReplaceAttachment(target, attachment);
    markDirty();
    const rawData = applyCurrentReportEditsToState();
    const token = await getAccessToken();
    await uploadReportsFile(token);
    state.isDirty = false;
    state.currentRaw = cloneJson(rawData);
    updateDirtyState();
    renderReportList();
    renderEditor();
    showToast("Attachment uploaded and saved.", "success");
  } catch (error) {
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

async function uploadAttachmentFile(file, sourceKey) {
  const token = await getAccessToken();
  await ensureSharedRoot(token);
  const report = selectedReport();
  const rawData = state.currentRaw || {};
  const reportType =
    valueToString(rawData.reportType).toLowerCase() === "standalone"
      ? "standalone"
      : "hypercharger";
  const siteFolder = safeDriveSegment(rawData.siteName || "Unknown Site");
  const reportFolder = safeDriveSegment(
    rawData.editingReportId || report?.id || rawData.reportNumber || "report",
  );
  const category = file.type === "application/pdf" ? "Certificates" : "Images";
  const fileName = `${Date.now()}_${safeDriveSegment(file.name || "attachment")}`;
  const folderPath = `${APP_ROOT_FOLDER}/reports/${reportType}/${siteFolder}/${reportFolder}/${category}`;
  const storagePath = `${folderPath}/${fileName}`;

  await ensureDriveFolderPath(token, folderPath);

  const response = await fetch(graphDriveFileContentUrl(storagePath), {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": file.type || "application/octet-stream",
    },
    body: await file.arrayBuffer(),
  });

  if (!response.ok) {
    throw new Error(await graphError(response, "Could not upload attachment"));
  }

  const driveItem = await response.json();
  return {
    id: `att_${Date.now()}`,
    fieldKey: sourceKey,
    kind: attachmentKindForFile(file),
    category,
    name: file.name,
    mimeType: file.type || mimeTypeForFileName(file.name),
    sizeBytes: file.size,
    localPath: file.name,
    storagePath,
    driveItemId: driveItem.id || "",
    webUrl: driveItem.webUrl || "",
    downloadUrl: driveItem["@microsoft.graph.downloadUrl"] || "",
    syncStatus: "uploaded",
    createdAt: new Date().toISOString(),
  };
}

function addOrReplaceAttachment(target, attachment) {
  const replacement = target.replaceAssetId ? assetById(target.replaceAssetId) : null;
  const sourceKey = target.sourceKey || replacement?.sourceKey || attachment.fieldKey;
  attachment.fieldKey = sourceKey;

  if (sourceKey?.startsWith("additional_observations.")) {
    const index = Number(sourceKey.split(".")[1]);
    const observations = ensureArray("additional_observations");
    if (!observations[index]) observations[index] = { observation: "", imagePaths: [] };
    if (!Array.isArray(observations[index].imageAttachments)) {
      observations[index].imageAttachments = [];
    }
    if (replacement && Number.isInteger(replacement.assetIndex)) {
      observations[index].imageAttachments[replacement.assetIndex] = attachment;
    } else {
      observations[index].imageAttachments.push(attachment);
    }
    return;
  }

  const attachmentKey = attachmentKeyForSource(sourceKey);
  if (!Array.isArray(state.currentRaw[attachmentKey])) {
    state.currentRaw[attachmentKey] = [];
  }

  if (replacement && Number.isInteger(replacement.assetIndex)) {
    state.currentRaw[attachmentKey][replacement.assetIndex] = attachment;
  } else {
    state.currentRaw[attachmentKey].push(attachment);
  }

  if (/_image$/i.test(sourceKey)) {
    state.currentRaw[sourceKey] = attachment.localPath;
  }

  if (isAssetPathField(sourceKey)) {
    state.currentRaw[sourceKey] = attachment.localPath;
    const nameKey = sourceKey.replace(/Path$/i, "").replace(/_path$/i, "_name");
    if (nameKey !== sourceKey) state.currentRaw[nameKey] = attachment.name;
  }
}

async function removeAssetById(assetId) {
  const asset = assetById(assetId);
  if (!asset) return;

  if (asset.sourceKey?.startsWith("additional_observations.")) {
    const index = Number(asset.sourceKey.split(".")[1]);
    const observation = ensureArray("additional_observations")[index];
    if (!observation || typeof observation !== "object") return;
    const list = Array.isArray(observation.imageAttachments)
      ? observation.imageAttachments
      : observation.imagePaths;
    if (Array.isArray(list)) list.splice(asset.assetIndex, 1);
  } else if (asset.attachmentKey && Array.isArray(state.currentRaw[asset.attachmentKey])) {
    state.currentRaw[asset.attachmentKey].splice(asset.assetIndex, 1);
  } else if (asset.sourceKey) {
    state.currentRaw[asset.sourceKey] = "";
  }

  if (asset.sourceKey && !asset.sourceKey.includes(".")) {
    const remaining = asset.attachmentKey && Array.isArray(state.currentRaw[asset.attachmentKey])
      ? state.currentRaw[asset.attachmentKey]
      : [];
    if (!remaining.length && (/_image$/i.test(asset.sourceKey) || isAssetPathField(asset.sourceKey))) {
      state.currentRaw[asset.sourceKey] = "";
    }
  }

  markDirty();
  setBusy(true, "Saving");

  try {
    const rawData = applyCurrentReportEditsToState();
    const token = await getAccessToken();
    await uploadReportsFile(token);
    state.isDirty = false;
    state.currentRaw = cloneJson(rawData);
    updateDirtyState();
    renderReportList();
    renderEditor();
    showToast("Attachment removed and saved.", "success");
  } catch (error) {
    renderEditor();
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

function assetById(assetId) {
  const index = Number(assetId);
  return Number.isInteger(index) ? state.renderedAssets[index] : null;
}

function handleAssetViewerClick(event) {
  const target = event.target.closest("[data-action]");
  if (!target) return;

  const action = target.dataset.action;
  if (action === "viewer-close") closeAssetViewer();
  if (action === "viewer-prev") moveAssetViewer(-1);
  if (action === "viewer-next") moveAssetViewer(1);
  if (action === "pdf-prev") movePdfPage(-1);
  if (action === "pdf-next") movePdfPage(1);
  if (action === "pdf-zoom-out") zoomPdf(-0.2);
  if (action === "pdf-zoom-in") zoomPdf(0.2);
}

function handleGlobalKeydown(event) {
  if (dom.assetViewer.hidden) return;
  if (event.key === "Escape") closeAssetViewer();
  if (event.key === "ArrowLeft") moveAssetViewer(-1);
  if (event.key === "ArrowRight") moveAssetViewer(1);
}

function movePdfPage(direction) {
  const viewer = state.pdfViewer;
  if (!viewer.document) return;
  viewer.page = Math.min(
    Math.max(viewer.page + direction, 1),
    viewer.document.numPages,
  );
  renderPdfPage();
}

function zoomPdf(delta) {
  const viewer = state.pdfViewer;
  if (!viewer.document) return;
  viewer.scale = Math.min(Math.max(viewer.scale + delta, 0.5), 2.4);
  renderPdfPage();
}

function openAssetViewer(assetId) {
  const asset = assetById(assetId);
  if (!asset) return;

  const currentIndex = state.viewerAssetIds.indexOf(String(assetId));
  state.viewerIndex = currentIndex >= 0 ? currentIndex : 0;
  dom.assetViewer.hidden = false;
  renderAssetViewer(asset);
}

function closeAssetViewer() {
  dom.assetViewer.hidden = true;
  resetPdfViewer();
  dom.assetViewerStage.innerHTML = "";
}

function moveAssetViewer(direction) {
  if (dom.assetViewer.hidden || !state.viewerAssetIds.length) return;
  state.viewerIndex =
    (state.viewerIndex + direction + state.viewerAssetIds.length) %
    state.viewerAssetIds.length;
  renderAssetViewer(assetById(state.viewerAssetIds[state.viewerIndex]));
}

function renderAssetViewer(asset) {
  if (!asset) return;

  resetPdfViewer();
  const kind = assetKind(asset);
  const url = assetOpenUrl(asset);
  const browserUrl = assetBrowserUrl(asset);
  dom.assetViewerTitle.textContent = assetLabel(asset);
  dom.assetViewerMeta.textContent = formatAttachmentSource(asset.sourceKey || "");
  dom.assetViewerDownload.hidden = !url;
  dom.assetViewerDownload.href = url || "#";

  if (kind === "image" && url) {
    dom.assetViewerStage.innerHTML = `<img src="${escapeAttr(url)}" alt="${escapeAttr(
      assetLabel(asset),
    )}" />`;
    return;
  }

  if (kind === "pdf" && (url || assetCanLoadFromGraph(asset))) {
    renderPdfViewer(asset);
    return;
  }

  dom.assetViewerStage.innerHTML = `
    <div class="pdf-fallback">
      ${renderFilePreview(kind)}
      <p>${escapeHtml(assetUnavailableMessage(asset))}</p>
      ${
        browserUrl
          ? `<a class="small-button" href="${escapeAttr(browserUrl)}" target="_blank" rel="noopener">Open File</a>`
          : ""
      }
    </div>
  `;
}

function assetUnavailableMessage(asset) {
  if (asset.syncStatus && asset.syncStatus !== "uploaded") {
    return asset.syncError
      ? `Attachment has not uploaded to OneDrive yet: ${asset.syncError}`
      : "Attachment has not uploaded to OneDrive yet. Sync the mobile app while signed in, then refresh the portal.";
  }

  return "No browser-viewable file URL is available for this attachment yet.";
}

async function renderPdfViewer(asset) {
  const session = Date.now();
  state.pdfViewer = {
    document: null,
    page: 1,
    scale: 1,
    session,
  };

  dom.assetViewerStage.innerHTML = `
    <div class="pdf-loading">
      ${renderFilePreview("pdf")}
      <p>Loading PDF preview...</p>
    </div>
  `;

  try {
    await loadPdfLibrary();
    const pdfSource = await pdfDocumentSource(asset);
    const task = window.pdfjsLib.getDocument(pdfSource);
    const pdfDocument = await task.promise;

    if (state.pdfViewer.session !== session) return;

    state.pdfViewer.document = pdfDocument;
    state.pdfViewer.page = 1;
    state.pdfViewer.scale = 1;
    await renderPdfPage();
  } catch (error) {
    if (state.pdfViewer.session !== session) return;
    renderPdfFallback(asset, error);
  }
}

async function loadPdfLibrary() {
  if (window.pdfjsLib?.getDocument) return;

  for (const url of PDFJS_SCRIPT_URLS) {
    try {
      await loadScript(url);
      if (window.pdfjsLib?.getDocument) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl(url);
        return;
      }
    } catch (_error) {
      // Try the next CDN.
    }
  }

  throw new Error("PDF viewer library could not be loaded.");
}

function pdfWorkerUrl(scriptUrl) {
  return scriptUrl.replace(/pdf(\.min)?\.js$/, "pdf.worker.min.js");
}

async function pdfDocumentSource(asset) {
  if (asset.driveItemId && state.account) {
    const token = await getAccessToken();
    await ensureSharedRoot(token);
    const response = await fetch(
      graphDriveItemContentByIdUrl(asset.driveItemId),
      {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      },
    );

    if (!response.ok) {
      throw new Error(await graphError(response, "Could not load PDF"));
    }

    return {
      data: new Uint8Array(await response.arrayBuffer()),
    };
  }

  const storagePath = normalizeAttachmentStoragePath(asset);
  if (storagePath && state.account) {
    const token = await getAccessToken();
    await ensureSharedRoot(token);
    const response = await fetch(graphDriveFileContentUrl(storagePath), {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      throw new Error(await graphError(response, "Could not load PDF"));
    }

    return {
      data: new Uint8Array(await response.arrayBuffer()),
    };
  }

  const response = await fetch(assetOpenUrl(asset));

  if (!response.ok) {
    throw new Error("Could not load PDF.");
  }

  return {
    data: new Uint8Array(await response.arrayBuffer()),
  };
}

async function renderPdfPage() {
  const viewer = state.pdfViewer;
  const pdfDocument = viewer.document;
  if (!pdfDocument) return;

  const pageNumber = Math.min(Math.max(viewer.page, 1), pdfDocument.numPages);
  viewer.page = pageNumber;
  const session = viewer.session;
  const page = await pdfDocument.getPage(pageNumber);
  if (state.pdfViewer.session !== session) return;

  const viewport = page.getViewport({ scale: viewer.scale * 1.35 });
  dom.assetViewerStage.innerHTML = `
    <div class="pdf-toolbar">
      <button class="small-button" type="button" data-action="pdf-prev" ${pageNumber <= 1 ? "disabled" : ""}>Previous</button>
      <span>Page ${pageNumber} of ${pdfDocument.numPages}</span>
      <button class="small-button" type="button" data-action="pdf-next" ${pageNumber >= pdfDocument.numPages ? "disabled" : ""}>Next</button>
      <button class="small-button" type="button" data-action="pdf-zoom-out">-</button>
      <button class="small-button" type="button" data-action="pdf-zoom-in">+</button>
    </div>
    <div class="pdf-canvas-wrap">
      <canvas id="pdfViewerCanvas" width="${Math.ceil(viewport.width)}" height="${Math.ceil(viewport.height)}"></canvas>
    </div>
  `;

  const canvas = window.document.getElementById("pdfViewerCanvas");
  const context = canvas.getContext("2d");
  await page.render({ canvasContext: context, viewport }).promise;
}

function renderPdfFallback(asset, error) {
  const browserUrl = assetBrowserUrl(asset);
  dom.assetViewerStage.innerHTML = `
    <div class="pdf-fallback">
      ${renderFilePreview("pdf")}
      <p>PDF preview could not be loaded${error?.message ? `: ${escapeHtml(error.message)}` : "."}</p>
      ${
        browserUrl
          ? `<a class="small-button" href="${escapeAttr(browserUrl)}" target="_blank" rel="noopener">Open PDF</a>`
          : ""
      }
    </div>
  `;
}

function resetPdfViewer() {
  state.pdfViewer.session += 1;
  state.pdfViewer.document = null;
  state.pdfViewer.page = 1;
  state.pdfViewer.scale = 1;
}

function safeDriveSegment(value) {
  return valueToString(value)
    .trim()
    .replace(/[\\/:*?"<>|#%{}~&]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 90) || "attachment";
}

function mimeTypeForFileName(name) {
  const value = valueToString(name).toLowerCase();
  if (value.endsWith(".pdf")) return "application/pdf";
  if (value.endsWith(".png")) return "image/png";
  if (value.endsWith(".webp")) return "image/webp";
  if (value.endsWith(".gif")) return "image/gif";
  if (value.endsWith(".heic")) return "image/heic";
  if (value.endsWith(".heif")) return "image/heif";
  return "image/jpeg";
}

function attachmentKindForFile(file) {
  const mimeType = file.type || mimeTypeForFileName(file.name);
  if (mimeType === "application/pdf") return "pdf";
  if (mimeType.startsWith("image/")) return "image";
  return "document";
}

async function saveCurrentReport() {
  if (!selectedReport() || !state.currentRaw) return;

  setBusy(true, "Saving");

  try {
    const rawData = applyCurrentReportEditsToState();

    const token = await getAccessToken();
    await uploadReportsFile(token);

    state.isDirty = false;
    state.currentRaw = cloneJson(rawData);
    updateDirtyState();
    renderReportList();
    renderEditor();
    showToast("Report saved to OneDrive.", "success");
  } catch (error) {
    showToast(error.message || String(error), "error");
  } finally {
    setBusy(false);
  }
}

function exportCurrentReportPdf() {
  const report = selectedReport();
  if (!report || !state.currentRaw) {
    showToast("Select a report first.", "warn");
    return;
  }

  const rawData = {
    ...cloneJson(state.currentRaw),
    editingReportId: report.id,
  };
  const printableReport = {
    ...report,
    title: buildReportTitle(rawData),
    location: rawData.siteName || rawData.siteId || report.location,
    rawData,
  };
  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    showToast("Allow pop-ups to export the report PDF.", "warn");
    return;
  }

  printWindow.document.open();
  printWindow.document.write(buildPrintableReportHtml(printableReport));
  printWindow.document.close();
  printWindow.focus();
  window.setTimeout(() => {
    printWindow.print();
  }, 250);
}

function buildPrintableReportHtml(report) {
  const rawData = report.rawData || {};
  const generated = new Date().toLocaleString();

  return `<!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(report.title)}</title>
        <style>
          body {
            margin: 0;
            color: #172033;
            font-family: Arial, Helvetica, sans-serif;
            font-size: 11px;
            line-height: 1.35;
          }
          .page {
            max-width: 840px;
            margin: 0 auto;
            padding: 28px;
          }
          .header {
            display: flex;
            justify-content: space-between;
            gap: 24px;
            padding-bottom: 16px;
            border-bottom: 2px solid #1e3a5f;
            margin-bottom: 18px;
          }
          h1 {
            margin: 0 0 6px;
            color: #1e3a5f;
            font-size: 22px;
          }
          h2 {
            margin: 18px 0 0;
            padding: 8px 10px;
            color: #ffffff;
            background: #1e3a5f;
            font-size: 13px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            page-break-inside: avoid;
          }
          th, td {
            vertical-align: top;
            border: 1px solid #d9dee8;
            padding: 7px 8px;
          }
          th {
            width: 34%;
            text-align: left;
            background: #f3f6f9;
          }
          .inspection th:nth-child(1) { width: 44%; }
          .inspection th:nth-child(2) { width: 16%; }
          .muted { color: #667085; }
          .footer {
            margin-top: 24px;
            padding-top: 10px;
            border-top: 1px solid #d9dee8;
            color: #667085;
            font-size: 10px;
          }
          @media print {
            @page { size: A4; margin: 12mm; }
            .page { max-width: none; padding: 0; }
            h2 { break-after: avoid; }
          }
        </style>
      </head>
      <body>
        <main class="page">
          <header class="header">
            <div>
              <h1>Voltempo - Commissioning Report</h1>
              <div class="muted">Generated: ${escapeHtml(generated)}</div>
              <div><strong>Site:</strong> ${escapeHtml(rawData.siteName || "Not specified")} (ID: ${escapeHtml(rawData.siteId || "Not specified")})</div>
            </div>
            <div>
              <strong>${escapeHtml(report.status || "Draft")}</strong><br />
              <span class="muted">${escapeHtml(report.id || "")}</span>
            </div>
          </header>
          ${printFieldSection("Site Information", FIELD_SECTIONS[0].fields, rawData)}
          ${printTestEquipmentSection(rawData)}
          ${printFieldSection("Systems and Components to be commissioned", FIELD_SECTIONS[1].fields, rawData)}
          ${printFieldSection("Incoming Supply Characteristics", FIELD_SECTIONS[2].fields, rawData)}
          ${printFieldSection("Transformer Details", FIELD_SECTIONS[3].fields, rawData)}
          ${INSPECTION_SECTIONS.map((section) => printInspectionSection(section, rawData)).join("")}
          ${printFieldSection("Pre-Checks", FIELD_SECTIONS[4].fields, rawData)}
          ${printObservationsSection(rawData)}
          ${printFieldSection("Sign Off", FIELD_SECTIONS[5].fields, rawData)}
          <footer class="footer">
            This report was generated by the Voltempo desktop portal. Images are referenced from report data and will be included once cloud asset storage is added.
          </footer>
        </main>
      </body>
    </html>`;
}

function printFieldSection(title, fields, rawData) {
  return `<h2>${escapeHtml(title)}</h2>
    <table>
      <tbody>
        ${fields
          .map(
            (field) =>
              `<tr><th>${escapeHtml(field.label)}</th><td>${escapeHtml(
                printableValue(rawData[field.key]),
              )}</td></tr>`,
          )
          .join("")}
      </tbody>
    </table>`;
}

function printTestEquipmentSection(rawData) {
  const equipment = Array.isArray(rawData.testEquipment)
    ? rawData.testEquipment
    : [];

  if (!equipment.length) {
    return `<h2>Record of Test Equipment</h2><table><tbody>
      <tr><th>Description</th><td>${escapeHtml(printableValue(rawData.description))}</td></tr>
      <tr><th>Serial Number</th><td>${escapeHtml(printableValue(rawData.serialNumber))}</td></tr>
      <tr><th>Calibration</th><td>${escapeHtml(printableValue(rawData.calibration))}</td></tr>
    </tbody></table>`;
  }

  return `<h2>Record of Test Equipment</h2><table><tbody>
    ${equipment
      .map(
        (item, index) => `
          <tr><th colspan="2">Equipment ${index + 1}</th></tr>
          <tr><th>Description</th><td>${escapeHtml(printableValue(item.description))}</td></tr>
          <tr><th>Serial Number</th><td>${escapeHtml(printableValue(item.serialNumber))}</td></tr>
          <tr><th>Calibration</th><td>${escapeHtml(printableValue(item.calibration))}</td></tr>
        `,
      )
      .join("")}
    </tbody></table>`;
}

function printInspectionSection(section, rawData) {
  return `<h2>${escapeHtml(section.title)}</h2>
    <table class="inspection">
      <thead>
        <tr><th>Check</th><th>Result</th><th>Comment / Evidence</th></tr>
      </thead>
      <tbody>
        ${section.rows
          .map((row) => {
            const result = formValueForData(
              rawData,
              `${row.key}_selection`,
              row.fallbackKey ? `${row.fallbackKey}_selection` : "",
            );
            const comment = formValueForData(
              rawData,
              `${row.key}_comment`,
              row.fallbackKey ? `${row.fallbackKey}_comment` : "",
            );
            const image = formValueForData(
              rawData,
              `${row.key}_image`,
              row.fallbackKey ? `${row.fallbackKey}_image` : "",
            );
            const evidence = [comment, assetText(image)]
              .map(printableValue)
              .filter((value) => value !== "Not specified")
              .join(" | ");
            return `<tr><td>${escapeHtml(row.label)}</td><td>${escapeHtml(
              printableValue(result),
            )}</td><td>${escapeHtml(evidence || "Not specified")}</td></tr>`;
          })
          .join("")}
      </tbody>
    </table>`;
}

function printObservationsSection(rawData) {
  const observations = Array.isArray(rawData.additional_observations)
    ? rawData.additional_observations
    : [];

  if (!observations.length) {
    return `<h2>Additional Observations</h2><table><tbody><tr><th>Observations</th><td>Not specified</td></tr></tbody></table>`;
  }

  return `<h2>Additional Observations</h2><table><tbody>
    ${observations
      .map((item, index) => {
        const text = typeof item === "object" ? item.observation : item;
        const images = typeof item === "object" ? item.imagePaths : "";
        const value = [text, assetText(images)]
          .map(printableValue)
          .filter((entry) => entry !== "Not specified")
          .join(" | ");
        return `<tr><th>Observation ${index + 1}</th><td>${escapeHtml(
          value || "Not specified",
        )}</td></tr>`;
      })
      .join("")}
    </tbody></table>`;
}

function applyCurrentReportEditsToState() {
  const report = selectedReport();
  if (!report || !state.currentRaw) return null;

  const now = new Date();
  const rawData = {
    ...state.currentRaw,
    editingReportId: report.id,
  };

  report.rawData = rawData;
  report.title = buildReportTitle(rawData);
  report.location = rawData.siteName || rawData.siteId || "Unknown site";
  report.siteId = rawData.siteId || "";
  report.savedDate = formatDateTime(now);
  report.updatedAt = now.toISOString();
  report.size = reportSize(rawData);
  report.isSaved = true;
  report.workflowHistory = Array.isArray(report.workflowHistory)
    ? report.workflowHistory
    : [];
  report.tags = Array.isArray(report.tags) && report.tags.length
    ? report.tags
    : ["saved"];

  state.reports = state.reports.map((item) =>
    item.id === report.id ? report : item,
  );
  state.reports.sort(sortReports);
  state.reportsFile.savedReports = state.reports;
  return rawData;
}

function selectedReport() {
  return state.reports.find((report) => report.id === state.selectedId) || null;
}

function markDirty() {
  state.isDirty = true;
  updateDirtyState();
}

function updateDirtyState() {
  dom.dirtyState.textContent = state.isDirty ? "Unsaved" : "";
  dom.uploadButton.disabled = !state.account || !state.reportsFile;
  dom.submitReportButton.disabled = !canSubmitSelectedReport();
  dom.saveButton.disabled = !state.account || !state.selectedId || !state.isDirty;
  dom.exportButton.disabled = !state.selectedId;
}

function setBusy(isBusy, label = "") {
  state.isBusy = isBusy;
  setControlLabel(
    dom.cloudState,
    isBusy ? label : state.account ? "Signed in" : "Offline",
  );
  dom.cloudState.className = `state-chip ${isBusy ? "busy" : state.account ? "online" : ""}`.trim();
  dom.syncButton.disabled = isBusy || !state.account;
  dom.importButton.disabled = isBusy;
  dom.uploadButton.disabled = isBusy || !state.account || !state.reportsFile;
  dom.submitReportButton.disabled = isBusy || !canSubmitSelectedReport();
  dom.signInButton.disabled = isBusy && !state.account;
  dom.saveButton.disabled =
    isBusy || !state.account || !state.selectedId || !state.isDirty;
  dom.exportButton.disabled = isBusy || !state.selectedId;
}

function ensureArray(key, emptyItem) {
  if (!Array.isArray(state.currentRaw[key])) {
    state.currentRaw[key] = emptyItem ? [cloneJson(emptyItem)] : [];
  }
  return state.currentRaw[key];
}

async function graphError(response, fallback) {
  try {
    const body = await response.json();
    return body?.error?.message || `${fallback} (${response.status})`;
  } catch (error) {
    return `${fallback} (${response.status})`;
  }
}

function isRequiredAccount(account) {
  return AUTHORISED_ACCOUNTS.includes(accountEmail(account));
}

function accountEmail(account = state.account) {
  return valueToString(
    account?.username || account?.idTokenClaims?.preferred_username,
  ).toLowerCase();
}

function currentUserRole() {
  const email = accountEmail();
  if (email === ENGINEER_EMAIL) return "engineer";
  if (email === APPROVER_1_EMAIL) return "approver1";
  if (email === APPROVER_2_EMAIL) return "approver2";
  return "";
}

function buildReportTitle(rawData) {
  const siteName = valueToString(rawData.siteName).trim();
  const siteId = valueToString(rawData.siteId).trim();
  if (siteName) return `Commissioning Report - ${siteName}`;
  if (siteId) return `Commissioning Report - ${siteId}`;
  return "Commissioning Report - Untitled Site";
}

function reportSize(rawData) {
  const bytes = new TextEncoder().encode(JSON.stringify(rawData || {})).length;
  return `${Math.min(bytes / 1024, 99999).toFixed(1)} KB`;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDateTime(date) {
  if (!(date instanceof Date)) {
    const parsed = new Date(date);
    if (!Number.isNaN(parsed.getTime())) {
      date = parsed;
    } else {
      return valueToString(date);
    }
  }
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${formatDate(date)} - ${hours}:${minutes}`;
}

function labelFromKey(key) {
  return key
    .replace(/_/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase())
    .replace(/\bLv\b/g, "LV")
    .replace(/\bHv\b/g, "HV")
    .replace(/\bMcb\b/g, "MCB")
    .replace(/\bPsu\b/g, "PSU")
    .replace(/\bVsecc\b/g, "VSECC");
}

function formValue(primaryKey, fallbackKey = "") {
  return formValueForData(state.currentRaw || {}, primaryKey, fallbackKey);
}

function formValueForData(data, primaryKey, fallbackKey = "") {
  const primaryValue = data[primaryKey];
  if (primaryValue !== undefined && primaryValue !== null && primaryValue !== "") {
    return primaryValue;
  }
  return fallbackKey ? data[fallbackKey] : primaryValue;
}

function assetText(value) {
  if (Array.isArray(value)) return value.filter(Boolean).join(", ");
  if (value && typeof value === "object") return JSON.stringify(value);
  return valueToString(value);
}

function printableValue(value) {
  const text = assetText(value).trim();
  return text || "Not specified";
}

function statusClassName(status) {
  const value = valueToString(status).toLowerCase();
  if (value.includes("submitted")) return "submitted";
  if (value.includes("review")) return "review";
  if (value.includes("change")) return "changes";
  if (value.includes("final")) return "final";
  if (value.includes("archive")) return "archived";
  if (value.includes("complete") || value.includes("approved")) return "completed";
  if (value.includes("progress")) return "progress";
  if (value.includes("draft")) return "draft";
  return "";
}

function optionHtml(option, selectedValue) {
  const selected = option === selectedValue ? " selected" : "";
  return `<option value="${escapeAttr(option)}"${selected}>${escapeHtml(
    option || "Not set",
  )}</option>`;
}

function valueToString(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  return String(value);
}

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function escapeHtml(value) {
  return valueToString(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/"/g, "&quot;");
}

function showToast(message, type = "") {
  dom.toast.textContent = message;
  dom.toast.className = `toast ${type}`.trim();
  dom.toast.hidden = false;
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    dom.toast.hidden = true;
  }, 4200);
}
