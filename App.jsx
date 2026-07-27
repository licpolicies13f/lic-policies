import {
  memo,
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  AlertTriangle,
  ArrowDownUp,
  BadgeCheck,
  BellRing,
  CalendarDays,
  Check,
  CheckCircle2,
  CircleDollarSign,
  Clock3,
  Download,
  FileCheck2,
  FileSpreadsheet,
  LayoutDashboard,
  MessageCircle,
  Pencil,
  Plus,
  Printer,
  Search,
  ShieldCheck,
  TableProperties,
  Trash2,
  Upload,
  WalletCards,
  X,
} from "lucide-react";
import {
  createPolicyWorkbookFile,
  exportPolicyWorkbook,
  mergePolicyCollections,
  readPolicyWorkbook,
} from "./excel.js";

const EMPTY_FORM = {
  serialNumber: "",
  clientName: "",
  dob: "",
  policyNumber: "",
  agentCode: "",
  commencementDate: "",
  plan: "",
  sumAssured: "",
  premium: "",
  mode: "Monthly",
  status: "Unpaid",
};

const NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "upload", label: "Upload", icon: Upload },
  { id: "dues", label: "Dues", icon: BellRing },
  { id: "paid", label: "Paid", icon: BadgeCheck },
  { id: "policies", label: "Policies", icon: TableProperties },
];

const MODES = ["Monthly", "Quarterly", "Half-Yearly", "Yearly"];
const MODE_INTERVALS = {
  Monthly: 1,
  Quarterly: 3,
  "Half-Yearly": 6,
  Yearly: 12,
};
const LEGACY_POLICY_STORAGE_KEY = "lic-policy-tracker-policies-v1";
const ACTIVE_AGENCY_STORAGE_KEY = "lic-policy-tracker-active-agency-v1";
const AGENCIES = [
  { id: "agency-1", label: "Old Agency", code: "0035713F" },
  { id: "agency-2", label: "New Agency", code: "0212313F" },
];
const FIXED_AGENCY_CODES = Object.freeze(
  Object.fromEntries(AGENCIES.map((agency) => [agency.id, agency.code])),
);
const PAGE_SIZE = 40;
const policySearchCache = new WeakMap();
const lastPersistedPolicySnapshots = new Map();

const COLUMNS = [
  { key: "serialNumber", label: "SNo" },
  { key: "clientName", label: "Name" },
  { key: "dob", label: "DOB" },
  { key: "policyNumber", label: "Policy" },
  { key: "agentCode", label: "Agcode" },
  { key: "commencementDate", label: "Com.Date" },
  { key: "plan", label: "P/T/PP" },
  { key: "sumAssured", label: "SumAssd" },
  { key: "mode", label: "Mode" },
  { key: "premium", label: "Premium" },
  { key: "status", label: "Status" },
];

const currencyFormatter = new Intl.NumberFormat("en-IN", {
  maximumFractionDigits: 2,
});

const MODE_CODES = {
  Monthly: "Mly",
  Quarterly: "Qly",
  "Half-Yearly": "Hly",
  Yearly: "Yly",
};

const newId = () =>
  globalThis.crypto?.randomUUID?.() ?? `policy-${Date.now()}-${Math.random()}`;

const formatDateValue = (value) => {
  if (!value) return "";
  const text = String(value).trim();
  const isoDate = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDate) return `${isoDate[3]}/${isoDate[2]}/${isoDate[1]}`;

  const slashDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashDate) {
    return `${slashDate[1].padStart(2, "0")}/${slashDate[2].padStart(2, "0")}/${slashDate[3]}`;
  }

  return text;
};

const formatDate = (value) => formatDateValue(value) || "—";

const toStoredDate = (value) => {
  const text = String(value ?? "").trim();
  if (!text) return "";

  const isoDate = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const slashDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  const year = Number(isoDate?.[1] ?? slashDate?.[3]);
  const month = Number(isoDate?.[2] ?? slashDate?.[2]);
  const day = Number(isoDate?.[3] ?? slashDate?.[1]);

  if (!year || !month || !day) return null;
  const parsed = new Date(Date.UTC(year, month - 1, day));
  if (
    parsed.getUTCFullYear() !== year ||
    parsed.getUTCMonth() !== month - 1 ||
    parsed.getUTCDate() !== day
  ) {
    return null;
  }

  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
};

const getCurrentMonthKey = () => {
  const today = new Date();
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
};

const getMonthLabel = (monthKey) => {
  const [year, month] = monthKey.split("-").map(Number);
  if (!year || !month) return monthKey;
  return new Date(year, month - 1, 1).toLocaleDateString("en-IN", {
    month: "long",
    year: "numeric",
  });
};

const getPolicyDueDate = (policy, monthKey) => {
  const interval = MODE_INTERVALS[policy.mode];
  const storedDate = toStoredDate(policy.commencementDate);
  if (!interval || !storedDate) return null;

  const [startYear, startMonth, startDay] = storedDate.split("-").map(Number);
  const [targetYear, targetMonth] = monthKey.split("-").map(Number);
  if (!targetYear || !targetMonth) return null;

  const monthsSinceStart =
    (targetYear - startYear) * 12 + (targetMonth - startMonth);
  if (monthsSinceStart < 0 || monthsSinceStart % interval !== 0) return null;

  const lastDay = new Date(targetYear, targetMonth, 0).getDate();
  const dueDay = Math.min(startDay, lastDay);
  return `${targetYear}-${String(targetMonth).padStart(2, "0")}-${String(dueDay).padStart(2, "0")}`;
};

const normalizePaidPeriods = (periods) =>
  Array.isArray(periods)
    ? Array.from(
        new Set(periods.filter((period) => /^\d{4}-\d{2}$/.test(String(period)))),
      )
    : [];

const compactPolicy = (policy) => [
  policy.id,
  policy.serialNumber,
  policy.clientName,
  policy.dob,
  policy.policyNumber,
  policy.agentCode,
  policy.commencementDate,
  policy.plan,
  policy.sumAssured,
  policy.premium,
  policy.mode,
  policy.status,
  normalizePaidPeriods(policy.paidPeriods),
];

const expandPolicy = (record) => ({
  id: record[0] || newId(),
  serialNumber: record[1] || "",
  clientName: record[2] || "",
  dob: record[3] || "",
  policyNumber: record[4] || "",
  agentCode: record[5] || "",
  commencementDate: record[6] || "",
  plan: record[7] || "",
  sumAssured: Number(record[8]) || 0,
  premium: Number(record[9]) || 0,
  mode: record[10] || "Monthly",
  status: record[11] || "Unpaid",
  paidPeriods: normalizePaidPeriods(record[12]),
});

const getAgencyStorageKey = (agencyId) =>
  `${LEGACY_POLICY_STORAGE_KEY}:${agencyId}`;

const isAgencyId = (value) =>
  AGENCIES.some((agency) => agency.id === value);

const normalizeAgencyCode = (value) =>
  String(value ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");

const getAgencyIdForCode = (agentCode, agencyCodes) => {
  const normalizedCode = normalizeAgencyCode(agentCode);
  if (!normalizedCode) return null;
  return (
    AGENCIES.find(
      (agency) =>
        normalizeAgencyCode(agencyCodes[agency.id]) === normalizedCode,
    )?.id ?? null
  );
};

const loadActiveAgency = () => {
  try {
    const stored = window.localStorage.getItem(ACTIVE_AGENCY_STORAGE_KEY);
    return isAgencyId(stored) ? stored : AGENCIES[0].id;
  } catch {
    return AGENCIES[0].id;
  }
};

const persistPolicies = (policies, agencyId) => {
  try {
    const storageKey = getAgencyStorageKey(agencyId);
    const snapshot = JSON.stringify({
      version: 2,
      records: policies.map(compactPolicy),
    });
    if (snapshot === lastPersistedPolicySnapshots.get(storageKey)) return true;
    window.localStorage.setItem(storageKey, snapshot);
    lastPersistedPolicySnapshots.set(storageKey, snapshot);
    return true;
  } catch {
    // The tracker continues in memory if device storage is unavailable or full.
    return false;
  }
};

const loadStoredPolicies = (agencyId) => {
  try {
    const storageKey = getAgencyStorageKey(agencyId);
    let stored = window.localStorage.getItem(storageKey);
    if (!stored && agencyId === AGENCIES[0].id) {
      stored = window.localStorage.getItem(LEGACY_POLICY_STORAGE_KEY);
      if (stored) window.localStorage.setItem(storageKey, stored);
    }
    const parsed = stored ? JSON.parse(stored) : null;
    if (parsed?.version === 2) {
      lastPersistedPolicySnapshots.set(storageKey, stored);
    }
    const records =
      parsed?.version === 2 && Array.isArray(parsed.records)
        ? parsed.records.map(expandPolicy)
        : Array.isArray(parsed)
          ? parsed
          : [];
    const currentMonth = getCurrentMonthKey();
    return records.map((policy) => ({
      ...policy,
      paidPeriods: Array.isArray(policy.paidPeriods)
        ? normalizePaidPeriods(policy.paidPeriods)
        : policy.status === "Paid"
          ? [currentMonth]
          : [],
    }));
  } catch {
    return [];
  }
};

const renumberPolicies = (policies) =>
  policies.map((policy, index) => ({
    ...policy,
    serialNumber: String(index + 1),
  }));

const redistributeAgencyPolicies = (collections, agencyCodes) => {
  const redistributed = Object.fromEntries(
    AGENCIES.map((agency) => [agency.id, []]),
  );
  const seenPolicyNumbers = Object.fromEntries(
    AGENCIES.map((agency) => [agency.id, new Set()]),
  );
  let moved = 0;
  let duplicates = 0;

  AGENCIES.forEach((agency) => {
    (collections[agency.id] ?? []).forEach((policy) => {
      const matchingAgency = getAgencyIdForCode(policy.agentCode, agencyCodes);
      const targetAgency = matchingAgency ?? agency.id;
      const policyNumber = String(policy.policyNumber ?? "").trim().toLowerCase();

      if (
        policyNumber &&
        seenPolicyNumbers[targetAgency].has(policyNumber)
      ) {
        duplicates += 1;
      }

      if (policyNumber) seenPolicyNumbers[targetAgency].add(policyNumber);
      redistributed[targetAgency].push(policy);
      if (targetAgency !== agency.id) moved += 1;
    });
  });

  AGENCIES.forEach((agency) => {
    redistributed[agency.id] = renumberPolicies(redistributed[agency.id]);
  });

  return { collections: redistributed, moved, duplicates };
};

const getClientName = (value) =>
  String(value || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find(Boolean) || "Client";

const getWhatsAppNumber = (value) => {
  const lines = String(value || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines.slice(1)) {
    let digits = line.replace(/\D/g, "");
    if (digits.length < 10 || digits.length > 13) continue;
    if (digits.length === 10) digits = `91${digits}`;
    if (digits.length === 11 && digits.startsWith("0")) digits = `91${digits.slice(1)}`;
    return digits;
  }

  return "";
};

const getPolicySearchText = (policy) => {
  const cached = policySearchCache.get(policy);
  if (cached) return cached;
  const searchText = [
    policy.clientName,
    policy.policyNumber,
    policy.agentCode,
  ]
    .map((value) => String(value || "").toLowerCase())
    .join("\n");
  policySearchCache.set(policy, searchText);
  return searchText;
};

function usePaginatedItems(items, resetKey, renderAll = false) {
  const [page, setPage] = useState(1);
  const totalPages = Math.max(1, Math.ceil(items.length / PAGE_SIZE));

  useEffect(() => {
    setPage(1);
  }, [resetKey]);

  useEffect(() => {
    setPage((current) => Math.min(current, totalPages));
  }, [totalPages]);

  const pageItems = useMemo(() => {
    if (renderAll) return items;
    const start = (page - 1) * PAGE_SIZE;
    return items.slice(start, start + PAGE_SIZE);
  }, [items, page, renderAll]);

  return {
    page,
    pageItems,
    setPage,
    totalPages,
  };
}

function Pagination({ page, totalPages, totalItems, onPageChange }) {
  if (totalItems <= PAGE_SIZE) return null;
  return (
    <div
      className="no-print pagination-bar flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 px-4 py-4 sm:px-5"
      aria-label="Pagination"
    >
      <p className="text-xs font-semibold text-slate-500">
        Page <span className="font-black text-slate-800">{page}</span> of{" "}
        <span className="font-black text-slate-800">{totalPages}</span>
        <span className="ml-2 text-slate-400">• {PAGE_SIZE} per page</span>
      </p>
      <div className="flex items-center gap-2">
        <button
          type="button"
          disabled={page <= 1}
          onClick={() => onPageChange(page - 1)}
          className="min-h-9 rounded-lg border border-slate-200 bg-white px-3 text-xs font-black text-slate-600 transition hover:border-[#0756a0]/30 hover:text-[#0756a0] disabled:cursor-not-allowed disabled:opacity-35"
        >
          Previous
        </button>
        <button
          type="button"
          disabled={page >= totalPages}
          onClick={() => onPageChange(page + 1)}
          className="min-h-9 rounded-lg bg-[#0756a0] px-3 text-xs font-black text-white transition hover:bg-[#064985] disabled:cursor-not-allowed disabled:opacity-35"
        >
          Next
        </button>
      </div>
    </div>
  );
}

const sendPolicyOnWhatsApp = (policy) => {
  const phone = getWhatsAppNumber(policy.clientName);
  if (!phone) return;
  const name = getClientName(policy.clientName);
  const policyNumber = policy.policyNumber || "your LIC policy";
  const message = `Hello ${name}, this message is regarding LIC Policy ${policyNumber}.`;
  window.open(
    `https://wa.me/${phone}?text=${encodeURIComponent(message)}`,
    "_blank",
    "noopener,noreferrer",
  );
};

function Brand({ agencyLabel }) {
  return (
    <div className="brand-mark flex items-center gap-3">
      <div className="brand-symbol grid size-11 place-items-center rounded-xl bg-[#0756a0] text-sm font-black tracking-tight text-white shadow-sm">
        LIC
      </div>
      <div>
        <p className="text-[11px] font-bold uppercase tracking-[0.18em] text-[#0756a0]">
          {agencyLabel} workspace
        </p>
        <p className="text-base font-bold text-slate-900">Policy Tracker</p>
      </div>
    </div>
  );
}

function AgencySwitcher({
  activeAgency,
  currentPolicyCount,
  onAgencyChange,
}) {
  return (
    <div className="floating-window agency-switcher mt-5 flex flex-col gap-4 rounded-3xl p-4 sm:flex-row sm:items-center sm:justify-between sm:p-5">
      <div>
        <p className="text-xs font-black uppercase tracking-[0.16em] text-[#0756a0]">
          Agency workspace
        </p>
        <p className="mt-1 text-sm text-slate-500">
          Agency codes are fixed. Excel policies are sorted automatically using
          the Agcode column.
        </p>
      </div>
      <div
        role="group"
        aria-label="Select agency"
        className="grid gap-2 sm:min-w-[420px] sm:grid-cols-2"
      >
        {AGENCIES.map((agency) => {
          const active = activeAgency === agency.id;
          return (
            <AgencyOption
              key={agency.id}
              agency={agency}
              active={active}
              currentPolicyCount={currentPolicyCount}
              onAgencyChange={onAgencyChange}
            />
          );
        })}
      </div>
    </div>
  );
}

function AgencyOption({
  agency,
  active,
  currentPolicyCount,
  onAgencyChange,
}) {
  return (
    <div
      className={`agency-option rounded-2xl border p-2 transition ${
        active
          ? "border-[#0756a0] bg-[#0756a0] text-white shadow-lg shadow-[#0756a0]/15"
          : "border-slate-200 bg-white text-slate-700"
      }`}
    >
      <button
        type="button"
        aria-pressed={active}
        onClick={() => onAgencyChange(agency.id)}
        className="min-h-12 w-full rounded-xl px-2 text-left"
      >
        <span className="flex items-center justify-between gap-2">
          <span className="text-sm font-black">{agency.label}</span>
          {active && <CheckCircle2 size={16} />}
        </span>
        <span
          className={`mt-1 block text-[10px] font-bold ${
            active ? "text-blue-100" : "text-slate-400"
          }`}
        >
          {active
            ? `${currentPolicyCount} ${currentPolicyCount === 1 ? "policy" : "policies"}`
            : "Select workspace"}
        </span>
      </button>
      <div
        aria-label={`${agency.label} fixed agency code ${agency.code}`}
        className={`mx-2 mt-1 flex items-center justify-between rounded-lg border px-3 py-2 ${
          active
            ? "border-white/20 bg-white/10"
            : "border-slate-200 bg-slate-50"
        }`}
      >
        <span
          className={`text-[9px] font-black uppercase tracking-[0.12em] ${
            active ? "text-blue-100" : "text-slate-400"
          }`}
        >
          Fixed Agcode
        </span>
        <span className="font-mono text-xs font-black tracking-wide">
          {agency.code}
        </span>
      </div>
    </div>
  );
}

function PrimaryButton({ children, className = "", ...props }) {
  return (
    <button
      className={`inline-flex min-h-11 items-center justify-center gap-2 rounded-xl bg-[#0756a0] px-4 text-sm font-bold text-white shadow-sm transition hover:bg-[#064985] focus-visible:outline-none focus-visible:ring-4 focus-visible:ring-[#0756a0]/20 disabled:cursor-not-allowed disabled:opacity-45 ${className}`}
      {...props}
    >
      {children}
    </button>
  );
}

function PageHeading({ eyebrow, title, description, actions }) {
  return (
    <div className="page-heading flex flex-col gap-5 sm:flex-row sm:items-end sm:justify-between">
      <div>
        <p className="mb-2 text-xs font-bold uppercase tracking-[0.2em] text-[#0756a0]">
          {eyebrow}
        </p>
        <h1 className="text-3xl font-black tracking-[-0.035em] text-slate-950 sm:text-4xl">
          {title}
        </h1>
        <p className="mt-2 max-w-2xl text-sm leading-6 text-slate-500">{description}</p>
      </div>
      {actions && <div className="flex shrink-0 flex-wrap gap-3">{actions}</div>}
    </div>
  );
}

function StatCard({ label, value, icon: Icon, tone = "blue" }) {
  return (
    <div
      data-tone={tone}
      className="stat-card group relative min-h-36 overflow-hidden rounded-3xl p-5 transition hover:-translate-y-1"
    >
      <div className="relative flex h-full flex-col justify-between gap-5">
        <div className="stat-icon grid size-10 place-items-center rounded-xl">
          <Icon size={19} strokeWidth={2.2} />
        </div>
        <div>
          <p className="text-3xl font-black tracking-tight text-slate-950">{value}</p>
          <p className="mt-1 text-sm font-semibold text-slate-500">
            {label}
          </p>
        </div>
      </div>
    </div>
  );
}

const Dashboard = memo(function Dashboard({
  policies,
  activeAgency,
  onAgencyChange,
  onAdd,
  onUpload,
}) {
  const activeAgencyLabel =
    AGENCIES.find((agency) => agency.id === activeAgency)?.label ??
    AGENCIES[0].label;
  const counts = useMemo(
    () =>
      policies.reduce(
        (totals, policy) => {
          totals.total += 1;
          if (policy.status === "Paid") totals.paid += 1;
          else totals.unpaid += 1;
          if (policy.mode === "Monthly") totals.monthly += 1;
          if (policy.mode === "Quarterly") totals.quarterly += 1;
          if (policy.mode === "Half-Yearly") totals.halfYearly += 1;
          if (policy.mode === "Yearly") totals.yearly += 1;
          return totals;
        },
        {
          total: 0,
          paid: 0,
          unpaid: 0,
          monthly: 0,
          quarterly: 0,
          halfYearly: 0,
          yearly: 0,
        },
      ),
    [policies],
  );

  const cards = [
    { label: "Total Policies", value: counts.total, icon: WalletCards, tone: "blue" },
    { label: "Paid Policies", value: counts.paid, icon: CheckCircle2, tone: "yellow" },
    { label: "Unpaid Policies", value: counts.unpaid, icon: Clock3, tone: "white" },
    { label: "Monthly Policies", value: counts.monthly, icon: CalendarDays, tone: "white" },
    {
      label: "Quarterly Policies",
      value: counts.quarterly,
      icon: CircleDollarSign,
      tone: "white",
    },
    {
      label: "Half-Yearly Policies",
      value: counts.halfYearly,
      icon: FileCheck2,
      tone: "white",
    },
    { label: "Yearly Policies", value: counts.yearly, icon: ShieldCheck, tone: "white" },
  ];

  return (
    <section>
      <PageHeading
        eyebrow="Overview"
        title={`${activeAgencyLabel} dashboard`}
        description={`A clear snapshot of every LIC policy saved in ${activeAgencyLabel}.`}
        actions={
          <>
            <button
              onClick={onUpload}
              className="inline-flex min-h-11 items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 text-sm font-bold text-slate-700 shadow-sm transition hover:border-[#0756a0]/30 hover:text-[#0756a0]"
            >
              <Upload size={17} />
              Upload Excel
            </button>
            <PrimaryButton onClick={onAdd}>
              <Plus size={18} />
              Add Policy
            </PrimaryButton>
          </>
        }
      />

      <AgencySwitcher
        activeAgency={activeAgency}
        currentPolicyCount={policies.length}
        onAgencyChange={onAgencyChange}
      />

      <div className="dashboard-grid mt-7 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        {cards.map((card, index) => (
          <StatCard key={card.label} {...card} className={index === 0 ? "lg:col-span-1" : ""} />
        ))}
      </div>

      <div className="floating-window privacy-window mt-7 flex flex-col gap-4 rounded-3xl p-5 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-start gap-3">
          <div className="mt-0.5 grid size-9 shrink-0 place-items-center rounded-xl bg-white text-[#0756a0] shadow-sm">
            <ShieldCheck size={19} />
          </div>
          <div>
            <p className="font-bold text-slate-900">Private by design</p>
            <p className="mt-1 text-sm leading-6 text-slate-600">
              Your workbook stays on this device. Nothing is uploaded to a server.
            </p>
          </div>
        </div>
        <span className="w-fit rounded-full bg-white px-3 py-1.5 text-xs font-bold text-[#0756a0] shadow-sm">
          Offline • Saved locally
        </span>
      </div>
    </section>
  );
});

const UploadExcel = memo(function UploadExcel({
  policies,
  onImported,
  onGoToPolicies,
}) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState(null);

  const loadFiles = async (fileList) => {
    const files = Array.from(fileList ?? []);
    if (!files.length) return;

    const invalidFiles = files.filter(
      (file) => !file.name.toLowerCase().endsWith(".xlsx"),
    );
    if (invalidFiles.length) {
      setMessage({ type: "error", text: "Please choose only Excel .xlsx files." });
      return;
    }

    setBusy(true);
    setMessage(null);
    try {
      const importedGroups = [];
      for (const file of files) {
        importedGroups.push(await readPolicyWorkbook(file));
        await new Promise((resolve) => window.setTimeout(resolve, 0));
      }
      const imported = importedGroups.flat();
      const sourceLabel =
        files.length === 1 ? files[0].name : `${files.length} Excel files`;
      const importResult = onImported(imported, sourceLabel);
      const added = importResult?.added ?? 0;
      const skipped = importResult?.skipped ?? 0;
      const destinationText = (importResult?.destinations ?? [])
        .filter((destination) => destination.added > 0)
        .map(
          (destination) =>
            `${destination.added} to ${destination.label}`,
        )
        .join(", ");
      const unmatchedText = importResult?.unmatched
        ? ` ${importResult.unmatched} ${importResult.unmatched === 1 ? "record had" : "records had"} an unknown Agcode and ${importResult.unmatched === 1 ? "was" : "were"} kept in the selected agency.`
        : "";
      setMessage({
        type: "success",
        text: `${added} ${added === 1 ? "policy" : "policies"} added from ${sourceLabel}${destinationText ? ` (${destinationText})` : ""}.${skipped ? ` ${skipped} duplicate ${skipped === 1 ? "was" : "records were"} skipped.` : ""}${unmatchedText}`,
      });
    } catch (error) {
      setMessage({
        type: "error",
        text: error instanceof Error ? error.message : "The workbook could not be read.",
      });
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  };

  return (
    <section>
        <PageHeading
          eyebrow="Local workbook"
          title="Upload your Excel file"
          description="Choose one or multiple LIC .xlsx workbooks. Policies are added to Old Agency or New Agency automatically when their Agcode matches a saved agency code."
      />

      <div className="mx-auto mt-8 max-w-3xl">
        <input
          ref={inputRef}
          type="file"
          multiple
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          className="sr-only"
          onChange={(event) => loadFiles(event.target.files)}
        />
        <button
          type="button"
          aria-label="Choose or drop one or multiple Excel files"
          onClick={() => inputRef.current?.click()}
          onDragEnter={(event) => {
            event.preventDefault();
            setDragging(true);
          }}
          onDragOver={(event) => event.preventDefault()}
          onDragLeave={(event) => {
            event.preventDefault();
            if (event.currentTarget === event.target) setDragging(false);
          }}
          onDrop={(event) => {
            event.preventDefault();
            setDragging(false);
            loadFiles(event.dataTransfer.files);
          }}
          disabled={busy}
          className={`upload-window group flex min-h-80 w-full flex-col items-center justify-center rounded-3xl border-2 border-dashed px-6 text-center transition ${
            dragging
              ? "border-[#0756a0] bg-[#eaf4ff]"
              : "border-slate-300 bg-white hover:border-[#0756a0]/60 hover:bg-[#f8fbff]"
          }`}
        >
          <span className="grid size-16 place-items-center rounded-2xl bg-[#0756a0] text-white shadow-lg shadow-[#0756a0]/20 transition group-hover:-translate-y-1">
            <FileSpreadsheet size={30} />
          </span>
          <span className="mt-6 text-xl font-black text-slate-900">
            {busy ? "Reading your workbooks…" : "Drop your Excel files here"}
          </span>
          <span className="mt-2 text-sm text-slate-500">
            or click to select one or multiple files
          </span>
          <span className="mt-5 rounded-full bg-[#ffcf2f] px-4 py-2 text-xs font-black text-[#172033]">
            .XLSX FILES ONLY
          </span>
        </button>

        {message && (
          <div
            role="status"
            className={`mt-4 flex items-start gap-3 rounded-2xl border p-4 text-sm ${
              message.type === "success"
                ? "border-emerald-200 bg-emerald-50 text-emerald-800"
                : "border-red-200 bg-red-50 text-red-700"
            }`}
          >
            {message.type === "success" ? (
              <CheckCircle2 className="mt-0.5 shrink-0" size={18} />
            ) : (
              <X className="mt-0.5 shrink-0" size={18} />
            )}
            <div className="flex-1">
              <p className="font-semibold">{message.text}</p>
              {message.type === "success" && (
                <button
                  onClick={onGoToPolicies}
                  className="mt-2 font-bold text-[#0756a0] underline decoration-[#0756a0]/30 underline-offset-4"
                >
                  View policies
                </button>
              )}
            </div>
          </div>
        )}

        <div className="mt-5 grid gap-3 sm:grid-cols-3">
          {[ 
            ["1", "Use the first worksheet"],
            ["2", "Sort records by Agcode"],
            ["3", "Keep existing data"],
          ].map(([number, text]) => (
            <div
              key={number}
              className="floating-window instruction-window flex items-center gap-3 rounded-2xl px-4 py-3"
            >
              <span className="grid size-7 shrink-0 place-items-center rounded-lg bg-[#eaf4ff] text-xs font-black text-[#0756a0]">
                {number}
              </span>
              <span className="text-xs font-semibold text-slate-600">{text}</span>
            </div>
          ))}
        </div>

        {policies.length > 0 && (
          <p className="mt-4 text-center text-xs leading-5 text-slate-500">
            You currently have {policies.length} records saved locally. Uploading another workbook
            will add new policies and keep these records.
          </p>
        )}
      </div>
    </section>
  );
});

function FilterButton({ active, children, ...props }) {
  return (
    <button
      className={`min-h-9 rounded-lg px-3 text-xs font-bold transition ${
        active
          ? "bg-[#0756a0] text-white shadow-sm"
          : "bg-white text-slate-600 ring-1 ring-inset ring-slate-200 hover:text-[#0756a0]"
      }`}
      {...props}
    >
      {children}
    </button>
  );
}

function StatusBadge({ status }) {
  const paid = status === "Paid";
  return (
    <span
      className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-bold ${
        paid ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"
      }`}
    >
      <span className={`size-1.5 rounded-full ${paid ? "bg-emerald-500" : "bg-amber-500"}`} />
      {status}
    </span>
  );
}

function ClientDetails({ value }) {
  const lines = String(value || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const [name, ...details] = lines;

  if (!name) return "—";

  return (
    <div className="min-w-64 max-w-80">
      <p className="font-bold text-slate-900">{name}</p>
      {details.map((detail, index) => (
        <p
          key={`${detail}-${index}`}
          className={`${index > 0 ? "no-print" : ""} ${
            index === 0 ? "mt-1" : ""
          } text-xs leading-5 text-slate-500`}
        >
          {detail}
        </p>
      ))}
    </div>
  );
}

const DuePolicyCard = memo(function DuePolicyCard({
  item,
  critical = false,
  monthKey,
  onMarkPaid,
  onMarkUnpaid,
}) {
  const { policy, dueDate, paid } = item;
  const phone = getWhatsAppNumber(policy.clientName);

  return (
    <article className={`due-card ${critical ? "due-card-critical" : ""}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <p className="truncate text-base font-black text-slate-950">
            {getClientName(policy.clientName)}
          </p>
          <p className="mt-1 text-xs font-semibold text-slate-500">
            Policy {policy.policyNumber || "—"}
          </p>
        </div>
        <span
          className={`shrink-0 rounded-full px-2.5 py-1 text-[10px] font-black uppercase tracking-wider ${
            paid
              ? "bg-emerald-100 text-emerald-700"
              : critical
                ? "bg-red-100 text-red-700"
                : "bg-amber-100 text-amber-700"
          }`}
        >
          {paid ? "Paid" : critical ? "Critical" : "Due"}
        </span>
      </div>

      <div className="mt-4 grid grid-cols-2 gap-3 text-sm">
        <div>
          <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Mode</p>
          <p className="mt-1 font-bold text-slate-700">{policy.mode}</p>
        </div>
        <div>
          <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">
            Due date
          </p>
          <p className="mt-1 font-bold text-slate-700">{formatDate(dueDate)}</p>
        </div>
        <div className="col-span-2">
          <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">
            Premium
          </p>
          <p className="mt-1 text-xl font-black text-[#0756a0]">
            ₹{currencyFormatter.format(Number(policy.premium) || 0)}
          </p>
        </div>
      </div>

      <div className="mt-4 flex items-center gap-2 border-t border-slate-200/70 pt-4">
        <button
          onClick={() => sendPolicyOnWhatsApp(policy)}
          disabled={!phone}
          title={phone ? "Send WhatsApp reminder" : "No phone number found"}
          className="grid size-10 shrink-0 place-items-center rounded-xl bg-emerald-50 text-emerald-600 transition hover:bg-emerald-100 disabled:cursor-not-allowed disabled:opacity-30"
        >
          <MessageCircle size={17} />
        </button>
        <button
          onClick={() =>
            paid
              ? onMarkUnpaid(policy.id, monthKey)
              : onMarkPaid(policy.id, monthKey)
          }
          className={`inline-flex min-h-10 flex-1 items-center justify-center gap-2 rounded-xl px-4 text-xs font-black transition ${
            paid
              ? "bg-red-50 text-red-700 hover:bg-red-100"
              : "bg-[#0756a0] text-white hover:bg-[#064985]"
          }`}
        >
          {paid ? <X size={16} /> : <CheckCircle2 size={16} />}
          {paid ? "Mark unpaid" : "Mark as paid"}
        </button>
      </div>
    </article>
  );
});

const PremiumDues = memo(function PremiumDues({
  policies,
  onMarkPaid,
  onMarkUnpaid,
}) {
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonthKey);
  const [modeFilter, setModeFilter] = useState("All");

  const dues = useMemo(
    () =>
      policies
        .map((policy) => {
          const dueDate = getPolicyDueDate(policy, selectedMonth);
          if (!dueDate) return null;
          return {
            policy,
            dueDate,
            paid: Array.isArray(policy.paidPeriods)
              ? policy.paidPeriods.includes(selectedMonth)
              : false,
          };
        })
        .filter(Boolean)
        .sort((left, right) => left.dueDate.localeCompare(right.dueDate)),
    [policies, selectedMonth],
  );

  const visibleDues = useMemo(
    () => dues.filter((item) => modeFilter === "All" || item.policy.mode === modeFilter),
    [dues, modeFilter],
  );
  const criticalDues = visibleDues.filter((item) => !item.paid);
  const paidDues = visibleDues.filter((item) => item.paid);
  const premiumDue = criticalDues.reduce(
    (total, item) => total + (Number(item.policy.premium) || 0),
    0,
  );
  const missingDates = useMemo(
    () =>
      policies.filter((policy) => !toStoredDate(policy.commencementDate))
        .length,
    [policies],
  );
  const duesResetKey = `${selectedMonth}|${modeFilter}`;
  const criticalPage = usePaginatedItems(
    criticalDues,
    `${duesResetKey}|critical`,
  );
  const scheduledPage = usePaginatedItems(
    visibleDues,
    `${duesResetKey}|scheduled`,
  );

  return (
    <section>
      <PageHeading
        eyebrow="Premium calendar"
        title="Premium dues"
        description="Recurring dues are calculated from each policy’s Com.Date and payment mode."
        actions={
          <label className="month-picker">
            <span>View month</span>
            <input
              type="month"
              value={selectedMonth}
              onChange={(event) => setSelectedMonth(event.target.value || getCurrentMonthKey())}
            />
          </label>
        }
      />

      <div className="floating-window dues-controls mt-5 flex flex-col gap-4 rounded-3xl p-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <p className="text-xs font-black uppercase tracking-[0.16em] text-[#0756a0]">
            {getMonthLabel(selectedMonth)}
          </p>
          <p className="mt-1 text-sm text-slate-500">Choose a premium frequency.</p>
        </div>
        <div className="policy-filters flex gap-2 overflow-x-auto pb-1">
          {["All", ...MODES].map((mode) => (
            <FilterButton
              key={mode}
              active={modeFilter === mode}
              onClick={() => setModeFilter(mode)}
            >
              {mode}
            </FilterButton>
          ))}
        </div>
      </div>

      <div className="dues-summary mt-5 grid grid-cols-2 gap-3 lg:grid-cols-4">
        {[
          ["Due policies", visibleDues.length, "blue"],
          ["Critical unpaid", criticalDues.length, "red"],
          ["Paid", paidDues.length, "green"],
          ["Premium pending", `₹${currencyFormatter.format(premiumDue)}`, "yellow"],
        ].map(([label, value, tone]) => (
          <div key={label} data-tone={tone} className="dues-summary-card">
            <p className="text-2xl font-black tracking-tight text-slate-950">{value}</p>
            <p className="mt-1 text-xs font-bold text-slate-500">{label}</p>
          </div>
        ))}
      </div>

      <div className="critical-window mt-6 rounded-3xl p-4 sm:p-5">
        <div className="flex items-start gap-3">
          <div className="grid size-10 shrink-0 place-items-center rounded-xl bg-red-100 text-red-600">
            <AlertTriangle size={19} />
          </div>
          <div>
            <p className="font-black text-red-950">Critical unpaid this month</p>
            <p className="mt-1 text-sm text-red-700/75">
              {criticalDues.length
                ? `${criticalDues.length} premium ${criticalDues.length === 1 ? "payment needs" : "payments need"} attention.`
                : "No unpaid premiums for this selection."}
            </p>
          </div>
        </div>

        {criticalDues.length > 0 && (
          <>
            <div className="dues-grid mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
              {criticalPage.pageItems.map((item) => (
                <DuePolicyCard
                  key={`critical-${item.policy.id}`}
                  item={item}
                  critical
                  monthKey={selectedMonth}
                  onMarkPaid={onMarkPaid}
                  onMarkUnpaid={onMarkUnpaid}
                />
              ))}
            </div>
            <Pagination
              page={criticalPage.page}
              totalPages={criticalPage.totalPages}
              totalItems={criticalDues.length}
              onPageChange={criticalPage.setPage}
            />
          </>
        )}
      </div>

      <div className="mt-7">
        <div className="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
          <div>
            <p className="text-xs font-black uppercase tracking-[0.16em] text-[#0756a0]">
              Dues by mode
            </p>
            <h2 className="mt-1 text-2xl font-black tracking-tight text-slate-950">
              {modeFilter === "All" ? "All scheduled premiums" : `${modeFilter} premiums`}
            </h2>
          </div>
          {missingDates > 0 && (
            <p className="text-xs font-semibold text-amber-700">
              {missingDates} {missingDates === 1 ? "policy has" : "policies have"} no Com.Date.
            </p>
          )}
        </div>

        {visibleDues.length > 0 ? (
          <div className="floating-window mt-4 overflow-hidden rounded-3xl">
            <div className="dues-grid grid gap-3 p-4 md:grid-cols-2 xl:grid-cols-3">
              {scheduledPage.pageItems.map((item) => (
                <DuePolicyCard
                  key={item.policy.id}
                  item={item}
                  monthKey={selectedMonth}
                  onMarkPaid={onMarkPaid}
                  onMarkUnpaid={onMarkUnpaid}
                />
              ))}
            </div>
            <Pagination
              page={scheduledPage.page}
              totalPages={scheduledPage.totalPages}
              totalItems={visibleDues.length}
              onPageChange={scheduledPage.setPage}
            />
          </div>
        ) : (
          <div className="floating-window mt-4 rounded-3xl px-6 py-12 text-center">
            <BellRing className="mx-auto text-slate-300" size={30} />
            <p className="mt-3 font-black text-slate-800">No premiums due</p>
            <p className="mt-1 text-sm text-slate-500">
              No matching policy is scheduled for {getMonthLabel(selectedMonth)}.
            </p>
          </div>
        )}
      </div>
    </section>
  );
});

const PaidHistory = memo(function PaidHistory({ policies, onMarkUnpaid }) {
  const [query, setQuery] = useState("");
  const [selectedMonth, setSelectedMonth] = useState("");
  const deferredQuery = useDeferredValue(query);

  const paidRecords = useMemo(
    () =>
      policies
        .flatMap((policy) =>
          (Array.isArray(policy.paidPeriods) ? policy.paidPeriods : []).map((monthKey) => ({
            policy,
            monthKey,
            dueDate: getPolicyDueDate(policy, monthKey) || `${monthKey}-01`,
          })),
        )
        .sort((left, right) => right.monthKey.localeCompare(left.monthKey)),
    [policies],
  );

  const visibleRecords = useMemo(() => {
    const search = deferredQuery.trim().toLowerCase();
    return paidRecords.filter(
      ({ policy, monthKey }) =>
        (!selectedMonth || monthKey === selectedMonth) &&
        (!search || getPolicySearchText(policy).includes(search)),
    );
  }, [deferredQuery, paidRecords, selectedMonth]);
  const paidPage = usePaginatedItems(
    visibleRecords,
    `${deferredQuery}|${selectedMonth}`,
  );

  const totalPaid = visibleRecords.reduce(
    (total, item) => total + (Number(item.policy.premium) || 0),
    0,
  );
  const paidPolicyCount = new Set(visibleRecords.map((item) => item.policy.id)).size;

  return (
    <section>
      <PageHeading
        eyebrow="Payment archive"
        title="Paid premium history"
        description="Every premium marked paid is stored here by policy and month on this device."
        actions={
          <label className="month-picker">
            <span>Filter month</span>
            <input
              type="month"
              value={selectedMonth}
              onChange={(event) => setSelectedMonth(event.target.value)}
            />
          </label>
        }
      />

      <div className="dues-summary mt-5 grid grid-cols-2 gap-3 lg:grid-cols-3">
        {[
          ["Paid entries", visibleRecords.length, "green"],
          ["Paid policies", paidPolicyCount, "blue"],
          ["Premium recorded", `₹${currencyFormatter.format(totalPaid)}`, "yellow"],
        ].map(([label, value, tone]) => (
          <div key={label} data-tone={tone} className="dues-summary-card">
            <p className="text-2xl font-black tracking-tight text-slate-950">{value}</p>
            <p className="mt-1 text-xs font-bold text-slate-500">{label}</p>
          </div>
        ))}
      </div>

      <div className="floating-window mt-5 rounded-3xl p-4 sm:p-5">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
          <label className="relative block flex-1">
            <Search
              size={18}
              className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <span className="sr-only">Search paid premiums</span>
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Search name, policy number or agent code"
              className="h-11 w-full rounded-xl border border-slate-200 bg-white/80 pl-11 pr-4 text-sm text-slate-800 outline-none transition placeholder:text-slate-400 focus:border-[#0756a0] focus:ring-4 focus:ring-[#0756a0]/10"
            />
          </label>
          {selectedMonth && (
            <button
              onClick={() => setSelectedMonth("")}
              className="min-h-11 rounded-xl border border-slate-200 bg-white px-4 text-xs font-black text-slate-600 transition hover:text-[#0756a0]"
            >
              Show all months
            </button>
          )}
        </div>
      </div>

      {visibleRecords.length > 0 ? (
        <div className="floating-window mt-5 overflow-hidden rounded-3xl">
          <div className="paid-history-grid grid gap-3 p-4 md:grid-cols-2 xl:grid-cols-3">
            {paidPage.pageItems.map(({ policy, monthKey, dueDate }) => {
              const phone = getWhatsAppNumber(policy.clientName);
              return (
                <article key={`${policy.id}-${monthKey}`} className="due-card paid-history-card">
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <p className="truncate text-base font-black text-slate-950">
                      {getClientName(policy.clientName)}
                    </p>
                    <p className="mt-1 text-xs font-semibold text-slate-500">
                      Policy {policy.policyNumber || "—"}
                    </p>
                  </div>
                  <span className="shrink-0 rounded-full bg-emerald-100 px-2.5 py-1 text-[10px] font-black uppercase tracking-wider text-emerald-700">
                    Paid
                  </span>
                </div>

                <div className="mt-4 grid grid-cols-2 gap-3">
                  <div>
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">
                      Paid month
                    </p>
                    <p className="mt-1 text-sm font-bold text-slate-700">
                      {getMonthLabel(monthKey)}
                    </p>
                  </div>
                  <div>
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">
                      Due date
                    </p>
                    <p className="mt-1 text-sm font-bold text-slate-700">
                      {formatDate(dueDate)}
                    </p>
                  </div>
                  <div>
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">
                      Mode
                    </p>
                    <p className="mt-1 text-sm font-bold text-slate-700">{policy.mode}</p>
                  </div>
                  <div>
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">
                      Premium
                    </p>
                    <p className="mt-1 text-sm font-black text-[#0756a0]">
                      ₹{currencyFormatter.format(Number(policy.premium) || 0)}
                    </p>
                  </div>
                </div>

                <div className="mt-4 flex items-center gap-2 border-t border-slate-200/70 pt-4">
                  <button
                    onClick={() => sendPolicyOnWhatsApp(policy)}
                    disabled={!phone}
                    title={phone ? "Send WhatsApp" : "No phone number found"}
                    className="grid size-10 shrink-0 place-items-center rounded-xl bg-emerald-50 text-emerald-600 transition hover:bg-emerald-100 disabled:cursor-not-allowed disabled:opacity-30"
                  >
                    <MessageCircle size={17} />
                  </button>
                  <button
                    onClick={() => onMarkUnpaid(policy.id, monthKey)}
                    className="inline-flex min-h-10 flex-1 items-center justify-center gap-2 rounded-xl bg-red-50 px-4 text-xs font-black text-red-700 transition hover:bg-red-100"
                  >
                    <X size={16} />
                    Mark unpaid
                  </button>
                </div>
                </article>
              );
            })}
          </div>
          <Pagination
            page={paidPage.page}
            totalPages={paidPage.totalPages}
            totalItems={visibleRecords.length}
            onPageChange={paidPage.setPage}
          />
        </div>
      ) : (
        <div className="floating-window mt-5 rounded-3xl px-6 py-14 text-center">
          <BadgeCheck className="mx-auto text-slate-300" size={32} />
          <p className="mt-3 font-black text-slate-800">No paid premiums found</p>
          <p className="mt-1 text-sm text-slate-500">
            Mark a premium paid on the Dues page and it will appear here.
          </p>
        </div>
      )}
    </section>
  );
});

const PolicyRow = memo(function PolicyRow({ policy, onEdit, onDelete }) {
  const whatsappNumber = getWhatsAppNumber(policy.clientName);

  return (
    <tr className="policy-row group border-b border-slate-100 transition last:border-0 hover:bg-[#f8fbff]">
      <td data-label="SNo" className="px-4 py-4 text-sm text-slate-500">
        {policy.serialNumber || "—"}
      </td>
      <td data-label="Name" className="policy-client-cell px-4 py-4 text-sm">
        <ClientDetails value={policy.clientName} />
      </td>
      <td data-label="DOB" className="whitespace-nowrap px-4 py-4 text-sm text-slate-600">
        {formatDate(policy.dob)}
      </td>
      <td data-label="Policy" className="px-4 py-4 text-base font-black text-[#0756a0]">
        {policy.policyNumber || "—"}
      </td>
      <td data-label="Agcode" className="px-4 py-4 text-sm text-slate-600">
        {policy.agentCode || "—"}
      </td>
      <td
        data-label="Com.Date"
        className="whitespace-nowrap px-4 py-4 text-sm text-slate-600"
      >
        {formatDate(policy.commencementDate)}
      </td>
      <td data-label="P/T/PP" className="px-4 py-4 text-sm font-semibold text-[#0756a0]">
        {policy.plan || "—"}
      </td>
      <td data-label="SumAssd" className="px-4 py-4 text-sm font-semibold text-slate-800">
        {currencyFormatter.format(Number(policy.sumAssured) || 0)}
      </td>
      <td data-label="Mode" className="px-4 py-4 text-sm font-bold text-slate-700">
        {MODE_CODES[policy.mode] || policy.mode}
      </td>
      <td data-label="Premium" className="px-4 py-4 text-sm font-semibold text-slate-800">
        {currencyFormatter.format(Number(policy.premium) || 0)}
      </td>
      <td data-label="Status" className="px-4 py-4">
        <StatusBadge status={policy.status} />
      </td>
      <td
        data-label="Actions"
        className="policy-actions no-print sticky right-0 bg-white px-4 py-4 transition group-hover:bg-[#f8fbff]"
      >
        <div className="flex justify-end gap-1">
          <button
            aria-label={`Send WhatsApp message to ${getClientName(policy.clientName)}`}
            title={whatsappNumber ? "Send WhatsApp" : "No phone number found"}
            onClick={() => sendPolicyOnWhatsApp(policy)}
            disabled={!whatsappNumber}
            className="grid size-9 place-items-center rounded-lg text-emerald-600 transition hover:bg-emerald-50 hover:text-emerald-700 disabled:cursor-not-allowed disabled:opacity-25"
          >
            <MessageCircle size={16} />
          </button>
          <button
            aria-label={`Edit ${policy.clientName || policy.policyNumber}`}
            title="Edit policy"
            onClick={() => onEdit(policy)}
            className="grid size-9 place-items-center rounded-lg text-slate-400 transition hover:bg-[#eaf4ff] hover:text-[#0756a0]"
          >
            <Pencil size={16} />
          </button>
          <button
            aria-label={`Delete ${policy.clientName || policy.policyNumber}`}
            title="Delete policy"
            onClick={() => onDelete(policy)}
            className="grid size-9 place-items-center rounded-lg text-slate-400 transition hover:bg-red-50 hover:text-red-600"
          >
            <Trash2 size={16} />
          </button>
        </div>
      </td>
    </tr>
  );
});

const PolicyTable = memo(function PolicyTable({
  policies,
  agencyLabel,
  onAdd,
  onEdit,
  onDelete,
  onExport,
  onUpload,
  onPrint,
  renderAll = false,
}) {
  const [query, setQuery] = useState("");
  const [modeFilter, setModeFilter] = useState("All");
  const [sort, setSort] = useState({ key: "clientName", direction: "asc" });
  const deferredQuery = useDeferredValue(query);

  const filteredPolicies = useMemo(() => {
    const search = deferredQuery.trim().toLowerCase();
    const filtered = policies.filter((policy) => {
      const matchesSearch =
        !search || getPolicySearchText(policy).includes(search);
      const matchesMode = modeFilter === "All" || policy.mode === modeFilter;
      return matchesSearch && matchesMode;
    });

    return filtered.sort((a, b) => {
      let left = a[sort.key];
      let right = b[sort.key];
      if (["serialNumber", "sumAssured", "premium"].includes(sort.key)) {
        left = Number(left) || 0;
        right = Number(right) || 0;
      } else {
        left = String(left ?? "").toLowerCase();
        right = String(right ?? "").toLowerCase();
      }
      if (left < right) return sort.direction === "asc" ? -1 : 1;
      if (left > right) return sort.direction === "asc" ? 1 : -1;
      return 0;
    });
  }, [deferredQuery, modeFilter, policies, sort]);
  const policyPage = usePaginatedItems(
    filteredPolicies,
    `${deferredQuery}|${modeFilter}|${sort.key}|${sort.direction}`,
    renderAll,
  );

  const changeSort = (key) => {
    setSort((current) => ({
      key,
      direction: current.key === key && current.direction === "asc" ? "desc" : "asc",
    }));
  };

  return (
    <section className="print-sheet">
      <div className="no-print">
        <PageHeading
          eyebrow="Policy register"
          title="All policies"
          description="Search, filter and manage the records currently loaded from your workbook."
          actions={
            <>
              <button
                onClick={onPrint}
                disabled={!filteredPolicies.length}
                className="inline-flex min-h-11 items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 text-sm font-bold text-slate-700 shadow-sm transition hover:border-[#0756a0]/30 hover:text-[#0756a0] disabled:cursor-not-allowed disabled:opacity-45"
              >
                <Printer size={17} />
                Save PDF / Print
              </button>
              <button
                onClick={onExport}
                disabled={!policies.length}
                className="inline-flex min-h-11 items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 text-sm font-bold text-slate-700 shadow-sm transition hover:border-[#0756a0]/30 hover:text-[#0756a0] disabled:cursor-not-allowed disabled:opacity-45"
              >
                <Download size={17} />
                Export Excel
              </button>
              <PrimaryButton onClick={onAdd}>
                <Plus size={18} />
                Add Policy
              </PrimaryButton>
            </>
          }
        />
      </div>

      <div className="print-only mb-4">
        <h1 className="text-xl font-black">{agencyLabel} LIC Policy Register</h1>
        <p className="mt-1 text-xs">
          {filteredPolicies.length} {filteredPolicies.length === 1 ? "policy" : "policies"} •
          Printed {new Date().toLocaleDateString("en-IN")}
        </p>
      </div>

      <div className="floating-window policy-window print-table-shell mt-7 rounded-3xl">
        <div className="policy-toolbar no-print border-b border-slate-100 p-4 sm:p-5">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
            <label className="relative block w-full xl:max-w-md">
              <Search
                size={18}
                className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400"
              />
              <span className="sr-only">Search policies</span>
              <input
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                placeholder="Search name, policy number or agent code"
                className="h-11 w-full rounded-xl border border-slate-200 bg-slate-50 pl-11 pr-4 text-sm text-slate-800 outline-none transition placeholder:text-slate-400 focus:border-[#0756a0] focus:bg-white focus:ring-4 focus:ring-[#0756a0]/10"
              />
            </label>

            <div className="policy-filters flex flex-wrap items-center gap-2">
              <span className="mr-1 text-xs font-bold uppercase tracking-wider text-slate-400">
                Mode
              </span>
              {["All", ...MODES].map((mode) => (
                <FilterButton
                  key={mode}
                  active={modeFilter === mode}
                  onClick={() => setModeFilter(mode)}
                >
                  {mode}
                </FilterButton>
              ))}
            </div>
          </div>
        </div>

        {policies.length === 0 ? (
          <div className="flex min-h-80 flex-col items-center justify-center px-6 text-center">
            <div className="grid size-14 place-items-center rounded-2xl bg-[#eaf4ff] text-[#0756a0]">
              <FileSpreadsheet size={27} />
            </div>
            <h2 className="mt-4 text-lg font-black text-slate-900">No policies yet</h2>
            <p className="mt-2 max-w-sm text-sm leading-6 text-slate-500">
              Upload an Excel workbook or add your first policy manually.
            </p>
            <div className="mt-5 flex flex-wrap justify-center gap-3">
              <button
                onClick={onUpload}
                className="min-h-10 rounded-xl border border-slate-200 px-4 text-sm font-bold text-slate-700"
              >
                Upload Excel
              </button>
              <PrimaryButton onClick={onAdd} className="min-h-10">
                <Plus size={17} />
                Add Policy
              </PrimaryButton>
            </div>
          </div>
        ) : (
          <>
            <div className="print-table-container overflow-x-auto">
              <table className="policy-table w-full min-w-[1380px] border-collapse">
                <thead>
                  <tr className="bg-slate-50/80">
                    {COLUMNS.map((column) => (
                      <th
                        key={column.key}
                        scope="col"
                        className="border-b border-slate-200 px-4 py-3 text-left"
                      >
                        <button
                          onClick={() => changeSort(column.key)}
                          className="inline-flex items-center gap-1.5 text-[11px] font-black uppercase tracking-[0.09em] text-slate-500 transition hover:text-[#0756a0]"
                        >
                          {column.label}
                          <ArrowDownUp
                            size={13}
                            className={
                              sort.key === column.key ? "text-[#0756a0]" : "text-slate-300"
                            }
                          />
                        </button>
                      </th>
                    ))}
                    <th
                      scope="col"
                      className="no-print sticky right-0 border-b border-slate-200 bg-slate-50 px-4 py-3 text-right text-[11px] font-black uppercase tracking-[0.09em] text-slate-500"
                    >
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {policyPage.pageItems.map((policy) => (
                    <PolicyRow
                      key={policy.id}
                      policy={policy}
                      onEdit={onEdit}
                      onDelete={onDelete}
                    />
                  ))}
                </tbody>
              </table>
            </div>

            {filteredPolicies.length === 0 && (
              <div className="px-6 py-14 text-center">
                <Search className="mx-auto text-slate-300" size={28} />
                <p className="mt-3 font-bold text-slate-700">No matching policies</p>
                <p className="mt-1 text-sm text-slate-500">Try a different search or filter.</p>
              </div>
            )}

            <Pagination
              page={policyPage.page}
              totalPages={policyPage.totalPages}
              totalItems={filteredPolicies.length}
              onPageChange={policyPage.setPage}
            />
          </>
        )}
      </div>
    </section>
  );
});

function Field({ label, error, children, wide = false }) {
  return (
    <label className={wide ? "sm:col-span-2" : ""}>
      <span className="mb-1.5 block text-xs font-bold text-slate-700">{label}</span>
      {children}
      {error && <span className="mt-1.5 block text-xs font-semibold text-red-600">{error}</span>}
    </label>
  );
}

const inputClass =
  "h-11 w-full rounded-xl border border-slate-200 bg-white px-3.5 text-sm text-slate-800 outline-none transition placeholder:text-slate-400 focus:border-[#0756a0] focus:ring-4 focus:ring-[#0756a0]/10";

function PolicyModal({ mode, policy, policies, onClose, onSave }) {
  const [form, setForm] = useState(() => {
    const initial = policy ? { ...EMPTY_FORM, ...policy } : { ...EMPTY_FORM };
    return {
      ...initial,
      dob: formatDateValue(initial.dob),
      commencementDate: formatDateValue(initial.commencementDate),
    };
  });
  const [errors, setErrors] = useState({});

  const update = (field, value) => {
    setForm((current) => ({ ...current, [field]: value }));
    if (errors[field]) setErrors((current) => ({ ...current, [field]: "" }));
  };

  const submit = (event) => {
    event.preventDefault();
    const nextErrors = {};
    const dob = toStoredDate(form.dob);
    const commencementDate = toStoredDate(form.commencementDate);
    if (!form.clientName.trim()) nextErrors.clientName = "Client name is required.";
    if (!form.policyNumber.trim()) {
      nextErrors.policyNumber = "Policy number is required.";
    } else {
      const duplicate = policies.some(
        (item) =>
          item.id !== policy?.id &&
          item.policyNumber.trim().toLowerCase() === form.policyNumber.trim().toLowerCase(),
      );
      if (duplicate) nextErrors.policyNumber = "This policy number already exists.";
    }
    if (form.premium === "" || Number(form.premium) < 0) {
      nextErrors.premium = "Enter a valid premium.";
    }
    if (dob === null) nextErrors.dob = "Use DD/MM/YYYY.";
    if (commencementDate === null) nextErrors.commencementDate = "Use DD/MM/YYYY.";

    if (Object.keys(nextErrors).length) {
      setErrors(nextErrors);
      return;
    }

    onSave({
      ...form,
      id: policy?.id ?? newId(),
      serialNumber: String(form.serialNumber || policies.length + 1).trim(),
      clientName: form.clientName.trim(),
      dob,
      policyNumber: form.policyNumber.trim(),
      agentCode: form.agentCode.trim(),
      commencementDate,
      plan: form.plan.trim(),
      sumAssured: Number(form.sumAssured) || 0,
      premium: Number(form.premium) || 0,
    });
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="policy-form-title"
      className="fixed inset-0 z-50 flex items-end justify-center bg-slate-950/45 p-0 backdrop-blur-[2px] sm:items-center sm:p-5"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <div className="modal-window max-h-[94vh] w-full overflow-y-auto rounded-t-3xl bg-white shadow-2xl sm:max-w-2xl sm:rounded-3xl">
        <div className="sticky top-0 z-10 flex items-start justify-between border-b border-slate-100 bg-white px-5 py-5 sm:px-7">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.18em] text-[#0756a0]">
              {mode === "edit" ? "Update record" : "New record"}
            </p>
            <h2 id="policy-form-title" className="mt-1 text-2xl font-black text-slate-950">
              {mode === "edit" ? "Edit policy" : "Add policy"}
            </h2>
          </div>
          <button
            onClick={onClose}
            aria-label="Close form"
            className="grid size-10 place-items-center rounded-xl text-slate-400 transition hover:bg-slate-100 hover:text-slate-700"
          >
            <X size={20} />
          </button>
        </div>

        <form onSubmit={submit} className="grid gap-4 px-5 py-6 sm:grid-cols-2 sm:px-7">
          <Field label="SNo">
            <input
              autoFocus
              value={form.serialNumber}
              onChange={(event) => update("serialNumber", event.target.value)}
              className={inputClass}
              placeholder={String(policies.length + 1)}
            />
          </Field>
          <Field label="Policy *" error={errors.policyNumber}>
            <input
              value={form.policyNumber}
              onChange={(event) => update("policyNumber", event.target.value)}
              className={inputClass}
              placeholder="e.g. 144735098"
            />
          </Field>
          <Field label="Name / Phone / Address *" error={errors.clientName} wide>
            <textarea
              rows="4"
              value={form.clientName}
              onChange={(event) => update("clientName", event.target.value)}
              className={`${inputClass} h-auto resize-y py-3`}
              placeholder={"RIYAZ AHMED\n9149997192\nS/O Mohd Rafiq, Rajouri"}
            />
          </Field>
          <Field label="DOB (DD/MM/YYYY)" error={errors.dob}>
            <input
              type="text"
              inputMode="numeric"
              value={form.dob}
              onChange={(event) => update("dob", event.target.value)}
              className={inputClass}
              placeholder="03/05/1988"
            />
          </Field>
          <Field label="Agcode">
            <input
              value={form.agentCode}
              onChange={(event) => update("agentCode", event.target.value)}
              className={inputClass}
              placeholder="e.g. 0212313F"
            />
          </Field>
          <Field label="Com.Date (DD/MM/YYYY)" error={errors.commencementDate}>
            <input
              type="text"
              inputMode="numeric"
              value={form.commencementDate}
              onChange={(event) => update("commencementDate", event.target.value)}
              className={inputClass}
              placeholder="07/02/2021"
            />
          </Field>
          <Field label="P/T/PP">
            <input
              value={form.plan}
              onChange={(event) => update("plan", event.target.value)}
              className={inputClass}
              placeholder="e.g. 933/21/18"
            />
          </Field>
          <Field label="SumAssd">
            <input
              type="number"
              min="0"
              step="0.01"
              value={form.sumAssured}
              onChange={(event) => update("sumAssured", event.target.value)}
              className={inputClass}
              placeholder="e.g. 300000"
            />
          </Field>
          <Field label="Premium *" error={errors.premium}>
            <input
              type="number"
              min="0"
              step="0.01"
              value={form.premium}
              onChange={(event) => update("premium", event.target.value)}
              className={inputClass}
              placeholder="e.g. 1454"
            />
          </Field>
          <Field label="Mode">
            <select
              value={form.mode}
              onChange={(event) => update("mode", event.target.value)}
              className={inputClass}
            >
              {MODES.map((item) => (
                <option key={item} value={item}>
                  {MODE_CODES[item]} — {item}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Status">
            <select
              value={form.status}
              onChange={(event) => update("status", event.target.value)}
              className={inputClass}
            >
              <option>Unpaid</option>
              <option>Paid</option>
            </select>
          </Field>

          <div className="mt-2 flex flex-col-reverse gap-3 border-t border-slate-100 pt-5 sm:col-span-2 sm:flex-row sm:justify-end">
            <button
              type="button"
              onClick={onClose}
              className="min-h-11 rounded-xl border border-slate-200 px-5 text-sm font-bold text-slate-600 transition hover:bg-slate-50"
            >
              Cancel
            </button>
            <PrimaryButton type="submit" className="px-6">
              <Check size={18} />
              {mode === "edit" ? "Save Changes" : "Add Policy"}
            </PrimaryButton>
          </div>
        </form>
      </div>
    </div>
  );
}

function DeleteDialog({ policy, onClose, onConfirm }) {
  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="delete-title"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/45 p-5 backdrop-blur-[2px]"
    >
      <div className="modal-window w-full max-w-md rounded-3xl bg-white p-6 shadow-2xl sm:p-7">
        <div className="grid size-12 place-items-center rounded-2xl bg-red-50 text-red-600">
          <Trash2 size={22} />
        </div>
        <h2 id="delete-title" className="mt-5 text-2xl font-black text-slate-950">
          Delete this policy?
        </h2>
        <p className="mt-2 text-sm leading-6 text-slate-500">
          <span className="font-bold text-slate-700">
            {policy.clientName || policy.policyNumber}
          </span>{" "}
          will be removed from the current records. Export your workbook to save this change.
        </p>
        <div className="mt-6 flex flex-col-reverse gap-3 sm:flex-row sm:justify-end">
          <button
            onClick={onClose}
            className="min-h-11 rounded-xl border border-slate-200 px-5 text-sm font-bold text-slate-600"
          >
            Keep Policy
          </button>
          <button
            onClick={onConfirm}
            className="min-h-11 rounded-xl bg-red-600 px-5 text-sm font-bold text-white transition hover:bg-red-700"
          >
            Delete Policy
          </button>
        </div>
      </div>
    </div>
  );
}

function Toast({ message }) {
  if (!message) return null;
  return (
    <div
      role="status"
      className="fixed bottom-5 left-1/2 z-[60] flex w-[calc(100%-2rem)] max-w-sm -translate-x-1/2 items-center gap-3 rounded-2xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-2xl"
    >
      <span className="grid size-7 shrink-0 place-items-center rounded-lg bg-emerald-500">
        <Check size={15} strokeWidth={3} />
      </span>
      {message}
    </div>
  );
}

export default function App() {
  const [activePage, setActivePage] = useState("dashboard");
  const [activeAgency, setActiveAgency] = useState(loadActiveAgency);
  const agencyCodes = FIXED_AGENCY_CODES;
  const [policies, setPolicies] = useState(() =>
    loadStoredPolicies(activeAgency),
  );
  const [fileName, setFileName] = useState("");
  const [modal, setModal] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [toast, setToast] = useState("");
  const [renderAllForPrint, setRenderAllForPrint] = useState(false);
  const policiesRef = useRef(policies);
  const activeAgencyRef = useRef(activeAgency);
  const storageReadyRef = useRef(false);
  const storageDirtyRef = useRef(false);
  const agencyStartupCheckedRef = useRef(false);
  policiesRef.current = policies;
  activeAgencyRef.current = activeAgency;
  const activeAgencyLabel =
    AGENCIES.find((agency) => agency.id === activeAgency)?.label ??
    AGENCIES[0].label;
  const getCurrentAgencyCollections = useCallback(
    () =>
      Object.fromEntries(
        AGENCIES.map((agency) => [
          agency.id,
          agency.id === activeAgencyRef.current
            ? policiesRef.current
            : loadStoredPolicies(agency.id),
        ]),
      ),
    [],
  );
  const saveAgencyCollections = useCallback((collections) => {
    AGENCIES.forEach((agency) => {
      persistPolicies(collections[agency.id] ?? [], agency.id);
    });
    const nextActivePolicies =
      collections[activeAgencyRef.current] ?? [];
    policiesRef.current = nextActivePolicies;
    storageDirtyRef.current = false;
    setPolicies(nextActivePolicies);
  }, []);

  useEffect(() => {
    if (agencyStartupCheckedRef.current) return;
    agencyStartupCheckedRef.current = true;
    const currentCollections = getCurrentAgencyCollections();
    const redistributed = redistributeAgencyPolicies(
      currentCollections,
      agencyCodes,
    );
    if (redistributed.moved > 0) {
      saveAgencyCollections(redistributed.collections);
    }
  }, [
    agencyCodes,
    getCurrentAgencyCollections,
    saveAgencyCollections,
  ]);
  const editPolicy = useCallback((policy) => setModal({ mode: "edit", policy }), []);
  const addPolicy = useCallback(() => setModal({ mode: "add" }), []);
  const showUpload = useCallback(() => setActivePage("upload"), []);
  const showPolicies = useCallback(() => setActivePage("policies"), []);
  const switchAgency = useCallback((nextAgency) => {
    if (!isAgencyId(nextAgency) || nextAgency === activeAgencyRef.current) return;

    persistPolicies(policiesRef.current, activeAgencyRef.current);
    const nextPolicies = loadStoredPolicies(nextAgency);
    try {
      window.localStorage.setItem(ACTIVE_AGENCY_STORAGE_KEY, nextAgency);
    } catch {
      // The agency still switches for this session if storage is unavailable.
    }
    storageDirtyRef.current = false;
    activeAgencyRef.current = nextAgency;
    policiesRef.current = nextPolicies;
    setActiveAgency(nextAgency);
    setPolicies(nextPolicies);
    setFileName("");
    setModal(null);
    setDeleteTarget(null);
    const label =
      AGENCIES.find((agency) => agency.id === nextAgency)?.label ??
      AGENCIES[0].label;
    setToast(`Switched to ${label}. Its records are separate.`);
  }, []);

  const printPolicies = useCallback(() => {
    if (!policies.length) return;
    setRenderAllForPrint(true);
    if (activePage !== "policies") setActivePage("policies");
    window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => window.print());
    });
  }, [activePage, policies.length]);

  const shareWorkbookOnWhatsApp = useCallback(async () => {
    if (!policies.length) return;

    try {
      const file = await createPolicyWorkbookFile(policies, activeAgencyLabel);
      const shareData = {
        title: `${activeAgencyLabel} LIC Policy Tracker`,
        text: `Updated LIC policy workbook for ${activeAgencyLabel}`,
        files: [file],
      };
      const canShareFile =
        typeof navigator.share === "function" &&
        (typeof navigator.canShare !== "function" || navigator.canShare(shareData));

      if (canShareFile) {
        await navigator.share(shareData);
        setToast("Workbook shared.");
        return;
      }

      await exportPolicyWorkbook(policies, activeAgencyLabel);
      window.open("https://web.whatsapp.com/", "_blank", "noopener,noreferrer");
      setToast("Workbook downloaded. Attach it in WhatsApp.");
    } catch (error) {
      if (error instanceof Error && error.name === "AbortError") return;
      setToast("Unable to share the workbook on this device.");
    }
  }, [activeAgencyLabel, policies]);

  useEffect(() => {
    if (!toast) return undefined;
    const timeout = window.setTimeout(() => setToast(""), 2600);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  useEffect(() => {
    const beforePrint = () => setRenderAllForPrint(true);
    const afterPrint = () => setRenderAllForPrint(false);
    window.addEventListener("beforeprint", beforePrint);
    window.addEventListener("afterprint", afterPrint);
    return () => {
      window.removeEventListener("beforeprint", beforePrint);
      window.removeEventListener("afterprint", afterPrint);
    };
  }, []);

  useEffect(() => {
    if (!storageReadyRef.current) {
      storageReadyRef.current = true;
      return undefined;
    }
    storageDirtyRef.current = true;
    const saveWhenIdle = () => {
      if (persistPolicies(policies, activeAgency)) {
        storageDirtyRef.current = false;
      }
    };

    if (typeof window.requestIdleCallback === "function") {
      const idleId = window.requestIdleCallback(saveWhenIdle, {
        timeout: 2500,
      });
      return () => window.cancelIdleCallback(idleId);
    }

    const timeout = window.setTimeout(saveWhenIdle, 700);
    return () => window.clearTimeout(timeout);
  }, [activeAgency, policies]);

  useEffect(() => {
    const flushStorage = () => {
      if (!storageDirtyRef.current) return;
      if (
        persistPolicies(policiesRef.current, activeAgencyRef.current)
      ) {
        storageDirtyRef.current = false;
      }
    };
    const flushWhenHidden = () => {
      if (document.visibilityState === "hidden") flushStorage();
    };
    window.addEventListener("pagehide", flushStorage);
    document.addEventListener("visibilitychange", flushWhenHidden);
    return () => {
      window.removeEventListener("pagehide", flushStorage);
      document.removeEventListener("visibilitychange", flushWhenHidden);
    };
  }, []);

  useEffect(() => {
    if (!modal && !deleteTarget) return undefined;
    const onKeyDown = (event) => {
      if (event.key === "Escape") {
        setModal(null);
        setDeleteTarget(null);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [deleteTarget, modal]);

  const savePolicy = useCallback(
    (policy) => {
      const paidPeriods = Array.isArray(policy.paidPeriods)
        ? policy.paidPeriods
        : policy.status === "Paid"
          ? [getCurrentMonthKey()]
          : [];
      const trackedPolicy = { ...policy, paidPeriods };
      const currentCollections = getCurrentAgencyCollections();
      const redistributed = redistributeAgencyPolicies(
        currentCollections,
        agencyCodes,
      ).collections;
      const targetAgency =
        getAgencyIdForCode(trackedPolicy.agentCode, agencyCodes) ??
        activeAgencyRef.current;
      const policyNumber = trackedPolicy.policyNumber.trim().toLowerCase();
      const duplicate = redistributed[targetAgency].some(
        (item) =>
          item.id !== trackedPolicy.id &&
          item.policyNumber.trim().toLowerCase() === policyNumber,
      );

      if (duplicate) {
        setToast("This policy number already exists in its agency.");
        return;
      }

      AGENCIES.forEach((agency) => {
        redistributed[agency.id] = redistributed[agency.id].filter(
          (item) => item.id !== trackedPolicy.id,
        );
      });
      redistributed[targetAgency] = [
        trackedPolicy,
        ...redistributed[targetAgency],
      ];
      AGENCIES.forEach((agency) => {
        redistributed[agency.id] = renumberPolicies(
          redistributed[agency.id],
        );
      });
      saveAgencyCollections(redistributed);
      setModal(null);

      const targetLabel =
        AGENCIES.find((agency) => agency.id === targetAgency)?.label ??
        activeAgencyLabel;
      const action = modal?.mode === "edit" ? "updated" : "added";
      setToast(
        targetAgency === activeAgencyRef.current
          ? `Policy ${action}.`
          : `Policy ${action} and moved to ${targetLabel} by Agcode.`,
      );
    },
    [
      activeAgencyLabel,
      agencyCodes,
      getCurrentAgencyCollections,
      modal?.mode,
      saveAgencyCollections,
    ],
  );

  const importPolicies = useCallback(
    (imported, name) => {
      const currentMonth = getCurrentMonthKey();
      const currentCollections = getCurrentAgencyCollections();
      const redistributed = redistributeAgencyPolicies(
        currentCollections,
        agencyCodes,
      ).collections;
      const importedByAgency = Object.fromEntries(
        AGENCIES.map((agency) => [agency.id, []]),
      );
      let unmatched = 0;

      imported.forEach((policy) => {
        const trackedPolicy = {
          ...policy,
          paidPeriods: policy.status === "Paid" ? [currentMonth] : [],
        };
        const matchingAgency = getAgencyIdForCode(
          trackedPolicy.agentCode,
          agencyCodes,
        );
        if (!matchingAgency) unmatched += 1;
        importedByAgency[matchingAgency ?? activeAgencyRef.current].push(
          trackedPolicy,
        );
      });

      let added = 0;
      let skipped = 0;
      const destinations = [];
      AGENCIES.forEach((agency) => {
        const result = mergePolicyCollections(
          redistributed[agency.id],
          importedByAgency[agency.id],
        );
        redistributed[agency.id] = result.policies;
        added += result.added;
        skipped += result.skipped;
        destinations.push({
          id: agency.id,
          label: agency.label,
          added: result.added,
        });
      });

      saveAgencyCollections(redistributed);
      setFileName(name);
      setToast(
        added
          ? `Excel data sorted by Agcode. ${added} ${added === 1 ? "policy was" : "policies were"} added and previous data was kept.`
          : "No new policies were added. Existing data was kept.",
      );
      return { added, skipped, unmatched, destinations };
    },
    [
      agencyCodes,
      getCurrentAgencyCollections,
      saveAgencyCollections,
    ],
  );

  const markPremiumPaid = useCallback((policyId, monthKey) => {
    setPolicies((current) =>
      current.map((policy) => {
        if (policy.id !== policyId) return policy;
        const paidPeriods = Array.from(
          new Set([...(Array.isArray(policy.paidPeriods) ? policy.paidPeriods : []), monthKey]),
        );
        return { ...policy, paidPeriods, status: "Paid" };
      }),
    );
    setToast(`Premium marked paid for ${getMonthLabel(monthKey)}.`);
  }, []);

  const markPremiumUnpaid = useCallback((policyId, monthKey) => {
    setPolicies((current) =>
      current.map((policy) => {
        if (policy.id !== policyId) return policy;
        const paidPeriods = (
          Array.isArray(policy.paidPeriods) ? policy.paidPeriods : []
        ).filter((period) => period !== monthKey);
        return {
          ...policy,
          paidPeriods,
          status: paidPeriods.length ? "Paid" : "Unpaid",
        };
      }),
    );
    setToast(`Premium marked unpaid for ${getMonthLabel(monthKey)}.`);
  }, []);

  const deletePolicy = useCallback(() => {
    if (!deleteTarget) return;
    setPolicies((current) => current.filter((item) => item.id !== deleteTarget.id));
    setDeleteTarget(null);
    setToast("Policy deleted.");
  }, [deleteTarget]);

  const exportPolicies = useCallback(
    () => exportPolicyWorkbook(policies, activeAgencyLabel),
    [activeAgencyLabel, policies],
  );

  return (
    <div className="app-shell min-h-screen">
      <div className="ambient-shape ambient-shape-one" aria-hidden="true" />
      <div className="ambient-shape ambient-shape-two" aria-hidden="true" />
      <div className="ambient-shape ambient-shape-three" aria-hidden="true" />
      <header className="app-header no-print sticky top-0 z-40 px-3 pt-3 sm:px-5">
        <div className="app-header-panel mx-auto flex max-w-[1480px] flex-wrap items-center justify-between gap-4 px-4 py-3 sm:px-6">
          <Brand agencyLabel={activeAgencyLabel} />

          <nav
            aria-label="Primary navigation"
            className="floating-nav order-3 flex w-full items-center gap-1 overflow-x-auto rounded-xl p-1 sm:order-none sm:w-auto"
          >
            {NAV_ITEMS.map(({ id, label, icon: Icon }) => (
              <button
                key={id}
                onClick={() => setActivePage(id)}
                className={`inline-flex min-h-9 flex-1 items-center justify-center gap-2 whitespace-nowrap rounded-lg px-3 text-xs font-bold transition sm:flex-none ${
                  activePage === id
                    ? "bg-white text-[#0756a0] shadow-sm"
                    : "text-slate-500 hover:text-slate-800"
                }`}
              >
                <Icon size={15} />
                {label}
              </button>
            ))}
          </nav>

          <div className="header-actions flex items-center gap-2">
            <button
              onClick={shareWorkbookOnWhatsApp}
              disabled={!policies.length}
              title="Share Excel file on WhatsApp"
              aria-label="Share Excel file on WhatsApp"
              className="whatsapp-file-button inline-flex min-h-10 items-center gap-2 rounded-xl bg-emerald-500 px-3 text-xs font-black text-white transition hover:bg-emerald-600 disabled:cursor-not-allowed disabled:opacity-45 sm:px-4"
            >
              <MessageCircle size={17} />
              <span className="header-button-label">WhatsApp</span>
            </button>
            <button
              onClick={printPolicies}
              disabled={!policies.length}
              title="Save PDF or print"
              className="inline-flex min-h-10 items-center gap-2 rounded-xl bg-[#ffcf2f] px-3 text-xs font-black text-[#172033] transition hover:bg-[#f8c719] disabled:cursor-not-allowed disabled:opacity-45 sm:px-4"
            >
              <Printer size={16} />
              <span className="header-button-label">Save PDF / Print</span>
            </button>
            <button
              onClick={exportPolicies}
              disabled={!policies.length}
              className="hidden min-h-10 items-center gap-2 rounded-xl bg-[#0756a0] px-4 text-xs font-black text-white transition hover:bg-[#064985] disabled:cursor-not-allowed disabled:opacity-45 lg:inline-flex"
            >
              <Download size={16} />
              Export Updated Excel
            </button>
          </div>
        </div>
      </header>

      <main className="app-main relative z-10 mx-auto min-h-[calc(100vh-136px)] max-w-[1480px] px-4 py-7 sm:px-6 sm:py-9 lg:px-8">
        {fileName && (
          <div className="no-print mb-5 flex items-center gap-2 text-xs font-semibold text-slate-500">
            <FileCheck2 size={15} className="text-emerald-600" />
            {activeAgencyLabel} loaded:{" "}
            <span className="max-w-[260px] truncate text-slate-700">{fileName}</span>
          </div>
        )}

        {activePage === "dashboard" && (
          <Dashboard
            policies={policies}
            activeAgency={activeAgency}
            onAgencyChange={switchAgency}
            onAdd={addPolicy}
            onUpload={showUpload}
          />
        )}
        {activePage === "upload" && (
          <UploadExcel
            policies={policies}
            onImported={importPolicies}
            onGoToPolicies={showPolicies}
          />
        )}
        {activePage === "dues" && (
          <PremiumDues
            policies={policies}
            onMarkPaid={markPremiumPaid}
            onMarkUnpaid={markPremiumUnpaid}
          />
        )}
        {activePage === "paid" && (
          <PaidHistory policies={policies} onMarkUnpaid={markPremiumUnpaid} />
        )}
        {activePage === "policies" && (
          <PolicyTable
            policies={policies}
            agencyLabel={activeAgencyLabel}
            onAdd={addPolicy}
            onEdit={editPolicy}
            onDelete={setDeleteTarget}
            onExport={exportPolicies}
            onUpload={showUpload}
            onPrint={printPolicies}
            renderAll={renderAllForPrint}
          />
        )}
      </main>

      <footer className="app-footer no-print relative z-10">
        <div className="mx-auto flex max-w-[1480px] flex-col gap-2 px-4 py-5 text-xs text-slate-400 sm:flex-row sm:items-center sm:justify-between sm:px-6 lg:px-8">
          <p className="font-semibold text-slate-500">LIC Policy Tracker</p>
          <p>No server • No account • Your data stays on this device</p>
        </div>
      </footer>

      {modal && (
        <PolicyModal
          mode={modal.mode}
          policy={modal.policy}
          policies={policies}
          onClose={() => setModal(null)}
          onSave={savePolicy}
        />
      )}
      {deleteTarget && (
        <DeleteDialog
          policy={deleteTarget}
          onClose={() => setDeleteTarget(null)}
          onConfirm={deletePolicy}
        />
      )}
      <Toast message={toast} />
    </div>
  );
}
