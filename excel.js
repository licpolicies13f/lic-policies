export const POLICY_FIELDS = [
  "serialNumber",
  "clientName",
  "dob",
  "policyNumber",
  "agentCode",
  "commencementDate",
  "plan",
  "sumAssured",
  "mode",
  "premium",
  "status",
];

const EXPORT_HEADERS = {
  serialNumber: "SNo",
  clientName: "Name",
  dob: "DOB",
  policyNumber: "Policy",
  agentCode: "Agcode",
  commencementDate: "Com.Date",
  plan: "P/T/PP",
  sumAssured: "SumAssd",
  mode: "Mode",
  premium: "Premium",
  status: "Status",
};

const HEADER_LOOKUP = {
  sno: "serialNumber",
  srno: "serialNumber",
  serialno: "serialNumber",
  serialnumber: "serialNumber",
  name: "clientName",
  clientname: "clientName",
  customername: "clientName",
  dob: "dob",
  dateofbirth: "dob",
  policy: "policyNumber",
  policynumber: "policyNumber",
  policyno: "policyNumber",
  policynum: "policyNumber",
  agcode: "agentCode",
  agentcode: "agentCode",
  comdate: "commencementDate",
  commdate: "commencementDate",
  commencementdate: "commencementDate",
  ptpp: "plan",
  plan: "plan",
  planname: "plan",
  sumassd: "sumAssured",
  sumassured: "sumAssured",
  assuredsum: "sumAssured",
  mode: "mode",
  premiummode: "mode",
  paymentmode: "mode",
  premium: "premium",
  premiumamount: "premium",
  status: "status",
  paymentstatus: "status",
  phonenumber: "phoneNumber",
  phone: "phoneNumber",
  mobile: "phoneNumber",
  mobilenumber: "phoneNumber",
  address: "address",
};

const cleanHeader = (value) =>
  String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const toText = (value) => String(value ?? "").trim();

const toAmount = (value) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const firstNumber = String(value ?? "").match(/-?[\d,]+(?:\.\d+)?/);
  const parsed = Number(firstNumber?.[0].replace(/,/g, "") ?? "");
  return Number.isFinite(parsed) ? parsed : 0;
};

const toMode = (value) => {
  const normalized = toText(value).toLowerCase().replace(/[\s_-]/g, "");
  if (["mly", "monthly", "month", "mnthly"].includes(normalized)) return "Monthly";
  if (["qly", "qtrly", "quarterly", "quarter", "qtr"].includes(normalized)) {
    return "Quarterly";
  }
  if (
    ["hly", "halfly", "halfyearly", "halfyear", "halfannual", "semiannual"].includes(
      normalized,
    )
  ) {
    return "Half-Yearly";
  }
  if (["yly", "yearly", "annual", "annually", "yrly"].includes(normalized)) {
    return "Yearly";
  }
  return toText(value) || "Monthly";
};

const toStatus = (value) =>
  toText(value).toLowerCase() === "paid" ? "Paid" : "Unpaid";

const toDateInput = (value, XLSX) => {
  if (!value) return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const text = toText(value);
  const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2].padStart(2, "0")}-${isoMatch[3].padStart(2, "0")}`;
  }

  const indianMatch = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (indianMatch) {
    return `${indianMatch[3]}-${indianMatch[2].padStart(2, "0")}-${indianMatch[1].padStart(2, "0")}`;
  }

  const parsedDate = new Date(text);
  return Number.isNaN(parsedDate.getTime()) ? "" : parsedDate.toISOString().slice(0, 10);
};

const rowToPolicy = (row, index, XLSX) => {
  const normalizedRow = {};
  Object.entries(row).forEach(([key, value]) => {
    const field = HEADER_LOOKUP[cleanHeader(key)];
    if (field) normalizedRow[field] = value;
  });

  const clientDetails = [
    toText(normalizedRow.clientName),
    toText(normalizedRow.phoneNumber),
    toText(normalizedRow.address),
  ]
    .filter(Boolean)
    .join("\n");

  return {
    id: globalThis.crypto?.randomUUID?.() ?? `policy-${Date.now()}-${index}`,
    serialNumber: toText(normalizedRow.serialNumber) || String(index + 1),
    clientName: clientDetails,
    dob: toDateInput(normalizedRow.dob, XLSX),
    policyNumber: toText(normalizedRow.policyNumber),
    agentCode: toText(normalizedRow.agentCode),
    commencementDate: toDateInput(normalizedRow.commencementDate, XLSX),
    plan: toText(normalizedRow.plan),
    sumAssured: toAmount(normalizedRow.sumAssured),
    mode: toMode(normalizedRow.mode),
    premium: toAmount(normalizedRow.premium),
    status: toStatus(normalizedRow.status),
  };
};

export async function readPolicyWorkbook(file) {
  const XLSX = await import("xlsx");
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error("This workbook does not contain a worksheet.");
  }

  const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
    defval: "",
    raw: false,
  });

  const policies = rawRows
    .map((row, index) => rowToPolicy(row, index, XLSX))
    .filter((policy) => policy.clientName || policy.policyNumber);

  if (!policies.length) {
    throw new Error("No policy rows were found. Check the column names and try again.");
  }

  return policies;
}

export function mergePolicyCollections(existingPolicies, importedPolicies) {
  const policyNumbers = new Set(
    existingPolicies
      .map((policy) => policy.policyNumber.trim().toLowerCase())
      .filter(Boolean),
  );
  const additions = [];
  let skipped = 0;

  importedPolicies.forEach((policy) => {
    const policyNumber = policy.policyNumber.trim().toLowerCase();
    if (policyNumber && policyNumbers.has(policyNumber)) {
      skipped += 1;
      return;
    }
    if (policyNumber) policyNumbers.add(policyNumber);
    additions.push(policy);
  });

  const policies = [...existingPolicies, ...additions].map((policy, index) => ({
    ...policy,
    serialNumber: String(index + 1),
  }));

  return { policies, added: additions.length, skipped };
}

export async function createPolicyWorkbookFile(policies, agencyLabel = "") {
  const XLSX = await import("xlsx");
  const modeCodes = {
    Monthly: "Mly",
    Quarterly: "Qly",
    "Half-Yearly": "Hly",
    Yearly: "Yly",
  };
  const toDisplayDate = (value) => {
    const text = toText(value);
    const isoDate = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoDate) return `${isoDate[3]}/${isoDate[2]}/${isoDate[1]}`;

    const slashDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashDate) {
      return `${slashDate[1].padStart(2, "0")}/${slashDate[2].padStart(2, "0")}/${slashDate[3]}`;
    }

    return text;
  };
  const toExportName = (value) =>
    toText(value)
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find(Boolean) || "";

  const rows = policies.map((policy) =>
    Object.fromEntries(
      POLICY_FIELDS.map((field) => [
        EXPORT_HEADERS[field],
        field === "clientName"
          ? toExportName(policy[field])
          : field === "premium" || field === "sumAssured"
          ? Number(policy[field]) || 0
          : field === "dob" || field === "commencementDate"
            ? toDisplayDate(policy[field])
          : field === "mode"
            ? modeCodes[policy[field]] || policy[field]
            : policy[field],
      ]),
    ),
  );

  const worksheet = XLSX.utils.json_to_sheet(rows, {
    header: POLICY_FIELDS.map((field) => EXPORT_HEADERS[field]),
  });
  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 42 },
    { wch: 14 },
    { wch: 18 },
    { wch: 14 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 12 },
    { wch: 14 },
    { wch: 12 },
  ];
  worksheet["!autofilter"] = { ref: worksheet["!ref"] };

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Policies");
  const safeAgencyLabel = toText(agencyLabel)
    .replace(/[^a-z0-9]+/gi, "_")
    .replace(/^_+|_+$/g, "");
  const fileName = `LIC_${safeAgencyLabel ? `${safeAgencyLabel}_` : ""}Policies_${new Date().toISOString().slice(0, 10)}.xlsx`;
  const workbookBytes = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
    compression: true,
  });
  return new File([workbookBytes], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function exportPolicyWorkbook(policies, agencyLabel = "") {
  const file = await createPolicyWorkbookFile(policies, agencyLabel);
  const downloadUrl = URL.createObjectURL(file);
  const link = document.createElement("a");
  link.href = downloadUrl;
  link.download = file.name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}
