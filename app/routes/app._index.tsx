import { useEffect, useMemo, useState } from "react";
import type {
  ActionFunctionArgs,
  HeadersFunction,
  LoaderFunctionArgs,
} from "react-router";
import { useFetcher, useLoaderData } from "react-router";
import { authenticate } from "../shopify.server";
import { boundary } from "@shopify/shopify-app-react-router/server";

type CsvRecord = Record<string, string>;

type ShippingAddress = {
  firstName?: string;
  lastName?: string;
  address1?: string;
  address2?: string;
  city?: string;
  provinceCode?: string;
  countryCode?: string;
  zip?: string;
};

type MoneyBagInput = {
  shopMoney: {
    amount: string;
    currencyCode: string;
  };
};

type OrderLineInput = {
  quantity: number;
  title?: string;
  requiresShipping?: boolean;
  priceSet?: MoneyBagInput;
};

type OrderCustomerInput = {
  toAssociate?: {
    id?: string;
  };
};

type OrderCreateInput = {
  email?: string;
  customer?: OrderCustomerInput;
  lineItems: OrderLineInput[];
  note?: string;
  tags?: string[];
  shippingAddress?: ShippingAddress;
};

type OrderDraft = {
  orderKey: string;
  rowNumbers: number[];
  input: OrderCreateInput;
};

type CsvPreviewRow = {
  rowNumber: number;
  recipient: string;
  address: string;
  address2: string;
  city: string;
  state: string;
  zipCode: string;
  email: string;
};

type InvalidRow = CsvPreviewRow & {
  errorMessage: string;
};

type ParseResult = {
  rowCount: number;
  orders: OrderDraft[];
  previewRows: CsvPreviewRow[];
  invalidRows: InvalidRow[];
  errors: string[];
};

type RowResult = CsvPreviewRow & {
  status: "success" | "failed";
  errorMessage: string;
};

type Summary = {
  total: number;
  success: number;
  failed: number;
  ordersCreated: number;
};

type ImportResult = {
  summary: Summary | null;
  results: RowResult[];
  error: string;
};

type CreateActionData =
  | {
      error: string;
    }
  | {
      summary: Summary;
      results: RowResult[];
    };

type ShippingReportOrder = {
  id: string;
  name: string;
  customerName: string;
  trackingNumbers: string[];
};

type ReportActionData =
  | {
      error: string;
    }
  | {
      reportOrders: ShippingReportOrder[];
      warning?: string;
    };

type ActionPayload = {
  rowCount: number;
  orders: OrderDraft[];
  previewRows: CsvPreviewRow[];
  invalidRows: InvalidRow[];
};

type NormalizedRow = {
  email?: string;
  note?: string;
  shippingAddress: ShippingAddress;
};

type OrderCreateResponse =
  | {
      ok: true;
      id: string;
      name: string;
    }
  | {
      ok: false;
      message: string;
    };

type CustomerOption = {
  id: string;
  displayName: string;
  email: string;
};

const REQUIRED_COLUMNS = [
  "first_name",
  "last_name",
  "address",
  "address2",
  "city",
  "state",
  "zip_code",
];

const DEFAULT_COUNTRY_CODE = "US";
const DEFAULT_LINE_ITEM_TITLE = "Corporate Gift";
const DEFAULT_LINE_ITEM_PRICE = "0.00";
const DEFAULT_LINE_ITEM_CURRENCY = "USD";
const DEFAULT_LINE_ITEM_QUANTITY = 1;

const ORDER_CREATE_MUTATION = `#graphql
  mutation OrderCreate($order: OrderCreateOrderInput!) {
    orderCreate(order: $order) {
      order {
        id
        name
      }
      userErrors {
        field
        message
      }
    }
  }
`;

const CUSTOMERS_QUERY = `#graphql
  query CustomersForSelect {
    customers(first: 250, sortKey: NAME) {
      edges {
        node {
          id
          displayName
          email
        }
      }
    }
  }
`;

const SHIPPING_REPORT_QUERY = `#graphql
  query ShippingReportOrders($query: String) {
    orders(first: 250, query: $query) {
      edges {
        node {
          id
          name
          customer {
            displayName
          }
          shippingAddress {
            firstName
            lastName
          }
          fulfillments(first: 20) {
            edges {
              node {
                trackingInfo {
                  number
                }
              }
            }
          }
        }
      }
    }
  }
`;

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);

  const customers: CustomerOption[] = [];
  let customerLoadWarning = "";

  try {
    const response = await admin.graphql(CUSTOMERS_QUERY);
    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        customers?: {
          edges?: Array<{
            node?: {
              id?: string;
              displayName?: string;
              email?: string | null;
            };
          }>;
        };
      };
    };

    if (json.errors && json.errors.length > 0) {
      const message = json.errors.map((error) => error.message).join("; ");
      if (message.toLowerCase().includes("access denied")) {
        customerLoadWarning =
          "Customer selector unavailable. Add read_customers scope and redeploy.";
      } else {
        customerLoadWarning = `Unable to load customers: ${message}`;
      }
    } else {
      for (const edge of json.data?.customers?.edges ?? []) {
        const node = edge.node;
        if (!node?.id) continue;

        customers.push({
          id: node.id,
          displayName: node.displayName ?? "Unnamed customer",
          email: node.email ?? "",
        });
      }
    }
  } catch {
    customerLoadWarning =
      "Unable to load customers right now. Customer selector may be incomplete.";
  }

  return { customers, customerLoadWarning };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const { admin } = await authenticate.admin(request);
  const formData = await request.formData();
  const intent = formData.get("intent");

  if (intent === "generate_report") {
    return generateShippingReport(admin, formData);
  }

  const payloadText = formData.get("payload");

  if (typeof payloadText !== "string") {
    return { error: "Invalid upload payload." } satisfies CreateActionData;
  }

  let payload: ActionPayload;
  try {
    payload = JSON.parse(payloadText) as ActionPayload;
  } catch {
    return { error: "Failed to parse uploaded payload." } satisfies CreateActionData;
  }

  if (!Array.isArray(payload.orders) || !Array.isArray(payload.previewRows)) {
    return { error: "Invalid order payload format." } satisfies CreateActionData;
  }

  const rowLookup = new Map<number, CsvPreviewRow>(
    payload.previewRows.map((row) => [row.rowNumber, row]),
  );

  const processedResults: RowResult[] = (payload.invalidRows ?? []).map((row) => ({
    ...row,
    status: "failed",
  }));

  let successCount = 0;
  let failedCount = payload.invalidRows?.length ?? 0;
  let ordersCreatedCount = 0;

  for (const order of payload.orders) {
    const created = await createOrder(admin, order.input);

    for (const rowNumber of order.rowNumbers) {
      const row = rowLookup.get(rowNumber);
      if (!row) continue;

      if (created.ok) {
        processedResults.push({
          ...row,
          status: "success",
          errorMessage: "",
        });
        successCount += 1;
      } else {
        processedResults.push({
          ...row,
          status: "failed",
          errorMessage: created.message,
        });
        failedCount += 1;
      }
    }

    if (created.ok) {
      ordersCreatedCount += 1;
    }
  }

  processedResults.sort((a, b) => a.rowNumber - b.rowNumber);

  return {
    summary: {
      total: payload.rowCount ?? processedResults.length,
      success: successCount,
      failed: failedCount,
      ordersCreated: ordersCreatedCount,
    },
    results: processedResults,
  } satisfies CreateActionData;
};

export default function Index() {
  const { customers, customerLoadWarning } = useLoaderData<typeof loader>();
  const createFetcher = useFetcher<CreateActionData>();
  const reportFetcher = useFetcher<ReportActionData>();

  const [fileName, setFileName] = useState("");
  const [selectedCustomerId, setSelectedCustomerId] = useState("");
  const [orderTagsInput, setOrderTagsInput] = useState("");
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [parseNotice, setParseNotice] = useState("");
  const [importResult, setImportResult] = useState<ImportResult | null>(null);

  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [orderNumbersInput, setOrderNumbersInput] = useState("");
  const [searchTagsInput, setSearchTagsInput] = useState("");
  const [reportOrders, setReportOrders] = useState<ShippingReportOrder[]>([]);
  const [reportNotice, setReportNotice] = useState("");
  const [reportError, setReportError] = useState("");

  const isCreating = createFetcher.state === "submitting";
  const isGeneratingReport = reportFetcher.state === "submitting";
  const canCreate =
    !isCreating && Boolean(parseResult) && (parseResult?.orders.length ?? 0) > 0;

  useEffect(() => {
    if (!createFetcher.data) return;

    if ("error" in createFetcher.data) {
      setImportResult({
        summary: null,
        results: [],
        error: createFetcher.data.error,
      });
      return;
    }

    setImportResult({
      summary: createFetcher.data.summary,
      results: createFetcher.data.results,
      error: "",
    });
  }, [createFetcher.data]);

  useEffect(() => {
    if (!reportFetcher.data) return;

    if ("error" in reportFetcher.data) {
      setReportError(reportFetcher.data.error);
      setReportNotice("");
      setReportOrders([]);
      return;
    }

    setReportError("");
    setReportOrders(reportFetcher.data.reportOrders);
    setReportNotice(
      reportFetcher.data.warning ||
        `Found ${reportFetcher.data.reportOrders.length} order(s) for report.`,
    );
  }, [reportFetcher.data]);

  const previewRows = useMemo(
    () => parseResult?.previewRows.slice(0, 5) ?? [],
    [parseResult],
  );

  const handleSelectedFile = async (file: File) => {
    setFileName(file.name);
    setParseResult(null);
    setParseNotice("");
    setImportResult(null);

    const content = await file.text();
    const parsed = parseCsvToOrders(content);
    setParseResult(parsed);

    if (parsed.orders.length === 0 && parsed.errors.length > 0) {
      setParseNotice(parsed.errors[0]);
      return;
    }

    if (parsed.errors.length > 0) {
      setParseNotice(
        `Loaded with ${parsed.errors.length} validation issue(s). ${parsed.orders.length} order(s) are still ready.`,
      );
      return;
    }

    setParseNotice("");
  };

  const onDropZoneInput = async (event: Event) => {
    const target = event.currentTarget as (HTMLElement & { files?: File[] }) | null;
    const file = target?.files?.[0];

    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".csv")) {
      setParseNotice("Please upload a .csv file.");
      return;
    }

    await handleSelectedFile(file);
  };

  const onDropZoneRejected = () => {
    setParseNotice("Please upload a .csv file.");
  };

  const onCreateOrders = () => {
    if (!parseResult || !canCreate) return;

    const globalTags = parseTags(orderTagsInput);

    const ordersWithSettings = parseResult.orders.map((order) => ({
      ...order,
      input: {
        ...order.input,
        customer: selectedCustomerId
          ? {
              toAssociate: {
                id: selectedCustomerId,
              },
            }
          : undefined,
        tags: mergeTags(order.input.tags, globalTags),
      },
    }));

    const formData = new FormData();
    formData.append(
      "payload",
      JSON.stringify({
        rowCount: parseResult.rowCount,
        orders: ordersWithSettings,
        previewRows: parseResult.previewRows,
        invalidRows: parseResult.invalidRows,
      } satisfies ActionPayload),
    );

    createFetcher.submit(formData, { method: "POST" });
  };

  const onGenerateReport = () => {
    const formData = new FormData();
    formData.append("intent", "generate_report");
    formData.append("orderNumbers", orderNumbersInput);
    formData.append("startDate", startDate);
    formData.append("endDate", endDate);
    formData.append("searchTags", searchTagsInput);

    reportFetcher.submit(formData, { method: "POST" });
  };

  const onExportReport = () => {
    if (reportOrders.length === 0) return;
    const csv = buildShippingReportCsv(reportOrders);
    const fileNameDate = new Date().toISOString().slice(0, 10);
    downloadCsv(`shipping-report-${fileNameDate}.csv`, csv);
  };

  const summary = importResult?.summary;
  const results = importResult?.results ?? [];

  return (
    <s-page heading="CSV Order Importer">
      {importResult?.error ? (
        <s-banner tone="critical">{importResult.error}</s-banner>
      ) : null}

      {parseNotice ? <s-banner tone="warning">{parseNotice}</s-banner> : null}

      {customerLoadWarning ? (
        <s-banner tone="warning">{customerLoadWarning}</s-banner>
      ) : null}

      {reportError ? <s-banner tone="critical">{reportError}</s-banner> : null}

      {reportNotice ? <s-banner tone="success">{reportNotice}</s-banner> : null}

      {summary ? (
        summary.failed === 0 ? (
          <s-banner tone="success">
            Processed {summary.success} rows and created {summary.ordersCreated}{" "}
            orders.
          </s-banner>
        ) : (
          <s-banner tone="warning">
            Processed {summary.total} rows: {summary.success} succeeded,{" "}
            {summary.failed} failed.
          </s-banner>
        )
      ) : null}

      <s-section heading="Upload CSV File">
        <s-stack direction="block" gap="base">
          <s-text color="subdued">
            Upload a CSV file with columns: first_name, last_name, address, address2,
            city, state, zip_code
          </s-text>

          <s-drop-zone
            label="Upload CSV"
            accept=".csv,text/csv"
            onInput={onDropZoneInput}
            onDropRejected={onDropZoneRejected}
          />

          {fileName ? (
            <s-text>
              <s-text type="strong">Selected file:</s-text> {fileName}
            </s-text>
          ) : null}
        </s-stack>
      </s-section>

      <s-section heading="Order Settings">
        <s-stack direction="block" gap="base">
          <s-select
            label="Customer (optional)"
            value={selectedCustomerId}
            onInput={(event: Event) => {
              const target = event.currentTarget as HTMLElement & { value?: string };
              setSelectedCustomerId(target.value || "");
            }}
          >
            <s-option value="">No customer</s-option>
            {customers.map((customer) => (
              <s-option key={customer.id} value={customer.id}>
                {customer.displayName}
                {customer.email ? ` (${customer.email})` : ""}
              </s-option>
            ))}
          </s-select>
          <s-text color="subdued">
            Assign all imported orders to one customer while shipping to each
            recipient.
          </s-text>

          <s-text-field
            label="Order Tags (optional)"
            placeholder="e.g., imported, bulk-order"
            value={orderTagsInput}
            onInput={(event: Event) => {
              const target = event.currentTarget as HTMLElement & { value?: string };
              setOrderTagsInput(target.value || "");
            }}
          />
          <s-text color="subdued">
            Add tags to all imported orders. Separate multiple tags with commas.
          </s-text>
        </s-stack>
      </s-section>

      {parseResult ? (
        <s-section heading="Parsed CSV Data" padding="none">
          <s-box padding="base">
            <s-stack gap="base" direction="block">
              <s-text color="subdued">
                {parseResult.rowCount} rows ready to import
              </s-text>
              <s-button
                variant="primary"
                onClick={onCreateOrders}
                {...(isCreating ? { loading: true } : {})}
                disabled={!canCreate}
              >
                Create Orders
              </s-button>
            </s-stack>
          </s-box>

          <s-table>
            <s-table-header-row>
              <s-table-header listSlot="kicker">Row</s-table-header>
              <s-table-header listSlot="primary">First Name</s-table-header>
              <s-table-header listSlot="labeled">Last Name</s-table-header>
              <s-table-header listSlot="labeled">Address</s-table-header>
              <s-table-header listSlot="labeled">City</s-table-header>
              <s-table-header listSlot="labeled">State</s-table-header>
              <s-table-header listSlot="labeled">Zip Code</s-table-header>
            </s-table-header-row>
            <s-table-body>
              {previewRows.map((row) => {
                const [firstName, ...lastParts] = row.recipient.split(" ");
                return (
                  <s-table-row key={row.rowNumber}>
                    <s-table-cell>{row.rowNumber}</s-table-cell>
                    <s-table-cell>{firstName || "-"}</s-table-cell>
                    <s-table-cell>{lastParts.join(" ") || "-"}</s-table-cell>
                    <s-table-cell>{row.address || "-"}</s-table-cell>
                    <s-table-cell>{row.city || "-"}</s-table-cell>
                    <s-table-cell>{row.state || "-"}</s-table-cell>
                    <s-table-cell>{row.zipCode || "-"}</s-table-cell>
                  </s-table-row>
                );
              })}
            </s-table-body>
          </s-table>
        </s-section>
      ) : null}

      {results.length > 0 ? (
        <s-section heading="Created Orders" padding="none">
          <s-table>
            <s-table-header-row>
              <s-table-header listSlot="kicker">Row</s-table-header>
              <s-table-header listSlot="primary">Recipient</s-table-header>
              <s-table-header listSlot="labeled">Status</s-table-header>
              <s-table-header listSlot="labeled">Error</s-table-header>
            </s-table-header-row>
            <s-table-body>
              {results.map((row) => (
                <s-table-row key={row.rowNumber}>
                  <s-table-cell>{row.rowNumber}</s-table-cell>
                  <s-table-cell>{row.recipient}</s-table-cell>
                  <s-table-cell>
                    <s-badge
                      tone={row.status === "success" ? "success" : "critical"}
                    >
                      {row.status === "success" ? "Created" : "Failed"}
                    </s-badge>
                  </s-table-cell>
                  <s-table-cell>{row.errorMessage || "-"}</s-table-cell>
                </s-table-row>
              ))}
            </s-table-body>
          </s-table>
        </s-section>
      ) : null}

      <s-section heading="Shipping Report">
        <s-stack gap="base" direction="block">
          <s-text color="subdued">
            Search by order numbers or date range. You can also filter by tags and
            export the results.
          </s-text>

          <s-text-field
            label="Order Numbers (optional)"
            placeholder="e.g., #1001, #1002"
            value={orderNumbersInput}
            onInput={(event: Event) => {
              const target = event.currentTarget as HTMLElement & { value?: string };
              setOrderNumbersInput(target.value || "");
            }}
          />

          <s-stack direction="inline" gap="base">
            <s-date-field
              label="Start Date (optional)"
              value={startDate}
              onInput={(event: Event) => {
                const target = event.currentTarget as HTMLElement & { value?: string };
                setStartDate(target.value || "");
              }}
            />
            <s-date-field
              label="End Date (optional)"
              value={endDate}
              onInput={(event: Event) => {
                const target = event.currentTarget as HTMLElement & { value?: string };
                setEndDate(target.value || "");
              }}
            />
          </s-stack>

          <s-text-field
            label="Tags (optional)"
            placeholder="e.g., imported, batch:sonya"
            value={searchTagsInput}
            onInput={(event: Event) => {
              const target = event.currentTarget as HTMLElement & { value?: string };
              setSearchTagsInput(target.value || "");
            }}
          />

          <s-stack direction="inline" gap="base">
            <s-button
              variant="primary"
              onClick={onGenerateReport}
              {...(isGeneratingReport ? { loading: true } : {})}
            >
              Generate Report
            </s-button>
            <s-button
              onClick={onExportReport}
              disabled={reportOrders.length === 0}
            >
              Export to CSV
            </s-button>
          </s-stack>
        </s-stack>
      </s-section>

      {reportOrders.length > 0 ? (
        <s-section heading="Shipping Report Results" padding="none">
          <s-table>
            <s-table-header-row>
              <s-table-header listSlot="primary">Order Number</s-table-header>
              <s-table-header listSlot="labeled">Customer Name</s-table-header>
              <s-table-header listSlot="labeled">Tracking Numbers</s-table-header>
            </s-table-header-row>
            <s-table-body>
              {reportOrders.map((order) => (
                <s-table-row key={order.id}>
                  <s-table-cell>{order.name}</s-table-cell>
                  <s-table-cell>{order.customerName}</s-table-cell>
                  <s-table-cell>
                    {order.trackingNumbers.length > 0
                      ? order.trackingNumbers.join(", ")
                      : "No tracking"}
                  </s-table-cell>
                </s-table-row>
              ))}
            </s-table-body>
          </s-table>
        </s-section>
      ) : null}
    </s-page>
  );
}

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};

function parseCsvToOrders(csvText: string): ParseResult {
  const rows = parseCsvRows(csvText);
  if (rows.length < 2) {
    return {
      rowCount: 0,
      orders: [],
      previewRows: [],
      invalidRows: [],
      errors: ["CSV must include a header row and at least one data row."],
    };
  }

  const headers = rows[0].map((header) => normalizeHeader(header));
  const missingHeaders = REQUIRED_COLUMNS.filter(
    (header) => !headers.includes(header),
  );

  if (missingHeaders.length > 0) {
    return {
      rowCount: rows.length - 1,
      orders: [],
      previewRows: [],
      invalidRows: [],
      errors: [`Missing required columns: ${missingHeaders.join(", ")}`],
    };
  }

  const errors: string[] = [];
  const invalidRows: InvalidRow[] = [];
  const previewRows: CsvPreviewRow[] = [];
  const orders: OrderDraft[] = [];

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const csvRow = toRecord(headers, rows[rowIndex]);
    if (isEmptyRecord(csvRow)) continue;

    const rowNumber = rowIndex + 1;
    const previewRow = buildPreviewRow(csvRow, rowNumber);
    previewRows.push(previewRow);

    const parsed = normalizeRow(csvRow, rowNumber);
    if (!parsed.ok) {
      const errorMessage = parsed.errors.join(" | ");
      errors.push(...parsed.errors);
      invalidRows.push({ ...previewRow, errorMessage });
      continue;
    }

    orders.push({
      orderKey: `row-${rowNumber}`,
      rowNumbers: [rowNumber],
      input: {
        email: parsed.value.email,
        note: parsed.value.note,
        shippingAddress: parsed.value.shippingAddress,
        lineItems: [
          {
            title: DEFAULT_LINE_ITEM_TITLE,
            quantity: DEFAULT_LINE_ITEM_QUANTITY,
            requiresShipping: true,
            priceSet: {
              shopMoney: {
                amount: DEFAULT_LINE_ITEM_PRICE,
                currencyCode: DEFAULT_LINE_ITEM_CURRENCY,
              },
            },
          },
        ],
      },
    });
  }

  if (previewRows.length === 0) {
    errors.push("No non-empty data rows found.");
  }

  return {
    rowCount: previewRows.length,
    orders,
    previewRows,
    invalidRows,
    errors,
  };
}

function parseCsvRows(csvText: string): string[][] {
  const text = csvText.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === "\"") {
      if (inQuotes && next === "\"") {
        field += "\"";
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(field);
      field = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(field);
      field = "";
      if (!isEmptyArrayRow(row)) {
        rows.push(row);
      }
      row = [];
      continue;
    }

    field += char;
  }

  row.push(field);
  if (!isEmptyArrayRow(row)) {
    rows.push(row);
  }

  return rows;
}

function normalizeHeader(value: string): string {
  return value.trim().toLowerCase();
}

function toRecord(headers: string[], row: string[]): CsvRecord {
  const record: CsvRecord = {};
  headers.forEach((header, index) => {
    record[header] = row[index] ?? "";
  });
  return record;
}

function isEmptyArrayRow(row: string[]): boolean {
  return row.every((cell) => cell.trim() === "");
}

function isEmptyRecord(record: CsvRecord): boolean {
  return Object.values(record).every((value) => value.trim() === "");
}

function readColumn(record: CsvRecord, key: string): string {
  return (record[key] ?? "").trim();
}

function buildPreviewRow(record: CsvRecord, rowNumber: number): CsvPreviewRow {
  const firstName = readColumn(record, "first_name");
  const lastName = readColumn(record, "last_name");

  return {
    rowNumber,
    recipient: [firstName, lastName].filter(Boolean).join(" "),
    address: readColumn(record, "address"),
    address2: readColumn(record, "address2"),
    city: readColumn(record, "city"),
    state: readColumn(record, "state"),
    zipCode: readColumn(record, "zip_code"),
    email: readColumn(record, "email"),
  };
}

function parseTags(value: string): string[] | undefined {
  if (!value) return undefined;

  const tags = value
    .split(",")
    .map((tag) => tag.trim())
    .filter(Boolean);

  return tags.length > 0 ? tags : undefined;
}

function mergeTags(...tagGroups: Array<string[] | undefined>): string[] | undefined {
  const merged = new Set<string>();

  for (const group of tagGroups) {
    if (!group) continue;
    for (const tag of group) {
      const trimmed = tag.trim();
      if (!trimmed) continue;
      merged.add(trimmed);
    }
  }

  return merged.size > 0 ? Array.from(merged) : undefined;
}

function normalizeRow(
  record: CsvRecord,
  rowNumber: number,
):
  | {
      ok: true;
      value: NormalizedRow;
    }
  | {
      ok: false;
      errors: string[];
    } {
  const rowErrors: string[] = [];

  const firstName = readColumn(record, "first_name");
  const lastName = readColumn(record, "last_name");
  const address1 = readColumn(record, "address");
  const address2 = readColumn(record, "address2") || undefined;
  const city = readColumn(record, "city");
  const state = readColumn(record, "state");
  const zip = readColumn(record, "zip_code");
  const countryCode =
    readColumn(record, "country_code").toUpperCase() || DEFAULT_COUNTRY_CODE;
  const email = readColumn(record, "email") || undefined;
  const note = readColumn(record, "note") || undefined;

  if (!firstName) rowErrors.push(`Row ${rowNumber}: first_name is required.`);
  if (!lastName) rowErrors.push(`Row ${rowNumber}: last_name is required.`);
  if (!address1) rowErrors.push(`Row ${rowNumber}: address is required.`);
  if (!city) rowErrors.push(`Row ${rowNumber}: city is required.`);
  if (!state) rowErrors.push(`Row ${rowNumber}: state is required.`);
  if (!zip) rowErrors.push(`Row ${rowNumber}: zip_code is required.`);

  if (rowErrors.length > 0) {
    return { ok: false, errors: rowErrors };
  }

  return {
    ok: true,
    value: {
      email,
      note,
      shippingAddress: {
        firstName,
        lastName,
        address1,
        address2,
        city,
        provinceCode: state,
        countryCode,
        zip,
      },
    },
  };
}

function compactOrderInput(input: OrderCreateInput): OrderCreateInput {
  const compacted: OrderCreateInput = {
    lineItems: input.lineItems,
  };

  if (input.email) compacted.email = input.email;
  if (input.customer) compacted.customer = input.customer;
  if (input.note) compacted.note = input.note;
  if (input.tags && input.tags.length > 0) compacted.tags = input.tags;
  if (input.shippingAddress) compacted.shippingAddress = input.shippingAddress;

  return compacted;
}

async function createOrder(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  input: OrderCreateInput,
): Promise<OrderCreateResponse> {
  try {
    const response = await admin.graphql(ORDER_CREATE_MUTATION, {
      variables: { order: compactOrderInput(input) },
    });

    const json = await response.json();
    const graphQLErrors = (json as { errors?: { message: string }[] }).errors;
    if (graphQLErrors && graphQLErrors.length > 0) {
      return {
        ok: false,
        message: graphQLErrors.map((error) => error.message).join("; "),
      };
    }

    const payload = (json as {
      data?: {
        orderCreate?: {
          order?: { id: string; name: string };
          userErrors?: { message: string }[];
        };
      };
    }).data?.orderCreate;

    const userErrors = payload?.userErrors ?? [];
    if (userErrors.length > 0) {
      return {
        ok: false,
        message: userErrors.map((error) => error.message).join("; "),
      };
    }

    if (!payload?.order?.id) {
      return { ok: false, message: "Unknown orderCreate response." };
    }

    return { ok: true, id: payload.order.id, name: payload.order.name };
  } catch (error) {
    return {
      ok: false,
      message:
        error instanceof Error
          ? error.message
          : "Unexpected error while creating order.",
    };
  }
}

async function generateShippingReport(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  formData: FormData,
): Promise<ReportActionData> {
  const orderNumbers = (formData.get("orderNumbers") as string | null)?.trim() || "";
  const startDate = (formData.get("startDate") as string | null)?.trim() || "";
  const endDate = (formData.get("endDate") as string | null)?.trim() || "";
  const searchTags = (formData.get("searchTags") as string | null)?.trim() || "";

  const query = buildOrderSearchQuery({
    orderNumbers,
    startDate,
    endDate,
    searchTags,
  });

  try {
    const response = await admin.graphql(SHIPPING_REPORT_QUERY, {
      variables: {
        query: query || null,
      },
    });

    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        orders?: {
          edges?: Array<{
            node?: {
              id?: string;
              name?: string;
              customer?: {
                displayName?: string;
              } | null;
              shippingAddress?: {
                firstName?: string;
                lastName?: string;
              } | null;
              fulfillments?: {
                edges?: Array<{
                  node?: {
                    trackingInfo?: Array<{
                      number?: string | null;
                    }>;
                  };
                }>;
              };
            };
          }>;
        };
      };
    };

    if (json.errors && json.errors.length > 0) {
      return {
        error: json.errors.map((error) => error.message).join("; "),
      };
    }

    const reportOrders: ShippingReportOrder[] = [];

    for (const edge of json.data?.orders?.edges ?? []) {
      const order = edge.node;
      if (!order?.id || !order.name) continue;

      const recipientName = [
        order.shippingAddress?.firstName?.trim() ?? "",
        order.shippingAddress?.lastName?.trim() ?? "",
      ]
        .filter(Boolean)
        .join(" ");

      const fallbackName = recipientName || "No customer";

      const trackingNumbers: string[] = [];
      for (const fulfillmentEdge of order.fulfillments?.edges ?? []) {
        for (const info of fulfillmentEdge.node?.trackingInfo ?? []) {
          const number = (info.number || "").trim();
          if (number) trackingNumbers.push(number);
        }
      }

      reportOrders.push({
        id: order.id,
        name: order.name,
        customerName: order.customer?.displayName || fallbackName,
        trackingNumbers,
      });
    }

    if (reportOrders.length === 0) {
      return {
        reportOrders: [],
        warning: "No orders found matching the report criteria.",
      };
    }

    return { reportOrders };
  } catch (error) {
    return {
      error:
        error instanceof Error
          ? error.message
          : "Failed to generate shipping report.",
    };
  }
}

function buildOrderSearchQuery(filters: {
  orderNumbers: string;
  startDate: string;
  endDate: string;
  searchTags: string;
}): string {
  const queryParts: string[] = [];

  const orderNumbers = splitCommaValues(filters.orderNumbers);
  const tags = splitCommaValues(filters.searchTags);

  if (orderNumbers.length > 0) {
    queryParts.push(
      `(${orderNumbers
        .map((number) => `name:'${escapeSearchValue(number)}'`)
        .join(" OR ")})`,
    );
  } else {
    if (filters.startDate) {
      queryParts.push(`created_at:>=${escapeSearchValue(filters.startDate)}`);
    }
    if (filters.endDate) {
      queryParts.push(`created_at:<=${escapeSearchValue(filters.endDate)}`);
    }
  }

  if (tags.length > 0) {
    queryParts.push(
      `(${tags.map((tag) => `tag:'${escapeSearchValue(tag)}'`).join(" OR ")})`,
    );
  }

  return queryParts.join(" AND ");
}

function splitCommaValues(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function escapeSearchValue(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

function buildShippingReportCsv(orders: ShippingReportOrder[]): string {
  const header = ["Order Number", "Customer Name", "Tracking Numbers"];
  const rows = orders.map((order) => [
    order.name,
    order.customerName,
    order.trackingNumbers.join("; ") || "No tracking",
  ]);

  return [header, ...rows]
    .map((row) => row.map(csvEscape).join(","))
    .join("\n");
}

function csvEscape(value: string): string {
  const normalized = value.replace(/\r?\n/g, " ");
  return `"${normalized.replace(/"/g, '""')}"`;
}

function downloadCsv(fileName: string, content: string): void {
  if (typeof window === "undefined") return;

  const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.URL.revokeObjectURL(url);
}
