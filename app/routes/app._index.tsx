import { useEffect, useMemo, useRef, useState } from "react";
import type { ChangeEvent, DragEvent } from "react";
import type {
  ActionFunctionArgs,
  HeadersFunction,
  LoaderFunctionArgs,
} from "react-router";
import { useFetcher } from "react-router";
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
  variantId?: string;
  title?: string;
  requiresShipping?: boolean;
  priceSet?: MoneyBagInput;
};

type OrderCreateInput = {
  email?: string;
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

type ActionData =
  | {
      error: string;
    }
  | {
      summary: Summary;
      results: RowResult[];
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

export const loader = async ({ request }: LoaderFunctionArgs) => {
  await authenticate.admin(request);
  return null;
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const { admin } = await authenticate.admin(request);
  const formData = await request.formData();
  const payloadText = formData.get("payload");

  if (typeof payloadText !== "string") {
    return { error: "Invalid upload payload." } satisfies ActionData;
  }

  let payload: ActionPayload;
  try {
    payload = JSON.parse(payloadText) as ActionPayload;
  } catch {
    return { error: "Failed to parse uploaded payload." } satisfies ActionData;
  }

  if (!Array.isArray(payload.orders) || !Array.isArray(payload.previewRows)) {
    return { error: "Invalid order payload format." } satisfies ActionData;
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
  } satisfies ActionData;
};

export default function Index() {
  const fetcher = useFetcher<ActionData>();
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [fileName, setFileName] = useState("");
  const [orderTagsInput, setOrderTagsInput] = useState("");
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [parseNotice, setParseNotice] = useState("");
  const [importResult, setImportResult] = useState<ImportResult | null>(null);

  const isCreating = fetcher.state === "submitting";
  const canCreate =
    !isCreating && Boolean(parseResult) && (parseResult?.orders.length ?? 0) > 0;

  useEffect(() => {
    if (!fetcher.data) return;

    if ("error" in fetcher.data) {
      setImportResult({
        summary: null,
        results: [],
        error: fetcher.data.error,
      });
      return;
    }

    setImportResult({
      summary: fetcher.data.summary,
      results: fetcher.data.results,
      error: "",
    });
  }, [fetcher.data]);

  const previewRows = useMemo(
    () => parseResult?.previewRows.slice(0, 5) ?? [],
    [parseResult],
  );

  const resetAll = () => {
    setFileName("");
    setParseResult(null);
    setParseNotice("");
    setImportResult(null);
  };

  const onOpenFilePicker = () => {
    fileInputRef.current?.click();
  };

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

  const onFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    await handleSelectedFile(file);
  };

  const onDrop = async (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".csv")) {
      setParseNotice("Please upload a .csv file.");
      return;
    }
    await handleSelectedFile(file);
  };

  const onCreateOrders = () => {
    if (!parseResult || !canCreate) return;

    const globalTags = parseTags(orderTagsInput);
    const ordersWithTags = parseResult.orders.map((order) => ({
      ...order,
      input: {
        ...order.input,
        tags: globalTags,
      },
    }));

    const formData = new FormData();
    formData.append(
      "payload",
      JSON.stringify({
        rowCount: parseResult.rowCount,
        orders: ordersWithTags,
        previewRows: parseResult.previewRows,
        invalidRows: parseResult.invalidRows,
      } satisfies ActionPayload),
    );

    fetcher.submit(formData, { method: "POST" });
  };

  const summary = importResult?.summary;
  const results = importResult?.results ?? [];

  return (
    <s-page heading="Corporate Addresses CSV Import" inlineSize="large">
      {importResult?.error ? (
        <s-banner tone="critical">{importResult.error}</s-banner>
      ) : null}

      {parseNotice ? <s-banner tone="warning">{parseNotice}</s-banner> : null}

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

      {results.length === 0 ? (
        <>
          <s-section heading="Upload CSV File">
            <s-text>
              Upload a CSV file with columns: <code>first_name</code>,{" "}
              <code>last_name</code>, <code>address</code>, <code>address2</code>,{" "}
              <code>city</code>, <code>state</code>, <code>zip_code</code>.
            </s-text>

            <input
              ref={fileInputRef}
              type="file"
              accept=".csv,text/csv"
              onChange={onFileChange}
              style={{ display: "none" }}
            />

            <div
              onDragOver={(event) => event.preventDefault()}
              onDrop={onDrop}
              style={{
                marginTop: "16px",
                border: "1px dashed var(--s-border-color)",
                borderRadius: "12px",
                padding: "28px",
                textAlign: "center",
                background: "var(--s-surface-subdued)",
              }}
            >
              <s-stack direction="block" gap="base" alignItems="center">
                <s-button variant="primary" onClick={onOpenFilePicker}>
                  Add files
                </s-button>
                <s-text>
                  {fileName ? `Selected file: ${fileName}` : "Drop CSV file here"}
                </s-text>
                <s-text color="subdued">Accepts .csv</s-text>
              </s-stack>
            </div>

            {parseResult ? (
              <>
                <s-paragraph>
                  {parseResult.rowCount} rows found. {parseResult.orders.length} order(s)
                  ready.
                </s-paragraph>

                <s-section heading="Preview (first 5 rows)">
                  <s-table>
                    <s-table-header-row>
                      <s-table-header listSlot="kicker">Row</s-table-header>
                      <s-table-header listSlot="primary">Recipient</s-table-header>
                      <s-table-header listSlot="labeled">Address</s-table-header>
                      <s-table-header listSlot="labeled">Address 2</s-table-header>
                      <s-table-header listSlot="labeled">City</s-table-header>
                      <s-table-header listSlot="labeled">State</s-table-header>
                      <s-table-header listSlot="labeled">ZIP</s-table-header>
                    </s-table-header-row>
                    <s-table-body>
                      {previewRows.map((row) => (
                        <s-table-row key={row.rowNumber}>
                          <s-table-cell>{row.rowNumber}</s-table-cell>
                          <s-table-cell>{row.recipient || "-"}</s-table-cell>
                          <s-table-cell>{row.address || "-"}</s-table-cell>
                          <s-table-cell>{row.address2 || "-"}</s-table-cell>
                          <s-table-cell>{row.city || "-"}</s-table-cell>
                          <s-table-cell>{row.state || "-"}</s-table-cell>
                          <s-table-cell>{row.zipCode || "-"}</s-table-cell>
                        </s-table-row>
                      ))}
                    </s-table-body>
                  </s-table>
                </s-section>

                <s-stack direction="inline" gap="base" justifyContent="end">
                  <s-button onClick={resetAll}>Reset</s-button>
                  <s-button
                    variant="primary"
                    onClick={onCreateOrders}
                    disabled={!canCreate}
                    {...(isCreating ? { loading: true } : {})}
                  >
                    Create Orders
                  </s-button>
                </s-stack>
              </>
            ) : null}
          </s-section>

          <s-section heading="Order Settings">
            <s-text-field
              label="Order Tags (optional)"
              value={orderTagsInput}
              onChange={(event) => setOrderTagsInput(event.currentTarget.value)}
            ></s-text-field>
            <s-text color="subdued">
              Add tags to all imported orders. Separate multiple tags with commas.
            </s-text>
          </s-section>
        </>
      ) : (
        <s-section heading="Results">
          <s-table>
            <s-table-header-row>
              <s-table-header listSlot="kicker">Row</s-table-header>
              <s-table-header listSlot="primary">Recipient</s-table-header>
              <s-table-header listSlot="labeled">Address</s-table-header>
              <s-table-header listSlot="labeled">City</s-table-header>
              <s-table-header listSlot="labeled">State</s-table-header>
              <s-table-header listSlot="labeled">ZIP</s-table-header>
              <s-table-header listSlot="inline">Status</s-table-header>
              <s-table-header listSlot="labeled">Error</s-table-header>
            </s-table-header-row>
            <s-table-body>
              {results.map((row) => (
                <s-table-row key={row.rowNumber}>
                  <s-table-cell>{row.rowNumber}</s-table-cell>
                  <s-table-cell>{row.recipient || "-"}</s-table-cell>
                  <s-table-cell>{row.address || "-"}</s-table-cell>
                  <s-table-cell>{row.city || "-"}</s-table-cell>
                  <s-table-cell>{row.state || "-"}</s-table-cell>
                  <s-table-cell>{row.zipCode || "-"}</s-table-cell>
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

          <s-stack direction="inline" justifyContent="end">
            <s-button variant="primary" onClick={resetAll}>
              Upload New CSV
            </s-button>
          </s-stack>
        </s-section>
      )}
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
