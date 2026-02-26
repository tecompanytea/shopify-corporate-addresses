import { render } from "preact";
import { useRef, useState } from "preact/hooks";

type CsvRecord = Record<string, string>;

type ShippingAddress = {
  firstName?: string;
  lastName?: string;
  address1?: string;
  city?: string;
  provinceCode?: string;
  countryCode?: string;
  zip?: string;
  phone?: string;
};

type OrderLineInput = {
  variantId: string;
  quantity: number;
};

type OrderCreateInput = {
  email: string;
  lineItems: OrderLineInput[];
  currency?: string;
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
  orderKey: string;
  email: string;
  variantId: string;
  quantity: string;
  recipient: string;
  destination: string;
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

type CreateResult =
  | {
      ok: true;
      id: string;
      name: string;
    }
  | {
      ok: false;
      message: string;
    };

const REQUIRED_COLUMNS = ["order_key", "email", "variant_id", "quantity"];

const ORDER_CREATE_MUTATION = `
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

export default async () => {
  render(<Extension />, document.body);
};

function Extension() {
  const { i18n } = shopify;
  const [fileName, setFileName] = useState("");
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [parseNotice, setParseNotice] = useState("");
  const [isCreating, setIsCreating] = useState(false);
  const [results, setResults] = useState<RowResult[]>([]);
  const [summary, setSummary] = useState<Summary | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const previewRows = parseResult?.previewRows.slice(0, 5) ?? [];
  const totalRows = parseResult?.rowCount ?? 0;
  const validRows = parseResult ? parseResult.rowCount - parseResult.invalidRows.length : 0;
  const uniqueOrderCount = parseResult?.orders.length ?? 0;

  const canCreate = !isCreating && Boolean(parseResult) && uniqueOrderCount > 0;

  const resetAll = () => {
    setFileName("");
    setParseResult(null);
    setParseNotice("");
    setResults([]);
    setSummary(null);
    setIsCreating(false);
  };

  const onOpenFilePicker = () => {
    fileInputRef.current?.click();
  };

  const onFileChange = async (event: Event) => {
    const input = event.target as HTMLInputElement;
    const file = input?.files?.[0];
    if (!file) return;

    resetAll();
    setFileName(file.name);

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

  const onCreateOrders = async () => {
    if (!parseResult || !canCreate) return;

    setIsCreating(true);
    setResults([]);
    setSummary(null);

    const rowLookup = new Map<number, CsvPreviewRow>(
      parseResult.previewRows.map((row) => [row.rowNumber, row]),
    );

    const processedResults: RowResult[] = parseResult.invalidRows.map((row) => ({
      ...row,
      status: "failed",
    }));

    let successCount = 0;
    let failedCount = parseResult.invalidRows.length;
    let ordersCreatedCount = 0;

    for (const order of parseResult.orders) {
      const result = await createOrder(order.input);

      for (const rowNumber of order.rowNumbers) {
        const row = rowLookup.get(rowNumber);
        if (!row) continue;

        if (result.ok) {
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
            errorMessage: result.message,
          });
          failedCount += 1;
        }
      }

      if (result.ok) {
        ordersCreatedCount += 1;
      }
    }

    processedResults.sort((a, b) => a.rowNumber - b.rowNumber);
    setResults(processedResults);
    setSummary({
      total: parseResult.rowCount,
      success: successCount,
      failed: failedCount,
      ordersCreated: ordersCreatedCount,
    });
    setIsCreating(false);
  };

  return (
    <s-admin-block heading={i18n.translate("name")}>
      <s-box>
        <div style={{ maxWidth: "760px", margin: "0 auto", padding: "4px 8px" }}>
          {parseNotice ? (
            <div
              style={{
                marginBottom: "12px",
                borderRadius: "8px",
                padding: "10px 12px",
                background: "#fff7ed",
                border: "1px solid #fdba74",
              }}
            >
              {parseNotice}
            </div>
          ) : null}

          {summary ? (
            <div
              style={{
                marginBottom: "12px",
                borderRadius: "8px",
                padding: "10px 12px",
                background: summary.failed === 0 ? "#ecfdf3" : "#fffbeb",
                border: summary.failed === 0 ? "1px solid #86efac" : "1px solid #fcd34d",
              }}
            >
              {summary.failed === 0
                ? `Processed ${summary.success} rows and created ${summary.ordersCreated} orders.`
                : `Processed ${summary.total} rows: ${summary.success} succeeded, ${summary.failed} failed.`}
            </div>
          ) : null}

          {results.length === 0 ? (
            <div>
              <h3 style={{ marginBottom: "8px" }}>Upload CSV File</h3>
              <p style={{ marginTop: 0, marginBottom: "4px", color: "#475467" }}>
                {i18n.translate("required-columns")}
              </p>
              <p style={{ marginTop: 0, marginBottom: "14px", color: "#475467" }}>
                One row per line item. Same <code>order_key</code> groups rows into one order.
              </p>

              <input
                ref={fileInputRef}
                type="file"
                accept=".csv,text/csv"
                onChange={onFileChange}
                disabled={isCreating}
                style={{ display: "none" }}
              />

              <div
                style={{
                  border: "1px dashed #98a2b3",
                  borderRadius: "10px",
                  padding: "22px 16px",
                  textAlign: "center",
                  marginBottom: "14px",
                }}
              >
                <s-button variant="primary" onClick={onOpenFilePicker} disabled={isCreating}>
                  {i18n.translate("upload-csv")}
                </s-button>
                <p style={{ marginBottom: 0, color: "#475467", marginTop: "10px" }}>
                  {fileName
                    ? `${i18n.translate("selected-file")} ${fileName}`
                    : "Choose a CSV file from your computer"}
                </p>
              </div>

              {parseResult ? (
                <>
                  <p style={{ marginTop: 0, color: "#475467" }}>
                    {totalRows} rows found. {validRows} valid rows. {uniqueOrderCount} order(s) ready.
                  </p>

                  {parseResult.errors.length > 0 ? (
                    <div
                      style={{
                        marginBottom: "12px",
                        borderRadius: "8px",
                        padding: "10px 12px",
                        background: "#fffbeb",
                        border: "1px solid #fcd34d",
                      }}
                    >
                      <strong>Validation issues</strong>
                      <ul style={{ marginBottom: 0 }}>
                        {parseResult.errors.slice(0, 8).map((error, index) => (
                          <li key={`parse-error-${index}`}>{error}</li>
                        ))}
                      </ul>
                    </div>
                  ) : null}

                  <h4 style={{ marginBottom: "8px" }}>Preview (first 5 rows)</h4>
                  <div
                    style={{
                      border: "1px solid #e4e7ec",
                      borderRadius: "8px",
                      overflowX: "auto",
                      marginBottom: "14px",
                    }}
                  >
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                      <thead>
                        <tr style={{ background: "#f9fafb" }}>
                          <th style={cellHeaderStyle}>Row</th>
                          <th style={cellHeaderStyle}>Order Key</th>
                          <th style={cellHeaderStyle}>Email</th>
                          <th style={cellHeaderStyle}>Recipient</th>
                          <th style={cellHeaderStyle}>Destination</th>
                          <th style={cellHeaderStyle}>Variant ID</th>
                          <th style={cellHeaderStyle}>Qty</th>
                        </tr>
                      </thead>
                      <tbody>
                        {previewRows.map((row) => (
                          <tr key={`preview-row-${row.rowNumber}`}>
                            <td style={cellValueStyle}>{row.rowNumber}</td>
                            <td style={cellValueStyle}>{row.orderKey || "-"}</td>
                            <td style={cellValueStyle}>{row.email || "-"}</td>
                            <td style={cellValueStyle}>{row.recipient || "-"}</td>
                            <td style={cellValueStyle}>{row.destination || "-"}</td>
                            <td style={cellValueStyle}>{row.variantId || "-"}</td>
                            <td style={cellValueStyle}>{row.quantity || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div style={{ display: "flex", justifyContent: "flex-end", gap: "8px" }}>
                    <s-button onClick={resetAll} disabled={isCreating}>
                      Reset
                    </s-button>
                    <s-button variant="primary" onClick={onCreateOrders} disabled={!canCreate}>
                      {isCreating
                        ? i18n.translate("creating-orders")
                        : i18n.translate("create-orders")}
                    </s-button>
                  </div>
                </>
              ) : null}
            </div>
          ) : (
            <div>
              <h3 style={{ marginBottom: "8px" }}>Results</h3>
              <div
                style={{
                  border: "1px solid #e4e7ec",
                  borderRadius: "8px",
                  overflowX: "auto",
                  marginBottom: "14px",
                }}
              >
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                  <thead>
                    <tr style={{ background: "#f9fafb" }}>
                      <th style={cellHeaderStyle}>Row</th>
                      <th style={cellHeaderStyle}>Order Key</th>
                      <th style={cellHeaderStyle}>Email</th>
                      <th style={cellHeaderStyle}>Variant ID</th>
                      <th style={cellHeaderStyle}>Qty</th>
                      <th style={cellHeaderStyle}>Status</th>
                      <th style={cellHeaderStyle}>Error</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.map((row) => (
                      <tr key={`result-row-${row.rowNumber}`}>
                        <td style={cellValueStyle}>{row.rowNumber}</td>
                        <td style={cellValueStyle}>{row.orderKey || "-"}</td>
                        <td style={cellValueStyle}>{row.email || "-"}</td>
                        <td style={cellValueStyle}>{row.variantId || "-"}</td>
                        <td style={cellValueStyle}>{row.quantity || "-"}</td>
                        <td style={cellValueStyle}>
                          <span
                            style={{
                              padding: "2px 8px",
                              borderRadius: "999px",
                              background: row.status === "success" ? "#ecfdf3" : "#fef2f2",
                              color: row.status === "success" ? "#166534" : "#991b1b",
                              border:
                                row.status === "success"
                                  ? "1px solid #86efac"
                                  : "1px solid #fca5a5",
                            }}
                          >
                            {row.status === "success" ? "Created" : "Failed"}
                          </span>
                        </td>
                        <td style={cellValueStyle}>{row.errorMessage || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div style={{ display: "flex", justifyContent: "flex-end" }}>
                <s-button variant="primary" onClick={resetAll}>
                  Upload New CSV
                </s-button>
              </div>
            </div>
          )}

          <div style={{ marginTop: "16px", color: "#667085", fontSize: "12px" }}>
            Tip: keep this block pinned on Order Details to access CSV import without a modal.
          </div>
        </div>
      </s-box>
    </s-admin-block>
  );
}

const cellHeaderStyle = {
  textAlign: "left",
  borderBottom: "1px solid #e4e7ec",
  padding: "8px",
  fontWeight: "600",
} as const;

const cellValueStyle = {
  textAlign: "left",
  borderBottom: "1px solid #f2f4f7",
  padding: "8px",
  verticalAlign: "top",
} as const;

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
  const missingHeaders = REQUIRED_COLUMNS.filter((header) => !headers.includes(header));

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
  const groupedOrders = new Map<string, OrderDraft>();

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

    const existing = groupedOrders.get(parsed.value.orderKey);
    if (!existing) {
      groupedOrders.set(parsed.value.orderKey, {
        orderKey: parsed.value.orderKey,
        rowNumbers: [rowNumber],
        input: {
          email: parsed.value.email,
          lineItems: [
            {
              variantId: parsed.value.variantId,
              quantity: parsed.value.quantity,
            },
          ],
          currency: parsed.value.currency,
          note: parsed.value.note,
          tags: parsed.value.tags,
          shippingAddress: parsed.value.shippingAddress,
        },
      });
      continue;
    }

    if (existing.input.email !== parsed.value.email) {
      const conflictMessage =
        `Row ${rowNumber}: email does not match earlier rows for order_key=${parsed.value.orderKey}`;
      errors.push(conflictMessage);
      invalidRows.push({ ...previewRow, errorMessage: conflictMessage });
      continue;
    }

    existing.input.lineItems.push({
      variantId: parsed.value.variantId,
      quantity: parsed.value.quantity,
    });
    existing.rowNumbers.push(rowNumber);
  }

  if (previewRows.length === 0) {
    errors.push("No non-empty data rows found.");
  }

  return {
    rowCount: previewRows.length,
    orders: Array.from(groupedOrders.values()),
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

    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
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
  const recipient = [readColumn(record, "shipping_first_name"), readColumn(record, "shipping_last_name")]
    .filter(Boolean)
    .join(" ");

  const destination = [
    readColumn(record, "shipping_city"),
    readColumn(record, "shipping_province_code"),
    readColumn(record, "shipping_country_code"),
  ]
    .filter(Boolean)
    .join(", ");

  return {
    rowNumber,
    orderKey: readColumn(record, "order_key"),
    email: readColumn(record, "email"),
    variantId: readColumn(record, "variant_id"),
    quantity: readColumn(record, "quantity"),
    recipient,
    destination,
  };
}

function normalizeVariantId(value: string): string {
  if (/^\d+$/.test(value)) {
    return `gid://shopify/ProductVariant/${value}`;
  }
  return value;
}

function parseTags(value: string): string[] | undefined {
  if (!value) return undefined;
  const tags = value
    .split(/[|,]/g)
    .map((tag) => tag.trim())
    .filter(Boolean);
  return tags.length > 0 ? tags : undefined;
}

function buildShippingAddress(record: CsvRecord): ShippingAddress | undefined {
  const shippingAddress: ShippingAddress = {
    firstName: readColumn(record, "shipping_first_name") || undefined,
    lastName: readColumn(record, "shipping_last_name") || undefined,
    address1: readColumn(record, "shipping_address1") || undefined,
    city: readColumn(record, "shipping_city") || undefined,
    provinceCode: readColumn(record, "shipping_province_code") || undefined,
    countryCode: readColumn(record, "shipping_country_code") || undefined,
    zip: readColumn(record, "shipping_zip") || undefined,
    phone: readColumn(record, "phone") || undefined,
  };

  const hasAnyField = Object.values(shippingAddress).some(Boolean);
  return hasAnyField ? shippingAddress : undefined;
}

function normalizeRow(
  record: CsvRecord,
  rowNumber: number,
):
  | {
      ok: true;
      value: {
        orderKey: string;
        email: string;
        variantId: string;
        quantity: number;
        currency?: string;
        note?: string;
        tags?: string[];
        shippingAddress?: ShippingAddress;
      };
    }
  | {
      ok: false;
      errors: string[];
    } {
  const rowErrors: string[] = [];
  const orderKey = readColumn(record, "order_key");
  const email = readColumn(record, "email");
  const variantIdRaw = readColumn(record, "variant_id");
  const quantityRaw = readColumn(record, "quantity");

  if (!orderKey) rowErrors.push(`Row ${rowNumber}: order_key is required.`);
  if (!email) rowErrors.push(`Row ${rowNumber}: email is required.`);
  if (!variantIdRaw) rowErrors.push(`Row ${rowNumber}: variant_id is required.`);

  const quantity = Number.parseInt(quantityRaw, 10);
  if (!Number.isInteger(quantity) || quantity <= 0) {
    rowErrors.push(`Row ${rowNumber}: quantity must be a positive integer.`);
  }

  if (rowErrors.length > 0) {
    return { ok: false, errors: rowErrors };
  }

  const currency = readColumn(record, "currency_code");
  const normalizedCurrency = currency ? currency.toUpperCase() : undefined;
  const note = readColumn(record, "note") || undefined;
  const tags = parseTags(readColumn(record, "tags"));
  const shippingAddress = buildShippingAddress(record);

  return {
    ok: true,
    value: {
      orderKey,
      email,
      variantId: normalizeVariantId(variantIdRaw),
      quantity,
      currency: normalizedCurrency,
      note,
      tags,
      shippingAddress,
    },
  };
}

async function createOrder(input: OrderCreateInput): Promise<CreateResult> {
  try {
    const response = await adminGraphql<{
      data?: {
        orderCreate?: {
          order?: {
            id: string;
            name: string;
          };
          userErrors?: { field?: string[]; message: string }[];
        };
      };
      errors?: { message: string }[];
    }>(ORDER_CREATE_MUTATION, { order: compactOrderInput(input) });

    if (response.errors && response.errors.length > 0) {
      return { ok: false, message: response.errors.map((error) => error.message).join("; ") };
    }

    const payload = response.data?.orderCreate;
    const userErrors = payload?.userErrors ?? [];
    if (userErrors.length > 0) {
      return { ok: false, message: userErrors.map((error) => error.message).join("; ") };
    }

    if (!payload?.order?.id) {
      return { ok: false, message: "Unknown orderCreate response." };
    }

    return {
      ok: true,
      id: payload.order.id,
      name: payload.order.name ?? "Order",
    };
  } catch (error) {
    return {
      ok: false,
      message: error instanceof Error ? error.message : "Unexpected error while creating order.",
    };
  }
}

function compactOrderInput(input: OrderCreateInput): OrderCreateInput {
  const compacted: OrderCreateInput = {
    email: input.email,
    lineItems: input.lineItems,
  };

  if (input.currency) compacted.currency = input.currency;
  if (input.note) compacted.note = input.note;
  if (input.tags && input.tags.length > 0) compacted.tags = input.tags;
  if (input.shippingAddress) compacted.shippingAddress = input.shippingAddress;

  return compacted;
}

async function adminGraphql<T>(query: string, variables: Record<string, unknown>): Promise<T> {
  const response = await fetch("shopify:admin/api/graphql.json", {
    method: "POST",
    body: JSON.stringify({ query, variables }),
  });

  if (!response.ok) {
    throw new Error(`GraphQL request failed with status ${response.status}`);
  }

  return response.json() as Promise<T>;
}
