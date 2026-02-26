import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
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

type CustomerActionData =
  | {
      type: "search";
      customers: CustomerOption[];
      warning?: string;
    }
  | {
      type: "create";
      customer: CustomerOption;
    }
  | {
      type: "error";
      error: string;
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

const CUSTOMER_PICKER_STYLES = `
.Polaris-Layout {
  display: grid;
  gap: 16px;
  grid-template-columns: minmax(0, 2fr) minmax(0, 1fr);
}

.Polaris-Layout__Section {
  min-width: 0;
  overflow: visible;
}

.Polaris-Layout__Section--oneThird {
  overflow: visible;
}

@container (inline-size <= 900px) {
  .Polaris-Layout {
    grid-template-columns: minmax(0, 1fr);
  }
}

.customer-picker {
  position: relative;
  z-index: 40;
}

.customer-picker__menu {
  position: fixed;
  z-index: 2147483647;
  background: var(--p-color-bg-surface, #fff);
  border: 1px solid var(--p-color-border, #d1d5db);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.12);
  overflow: hidden;
  max-height: min(320px, calc(100vh - 24px));
}

.customer-picker__list {
  max-height: 280px;
  overflow-y: auto;
}
`;

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

const SEARCH_CUSTOMERS_QUERY = `#graphql
  query SearchCustomers($query: String!) {
    customers(first: 50, query: $query) {
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

const CUSTOMER_CREATE_MUTATION = `#graphql
  mutation CustomerCreateForImporter($input: CustomerInput!) {
    customerCreate(input: $input) {
      customer {
        id
        displayName
        email
      }
      userErrors {
        field
        message
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
          fulfillments {
            trackingInfo {
              number
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

  if (intent === "search_customers") {
    return searchCustomers(admin, formData);
  }

  if (intent === "create_customer") {
    return createCustomer(admin, formData);
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
  const customerSearchFetcher = useFetcher<CustomerActionData>();
  const customerCreateFetcher = useFetcher<CustomerActionData>();

  const [fileName, setFileName] = useState("");
  const [selectedCustomerId, setSelectedCustomerId] = useState("");
  const [customerOptions, setCustomerOptions] = useState<CustomerOption[]>(() =>
    sortCustomerOptions(customers),
  );
  const [customerInputValue, setCustomerInputValue] = useState("");
  const [isCustomerDropdownOpen, setIsCustomerDropdownOpen] = useState(false);
  const [createCustomerFirstNameInput, setCreateCustomerFirstNameInput] =
    useState("");
  const [createCustomerLastNameInput, setCreateCustomerLastNameInput] =
    useState("");
  const [createCustomerEmailInput, setCreateCustomerEmailInput] = useState("");
  const [customerMessage, setCustomerMessage] = useState("");
  const [customerMessageTone, setCustomerMessageTone] = useState<
    "success" | "warning"
  >("success");
  const [customerError, setCustomerError] = useState("");
  const [isHydrated, setIsHydrated] = useState(false);
  const [orderTagsInput, setOrderTagsInput] = useState("");
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [parseNotice, setParseNotice] = useState("");
  const [importResult, setImportResult] = useState<ImportResult | null>(null);

  const [orderNumbersInput, setOrderNumbersInput] = useState("");
  const [searchTagsInput, setSearchTagsInput] = useState("");
  const [reportOrders, setReportOrders] = useState<ShippingReportOrder[]>([]);
  const [reportNotice, setReportNotice] = useState("");
  const [reportError, setReportError] = useState("");
  const [customerMenuRect, setCustomerMenuRect] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);
  const createCustomerModalRef = useRef<HTMLElementTagNameMap["s-modal"] | null>(
    null,
  );
  const customerPickerRef = useRef<HTMLDivElement | null>(null);
  const customerFieldElementRef = useRef<HTMLElement | null>(null);
  const customerMenuRef = useRef<HTMLDivElement | null>(null);

  const isCreating = createFetcher.state === "submitting";
  const isGeneratingReport = reportFetcher.state === "submitting";
  const isCreatingCustomer = customerCreateFetcher.state === "submitting";
  const canCreate =
    !isCreating && Boolean(parseResult) && (parseResult?.orders.length ?? 0) > 0;

  useEffect(() => {
    setCustomerOptions((existing) =>
      mergeCustomerOptions([...existing, ...sortCustomerOptions(customers)]),
    );
  }, [customers]);

  useEffect(() => {
    setIsHydrated(true);
  }, []);

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

  useEffect(() => {
    const searchData = customerSearchFetcher.data;
    if (!searchData) return;

    if (searchData.type === "error") {
      setCustomerError(searchData.error);
      setCustomerMessage("");
      return;
    }

    if (searchData.type !== "search") return;

    setCustomerError("");
    setCustomerOptions((existing) =>
      mergeCustomerOptions([...existing, ...searchData.customers]),
    );
    if (searchData.warning) {
      setCustomerMessage(searchData.warning);
      setCustomerMessageTone("warning");
    } else {
      setCustomerMessage("");
    }
  }, [customerSearchFetcher.data]);

  useEffect(() => {
    const createData = customerCreateFetcher.data;
    if (!createData) return;

    if (createData.type === "error") {
      setCustomerError(createData.error);
      setCustomerMessage("");
      return;
    }

    if (createData.type !== "create") return;

    const createdCustomer = createData.customer;
    setCustomerError("");
    setCustomerMessage(`Created customer ${createdCustomer.displayName}.`);
    setCustomerMessageTone("success");
    setCustomerOptions((existing) =>
      mergeCustomerOptions([createdCustomer, ...existing]),
    );
    setSelectedCustomerId(createdCustomer.id);
    setCustomerInputValue(formatCustomerOption(createdCustomer));
    setCreateCustomerFirstNameInput("");
    setCreateCustomerLastNameInput("");
    setCreateCustomerEmailInput("");
    setIsCustomerDropdownOpen(false);
    setCustomerMenuRect(null);
    createCustomerModalRef.current?.hideOverlay?.();
  }, [customerCreateFetcher.data]);

  useEffect(() => {
    const query = customerInputValue.trim();
    if (query.length < 2) return;

    const timeout = window.setTimeout(() => {
      const formData = new FormData();
      formData.append("intent", "search_customers");
      formData.append("query", query);
      customerSearchFetcher.submit(formData, { method: "POST" });
    }, 250);

    return () => {
      window.clearTimeout(timeout);
    };
  }, [customerInputValue]);

  useEffect(() => {
    if (!isCustomerDropdownOpen) return;

    updateCustomerMenuPosition();

    const onWindowChange = () => {
      updateCustomerMenuPosition();
    };

    window.addEventListener("resize", onWindowChange);
    window.addEventListener("scroll", onWindowChange, true);

    return () => {
      window.removeEventListener("resize", onWindowChange);
      window.removeEventListener("scroll", onWindowChange, true);
    };
  }, [isCustomerDropdownOpen, customerInputValue]);

  useEffect(() => {
    if (!isCustomerDropdownOpen) return;

    const onPointerDown = (event: PointerEvent) => {
      const picker = customerPickerRef.current;
      const menu = customerMenuRef.current;
      const target = event.target as Node | null;
      const path = typeof event.composedPath === "function" ? event.composedPath() : [];
      const insidePicker = picker
        ? path.includes(picker) || (target ? picker.contains(target) : false)
        : false;
      const insideMenu = menu
        ? path.includes(menu) || (target ? menu.contains(target) : false)
        : false;

      if (insidePicker || insideMenu) {
        return;
      }

      setIsCustomerDropdownOpen(false);
      setCustomerMenuRect(null);
    };

    document.addEventListener("pointerdown", onPointerDown, true);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown, true);
    };
  }, [isCustomerDropdownOpen]);

  const filteredCustomers = useMemo(() => {
    const query = customerInputValue.trim().toLowerCase();
    const list = sortCustomerOptions(customerOptions);

    if (!query) {
      return list.slice(0, 50);
    }

    return list
      .filter((customer) => {
        const name = customer.displayName.toLowerCase();
        const email = customer.email.toLowerCase();
        return name.includes(query) || email.includes(query);
      })
      .slice(0, 50);
  }, [customerInputValue, customerOptions]);

  const resolveCustomerAnchor = () => {
    if (customerFieldElementRef.current) {
      return customerFieldElementRef.current;
    }

    return customerPickerRef.current?.querySelector<HTMLElement>(
      "s-search-field",
    );
  };

  const updateCustomerMenuPosition = () => {
    const field = resolveCustomerAnchor();
    if (!field) {
      setCustomerMenuRect(null);
      return;
    }

    customerFieldElementRef.current = field;

    const rect = field.getBoundingClientRect();
    const viewportPadding = 8;
    const menuMaxHeight = Math.min(320, window.innerHeight - 24);
    const width = Math.max(rect.width, 280);
    const maxLeft = Math.max(
      viewportPadding,
      window.innerWidth - width - viewportPadding,
    );
    const left = Math.min(Math.max(rect.left, viewportPadding), maxLeft);
    const shouldOpenUpwards =
      rect.bottom + 4 + menuMaxHeight > window.innerHeight - viewportPadding &&
      rect.top - 4 - menuMaxHeight > viewportPadding;
    const top = shouldOpenUpwards
      ? Math.max(viewportPadding, rect.top - menuMaxHeight - 4)
      : Math.min(
          window.innerHeight - menuMaxHeight - viewportPadding,
          rect.bottom + 4,
        );

    setCustomerMenuRect({
      top,
      left,
      width,
    });
  };

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

  const onCustomerFieldInput = (event: Event) => {
    const target = event.currentTarget as HTMLElement & { value?: string };
    const value = target.value || "";
    customerFieldElementRef.current = target;
    setCustomerInputValue(value);
    setIsCustomerDropdownOpen(true);
    updateCustomerMenuPosition();
    setCustomerError("");
    setCustomerMessage("");

    if (selectedCustomerId) {
      const selected = customerOptions.find(
        (customer) => customer.id === selectedCustomerId,
      );
      if (!selected || formatCustomerOption(selected) !== value) {
        setSelectedCustomerId("");
      }
    }
  };

  const onCustomerFieldFocus = (event: Event) => {
    customerFieldElementRef.current = event.currentTarget as HTMLElement;
    setIsCustomerDropdownOpen(true);
    updateCustomerMenuPosition();
  };

  const onSelectCustomer = (customer: CustomerOption) => {
    setSelectedCustomerId(customer.id);
    setCustomerInputValue(formatCustomerOption(customer));
    setIsCustomerDropdownOpen(false);
    setCustomerMenuRect(null);
    setCustomerError("");
    setCustomerMessage("");
  };

  const onOpenCreateCustomerModal = () => {
    const typedValue = customerInputValue.trim();
    const extractedEmail = extractEmailCandidate(typedValue);
    setCreateCustomerFirstNameInput("");
    setCreateCustomerLastNameInput("");
    setCreateCustomerEmailInput(extractedEmail);
    setCustomerError("");
    setCustomerMessage("");
    setIsCustomerDropdownOpen(false);
    setCustomerMenuRect(null);
    createCustomerModalRef.current?.showOverlay?.();
  };

  const onCreateCustomerFromModal = () => {
    const email = createCustomerEmailInput.trim();
    if (!email) {
      setCustomerError("Email is required to create a customer.");
      setCustomerMessage("");
      return;
    }

    const formData = new FormData();
    formData.append("intent", "create_customer");
    formData.append("email", email);
    formData.append("firstName", createCustomerFirstNameInput.trim());
    formData.append("lastName", createCustomerLastNameInput.trim());
    customerCreateFetcher.submit(formData, { method: "POST" });
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
      <style>{CUSTOMER_PICKER_STYLES}</style>

      {importResult?.error ? (
        <s-banner tone="critical">{importResult.error}</s-banner>
      ) : null}

      {parseNotice ? <s-banner tone="warning">{parseNotice}</s-banner> : null}

      {customerLoadWarning ? (
        <s-banner tone="warning">{customerLoadWarning}</s-banner>
      ) : null}

      {customerError ? <s-banner tone="critical">{customerError}</s-banner> : null}

      {customerMessage ? (
        <s-banner tone={customerMessageTone}>{customerMessage}</s-banner>
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

      <div className="Polaris-Layout">
        <div className="Polaris-Layout__Section">
          <s-stack direction="block" gap="base">
            <s-section heading="Upload CSV File">
              <s-stack direction="block" gap="base">
                <s-text color="subdued">
                  Upload a CSV file with columns: first_name, last_name, address,
                  address2, city, state, zip_code
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
          </s-stack>
        </div>

        <div className="Polaris-Layout__Section Polaris-Layout__Section--oneThird">
          <s-stack direction="block" gap="base">
            <s-section heading="Order Settings">
              <s-stack direction="block" gap="base">
                <div className="customer-picker" ref={customerPickerRef}>
                  <s-search-field
                    label="Search or create a customer"
                    placeholder="Search or create a customer"
                    value={customerInputValue}
                    onInput={onCustomerFieldInput}
                    onFocus={onCustomerFieldFocus}
                  />
                </div>

                <s-text color="subdued">
                  Pick one company customer for all imported orders while shipping
                  to each recipient address.
                </s-text>

                <s-text-field
                  label="Order Tags (optional)"
                  placeholder="e.g., imported, bulk-order"
                  value={orderTagsInput}
                  onInput={(event: Event) => {
                    const target = event.currentTarget as HTMLElement & {
                      value?: string;
                    };
                    setOrderTagsInput(target.value || "");
                  }}
                />
                <s-text color="subdued">
                  Add tags to all imported orders. Separate multiple tags with
                  commas.
                </s-text>
              </s-stack>
            </s-section>

            <s-section heading="Shipping Report">
              <s-stack gap="base" direction="block">
                <s-text color="subdued">
                  Search by order numbers and tags, then export the results.
                </s-text>

                <s-text-field
                  label="Order Numbers (optional)"
                  placeholder="e.g., #1001, #1002"
                  value={orderNumbersInput}
                  onInput={(event: Event) => {
                    const target = event.currentTarget as HTMLElement & {
                      value?: string;
                    };
                    setOrderNumbersInput(target.value || "");
                  }}
                />

                <s-text-field
                  label="Tags (optional)"
                  placeholder="e.g., imported, batch:sonya"
                  value={searchTagsInput}
                  onInput={(event: Event) => {
                    const target = event.currentTarget as HTMLElement & {
                      value?: string;
                    };
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
                    <s-table-header listSlot="labeled">
                      Tracking Numbers
                    </s-table-header>
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
          </s-stack>
        </div>
      </div>

      {isHydrated && isCustomerDropdownOpen && customerMenuRect ? (
        createPortal(
          <div
            ref={customerMenuRef}
            className="customer-picker__menu"
            style={{
              top: `${customerMenuRect.top}px`,
              left: `${customerMenuRect.left}px`,
              width: `${customerMenuRect.width}px`,
            }}
            onMouseDown={(event) => {
              event.preventDefault();
            }}
          >
            <s-clickable onClick={onOpenCreateCustomerModal} padding="small">
              <s-text type="strong">+ Create a new customer</s-text>
            </s-clickable>
            <s-divider />
            <div className="customer-picker__list">
              {filteredCustomers.length === 0 ? (
                <s-box padding="small">
                  <s-text color="subdued">No matching customers.</s-text>
                </s-box>
              ) : (
                filteredCustomers.map((customer, index) => (
                  <s-box key={customer.id}>
                    <s-clickable
                      onClick={() => onSelectCustomer(customer)}
                      padding="small"
                    >
                      <s-stack direction="block" gap="none">
                        <s-text type="strong">{customer.displayName}</s-text>
                        <s-text color="subdued">
                          {customer.email || "No email"}
                        </s-text>
                      </s-stack>
                    </s-clickable>
                    {index < filteredCustomers.length - 1 ? <s-divider /> : null}
                  </s-box>
                ))
              )}
            </div>
          </div>,
          document.body,
        )
      ) : null}

      <s-modal
        id="create-customer-modal"
        ref={createCustomerModalRef}
        heading="Create a new customer"
        accessibilityLabel="Create a new customer"
        size="base"
      >
        <s-stack direction="block" gap="base">
          <s-stack direction="inline" gap="base">
            <s-text-field
              label="First name"
              value={createCustomerFirstNameInput}
              onInput={(event: Event) => {
                const target = event.currentTarget as HTMLElement & {
                  value?: string;
                };
                setCreateCustomerFirstNameInput(target.value || "");
              }}
            />
            <s-text-field
              label="Last name"
              value={createCustomerLastNameInput}
              onInput={(event: Event) => {
                const target = event.currentTarget as HTMLElement & {
                  value?: string;
                };
                setCreateCustomerLastNameInput(target.value || "");
              }}
            />
          </s-stack>
          <s-text-field
            label="Email"
            value={createCustomerEmailInput}
            onInput={(event: Event) => {
              const target = event.currentTarget as HTMLElement & { value?: string };
              setCreateCustomerEmailInput(target.value || "");
            }}
          />
          <s-stack direction="inline" gap="base" justifyContent="end">
            <s-button
              variant="secondary"
              commandFor="create-customer-modal"
              command="--hide"
            >
              Cancel
            </s-button>
            <s-button
              variant="primary"
              onClick={onCreateCustomerFromModal}
              {...(isCreatingCustomer ? { loading: true } : {})}
              disabled={!createCustomerEmailInput.trim()}
            >
              Save
            </s-button>
          </s-stack>
        </s-stack>
      </s-modal>

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
              fulfillments?: Array<{
                trackingInfo?: Array<{
                  number?: string | null;
                }>;
              }> | null;
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
      for (const fulfillment of order.fulfillments ?? []) {
        for (const info of fulfillment.trackingInfo ?? []) {
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

async function searchCustomers(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  formData: FormData,
): Promise<CustomerActionData> {
  const queryRaw = (formData.get("query") as string | null)?.trim() || "";
  if (queryRaw.length < 2) {
    return {
      type: "search",
      customers: [],
      warning: "Type at least 2 characters to search.",
    };
  }

  const escaped = escapeSearchValue(queryRaw);
  const tokenQueries = queryRaw
    .split(/\s+/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const escapedPart = escapeSearchValue(part);
      return `(name:${escapedPart}* OR email:${escapedPart}*)`;
    });
  const searchQuery = [...tokenQueries, escaped].join(" OR ");

  try {
    const response = await admin.graphql(SEARCH_CUSTOMERS_QUERY, {
      variables: { query: searchQuery },
    });

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
      return {
        type: "error",
        error: json.errors.map((error) => error.message).join("; "),
      };
    }

    const customers: CustomerOption[] = [];
    for (const edge of json.data?.customers?.edges ?? []) {
      const node = edge.node;
      if (!node?.id) continue;
      customers.push({
        id: node.id,
        displayName: node.displayName || node.email || "Unnamed customer",
        email: node.email || "",
      });
    }

    if (customers.length === 0) {
      return {
        type: "search",
        customers: [],
        warning: "No customer found. Create one below.",
      };
    }

    return {
      type: "search",
      customers,
    };
  } catch (error) {
    return {
      type: "error",
      error:
        error instanceof Error
          ? error.message
          : "Failed to search customers.",
    };
  }
}

async function createCustomer(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  formData: FormData,
): Promise<CustomerActionData> {
  const email = (formData.get("email") as string | null)?.trim() || "";
  const firstNameInput =
    (formData.get("firstName") as string | null)?.trim() || "";
  const lastNameInput = (formData.get("lastName") as string | null)?.trim() || "";
  const fullName = (formData.get("name") as string | null)?.trim() || "";

  if (!email) {
    return {
      type: "error",
      error: "Customer email is required.",
    };
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return {
      type: "error",
      error: "Customer email format is invalid.",
    };
  }

  let firstName = firstNameInput;
  let lastName = lastNameInput;

  if (!firstName && !lastName && fullName) {
    const [first, ...rest] = fullName.split(/\s+/).filter(Boolean);
    firstName = first || "";
    lastName = rest.join(" ");
  }

  try {
    const response = await admin.graphql(CUSTOMER_CREATE_MUTATION, {
      variables: {
        input: {
          email,
          firstName: firstName || undefined,
          lastName: lastName || undefined,
        },
      },
    });

    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        customerCreate?: {
          customer?: {
            id?: string;
            displayName?: string;
            email?: string | null;
          } | null;
          userErrors?: Array<{ message: string }>;
        };
      };
    };

    if (json.errors && json.errors.length > 0) {
      return {
        type: "error",
        error: json.errors.map((error) => error.message).join("; "),
      };
    }

    const payload = json.data?.customerCreate;
    const userErrors = payload?.userErrors ?? [];
    if (userErrors.length > 0) {
      return {
        type: "error",
        error: userErrors.map((error) => error.message).join("; "),
      };
    }

    const createdCustomer = payload?.customer;
    if (!createdCustomer?.id) {
      return {
        type: "error",
        error: "Customer was not created.",
      };
    }

    return {
      type: "create",
      customer: {
        id: createdCustomer.id,
        displayName:
          createdCustomer.displayName ||
          [firstName, lastName].filter(Boolean).join(" ") ||
          createdCustomer.email ||
          email,
        email: createdCustomer.email || email,
      },
    };
  } catch (error) {
    return {
      type: "error",
      error:
        error instanceof Error
          ? error.message
          : "Failed to create customer.",
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

function mergeCustomerOptions(customers: CustomerOption[]): CustomerOption[] {
  const byId = new Map<string, CustomerOption>();

  for (const customer of customers) {
    if (!customer.id) continue;
    if (!byId.has(customer.id)) {
      byId.set(customer.id, customer);
    }
  }

  return sortCustomerOptions(Array.from(byId.values()));
}

function sortCustomerOptions(customers: CustomerOption[]): CustomerOption[] {
  return [...customers].sort((a, b) => {
    const aLabel = (a.displayName || a.email || "").toLowerCase();
    const bLabel = (b.displayName || b.email || "").toLowerCase();
    return aLabel.localeCompare(bLabel, "en");
  });
}

function formatCustomerOption(customer: CustomerOption): string {
  const name = customer.displayName || customer.email || "Unnamed customer";
  return customer.email ? `${name} (${customer.email})` : name;
}

function extractEmailCandidate(value: string): string {
  const match = value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0].toLowerCase() : "";
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
