import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type {
  ActionFunctionArgs,
  ClientLoaderFunctionArgs,
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

type OrderTagSearchActionData =
  | {
      type: "search_order_tags";
      query: string;
      tags: string[];
    }
  | {
      type: "error";
      query: string;
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
const ORDER_TAG_SCAN_MAX_PAGES = 10;

const CUSTOMER_PICKER_STYLES = `
.Polaris-Layout {
  display: grid;
  gap: 16px;
  grid-template-columns: minmax(0, 1fr);
}

.Polaris-Layout__Section {
  min-width: 0;
  overflow: visible;
}

.Polaris-Layout__Section--oneThird {
  overflow: visible;
  order: -1;
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

.customer-picker__loading {
  padding: 8px 12px;
}

.customer-picker__menu [data-customer-row="true"][data-active="true"] {
  background: var(--p-color-bg-fill-tertiary, #f3f4f6);
}

.customer-picker__create-content {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.customer-picker__create-icon {
  width: 16px;
  height: 16px;
  color: var(--p-color-text-subdued, #6b7280);
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.customer-picker__create-icon svg {
  width: 16px;
  height: 16px;
  display: block;
  fill: currentColor;
}

.order-tag-picker {
  position: relative;
  z-index: 35;
}

.order-tag-picker__menu {
  position: fixed;
  z-index: 2147483646;
  background: var(--p-color-bg-surface, #fff);
  border: 1px solid var(--p-color-border, #d1d5db);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.12);
  overflow: hidden;
  max-height: min(320px, calc(100vh - 24px));
}

.order-tag-picker__list {
  max-height: 280px;
  overflow-y: auto;
}

.order-tag-picker__menu [data-order-tag-row="true"]:hover {
  background: var(--p-color-bg-fill-tertiary, #f3f4f6);
}

.order-tag-picker__action-content {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.order-tag-picker__icon {
  width: 16px;
  height: 16px;
  color: var(--p-color-text-subdued, #6b7280);
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.order-tag-picker__icon svg {
  width: 16px;
  height: 16px;
  display: block;
  fill: currentColor;
}

.order-tag-picker__checkbox {
  width: 16px;
  height: 16px;
  border: 1px solid var(--p-color-border, #d1d5db);
  border-radius: 4px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex: 0 0 auto;
}

.order-tag-picker__checkbox--checked {
  border-color: var(--p-color-bg-fill-emphasis, #111827);
  background: var(--p-color-bg-fill-emphasis, #111827);
  color: var(--p-color-text-on-fill, #fff);
}

.order-tag-picker__checkbox svg {
  width: 12px;
  height: 12px;
  display: block;
  fill: currentColor;
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

const ORDER_TAG_SUGGESTIONS_QUERY = `#graphql
  query OrderTagSuggestions {
    shop {
      orderTags(first: 250, sortKey: POPULAR) {
        edges {
          node
        }
      }
    }
  }
`;

const DRAFT_ORDER_TAG_SUGGESTIONS_QUERY = `#graphql
  query DraftOrderTagSuggestions {
    shop {
      draftOrderTags(first: 250, sortKey: POPULAR) {
        edges {
          node {
            title
          }
        }
      }
    }
  }
`;

const SEARCH_ORDER_TAGS_QUERY = `#graphql
  query SearchOrderTags($query: String!) {
    shop {
      orderTags(first: 250, sortKey: POPULAR, query: $query) {
        edges {
          node
        }
      }
    }
  }
`;

const SEARCH_DRAFT_ORDER_TAGS_QUERY = `#graphql
  query SearchDraftOrderTags($query: String!) {
    shop {
      draftOrderTags(first: 250, sortKey: POPULAR, query: $query) {
        edges {
          node {
            title
          }
        }
      }
    }
  }
`;

const ORDER_TAG_SUGGESTIONS_BY_DRAFT_ORDERS_QUERY = `#graphql
  query OrderTagSuggestionsByDraftOrders($after: String) {
    draftOrders(first: 250, after: $after, sortKey: UPDATED_AT, reverse: true) {
      edges {
        node {
          tags
        }
      }
      pageInfo {
        hasNextPage
        endCursor
      }
    }
  }
`;

const SEARCH_ORDER_TAGS_BY_DRAFT_ORDERS_QUERY = `#graphql
  query SearchOrderTagsByDraftOrders($query: String!, $after: String) {
    draftOrders(
      first: 250
      query: $query
      after: $after
      sortKey: UPDATED_AT
      reverse: true
    ) {
      edges {
        node {
          tags
        }
      }
      pageInfo {
        hasNextPage
        endCursor
      }
    }
  }
`;

const ORDER_TAG_SUGGESTIONS_BY_ORDERS_QUERY = `#graphql
  query OrderTagSuggestionsByOrders($after: String) {
    orders(
      first: 250
      after: $after
      sortKey: CREATED_AT
      reverse: true
    ) {
      edges {
        node {
          tags
        }
      }
      pageInfo {
        hasNextPage
        endCursor
      }
    }
  }
`;

const SEARCH_ORDER_TAGS_BY_ORDERS_QUERY = `#graphql
  query SearchOrderTagsByOrders($query: String!, $after: String) {
    orders(
      first: 250
      query: $query
      after: $after
      sortKey: CREATED_AT
      reverse: true
    ) {
      edges {
        node {
          tags
        }
      }
      pageInfo {
        hasNextPage
        endCursor
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

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);

  const customers: CustomerOption[] = [];
  let customerLoadWarning = "";
  const orderTagSuggestions: string[] = [];
  let orderTagLoadWarning = "";

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

  try {
    const tagResult = await loadShopOrderTags(admin);
    if (tagResult.error) {
      orderTagLoadWarning = `Unable to load order tag suggestions: ${tagResult.error}`;
    } else {
      orderTagSuggestions.push(...tagResult.tags);
    }
  } catch {
    orderTagLoadWarning =
      "Unable to load order tag suggestions right now. You can still add tags manually.";
  }

  return {
    customers,
    customerLoadWarning,
    orderTagSuggestions,
    orderTagLoadWarning,
  };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const { admin } = await authenticate.admin(request);
  const formData = await request.formData();
  const intent = formData.get("intent");

  if (intent === "search_customers") {
    return searchCustomers(admin, formData);
  }

  if (intent === "create_customer") {
    return createCustomer(admin, formData);
  }

  if (intent === "search_order_tags") {
    return searchOrderTags(admin, formData);
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

export async function clientLoader({ serverLoader }: ClientLoaderFunctionArgs) {
  return await serverLoader();
}

export function HydrateFallback() {
  return null;
}

export default function Index() {
  const {
    customers,
    customerLoadWarning,
    orderTagSuggestions,
    orderTagLoadWarning,
  } = useLoaderData<typeof loader>();
  const createFetcher = useFetcher<CreateActionData>();
  const customerSearchFetcher = useFetcher<CustomerActionData>();
  const customerCreateFetcher = useFetcher<CustomerActionData>();
  const orderTagSearchFetcher = useFetcher<OrderTagSearchActionData>();

  const [fileName, setFileName] = useState("");
  const [selectedCustomerId, setSelectedCustomerId] = useState("");
  const [customerOptions, setCustomerOptions] = useState<CustomerOption[]>(() =>
    sortCustomerOptions(customers),
  );
  const [customerInputValue, setCustomerInputValue] = useState("");
  const [isCustomerDropdownOpen, setIsCustomerDropdownOpen] = useState(false);
  const [activeCustomerMenuIndex, setActiveCustomerMenuIndex] = useState(0);
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
  const [isCustomerSearchQueued, setIsCustomerSearchQueued] = useState(false);
  const [isHydrated, setIsHydrated] = useState(false);
  const [selectedOrderTags, setSelectedOrderTags] = useState<string[]>([]);
  const [orderTagInputValue, setOrderTagInputValue] = useState("");
  const [isOrderTagDropdownOpen, setIsOrderTagDropdownOpen] = useState(false);
  const [isOrderTagSearchQueued, setIsOrderTagSearchQueued] = useState(false);
  const [orderTagSearchQuery, setOrderTagSearchQuery] = useState("");
  const [orderTagSearchResults, setOrderTagSearchResults] = useState<string[]>(
    [],
  );
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [parseNotice, setParseNotice] = useState("");
  const [importResult, setImportResult] = useState<ImportResult | null>(null);
  const [parsedTableQuery, setParsedTableQuery] = useState("");
  const [createdTableQuery, setCreatedTableQuery] = useState("");
  const [customerMenuRect, setCustomerMenuRect] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);
  const [orderTagMenuRect, setOrderTagMenuRect] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);
  const createCustomerModalRef = useRef<HTMLElementTagNameMap["s-modal"] | null>(
    null,
  );
  const customerPickerRef = useRef<HTMLDivElement | null>(null);
  const customerMenuRef = useRef<HTMLDivElement | null>(null);
  const orderTagPickerRef = useRef<HTMLDivElement | null>(null);
  const orderTagMenuRef = useRef<HTMLDivElement | null>(null);

  const isCreating = createFetcher.state === "submitting";
  const isCreatingCustomer = customerCreateFetcher.state === "submitting";
  const shouldSearchCustomers = customerInputValue.trim().length >= 2;
  const isSearchingCustomers =
    shouldSearchCustomers &&
    (isCustomerSearchQueued || customerSearchFetcher.state !== "idle");
  const trimmedOrderTagQuery = orderTagInputValue.trim();
  const shouldSearchOrderTags = trimmedOrderTagQuery.length >= 1;
  const isSearchingOrderTags =
    shouldSearchOrderTags &&
    (isOrderTagSearchQueued ||
      orderTagSearchFetcher.state !== "idle" ||
      orderTagSearchQuery !== trimmedOrderTagQuery);
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
    const searchData = customerSearchFetcher.data;
    if (!searchData) return;
    setIsCustomerSearchQueued(false);

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
    if (query.length < 2) {
      setIsCustomerSearchQueued(false);
      return;
    }
    setIsCustomerSearchQueued(true);

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
    const query = orderTagInputValue.trim();

    if (query.length < 1) {
      setIsOrderTagSearchQueued(false);
      setOrderTagSearchQuery("");
      setOrderTagSearchResults([]);
      return;
    }

    setIsOrderTagSearchQueued(true);

    const timeout = window.setTimeout(() => {
      const formData = new FormData();
      formData.append("intent", "search_order_tags");
      formData.append("query", query);
      orderTagSearchFetcher.submit(formData, { method: "POST" });
    }, 250);

    return () => {
      window.clearTimeout(timeout);
    };
  }, [orderTagInputValue]);

  useEffect(() => {
    const searchData = orderTagSearchFetcher.data;
    if (!searchData) return;

    setIsOrderTagSearchQueued(false);
    setOrderTagSearchQuery(searchData.query);

    if (searchData.type === "error") {
      setOrderTagSearchResults([]);
      return;
    }

    setOrderTagSearchResults(searchData.tags);
  }, [orderTagSearchFetcher.data]);

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

  useEffect(() => {
    if (!isOrderTagDropdownOpen) return;

    updateOrderTagMenuPosition();

    const onWindowChange = () => {
      updateOrderTagMenuPosition();
    };

    window.addEventListener("resize", onWindowChange);
    window.addEventListener("scroll", onWindowChange, true);

    return () => {
      window.removeEventListener("resize", onWindowChange);
      window.removeEventListener("scroll", onWindowChange, true);
    };
  }, [isOrderTagDropdownOpen, orderTagInputValue]);

  useEffect(() => {
    if (!isOrderTagDropdownOpen) return;

    const onPointerDown = (event: PointerEvent) => {
      const picker = orderTagPickerRef.current;
      const menu = orderTagMenuRef.current;
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

      setIsOrderTagDropdownOpen(false);
      setOrderTagMenuRect(null);
    };

    document.addEventListener("pointerdown", onPointerDown, true);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown, true);
    };
  }, [isOrderTagDropdownOpen]);

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
  const customerMenuOptionCount = filteredCustomers.length + 1;

  useEffect(() => {
    if (!isCustomerDropdownOpen) return;

    setActiveCustomerMenuIndex((current) =>
      Math.max(0, Math.min(current, customerMenuOptionCount - 1)),
    );
  }, [isCustomerDropdownOpen, customerMenuOptionCount]);

  const filteredOrderTagSuggestions = useMemo(() => {
    const query = orderTagInputValue.trim().toLowerCase();

    if (!query) {
      return [];
    }

    const localRanked = rankOrderTagList(query, orderTagSuggestions, 200);
    const hasFreshServerResults =
      query.length >= 1 &&
      orderTagSearchQuery === orderTagInputValue.trim() &&
      orderTagSearchResults.length > 0;

    if (!hasFreshServerResults) {
      return localRanked.slice(0, 100);
    }

    const merged = mergeUniqueTags(orderTagSearchResults, localRanked);
    return rankOrderTagList(query, merged, 100);
  }, [
    orderTagInputValue,
    orderTagSearchQuery,
    orderTagSearchResults,
    orderTagSuggestions,
  ]);

  const canAddTypedOrderTag = useMemo(() => {
    const value = orderTagInputValue.trim();
    if (!value) return false;

    const normalized = value.toLowerCase();
    const alreadySelected = selectedOrderTags.some(
      (tag) => tag.toLowerCase() === normalized,
    );
    if (alreadySelected) return false;

    const alreadyExists =
      orderTagSuggestions.some((tag) => tag.toLowerCase() === normalized) ||
      orderTagSearchResults.some((tag) => tag.toLowerCase() === normalized);
    return !alreadyExists;
  }, [
    orderTagInputValue,
    orderTagSearchResults,
    orderTagSuggestions,
    selectedOrderTags,
  ]);

  const updateCustomerMenuPosition = () => {
    const anchor = customerPickerRef.current;
    if (!anchor) {
      setCustomerMenuRect(null);
      return;
    }

    const rect = anchor.getBoundingClientRect();
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

  const updateOrderTagMenuPosition = () => {
    const anchor = orderTagPickerRef.current;
    if (!anchor) {
      setOrderTagMenuRect(null);
      return;
    }

    const rect = anchor.getBoundingClientRect();
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

    setOrderTagMenuRect({
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
    setParsedTableQuery("");
    setCreatedTableQuery("");

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
    const query = value.trim();
    setCustomerInputValue(value);
    setIsCustomerSearchQueued(query.length >= 2);
    if (query.length === 0) {
      setIsCustomerDropdownOpen(false);
      setCustomerMenuRect(null);
    } else {
      setIsCustomerDropdownOpen(true);
      setActiveCustomerMenuIndex(0);
      updateCustomerMenuPosition();
    }
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

  const onCustomerFieldFocus = () => {
    setIsCustomerDropdownOpen(true);
    setActiveCustomerMenuIndex(0);
    updateCustomerMenuPosition();
  };

  const onCustomerFieldKeyDown = (event: KeyboardEvent) => {
    const key = event?.key;
    if (!key) return;

    if (key === "ArrowDown") {
      event.preventDefault();
      if (!isCustomerDropdownOpen) {
        setIsCustomerDropdownOpen(true);
        updateCustomerMenuPosition();
      }
      setActiveCustomerMenuIndex((current) =>
        Math.min(current + 1, customerMenuOptionCount - 1),
      );
      return;
    }

    if (key === "ArrowUp") {
      event.preventDefault();
      if (!isCustomerDropdownOpen) {
        setIsCustomerDropdownOpen(true);
        updateCustomerMenuPosition();
      }
      setActiveCustomerMenuIndex((current) => Math.max(current - 1, 0));
      return;
    }

    if (key === "Enter" && isCustomerDropdownOpen) {
      event.preventDefault();
      if (activeCustomerMenuIndex === 0) {
        onOpenCreateCustomerModal();
        return;
      }

      const selected = filteredCustomers[activeCustomerMenuIndex - 1];
      if (selected) {
        onSelectCustomer(selected);
      }
      return;
    }

    if (key === "Escape" && isCustomerDropdownOpen) {
      event.preventDefault();
      setIsCustomerDropdownOpen(false);
      setCustomerMenuRect(null);
    }
  };

  useEffect(() => {
    const onDocumentKeyDown = (event: KeyboardEvent) => {
      const picker = customerPickerRef.current;
      const target = event.target as Node | null;

      if (!picker || !target || !picker.contains(target)) {
        return;
      }

      onCustomerFieldKeyDown(event);
    };

    document.addEventListener("keydown", onDocumentKeyDown);
    return () => {
      document.removeEventListener("keydown", onDocumentKeyDown);
    };
  }, [
    isCustomerDropdownOpen,
    customerMenuOptionCount,
    activeCustomerMenuIndex,
    filteredCustomers,
  ]);

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

  const onOrderTagFieldInput = (event: Event) => {
    const target = event.currentTarget as HTMLElement & { value?: string };
    const value = target.value || "";
    const query = value.trim();
    setOrderTagInputValue(value);
    const shouldOpen = query.length > 0;
    setIsOrderTagDropdownOpen(shouldOpen);
    if (shouldOpen) {
      updateOrderTagMenuPosition();
    } else {
      setOrderTagMenuRect(null);
    }
  };

  const onOrderTagFieldFocus = () => {
    if (orderTagInputValue.trim().length > 0) {
      setIsOrderTagDropdownOpen(true);
      updateOrderTagMenuPosition();
    }
  };

  const onToggleOrderTag = (tag: string) => {
    setSelectedOrderTags((existing) => {
      const existingIndex = existing.findIndex(
        (item) => item.toLowerCase() === tag.toLowerCase(),
      );
      if (existingIndex >= 0) {
        return existing.filter((_, index) => index !== existingIndex);
      }
      return [...existing, tag];
    });
  };

  const onRemoveOrderTag = (tag: string) => {
    setSelectedOrderTags((existing) =>
      existing.filter((item) => item.toLowerCase() !== tag.toLowerCase()),
    );
  };

  const onAddTypedOrderTag = () => {
    const value = orderTagInputValue.trim();
    if (!value) return;

    setSelectedOrderTags((existing) => {
      if (existing.some((item) => item.toLowerCase() === value.toLowerCase())) {
        return existing;
      }
      return [...existing, value];
    });
    setOrderTagInputValue("");
    setIsOrderTagDropdownOpen(false);
    setOrderTagMenuRect(null);
  };

  useEffect(() => {
    const onDocumentKeyDown = (event: KeyboardEvent) => {
      const picker = orderTagPickerRef.current;
      const target = event.target as Node | null;

      if (!picker || !target || !picker.contains(target)) {
        return;
      }

      if (event.key === "Escape" && isOrderTagDropdownOpen) {
        event.preventDefault();
        setIsOrderTagDropdownOpen(false);
        setOrderTagMenuRect(null);
      }

      if (event.key === "Enter") {
        if (canAddTypedOrderTag) {
          event.preventDefault();
          onAddTypedOrderTag();
        }
      }
    };

    document.addEventListener("keydown", onDocumentKeyDown);
    return () => {
      document.removeEventListener("keydown", onDocumentKeyDown);
    };
  }, [canAddTypedOrderTag, isOrderTagDropdownOpen, onAddTypedOrderTag]);

  const onCreateOrders = () => {
    if (!parseResult || !canCreate) return;

    const globalTags = selectedOrderTags.length > 0 ? selectedOrderTags : undefined;

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

  const summary = importResult?.summary;
  const results = importResult?.results ?? [];
  const filteredPreviewRows = useMemo(
    () => filterPreviewRows(previewRows, parsedTableQuery),
    [previewRows, parsedTableQuery],
  );
  const filteredResultRows = useMemo(
    () => filterResultRows(results, createdTableQuery),
    [results, createdTableQuery],
  );

  if (!isHydrated) {
    return <div>Loading CSV importer...</div>;
  }

  return (
    <s-page heading="CSV Order Importer">
      {parseResult ? (
        <s-button
          slot="primary-action"
          variant="primary"
          onClick={onCreateOrders}
          {...(isCreating ? { loading: true } : {})}
          disabled={!canCreate}
        >
          Create Orders
        </s-button>
      ) : null}

      <style>{CUSTOMER_PICKER_STYLES}</style>

      {importResult?.error ? (
        <s-banner tone="critical">{importResult.error}</s-banner>
      ) : null}

      {parseNotice ? <s-banner tone="warning">{parseNotice}</s-banner> : null}

      {customerLoadWarning ? (
        <s-banner tone="warning">{customerLoadWarning}</s-banner>
      ) : null}

      {orderTagLoadWarning ? (
        <s-banner tone="warning">{orderTagLoadWarning}</s-banner>
      ) : null}

      {customerError ? <s-banner tone="critical">{customerError}</s-banner> : null}

      {customerMessage ? (
        <s-banner tone={customerMessageTone}>{customerMessage}</s-banner>
      ) : null}

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
                  <s-text color="subdued">
                    {filteredPreviewRows.length} of {parseResult.rowCount} rows
                    shown
                  </s-text>
                </s-box>
                <s-table>
                  <s-search-field
                    slot="filters"
                    label="Search parsed rows"
                    placeholder="Search by name, address, city, or ZIP"
                    value={parsedTableQuery}
                    onInput={(event: Event) => {
                      const target = event.currentTarget as HTMLElement & {
                        value?: string;
                      };
                      setParsedTableQuery(target.value || "");
                    }}
                  />
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
                    {filteredPreviewRows.length === 0 ? (
                      <s-table-row>
                        <s-table-cell>No matching rows.</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                      </s-table-row>
                    ) : (
                      filteredPreviewRows.map((row) => {
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
                      })
                    )}
                  </s-table-body>
                </s-table>
              </s-section>
            ) : null}

            {results.length > 0 ? (
              <s-section heading="Created Orders" padding="none">
                <s-table>
                  <s-search-field
                    slot="filters"
                    label="Search created orders"
                    placeholder="Search by recipient, status, or error"
                    value={createdTableQuery}
                    onInput={(event: Event) => {
                      const target = event.currentTarget as HTMLElement & {
                        value?: string;
                      };
                      setCreatedTableQuery(target.value || "");
                    }}
                  />
                  <s-table-header-row>
                    <s-table-header listSlot="kicker">Row</s-table-header>
                    <s-table-header listSlot="primary">Recipient</s-table-header>
                    <s-table-header listSlot="labeled">Status</s-table-header>
                    <s-table-header listSlot="labeled">Error</s-table-header>
                  </s-table-header-row>
                  <s-table-body>
                    {filteredResultRows.length === 0 ? (
                      <s-table-row>
                        <s-table-cell>No matching rows.</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                        <s-table-cell>-</s-table-cell>
                      </s-table-row>
                    ) : (
                      filteredResultRows.map((row) => (
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
                      ))
                    )}
                  </s-table-body>
                </s-table>
              </s-section>
            ) : null}

          </s-stack>
        </div>

        <div className="Polaris-Layout__Section Polaris-Layout__Section--oneThird">
          <s-stack direction="block" gap="base">
            <s-section heading="Shipping Report">
              <s-stack gap="base" direction="block">
                <s-text color="subdued">
                  You will be able to select the particular data items to export.
                </s-text>
                <s-button variant="primary" href="/app/report">
                  New Report
                </s-button>
              </s-stack>
            </s-section>

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

                <div className="order-tag-picker" ref={orderTagPickerRef}>
                  <s-search-field
                    label="Order Tags (optional)"
                    placeholder="Search tags or add a new tag"
                    value={orderTagInputValue}
                    onInput={onOrderTagFieldInput}
                    onFocus={onOrderTagFieldFocus}
                  />
                </div>
                {selectedOrderTags.length > 0 ? (
                  <s-stack direction="inline" gap="small">
                    {selectedOrderTags.map((tag) => (
                      <s-clickable-chip
                        key={tag}
                        color="strong"
                        accessibilityLabel={`Remove tag ${tag}`}
                        removable
                        onRemove={() => onRemoveOrderTag(tag)}
                      >
                        {tag}
                      </s-clickable-chip>
                    ))}
                  </s-stack>
                ) : null}
                <s-text color="subdued">
                  Add tags to all imported orders.
                </s-text>
              </s-stack>
            </s-section>
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
            <div
              data-customer-row="true"
              data-active={activeCustomerMenuIndex === 0 ? "true" : "false"}
              onMouseEnter={() => setActiveCustomerMenuIndex(0)}
            >
              <s-clickable onClick={onOpenCreateCustomerModal} padding="small">
                <span className="customer-picker__create-content">
                  <span className="customer-picker__create-icon" aria-hidden="true">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">
                      <path d="M4.25 8a.75.75 0 0 1 .75-.75h2.25v-2.25a.75.75 0 0 1 1.5 0v2.25h2.25a.75.75 0 0 1 0 1.5h-2.25v2.25a.75.75 0 0 1-1.5 0v-2.25h-2.25a.75.75 0 0 1-.75-.75" />
                      <path
                        fillRule="evenodd"
                        d="M8 15a7 7 0 1 0 0-14 7 7 0 0 0 0 14m0-1.5a5.5 5.5 0 1 0 0-11 5.5 5.5 0 1 0 0 11"
                      />
                    </svg>
                  </span>
                  <s-text>Create a new customer</s-text>
                </span>
              </s-clickable>
            </div>
            <s-divider />
            <div className="customer-picker__list">
              {isSearchingCustomers ? (
                <div className="customer-picker__loading">
                  <s-stack direction="inline" gap="small">
                    <s-spinner
                      accessibilityLabel="Searching customers"
                      size="base"
                    />
                    <s-text color="subdued">Searching customers...</s-text>
                  </s-stack>
                </div>
              ) : null}
              {!isSearchingCustomers && filteredCustomers.length === 0 ? (
                <s-box padding="small">
                  <s-text color="subdued">No matching customers.</s-text>
                </s-box>
              ) : (
                filteredCustomers.map((customer, index) => {
                  const customerMenuIndex = index + 1;
                  return (
                  <s-box key={customer.id}>
                    <div
                      data-customer-row="true"
                      data-active={
                        activeCustomerMenuIndex === customerMenuIndex ? "true" : "false"
                      }
                      onMouseEnter={() =>
                        setActiveCustomerMenuIndex(customerMenuIndex)
                      }
                    >
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
                    </div>
                  </s-box>
                  );
                })
              )}
            </div>
          </div>,
          document.body,
        )
      ) : null}

      {isHydrated && isOrderTagDropdownOpen && orderTagMenuRect ? (
        createPortal(
          <div
            ref={orderTagMenuRef}
            className="order-tag-picker__menu"
            style={{
              top: `${orderTagMenuRect.top}px`,
              left: `${orderTagMenuRect.left}px`,
              width: `${orderTagMenuRect.width}px`,
            }}
            onMouseDown={(event) => {
              event.preventDefault();
            }}
          >
            {canAddTypedOrderTag ? (
              <>
                <div data-order-tag-row="true">
                  <s-clickable onClick={onAddTypedOrderTag} padding="small">
                    <span className="order-tag-picker__action-content">
                      <span className="order-tag-picker__icon" aria-hidden="true">
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">
                          <path d="M4.25 8a.75.75 0 0 1 .75-.75h2.25v-2.25a.75.75 0 0 1 1.5 0v2.25h2.25a.75.75 0 0 1 0 1.5h-2.25v2.25a.75.75 0 0 1-1.5 0v-2.25h-2.25a.75.75 0 0 1-.75-.75" />
                          <path
                            fillRule="evenodd"
                            d="M8 15a7 7 0 1 0 0-14 7 7 0 0 0 0 14m0-1.5a5.5 5.5 0 1 0 0-11 5.5 5.5 0 1 0 0 11"
                          />
                        </svg>
                      </span>
                      <s-text>
                        <s-text type="strong">Add</s-text> {orderTagInputValue.trim()}
                      </s-text>
                    </span>
                  </s-clickable>
                </div>
                <s-divider />
              </>
            ) : null}
            <div className="order-tag-picker__list">
              {isSearchingOrderTags ? (
                <div className="customer-picker__loading">
                  <s-stack direction="inline" gap="small">
                    <s-spinner
                      accessibilityLabel="Searching order tags"
                      size="base"
                    />
                    <s-text color="subdued">Searching tags...</s-text>
                  </s-stack>
                </div>
              ) : null}
              {!isSearchingOrderTags && filteredOrderTagSuggestions.length === 0 ? (
                <s-box padding="small">
                  <s-text color="subdued">No matching tags.</s-text>
                </s-box>
              ) : (
                filteredOrderTagSuggestions.map((tag) => {
                  const isSelected = selectedOrderTags.some(
                    (value) => value.toLowerCase() === tag.toLowerCase(),
                  );
                  return (
                    <div key={tag} data-order-tag-row="true">
                      <s-clickable onClick={() => onToggleOrderTag(tag)} padding="small">
                        <s-stack direction="inline" gap="small">
                          <span
                            className={`order-tag-picker__checkbox${isSelected ? " order-tag-picker__checkbox--checked" : ""}`}
                            aria-hidden="true"
                          >
                            {isSelected ? (
                              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">
                                <path d="M6.53 10.78a.75.75 0 0 1-1.06 0l-1.75-1.75a.75.75 0 0 1 1.06-1.06l1.22 1.22 3.22-3.22a.75.75 0 0 1 1.06 1.06z" />
                              </svg>
                            ) : null}
                          </span>
                          <s-text>{tag}</s-text>
                        </s-stack>
                      </s-clickable>
                    </div>
                  );
                })
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

async function searchOrderTags(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  formData: FormData,
): Promise<OrderTagSearchActionData> {
  const queryRaw = (formData.get("query") as string | null)?.trim() || "";
  if (queryRaw.length < 1) {
    return {
      type: "search_order_tags",
      query: queryRaw,
      tags: [],
    };
  }

  try {
    const tagResult = await loadShopOrderTags(admin, queryRaw);
    if (tagResult.error) {
      return {
        type: "error",
        query: queryRaw,
        error: tagResult.error,
      };
    }

    return {
      type: "search_order_tags",
      query: queryRaw,
      tags: rankOrderTagList(queryRaw, tagResult.tags, 100),
    };
  } catch (error) {
    return {
      type: "error",
      query: queryRaw,
      error:
        error instanceof Error ? error.message : "Failed to search order tags.",
    };
  }
}

async function loadShopOrderTags(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  rawQuery?: string,
): Promise<{ tags: string[]; error?: string }> {
  const query = rawQuery?.trim() || "";
  if (query) {
    let draftSearchTags: string[] = [];
    let draftBaselineTags: string[] = [];
    let searchTags: string[] = [];
    let baselineTags: string[] = [];
    const searchQuery = `title:${escapeSearchToken(query.toLowerCase())}*`;

    try {
      const searchedDraft = await admin.graphql(SEARCH_DRAFT_ORDER_TAGS_QUERY, {
        variables: { query: searchQuery },
      });
      const searchedDraftJson = (await searchedDraft.json()) as {
        errors?: Array<{ message: string }>;
        data?: {
          shop?: {
            draftOrderTags?: {
              edges?: Array<{
                node?: {
                  title?: string | null;
                } | null;
              }>;
            } | null;
          } | null;
        };
      };

      if (!searchedDraftJson.errors || searchedDraftJson.errors.length === 0) {
        draftSearchTags = readDraftTagConnectionValues(
          searchedDraftJson.data?.shop?.draftOrderTags,
        );
      }
    } catch {
      // Continue to fallback logic below.
    }

    try {
      const baselineDraftResponse = await admin.graphql(
        DRAFT_ORDER_TAG_SUGGESTIONS_QUERY,
      );
      const baselineDraftJson = (await baselineDraftResponse.json()) as {
        errors?: Array<{ message: string }>;
        data?: {
          shop?: {
            draftOrderTags?: {
              edges?: Array<{
                node?: {
                  title?: string | null;
                } | null;
              }>;
            } | null;
          } | null;
        };
      };

      if (!baselineDraftJson.errors || baselineDraftJson.errors.length === 0) {
        draftBaselineTags = readDraftTagConnectionValues(
          baselineDraftJson.data?.shop?.draftOrderTags,
        );
      }
    } catch {
      // Continue to fallback logic below.
    }

    try {
      const searched = await admin.graphql(SEARCH_ORDER_TAGS_QUERY, {
        variables: { query: searchQuery },
      });
      const searchedJson = (await searched.json()) as {
        errors?: Array<{ message: string }>;
        data?: {
          shop?: {
            orderTags?: {
              edges?: Array<{
                node?: string | null;
              }>;
            } | null;
          } | null;
        };
      };

      if (!searchedJson.errors || searchedJson.errors.length === 0) {
        searchTags = readTagConnectionValues(searchedJson.data?.shop?.orderTags);
      }
    } catch {
      // Continue to baseline/fallback logic below.
    }

    try {
      const baselineResponse = await admin.graphql(ORDER_TAG_SUGGESTIONS_QUERY);
      const baselineJson = (await baselineResponse.json()) as {
        errors?: Array<{ message: string }>;
        data?: {
          shop?: {
            orderTags?: {
              edges?: Array<{
                node?: string | null;
              }>;
            } | null;
          } | null;
        };
      };

      if (!baselineJson.errors || baselineJson.errors.length === 0) {
        baselineTags = readTagConnectionValues(baselineJson.data?.shop?.orderTags);
      }
    } catch {
      // Continue to fallback logic below.
    }

    const draftOrderTagsFromDraftOrders = await loadOrderTagsFromDraftOrders(
      admin,
      query,
    );
    const merged = mergeUniqueTags(
      draftSearchTags,
      draftBaselineTags,
      searchTags,
      baselineTags,
      draftOrderTagsFromDraftOrders.tags,
    );
    const ranked = rankOrderTagList(query, merged, 100);
    if (ranked.length > 0) {
      return { tags: ranked };
    }

    const orderTagsFromOrders = await loadOrderTagsFromOrders(admin, query);
    if (orderTagsFromOrders.tags.length > 0) {
      return {
        tags: rankOrderTagList(
          query,
          mergeUniqueTags(merged, orderTagsFromOrders.tags),
          100,
        ),
      };
    }

    if (draftOrderTagsFromDraftOrders.error && orderTagsFromOrders.error) {
      return {
        tags: [],
        error: `${draftOrderTagsFromDraftOrders.error}; ${orderTagsFromOrders.error}`,
      };
    }

    if (orderTagsFromOrders.error) {
      return { tags: [], error: orderTagsFromOrders.error };
    }

    if (draftOrderTagsFromDraftOrders.error) {
      return { tags: [], error: draftOrderTagsFromDraftOrders.error };
    }

    return { tags: [] };
  }

  try {
    const draftResponse = await admin.graphql(DRAFT_ORDER_TAG_SUGGESTIONS_QUERY);
    const draftJson = (await draftResponse.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        shop?: {
          draftOrderTags?: {
            edges?: Array<{
              node?: {
                title?: string | null;
              } | null;
            }>;
          } | null;
        } | null;
      };
    };

    if (!draftJson.errors || draftJson.errors.length === 0) {
      const tags = readDraftTagConnectionValues(draftJson.data?.shop?.draftOrderTags);
      if (tags.length > 0) {
        return { tags };
      }
    }
  } catch {
    // Fall through to order-tags logic below.
  }

  try {
    const response = await admin.graphql(ORDER_TAG_SUGGESTIONS_QUERY);
    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        shop?: {
          orderTags?: {
            edges?: Array<{
              node?: string | null;
            }>;
          } | null;
        } | null;
      };
    };

    if (!json.errors || json.errors.length === 0) {
      return {
        tags: readTagConnectionValues(json.data?.shop?.orderTags),
      };
    }
  } catch {
    // Fall through to orders-based fallback below.
  }

  const draftOrderTagsFromDraftOrders = await loadOrderTagsFromDraftOrders(admin);
  if (draftOrderTagsFromDraftOrders.tags.length > 0) {
    return draftOrderTagsFromDraftOrders;
  }

  const orderTagsFromOrders = await loadOrderTagsFromOrders(admin);
  if (orderTagsFromOrders.tags.length > 0) {
    return orderTagsFromOrders;
  }

  if (draftOrderTagsFromDraftOrders.error && orderTagsFromOrders.error) {
    return {
      tags: [],
      error: `${draftOrderTagsFromDraftOrders.error}; ${orderTagsFromOrders.error}`,
    };
  }

  if (orderTagsFromOrders.error) {
    return { tags: [], error: orderTagsFromOrders.error };
  }

  if (draftOrderTagsFromDraftOrders.error) {
    return { tags: [], error: draftOrderTagsFromDraftOrders.error };
  }

  return { tags: [] };
}

async function loadOrderTagsFromDraftOrders(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  rawQuery?: string,
): Promise<{ tags: string[]; error?: string }> {
  const query = rawQuery?.trim() || "";
  const searchQuery =
    query.length > 0
      ? query
          .split(/\s+/)
          .map((token) => token.trim())
          .filter(Boolean)
          .map((token) => `tag:${escapeSearchToken(token)}*`)
          .join(" AND ")
      : "";

  const tagCounts = new Map<string, number>();
  let after: string | null = null;
  let hasNextPage = true;
  let pagesLoaded = 0;

  while (hasNextPage && pagesLoaded < ORDER_TAG_SCAN_MAX_PAGES) {
    const response = await admin.graphql(
      query
        ? SEARCH_ORDER_TAGS_BY_DRAFT_ORDERS_QUERY
        : ORDER_TAG_SUGGESTIONS_BY_DRAFT_ORDERS_QUERY,
      {
        variables: query ? { query: searchQuery, after } : { after },
      },
    );
    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        draftOrders?: {
          edges?: Array<{
            node?: {
              tags?: string[] | null;
            };
          }>;
          pageInfo?: {
            hasNextPage?: boolean;
            endCursor?: string | null;
          };
        };
      };
    };

    if (json.errors && json.errors.length > 0) {
      return {
        tags: [],
        error: json.errors.map((error) => error.message).join("; "),
      };
    }

    for (const edge of json.data?.draftOrders?.edges ?? []) {
      for (const tag of edge.node?.tags ?? []) {
        const normalized = tag.trim();
        if (!normalized) continue;
        tagCounts.set(normalized, (tagCounts.get(normalized) ?? 0) + 1);
      }
    }

    pagesLoaded += 1;
    const pageInfo = json.data?.draftOrders?.pageInfo;
    hasNextPage = Boolean(pageInfo?.hasNextPage);
    after = pageInfo?.endCursor ?? null;
    if (!after) {
      hasNextPage = false;
    }
  }

  const sorted = Array.from(tagCounts.entries())
    .sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1];
      return a[0].localeCompare(b[0], "en");
    })
    .map(([tag]) => tag);

  return {
    tags: query ? rankOrderTagList(query, sorted, 100) : sorted.slice(0, 500),
  };
}

async function loadOrderTagsFromOrders(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  rawQuery?: string,
): Promise<{ tags: string[]; error?: string }> {
  const query = rawQuery?.trim() || "";
  const searchQuery =
    query.length > 0
      ? query
          .split(/\s+/)
          .map((token) => token.trim())
          .filter(Boolean)
          .map((token) => `tag:${escapeSearchToken(token)}*`)
          .join(" AND ")
      : "";

  const tagCounts = new Map<string, number>();
  let after: string | null = null;
  let hasNextPage = true;
  let pagesLoaded = 0;

  while (hasNextPage && pagesLoaded < ORDER_TAG_SCAN_MAX_PAGES) {
    const response = await admin.graphql(
      query ? SEARCH_ORDER_TAGS_BY_ORDERS_QUERY : ORDER_TAG_SUGGESTIONS_BY_ORDERS_QUERY,
      {
        variables: query ? { query: searchQuery, after } : { after },
      },
    );
    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        orders?: {
          edges?: Array<{
            node?: {
              tags?: string[] | null;
            };
          }>;
          pageInfo?: {
            hasNextPage?: boolean;
            endCursor?: string | null;
          };
        };
      };
    };

    if (json.errors && json.errors.length > 0) {
      return {
        tags: [],
        error: json.errors.map((error) => error.message).join("; "),
      };
    }

    for (const edge of json.data?.orders?.edges ?? []) {
      for (const tag of edge.node?.tags ?? []) {
        const normalized = tag.trim();
        if (!normalized) continue;
        tagCounts.set(normalized, (tagCounts.get(normalized) ?? 0) + 1);
      }
    }

    pagesLoaded += 1;
    const pageInfo = json.data?.orders?.pageInfo;
    hasNextPage = Boolean(pageInfo?.hasNextPage);
    after = pageInfo?.endCursor ?? null;
    if (!after) {
      hasNextPage = false;
    }
  }

  const sorted = Array.from(tagCounts.entries())
    .sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1];
      return a[0].localeCompare(b[0], "en");
    })
    .map(([tag]) => tag);

  return {
    tags: query ? rankOrderTagList(query, sorted, 100) : sorted.slice(0, 500),
  };
}

function readTagConnectionValues(
  connection:
    | {
        edges?: Array<{
          node?: string | null;
        }>;
      }
    | null
    | undefined,
): string[] {
  const seen = new Set<string>();
  const values: string[] = [];

  for (const edge of connection?.edges ?? []) {
    const tag = edge.node?.trim();
    if (!tag) continue;
    const key = tag.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    values.push(tag);
  }

  return values;
}

function readDraftTagConnectionValues(
  connection:
    | {
        edges?: Array<{
          node?: {
            title?: string | null;
          } | null;
        }>;
      }
    | null
    | undefined,
): string[] {
  const seen = new Set<string>();
  const values: string[] = [];

  for (const edge of connection?.edges ?? []) {
    const tag = edge.node?.title?.trim();
    if (!tag) continue;
    const key = tag.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    values.push(tag);
  }

  return values;
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

function filterPreviewRows(rows: CsvPreviewRow[], query: string): CsvPreviewRow[] {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) return rows;

  return rows.filter((row) =>
    buildSearchText([
      row.rowNumber,
      row.recipient,
      row.address,
      row.address2,
      row.city,
      row.state,
      row.zipCode,
      row.email,
    ]).includes(normalizedQuery),
  );
}

function filterResultRows(rows: RowResult[], query: string): RowResult[] {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) return rows;

  return rows.filter((row) =>
    buildSearchText([
      row.rowNumber,
      row.recipient,
      row.address,
      row.address2,
      row.city,
      row.state,
      row.zipCode,
      row.email,
      row.status,
      row.errorMessage,
    ]).includes(normalizedQuery),
  );
}

function buildSearchText(values: Array<string | number | undefined>): string {
  return values
    .filter((value) => value !== undefined && value !== null)
    .map((value) => String(value).trim().toLowerCase())
    .join(" ");
}

function escapeSearchValue(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

function escapeSearchToken(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/([:\\()])/g, "\\$1");
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

function rankOrderTagList(query: string, tags: string[], max: number): string[] {
  const normalizedQuery = query.trim().toLowerCase();
  const unique = new Set<string>();
  const deduped: string[] = [];

  for (const rawTag of tags) {
    const tag = rawTag.trim();
    if (!tag) continue;
    const key = tag.toLowerCase();
    if (unique.has(key)) continue;
    unique.add(key);
    deduped.push(tag);
  }

  if (!normalizedQuery) {
    return deduped.slice(0, max);
  }

  const startsWith: string[] = [];

  for (const tag of deduped) {
    const lower = tag.toLowerCase();
    if (lower.startsWith(normalizedQuery)) {
      startsWith.push(tag);
    }
  }

  return startsWith.slice(0, max);
}

function mergeUniqueTags(...groups: string[][]): string[] {
  const seen = new Set<string>();
  const merged: string[] = [];

  for (const group of groups) {
    for (const rawTag of group) {
      const tag = rawTag.trim();
      if (!tag) continue;
      const key = tag.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      merged.push(tag);
    }
  }

  return merged;
}

function formatCustomerOption(customer: CustomerOption): string {
  const name = customer.displayName || customer.email || "Unnamed customer";
  return customer.email ? `${name} (${customer.email})` : name;
}

function extractEmailCandidate(value: string): string {
  const match = value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0].toLowerCase() : "";
}
