import { useEffect, useMemo, useState } from "react";
import type {
  ActionFunctionArgs,
  HeadersFunction,
  LoaderFunctionArgs,
} from "react-router";
import { useFetcher, useLoaderData } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";
import { authenticate } from "../shopify.server";

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

const REPORT_TAG_SUGGESTIONS_QUERY = `#graphql
  query ReportTagSuggestions {
    orders(first: 100, sortKey: CREATED_AT, reverse: true) {
      edges {
        node {
          tags
        }
      }
    }
  }
`;

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);

  const reportTagSuggestions: string[] = [];
  let reportTagLoadWarning = "";

  try {
    const response = await admin.graphql(REPORT_TAG_SUGGESTIONS_QUERY);
    const json = (await response.json()) as {
      errors?: Array<{ message: string }>;
      data?: {
        orders?: {
          edges?: Array<{
            node?: {
              tags?: string[] | null;
            };
          }>;
        };
      };
    };

    if (json.errors && json.errors.length > 0) {
      reportTagLoadWarning = `Unable to load report tag suggestions: ${json.errors
        .map((error) => error.message)
        .join("; ")}`;
    } else {
      const tagCounts = new Map<string, number>();
      for (const edge of json.data?.orders?.edges ?? []) {
        for (const tag of edge.node?.tags ?? []) {
          const normalized = tag.trim();
          if (!normalized) continue;
          tagCounts.set(normalized, (tagCounts.get(normalized) ?? 0) + 1);
        }
      }

      const sortedTags = Array.from(tagCounts.entries())
        .sort((a, b) => {
          if (b[1] !== a[1]) return b[1] - a[1];
          return a[0].localeCompare(b[0], "en");
        })
        .map(([tag]) => tag)
        .slice(0, 100);

      reportTagSuggestions.push(...sortedTags);
    }
  } catch {
    reportTagLoadWarning =
      "Unable to load report tag suggestions right now. You can still type tags manually.";
  }

  return {
    reportTagSuggestions,
    reportTagLoadWarning,
  };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const { admin } = await authenticate.admin(request);
  const formData = await request.formData();
  const intent = formData.get("intent");

  if (intent !== "generate_report") {
    return { error: "Invalid report request." } satisfies ReportActionData;
  }

  return generateShippingReport(admin, formData);
};

export default function ReportPage() {
  const { reportTagSuggestions, reportTagLoadWarning } =
    useLoaderData<typeof loader>();
  const reportFetcher = useFetcher<ReportActionData>();

  const [orderNumbersInput, setOrderNumbersInput] = useState("");
  const [searchTagsInput, setSearchTagsInput] = useState("");
  const [selectedReportTagSuggestion, setSelectedReportTagSuggestion] =
    useState("");
  const [reportOrders, setReportOrders] = useState<ShippingReportOrder[]>([]);
  const [reportNotice, setReportNotice] = useState("");
  const [reportError, setReportError] = useState("");
  const [reportTableQuery, setReportTableQuery] = useState("");

  const isGeneratingReport = reportFetcher.state === "submitting";

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

  const filteredReportRows = useMemo(
    () => filterReportRows(reportOrders, reportTableQuery),
    [reportOrders, reportTableQuery],
  );

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

  const onSelectReportTagSuggestion = (event: Event) => {
    const target = event.currentTarget as HTMLElement & { value?: string };
    const selectedTag = (target.value || "").trim();
    setSelectedReportTagSuggestion(selectedTag);
    if (!selectedTag) return;

    setSearchTagsInput((existing) => appendUniqueCommaValue(existing, selectedTag));
    setSelectedReportTagSuggestion("");
  };

  return (
    <s-page heading="Shipping Report">
      <s-button
        slot="primary-action"
        variant="primary"
        onClick={onExportReport}
        disabled={reportOrders.length === 0}
      >
        Export
      </s-button>

      {reportTagLoadWarning ? (
        <s-banner tone="warning">{reportTagLoadWarning}</s-banner>
      ) : null}

      {reportError ? <s-banner tone="critical">{reportError}</s-banner> : null}

      {reportNotice ? <s-banner tone="success">{reportNotice}</s-banner> : null}

      <s-stack direction="block" gap="base">
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

            {reportTagSuggestions.length > 0 ? (
              <>
                <s-select
                  label="Available tags"
                  value={selectedReportTagSuggestion}
                  onInput={onSelectReportTagSuggestion}
                >
                  <s-option value="">Select a tag</s-option>
                  {reportTagSuggestions.map((tag) => (
                    <s-option key={tag} value={tag}>
                      {tag}
                    </s-option>
                  ))}
                </s-select>
                <s-text color="subdued">
                  Selecting a tag adds it to the report tag filter.
                </s-text>
              </>
            ) : null}

            {searchTagsInput ? (
              <s-stack direction="inline" gap="base" justifyContent="space-between">
                <s-text color="subdued">Selected tags: {searchTagsInput}</s-text>
                <s-button
                  variant="secondary"
                  onClick={() => {
                    setSearchTagsInput("");
                    setSelectedReportTagSuggestion("");
                  }}
                >
                  Clear tags
                </s-button>
              </s-stack>
            ) : null}

            <s-stack direction="inline" gap="base">
              <s-button
                variant="primary"
                onClick={onGenerateReport}
                {...(isGeneratingReport ? { loading: true } : {})}
              >
                Generate Report
              </s-button>
            </s-stack>
          </s-stack>
        </s-section>

        {reportOrders.length > 0 ? (
          <s-section heading="Shipping Report Results" padding="none">
            <s-table>
              <s-search-field
                slot="filters"
                label="Search shipping report results"
                placeholder="Search by order number, customer, or tracking"
                value={reportTableQuery}
                onInput={(event: Event) => {
                  const target = event.currentTarget as HTMLElement & {
                    value?: string;
                  };
                  setReportTableQuery(target.value || "");
                }}
              />
              <s-table-header-row>
                <s-table-header listSlot="primary">Order Number</s-table-header>
                <s-table-header listSlot="labeled">Customer Name</s-table-header>
                <s-table-header listSlot="labeled">Tracking Numbers</s-table-header>
              </s-table-header-row>
              <s-table-body>
                {filteredReportRows.length === 0 ? (
                  <s-table-row>
                    <s-table-cell>No matching orders.</s-table-cell>
                    <s-table-cell>-</s-table-cell>
                    <s-table-cell>-</s-table-cell>
                  </s-table-row>
                ) : (
                  filteredReportRows.map((order) => (
                    <s-table-row key={order.id}>
                      <s-table-cell>{order.name}</s-table-cell>
                      <s-table-cell>{order.customerName}</s-table-cell>
                      <s-table-cell>
                        {order.trackingNumbers.length > 0
                          ? order.trackingNumbers.join(", ")
                          : "No tracking"}
                      </s-table-cell>
                    </s-table-row>
                  ))
                )}
              </s-table-body>
            </s-table>
          </s-section>
        ) : null}

        <s-stack direction="inline" gap="base" justifyContent="end">
          <s-button variant="secondary" tone="critical" href="/app">
            Cancel
          </s-button>
          <s-button
            variant="primary"
            onClick={onExportReport}
            disabled={reportOrders.length === 0}
          >
            Export
          </s-button>
        </s-stack>
      </s-stack>
    </s-page>
  );
}

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};

async function generateShippingReport(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  formData: FormData,
): Promise<ReportActionData> {
  const orderNumbers = (formData.get("orderNumbers") as string | null)?.trim() || "";
  const searchTags = (formData.get("searchTags") as string | null)?.trim() || "";

  const query = buildOrderSearchQuery({
    orderNumbers,
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

function buildOrderSearchQuery(filters: {
  orderNumbers: string;
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

function appendUniqueCommaValue(existing: string, nextValue: string): string {
  const currentValues = splitCommaValues(existing);
  const nextNormalized = nextValue.trim();
  if (!nextNormalized) return existing;

  const hasValue = currentValues.some(
    (item) => item.toLowerCase() === nextNormalized.toLowerCase(),
  );

  if (hasValue) return currentValues.join(", ");
  return [...currentValues, nextNormalized].join(", ");
}

function filterReportRows(
  rows: ShippingReportOrder[],
  query: string,
): ShippingReportOrder[] {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) return rows;

  return rows.filter((row) =>
    [row.name, row.customerName, ...row.trackingNumbers]
      .join(" ")
      .toLowerCase()
      .includes(normalizedQuery),
  );
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
