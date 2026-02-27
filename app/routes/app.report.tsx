import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
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

type ReportTagSearchActionData =
  | {
      type: "search_report_tags";
      query: string;
      tags: string[];
    }
  | {
      type: "error";
      query: string;
      error: string;
    };

const REPORT_TAG_SCAN_MAX_PAGES = 10;

const REPORT_TAG_PICKER_STYLES = `
.report-tag-picker {
  position: relative;
  z-index: 35;
}

.report-tag-picker__menu {
  position: fixed;
  z-index: 2147483646;
  background: var(--p-color-bg-surface, #fff);
  border: 1px solid var(--p-color-border, #d1d5db);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.12);
  overflow: hidden;
  max-height: min(320px, calc(100vh - 24px));
}

.report-tag-picker__list {
  max-height: 280px;
  overflow-y: auto;
}

.report-tag-picker__loading {
  padding: 8px 12px;
}

.report-tag-picker__menu [data-report-tag-row="true"]:hover {
  background: var(--p-color-bg-fill-tertiary, #f3f4f6);
}

.report-tag-picker__checkbox {
  width: 16px;
  height: 16px;
  border: 1px solid var(--p-color-border, #d1d5db);
  border-radius: 4px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex: 0 0 auto;
}

.report-tag-picker__checkbox--checked {
  border-color: var(--p-color-bg-fill-emphasis, #111827);
  background: var(--p-color-bg-fill-emphasis, #111827);
  color: var(--p-color-text-on-fill, #fff);
}

.report-tag-picker__checkbox svg {
  width: 12px;
  height: 12px;
  display: block;
  fill: currentColor;
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

const REPORT_TAG_SUGGESTIONS_QUERY = `#graphql
  query ReportTagSuggestions {
    shop {
      orderTags(first: 250, sortKey: POPULAR) {
        edges {
          node
        }
      }
    }
  }
`;

const SEARCH_REPORT_TAGS_QUERY = `#graphql
  query SearchReportTags($query: String!) {
    shop {
      orderTags(first: 250, sortKey: POPULAR, query: $query) {
        edges {
          node
        }
      }
    }
  }
`;

const REPORT_TAG_SUGGESTIONS_BY_ORDERS_QUERY = `#graphql
  query ReportTagSuggestionsByOrders($after: String) {
    orders(first: 250, after: $after, sortKey: CREATED_AT, reverse: true) {
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

const SEARCH_REPORT_TAGS_BY_ORDERS_QUERY = `#graphql
  query SearchReportTagsByOrders($query: String!, $after: String) {
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

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);

  const reportTagSuggestions: string[] = [];
  let reportTagLoadWarning = "";

  try {
    const tagResult = await loadShopReportTags(admin);
    if (tagResult.error) {
      reportTagLoadWarning = `Unable to load report tag suggestions: ${tagResult.error}`;
    } else {
      reportTagSuggestions.push(...tagResult.tags);
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

  if (intent === "search_report_tags") {
    return searchReportTags(admin, formData);
  }

  if (intent !== "generate_report") {
    return { error: "Invalid report request." } satisfies ReportActionData;
  }

  const orderNumbers =
    (formData.get("orderNumbers") as string | null)?.trim() || "";
  const searchTags = (formData.get("searchTags") as string | null)?.trim() || "";

  if (
    splitCommaValues(orderNumbers).length === 0 &&
    splitCommaValues(searchTags).length === 0
  ) {
    return {
      error: "Enter at least one order number or choose at least one tag.",
    } satisfies ReportActionData;
  }

  return generateShippingReport(admin, formData);
};

export default function ReportPage() {
  const { reportTagSuggestions, reportTagLoadWarning } =
    useLoaderData<typeof loader>();
  const reportFetcher = useFetcher<ReportActionData>();
  const reportTagSearchFetcher = useFetcher<ReportTagSearchActionData>();

  const [orderNumbersInput, setOrderNumbersInput] = useState("");
  const [selectedReportTags, setSelectedReportTags] = useState<string[]>([]);
  const [reportTagInputValue, setReportTagInputValue] = useState("");
  const [isReportTagDropdownOpen, setIsReportTagDropdownOpen] = useState(false);
  const [isReportTagSearchQueued, setIsReportTagSearchQueued] = useState(false);
  const [reportTagSearchQuery, setReportTagSearchQuery] = useState("");
  const [reportTagSearchResults, setReportTagSearchResults] = useState<string[]>(
    [],
  );
  const [reportOrders, setReportOrders] = useState<ShippingReportOrder[]>([]);
  const [reportNotice, setReportNotice] = useState("");
  const [reportError, setReportError] = useState("");
  const [reportTableQuery, setReportTableQuery] = useState("");
  const [isHydrated, setIsHydrated] = useState(false);
  const [reportTagMenuRect, setReportTagMenuRect] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);
  const reportTagPickerRef = useRef<HTMLDivElement | null>(null);
  const reportTagMenuRef = useRef<HTMLDivElement | null>(null);

  const isGeneratingReport = reportFetcher.state === "submitting";
  const trimmedReportTagQuery = reportTagInputValue.trim();
  const shouldSearchReportTags = trimmedReportTagQuery.length >= 1;
  const isSearchingReportTags =
    shouldSearchReportTags &&
    (isReportTagSearchQueued ||
      reportTagSearchFetcher.state !== "idle" ||
      reportTagSearchQuery !== trimmedReportTagQuery);
  const hasReportCriteria =
    splitCommaValues(orderNumbersInput).length > 0 ||
    selectedReportTags.length > 0;

  useEffect(() => {
    setIsHydrated(true);
  }, []);

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
    const query = reportTagInputValue.trim();

    if (query.length < 1) {
      setIsReportTagSearchQueued(false);
      setReportTagSearchQuery("");
      setReportTagSearchResults([]);
      return;
    }

    setIsReportTagSearchQueued(true);

    const timeout = window.setTimeout(() => {
      const formData = new FormData();
      formData.append("intent", "search_report_tags");
      formData.append("query", query);
      reportTagSearchFetcher.submit(formData, { method: "POST" });
    }, 250);

    return () => {
      window.clearTimeout(timeout);
    };
  }, [reportTagInputValue]);

  useEffect(() => {
    const searchData = reportTagSearchFetcher.data;
    if (!searchData) return;

    setIsReportTagSearchQueued(false);
    setReportTagSearchQuery(searchData.query);

    if (searchData.type === "error") {
      setReportTagSearchResults([]);
      return;
    }

    setReportTagSearchResults(searchData.tags);
  }, [reportTagSearchFetcher.data]);

  const filteredReportRows = useMemo(
    () => filterReportRows(reportOrders, reportTableQuery),
    [reportOrders, reportTableQuery],
  );

  const filteredReportTagSuggestions = useMemo(() => {
    const query = reportTagInputValue.trim().toLowerCase();

    if (!query) {
      return [];
    }

    const localRanked = rankTagListByPrefix(query, reportTagSuggestions, 200);
    const hasFreshServerResults =
      query.length >= 1 &&
      reportTagSearchQuery === reportTagInputValue.trim() &&
      reportTagSearchResults.length > 0;

    if (!hasFreshServerResults) {
      return localRanked.slice(0, 100);
    }

    const merged = mergeUniqueTags(reportTagSearchResults, localRanked);
    return rankTagListByPrefix(query, merged, 100);
  }, [
    reportTagInputValue,
    reportTagSearchQuery,
    reportTagSearchResults,
    reportTagSuggestions,
  ]);

  const onGenerateReport = () => {
    const formData = new FormData();
    formData.append("intent", "generate_report");
    formData.append("orderNumbers", orderNumbersInput);
    formData.append("searchTags", selectedReportTags.join(", "));
    reportFetcher.submit(formData, { method: "POST" });
  };

  const onExportReport = () => {
    if (reportOrders.length === 0) return;
    const csv = buildShippingReportCsv(reportOrders);
    const fileNameDate = new Date().toISOString().slice(0, 10);
    downloadCsv(`shipping-report-${fileNameDate}.csv`, csv);
  };

  const updateReportTagMenuPosition = () => {
    const anchor = reportTagPickerRef.current;
    if (!anchor) {
      setReportTagMenuRect(null);
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

    setReportTagMenuRect({
      top: Math.round(top),
      left: Math.round(left),
      width: Math.round(width),
    });
  };

  useEffect(() => {
    if (!isReportTagDropdownOpen) return;

    updateReportTagMenuPosition();

    const onWindowChange = () => {
      updateReportTagMenuPosition();
    };

    window.addEventListener("resize", onWindowChange);
    window.addEventListener("scroll", onWindowChange, true);

    return () => {
      window.removeEventListener("resize", onWindowChange);
      window.removeEventListener("scroll", onWindowChange, true);
    };
  }, [isReportTagDropdownOpen, reportTagInputValue]);

  useEffect(() => {
    if (!isReportTagDropdownOpen) return;

    const onPointerDown = (event: PointerEvent) => {
      const picker = reportTagPickerRef.current;
      const menu = reportTagMenuRef.current;
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

      setIsReportTagDropdownOpen(false);
      setReportTagMenuRect(null);
    };

    document.addEventListener("pointerdown", onPointerDown, true);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown, true);
    };
  }, [isReportTagDropdownOpen]);

  const onReportTagFieldInput = (event: Event) => {
    const target = event.currentTarget as HTMLElement & { value?: string };
    const value = target.value || "";
    const query = value.trim();
    setReportTagInputValue(value);
    const shouldOpen = query.length > 0;
    setIsReportTagDropdownOpen(shouldOpen);
    if (shouldOpen) {
      updateReportTagMenuPosition();
    } else {
      setReportTagMenuRect(null);
    }
  };

  const onReportTagFieldFocus = () => {
    if (reportTagInputValue.trim().length > 0) {
      setIsReportTagDropdownOpen(true);
      updateReportTagMenuPosition();
    }
  };

  const onToggleReportTag = (tag: string) => {
    setSelectedReportTags((existing) => {
      const existingIndex = existing.findIndex(
        (item) => item.toLowerCase() === tag.toLowerCase(),
      );
      if (existingIndex >= 0) {
        return existing.filter((_, index) => index !== existingIndex);
      }
      return [...existing, tag];
    });
  };

  const onRemoveReportTag = (tag: string) => {
    setSelectedReportTags((existing) =>
      existing.filter((item) => item.toLowerCase() !== tag.toLowerCase()),
    );
  };

  useEffect(() => {
    const onDocumentKeyDown = (event: KeyboardEvent) => {
      const picker = reportTagPickerRef.current;
      const target = event.target as Node | null;

      if (!picker || !target || !picker.contains(target)) {
        return;
      }

      if (event.key === "Escape" && isReportTagDropdownOpen) {
        event.preventDefault();
        setIsReportTagDropdownOpen(false);
        setReportTagMenuRect(null);
      }

    };

    document.addEventListener("keydown", onDocumentKeyDown);
    return () => {
      document.removeEventListener("keydown", onDocumentKeyDown);
    };
  }, [isReportTagDropdownOpen]);

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

      <style>{REPORT_TAG_PICKER_STYLES}</style>

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

            <div className="report-tag-picker" ref={reportTagPickerRef}>
              <s-search-field
                label="Available tags"
                placeholder="Search tags or add a new tag"
                value={reportTagInputValue}
                onInput={onReportTagFieldInput}
                onFocus={onReportTagFieldFocus}
              />
            </div>

            {selectedReportTags.length > 0 ? (
              <s-stack direction="block" gap="small">
                <s-stack direction="inline" gap="small">
                  {selectedReportTags.map((tag) => (
                    <s-clickable-chip
                      key={tag}
                      color="strong"
                      accessibilityLabel={`Remove tag ${tag}`}
                      removable
                      onRemove={() => onRemoveReportTag(tag)}
                    >
                      {tag}
                    </s-clickable-chip>
                  ))}
                </s-stack>
              </s-stack>
            ) : null}

            <s-stack direction="inline" gap="base">
              <s-button
                variant="primary"
                onClick={onGenerateReport}
                {...(isGeneratingReport ? { loading: true } : {})}
                disabled={!hasReportCriteria}
              >
                Generate Report
              </s-button>
            </s-stack>
          </s-stack>
        </s-section>

        {reportOrders.length > 0 ? (
          <s-section padding="none">
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
          <s-button variant="secondary" href="/app">
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

      {isHydrated && isReportTagDropdownOpen && reportTagMenuRect ? (
        createPortal(
          <div
            ref={reportTagMenuRef}
            className="report-tag-picker__menu"
            style={{
              top: `${reportTagMenuRect.top}px`,
              left: `${reportTagMenuRect.left}px`,
              width: `${reportTagMenuRect.width}px`,
            }}
            onMouseDown={(event) => {
              event.preventDefault();
            }}
          >
            <div className="report-tag-picker__list">
              {isSearchingReportTags ? (
                <div className="report-tag-picker__loading">
                  <s-stack direction="inline" gap="small">
                    <s-spinner
                      accessibilityLabel="Searching report tags"
                      size="base"
                    />
                    <s-text color="subdued">Searching tags...</s-text>
                  </s-stack>
                </div>
              ) : null}
              {!isSearchingReportTags && filteredReportTagSuggestions.length === 0 ? (
                <s-box padding="small">
                  <s-text color="subdued">No matching tags.</s-text>
                </s-box>
              ) : (
                filteredReportTagSuggestions.map((tag) => {
                  const isSelected = selectedReportTags.some(
                    (value) => value.toLowerCase() === tag.toLowerCase(),
                  );
                  return (
                    <div key={tag} data-report-tag-row="true">
                      <s-clickable onClick={() => onToggleReportTag(tag)} padding="small">
                        <s-stack direction="inline" gap="small">
                          <span
                            className={`report-tag-picker__checkbox${isSelected ? " report-tag-picker__checkbox--checked" : ""}`}
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
    </s-page>
  );
}

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};

async function searchReportTags(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  formData: FormData,
): Promise<ReportTagSearchActionData> {
  const queryRaw = (formData.get("query") as string | null)?.trim() || "";
  if (queryRaw.length < 1) {
    return {
      type: "search_report_tags",
      query: queryRaw,
      tags: [],
    };
  }

  try {
    const tagResult = await loadShopReportTags(admin, queryRaw);
    if (tagResult.error) {
      return {
        type: "error",
        query: queryRaw,
        error: tagResult.error,
      };
    }

    return {
      type: "search_report_tags",
      query: queryRaw,
      tags: rankTagListByPrefix(queryRaw, tagResult.tags, 100),
    };
  } catch (error) {
    return {
      type: "error",
      query: queryRaw,
      error:
        error instanceof Error ? error.message : "Failed to search report tags.",
    };
  }
}

async function loadShopReportTags(
  admin: Awaited<ReturnType<typeof authenticate.admin>>["admin"],
  rawQuery?: string,
): Promise<{ tags: string[]; error?: string }> {
  const query = rawQuery?.trim() || "";
  if (query) {
    try {
      const searchQuery = `title:${escapeSearchToken(query)}*`;
      const searched = await admin.graphql(SEARCH_REPORT_TAGS_QUERY, {
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
        return {
          tags: readReportTagConnectionValues(searchedJson.data?.shop?.orderTags),
        };
      }
    } catch {
      // Fall through to orders-based fallback below.
    }

    return loadReportTagsFromOrders(admin, query);
  }

  try {
    const response = await admin.graphql(REPORT_TAG_SUGGESTIONS_QUERY);
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
        tags: readReportTagConnectionValues(json.data?.shop?.orderTags),
      };
    }
  } catch {
    // Fall through to orders-based fallback below.
  }

  return loadReportTagsFromOrders(admin);
}

async function loadReportTagsFromOrders(
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

  while (hasNextPage && pagesLoaded < REPORT_TAG_SCAN_MAX_PAGES) {
    const response = await admin.graphql(
      query ? SEARCH_REPORT_TAGS_BY_ORDERS_QUERY : REPORT_TAG_SUGGESTIONS_BY_ORDERS_QUERY,
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
    tags: query ? rankTagListByPrefix(query, sorted, 100) : sorted.slice(0, 500),
  };
}

function readReportTagConnectionValues(
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

function escapeSearchToken(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/([:\\()])/g, "\\$1");
}

function rankTagListByPrefix(query: string, tags: string[], max: number): string[] {
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

  return deduped
    .filter((tag) => tag.toLowerCase().startsWith(normalizedQuery))
    .slice(0, max);
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
