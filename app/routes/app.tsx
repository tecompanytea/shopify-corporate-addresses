import { useEffect, useState } from "react";
import type { HeadersFunction, LoaderFunctionArgs } from "react-router";
import { Outlet, redirect, useLoaderData, useRouteError } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";
import { AppProvider } from "@shopify/shopify-app-react-router/react";

import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { session } = await authenticate.admin(request);
  const url = new URL(request.url);
  const hostFromQuery = url.searchParams.get("host");

  // App Bridge requires a valid `host` query param in embedded context.
  // Normalize URLs that arrive without it so parent postMessage origin stays correct.
  if (!hostFromQuery) {
    const shopDomain = url.searchParams.get("shop") || session.shop;
    const storeHandle = shopDomain.replace(".myshopify.com", "");
    const host = Buffer.from(
      `admin.shopify.com/store/${storeHandle}`,
      "utf8",
    ).toString("base64");

    url.searchParams.set("host", host);
    url.searchParams.set("shop", shopDomain);
    url.searchParams.set("embedded", "1");
    throw redirect(`${url.pathname}?${url.searchParams.toString()}`);
  }

  // eslint-disable-next-line no-undef
  return { apiKey: process.env.SHOPIFY_API_KEY || "" };
};

export default function App() {
  const { apiKey } = useLoaderData<typeof loader>();
  const [isHydrated, setIsHydrated] = useState(false);

  useEffect(() => {
    setIsHydrated(true);
  }, []);

  if (!isHydrated) {
    return <div style={{ minHeight: "100vh" }} />;
  }

  return (
    <AppProvider embedded apiKey={apiKey}>
      <Outlet />
    </AppProvider>
  );
}

// Shopify needs React Router to catch some thrown responses, so that their headers are included in the response.
export function ErrorBoundary() {
  return boundary.error(useRouteError());
}

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
