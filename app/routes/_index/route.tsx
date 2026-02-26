import type { LoaderFunctionArgs } from "react-router";
import { redirect, Form, useLoaderData } from "react-router";

import { login } from "../../shopify.server";

import styles from "./styles.module.css";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);

  if (url.searchParams.get("shop")) {
    throw redirect(`/app?${url.searchParams.toString()}`);
  }

  return { showForm: Boolean(login) };
};

export default function App() {
  const { showForm } = useLoaderData<typeof loader>();

  return (
    <div className={styles.root}>
      <div className={styles.container}>
        <div className={styles.badge}>Internal</div>
        <h1 className={styles.heading}>
          Corporate<br />Addresses
        </h1>
        <p className={styles.tagline}>
          Bulk-create Shopify orders from a CSV — ship to hundreds of addresses in minutes.
        </p>
        {showForm && (
          <Form className={styles.form} method="post" action="/auth/login">
            <div className={styles.inputRow}>
              <input
                className={styles.input}
                type="text"
                name="shop"
                placeholder="my-shop.myshopify.com"
                autoComplete="on"
              />
              <button className={styles.button} type="submit">
                Log in
              </button>
            </div>
          </Form>
        )}
        <ul className={styles.features}>
          <li className={styles.feature}>
            <span className={styles.num}>01</span>
            <div>
              <strong>CSV import</strong>
              <span>Upload a spreadsheet with recipient names, addresses, and line items — no manual entry.</span>
            </div>
          </li>
          <li className={styles.feature}>
            <span className={styles.num}>02</span>
            <div>
              <strong>Bulk order creation</strong>
              <span>Generate hundreds of Shopify orders at once, each shipped to a unique address.</span>
            </div>
          </li>
          <li className={styles.feature}>
            <span className={styles.num}>03</span>
            <div>
              <strong>Batch export</strong>
              <span>Download a CSV with all order IDs and tracking details for easy reconciliation.</span>
            </div>
          </li>
        </ul>
      </div>
    </div>
  );
}
