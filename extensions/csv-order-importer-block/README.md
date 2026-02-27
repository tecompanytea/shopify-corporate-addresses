# Corporate Addresses

Admin action extension for importing CSV rows and creating Shopify orders.

Target:
- `admin.order-index.action.render`

Required CSV columns:
- `order_key`
- `email`
- `variant_id` (numeric variant ID or full `gid://shopify/ProductVariant/...`)
- `quantity`

Optional CSV columns:
- `currency_code`
- `note`
- `tags` (comma or pipe separated)
- `shipping_first_name`
- `shipping_last_name`
- `shipping_address1`
- `shipping_city`
- `shipping_province_code`
- `shipping_country_code`
- `shipping_zip`
- `phone`

Example:

```csv
order_key,email,variant_id,quantity,currency_code,note,tags,shipping_first_name,shipping_last_name,shipping_address1,shipping_city,shipping_province_code,shipping_country_code,shipping_zip
1001,alice@example.com,12345678901234,2,USD,First test order,new|csv,Alice,Ng,123 Main St,San Francisco,CA,US,94105
1001,alice@example.com,12345678904567,1,USD,First test order,new|csv,Alice,Ng,123 Main St,San Francisco,CA,US,94105
1002,bob@example.com,gid://shopify/ProductVariant/12345678907890,1,USD,,bulk,Bob,Lee,88 Oak Ave,Los Angeles,CA,US,90001
```
