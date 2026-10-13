# Order Desk — setup (about 20 minutes, $0)

Customers order on a web page (works with no signal and sends later). Every order lands in your Google Sheet with an order number, a permanent customer number, and separate **Ready**, **Delivered** and **Paid** tick boxes.

## Files

| File | What it is |
|---|---|
| `Code.gs` | Google Sheet backend — paste into Apps Script |
| `index.html` | The order page customers open |
| `sw.js`, `manifest.json`, `icon-*.png` | Make the page work offline and installable on phones |

## 1. Create the Sheet and backend

1. Create a new Google Sheet (e.g. "Orders 2026").
2. **Extensions → Apps Script**. Delete what's there, paste all of `Code.gs`.
3. At the top, change `ORDER_KEY: 'change-this-key'` to something only you know (e.g. `meat-7Qx29`). Set `SHOP_NAME`. Optionally put your email in `NOTIFY_EMAIL`.
4. Save. In the function dropdown choose **setup** → **Run** → approve the permissions (Advanced → Go to project).
5. Back in the Sheet you now have tabs: **Orders, Lines, Customers, Products, Pick List**.

## 2. Publish the backend

1. In Apps Script: **Deploy → New deployment → ⚙ → Web app**.
2. Execute as: **Me**. Who has access: **Anyone**. → Deploy.
3. Copy the **Web app URL** (ends in `/exec`). Open it in a browser — you should see `"Order Desk is running."`

> When you change `Code.gs` later: **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**. This keeps the same URL.

## 3. Point the page at your Sheet

Open `index.html` in any text editor and edit the SETTINGS block:

```js
ENDPOINT: 'https://script.google.com/macros/s/AKfy..../exec',
ORDER_KEY: 'meat-7Qx29',          // same as Code.gs
SHOP_NAME: 'Your Shop Name',
SHOP_WHATSAPP: '14025550123',     // optional: "Share on WhatsApp" goes straight to you
```

## 4. Host it free on GitHub Pages

1. Create a free account at github.com → **New repository** (e.g. `order`) → Public.
2. **Add file → Upload files** → drag in `index.html`, `sw.js`, `manifest.json`, `icon-192.png`, `icon-512.png` → Commit.
3. **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / root → Save.
4. After ~1 minute your link is `https://<your-username>.github.io/order/`.

## 5. Share it in WhatsApp

Pin a message in the group and put it in the group description:

> 🛒 Order here: https://yourname.github.io/order/
> You'll get an order number instantly. Tip: tap Share → **Add to Home Screen** so it works even with weak signal.

## Running the day

| Tab | Use it for |
|---|---|
| **Orders** | One row per order. Tick **Ready** when packed, **Delivered** when handed over, **Paid** when paid. Ready/Delivered times fill in automatically. Status turns New (red) → Ready (amber) → Delivered (green). Paid cell goes red if delivered but unpaid. |
| **Pick List** | Change the yellow date in B1 → see total lb/pieces per product to pack, plus every order for that day. |
| **Customers** | Each phone number gets a permanent number (C-001, C-002…) on first order, with order count and last order date. Same phone = same number forever. |
| **Products** | What appears on the page. Add rows, untick **Active** to hide, change **Sort** to reorder, Unit = `lb` or `each`. The page picks up changes next time it opens. |
| **Lines** | One row per item (feeds the Pick List). Don't edit. |

Use the filter on the Orders header to show e.g. only Status = New, or only today's For Date.

**Phone orders / WhatsApp text orders:** open the same page yourself and enter them under the customer's name and phone — they get the same customer number and the Sheet stays complete.

## Good to know

- **Offline:** if there's no signal, the order is saved on the phone and sent automatically the next time the page is open with signal. On iPhone the page must be opened again to send — that's an Apple limit — so ask customers to add it to their Home Screen.
- **Duplicates are blocked:** the page retries until the Sheet confirms, and each order carries a unique ID, so a retry never creates a second row.
- **Spam:** the shop key, a hidden spam trap, and a limit of 10 orders per phone per hour. Anyone who sees the page code can find the key — fine for a shop order form; change it in both files if you ever get junk orders.
- **Limits:** Google allows 30 orders being saved at the same moment and far more per day than a shop will send. Email alerts are capped at 100/day on a free Gmail account.
- **Don't rename** the tabs or reorder columns on the Orders tab — the script relies on them. Adding columns to the right of *Client ID* (hidden column Q) is fine.
