# End-to-End Digital Conversion & Sales Dashboard (Toy Store)

## Project Overview

### The Problem
A toy store client came to me with a number that didn't add up: traffic was healthy, ads were running, top-line sales looked fine — but net profit kept shrinking. Nobody could point to exactly why. So I dug in, and found three separate leaks hiding behind the "everything looks fine" surface:

* People were dropping off somewhere in the buying journey, but nobody knew exactly where.
* A handful of products had quietly high return rates that were chewing through margin.
* Some ad campaigns and devices were burning budget without anyone noticing.

### The Project
I built the whole thing in Excel — no Power BI, no external tool, no license fee. Just Power Query for ingestion, Power Pivot for the data model, and DAX for the actual math. The pipeline pulls in 7 raw sources (ad spend, web traffic, orders, refunds, and so on), cleans them automatically on refresh, and feeds a Star Schema model. From there, DAX measures connect what people actually did on the site to what it cost or earned the business — not just revenue, but real profit after COGS and refunds. The final output is a dark, Bento-style dashboard meant to be read in under a minute, not studied for an hour.

### The Outcome
One connected view, from the first ad click all the way to net profit, that closed all three leaks at once. It pointed to one specific checkout problem on mobile, and one specific pricing tier that was losing value before people even hit the cart. Instead of guessing, the client walked away with three things to do: kill a specific campaign, fix a specific product page problem, and audit a specific product's quality — each one backed by a number, not a hunch.

---

## Project Details

### Step 1 — Breaking Down the Brief

I split the problem into three things I could actually investigate on their own: where the funnel was leaking, where refunds were eating margin, and which campaigns/devices were wasting money. Before touching any raw file, I went through them and pulled only the columns I actually needed for these three questions — no point dragging in a whole table just because it's there.

<img width="1004" height="921" alt="image" src="https://github.com/user-attachments/assets/5ff35ba3-83fe-4244-b932-648d9a7756d4" />

---

### Step 2 — Ingestion, Dynamic Pathing & a Pipeline That Doesn't Break

This part had to handle around **1.7 million rows** across 7 fragmented source files — sessions, pageviews, orders, order items, refunds, products, and ad spend logs — without the workbook choking.

**1. Folder structure and how I actually read the files.**
Raw files sit in their own `Resources\Data\` subfolder next to the workbook, split into one subfolder per table (`orders`, `order_items`, `website_pageviews`, and so on). My first attempt pointed a single folder connector at the whole directory — it worked, but it forced Power Query to re-index every file's metadata across all 6 tables on every refresh, even when only one had changed. I split that into 6 separate folder queries instead, one per table, each scoped to just its own subfolder. Same folder-combine pattern under the hood, but narrower — a refresh only re-scans the table that actually moved, not the whole tree.

<img width="1304" height="168" alt="Folders and files" src="https://github.com/user-attachments/assets/dfb74d46-a065-43b9-ba4e-680fce1ce6cd" />

**2. A file path that doesn't break the moment you move the file.**
Hardcoding `C:\Users\Username\...` is a trap — the second this workbook gets shared or moved to another machine, every query breaks. So I built a small chain instead:

1. An Excel formula grabs the workbook's own folder at runtime:
   ```excel
   =LEFT(CELL("filename"), SEARCH("[", CELL("filename")) - 1)
   ```
2. That cell becomes a named range, `FolderPath`, readable anywhere in the workbook.
3. A helper query in Power Query caches that value:
   ```powerquery
   FolderPathValue = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content]{0}[Column1]
   ```
4. Every source query builds its path off that:
   ```powerquery
   Source = Csv.Document(File.Contents(FolderPathValue & "Resources\Data\f_Web_Sessions.csv"), [Delimiter=",", Encoding=65001])
   ```
   <img width="902" height="329" alt="Dynamic Paths" src="https://github.com/user-attachments/assets/c75e2d1b-a438-4fd6-a343-6f4f27e8a6bb" />

**3. Keeping everything out of the worksheet grid.**
Excel's grid caps at just over a million rows, which this dataset blows past easily. So all 7 pipelines load as **Connection Only** — nothing ever touches a worksheet cell. Data streams straight into Power Pivot's VertiPaq engine, and its columnar compression cut the memory footprint by roughly 80%, which is the only reason the dashboard stays fast at this volume.

<img width="1855" height="800" alt="Queries tree" src="https://github.com/user-attachments/assets/c3928791-1b5c-40a1-9415-d6270bd5f4bf" />

---

### Step 3 — Profiling the Data Before Touching It

> Full baseline audit: [DataProfiling.txt](Resources/DataProfiling.txt)

Before writing a single transformation, I went through every raw table to understand its shape, its keys, and its gaps. Skipping this step doesn't save time — it just means finding these problems later, when they're more expensive to fix.

**What I found:**
* The full dataset comes to about **1.735 million rows**: `website_pageviews` (1,188,124), `website_sessions` (472,871, across 394,318 distinct users), `orders` (32,313), `order_items` (40,025), `order_items_refunds` (1,731).
* Primary keys (`order_id`, `order_item_id`, `website_session_id`) were 100% unique, no nulls — one less thing to worry about in the joins later.
* I mapped out a rough set of relationships — `orders → order_items → order_items_refunds`, and both `orders` and `sessions` feeding into `website_pageviews` — but treated these as hypotheses at this stage. I didn't confirm them for real until the modeling step.
* 83,328 rows had a null `utm_source`/`utm_campaign`/`utm_content`. I almost assumed this was broken tracking, but it turned out to be organic, non-campaign traffic — so instead of dropping those rows, I remapped them to `"Direct search"`.
* One product, *The Original Mr. Fuzzy*, made up 73.8% of order volume on its own (23,861 of 32,313 orders) — I didn't think much of this number at the time, but it turned out to matter a lot once I got to the refund analysis.
* Data spans 1,096 days (2012–2015), with most traffic hitting between 11 AM and 3 PM.

**What that meant for the engine:**
* Splitting `created_at` into a `Date` column and an integer `Hour` (0–23) column dropped the time field's cardinality down to 24 values, so VertiPaq can compress it with run-length encoding instead of storing near-unique values row by row.
* Joining a dimension table against 1.18 million pageview rows the normal way was slow on every refresh. Wrapping the dimension side in `Table.Buffer` forces it into memory once, so the join becomes a single in-memory lookup instead of repeated disk reads.
* I standardized text fields with `Text.Clean`, `Text.Trim`, and `Text.Proper`, cast every column explicitly, and added `Table.ReplaceErrorValues` everywhere — so one bad row in a CSV doesn't take down a scheduled refresh.
* Clickstream data only tracked products by URL slug, no numeric ID attached. I built a synthetic `MatchKey` (product name with hyphens instead of spaces) to join those pageviews back to a real `product_id`.

<details>
<summary><b>View Production M-Code Implementation</b></summary>

```powerquery
let
    Source = Folder.Files(GetPathWebsite_pageviews),
    Defult1 = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    Defult2 = Table.AddColumn(Defult1, "Transform File (5)", each #"Transform File (5)"([Content])),
    Defult3 = Table.RenameColumns(Defult2, {"Name", "Source.Name"}),
    Defult4 = Table.SelectColumns(Defult3, {"Source.Name", "Transform File (5)"}),
    Defult5 = Table.ExpandTableColumn(Defult4, "Transform File (5)", Table.ColumnNames(#"Transform File (5)"(#"Sample File (5)"))),

    #"Delete file name" = Table.RemoveColumns(Defult5, {"Source.Name"}),
    #"Date-time column split" = Table.SplitColumn(#"Delete file name", "created_at", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"created_at_date", "created_at_time_hour"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Date-time column split",
    {{"website_pageview_id", Int64.Type}, 
    {"created_at_date", type date}, {"created_at_time_hour", type time}, {"website_session_id", Int64.Type},
    {"pageview_url", type text}}),
    #"Extracted Hour" = Table.TransformColumns(#"Changed Type", {{"created_at_time_hour", Time.Hour, Int64.Type}}),
    
    #"Cleaned Product Name" = Table.TransformColumns(#"Extracted Hour", {
    {"pageview_url", each Text.Proper(Text.Trim(Text.Clean(_))), type text}}),
    #"Errors replacer" = Table.ReplaceErrorValues(#"Cleaned Product Name", 
    {{"website_pageview_id", Int64.Type}, 
    {"created_at_date", null}, {"created_at_time_hour", null}, {"website_session_id", null},
    {"pageview_url", null}}),
    #"""\"" delete in url" = Table.ReplaceValue(#"Errors replacer", "/", "", Replacer.ReplaceText, {"pageview_url"}),
    #"Reordered Columns" = Table.ReorderColumns(#"""\"" delete in url", {"website_session_id", "website_pageview_id", "created_at_date", "created_at_time_hour", "pageview_url"}),

    // In-memory hash mapping to vectorize nested join across 1.18M rows
    ProductsLookup = Table.Buffer(
        Table.AddColumn(
            Table.SelectColumns(products, {"product_id", "product_name"}),
            "MatchKey",
            each Text.Replace(Text.Replace([product_name], " ", "-"), ".", ""),
            type text
        )
    ),

    #"Merged Products" = Table.NestedJoin(#"Reordered Columns", {"pageview_url"}, ProductsLookup, {"MatchKey"}, "MatchedProduct", JoinKind.LeftOuter),
    #"Expanded Product_ID" = Table.ExpandTableColumn(#"Merged Products", "MatchedProduct", {"product_id"}, {"product_id"}),
    #"Set Product_ID Type" = Table.TransformColumnTypes(#"Expanded Product_ID", {{"product_id", Int64.Type}})
in
    #"Set Product_ID Type"
```
</details>

---

### Step 4 — Star Schema & Getting to the Root Cause

With all 7 streams cleaned and keyed, I moved into Power Pivot. The goal here wasn't just to build a model — it was to build one clean enough that I could actually test ideas against it and trust the answer.

**Schema layout:**
* **Dimensions:** `Calender` (Date, Year, MonthNo, MonthName, YearMonth) as the time anchor; `products` as the catalog.
* **Facts:** `website_sessions` as the attribution backbone (utm source, campaign, device type); `website_pageviews` for clickstream; `orders`/`order_items` for transactions (price, COGS, primary vs. add-on flag); `order_items_refunds` for returns.
* **A disconnected `Measure` table** holding every DAX formula — kept separate from the actual data on purpose, since it's calculation logic, not a data source.

**Relationships:**
* `Calender[Date] (1) → (Many) website_sessions[created_at_date]`
* `products[product_id] (1) → (Many) order_items[product_id]`
* `products[product_id] (1) → (Many) website_pageviews[product_id1]`
* `website_sessions[website_session_id] (1) → (Many) website_pageviews[website_session_id]`
* `website_sessions[website_session_id] (1) → (1) orders[website_session_id]`
* `orders[order_id] (1) → (Many) order_items[order_id]`
* `order_items[order_item_id] (1) → (1) order_items_refunds[order_item_id]`

<img width="3828" height="1936" alt="Power Pivot relations" src="https://github.com/user-attachments/assets/5dc5f06e-1b56-4caa-90b3-c89c49376068" />

<details>
<summary><b>Full list of DAX measures (34)</b></summary>

**Funnel volume:**
`Funnel 2 - Products Page Sessions`, `Funnel 3 - Product Details Sessions`, `Funnel 4 - Cart Sessions`, `Funnel 5 - Shipping Sessions`, `Funnel 6 - Billing Sessions`, `Funnel 7 - Order Completed Sessions`

**Conversion rates:**
`CR - Entry to Products`, `CR - Products to Details`, `CR - Details to Cart`, `CR - Cart to Shipping`, `CR - Shipping to Billing`, `CR - Billing to Order`, `CR - Entry to Cart`, `Overall Store CR`, `Funnel_conv_rate`

**SKU-level behavior:**
`Views - Sugar Panda`, `Views - Mini Bear`, `Views - Mr. Fuzzy`, `Views - Love Bear`, `Cart Adds - Sugar Panda`, `Cart Adds - Mini Bear`, `Cart Adds - Mr. Fuzzy`, `Cart Adds - Love Bear`, `CR - Sugar Panda`, `CR - Mini Bear`, `CR - Mr. Fuzzy`, `CR - Love Bear`, `Mid Tier Traffic Share`

**Unit economics & attribution:**
`Total Units Sold`, `Refunded Units`, `Store Refund Rate`, `Total Sessions`, `Gross Revenue`, `Net Profit`

</details>

---

#### Diagnostic 1 — Where the Funnel Actually Leaks

The headline number first: store-wide conversion is **6.83%**, more than double the usual e-commerce benchmark of 2–3%. So the store, overall, is healthy. The interesting part is where the 45% landing-page drop-off comes from — and it turned out to be nothing. Bounce rate holds at roughly 45% across Gsearch, Bsearch, and Socialbook alike, which rules out a targeting problem — if one channel were sending the wrong people, that number would move.

I also checked device as a suspect, since a broken mobile checkout is a common culprit. It's not: desktop converts at 46.1%, mobile at 42.1%, a gap of exactly 4 points. Not nothing, but not the story either.

The real leak is between the product page and the cart, where 55% of shoppers disappear — and it doesn't line up with price the way I expected. The cheapest item ($29.99) converts at 65.13%. The most expensive ($59.99) still holds at 55.64%. It's the middle of the catalog, $45.99–$49.99, that's bleeding — 86.4% of all traffic lands there, and over half of it walks away before adding to cart. So this isn't a pricing problem or a technical one. It's a perceived-value gap: the mid-tier products aren't giving people a reason to say yes, and no amount of UX polish fixes that. The fix is better product storytelling and bundling on those specific SKUs.

---

#### Diagnostic 2 — The Refund Numbers Hide a Two-Sided Problem

Overall return rate looks fine at a glance — 4.32% (1,731 of 40,025 units), nothing alarming for a physical product. But averaging across the whole catalog buries what's actually happening.

Two products are doing almost all the damage: *The Original Mr. Fuzzy* is behind 71.46% of every refund in the store (1,237 units, a 5.11% return rate), and *The Birthday Sugar Panda* has the worst individual defect rate at 6.04%. Everything else is fine — the cheapest and priciest tiers return at 1.28% and 2.23%.

Before landing there, I ran down two ideas that didn't pan out. First, I figured cross-sell add-ons — the impulse buys people tack on at checkout — would have worse return behavior than deliberate purchases. Wrong: add-ons return at 2.52% versus 4.76% for primary items, while still making up almost a fifth of unit sales. Cross-selling isn't a risk here at all. Second, I tried to test whether shipping delays were driving returns, using days-to-refund as a rough stand-in since there's no actual delivery-date field in the data. Returns come in at a flat 0.25–0.37% a day for two weeks, then drop off hard after day 15 — which looks like a normal return window, not a shipping problem, but I can't say that for certain without the delivery data. That one stays open; I'm not going to force a conclusion the data doesn't support.

Put together: the mid-tier products get hit twice — high drop-off before the sale, then a high return rate after it. That's not a policy issue, it's a quality-control issue, and it points straight at *Mr. Fuzzy* and *Sugar Panda* specifically.

---

#### Diagnostic 3 — Following the Ad Money to Where It Actually Goes (or Doesn't)

I broke attribution down three levels — source, campaign, and the specific campaign-device combination — because looking at just the top level lies to you, as it turns out.

Mobile is the clear systemic weak point: conversion sits at 2.29–3.28% across every channel, roughly a third of what desktop gets (7.56–10.54%). `Gsearch Nonbrand Mobile` alone is 87,551 sessions converting at 3.18% — a lot of spend for very little return.

The more interesting case is Socialbook. At the aggregate level, it looks like a losing channel, full stop — the kind of number that gets a whole campaign killed. But breaking it down by device tells a different story: the entire loss is sitting in one specific segment, `Pilot Mobile` (0.83% conversion on 4,573 sessions). `Desktop_Targeted`, on the same platform, is actually one of the stronger performers — 5.15% conversion and $10,834.19 in net profit. Judging Socialbook by its average would have meant cutting a channel that's actually working, for the wrong reason.

The real profit engines: `Gsearch Nonbrand Desktop` generates $556,351.65 on its own — close to half the company's total profit. `Bsearch Brand Desktop` has the best conversion rate anywhere in the account, 9.86%. And direct/organic traffic, which costs nothing to acquire, still brings in $217,751.70 at a 7.34% conversion rate — proof the brand itself has real pull, independent of ad spend.

---

### Strategic Roadmap
* **Cut `Socialbook Pilot Mobile`**, move that budget into `Desktop_Targeted` and `Bsearch Brand/Nonbrand Desktop`, and apply negative mobile bid adjustments on generic search.
* **Fix mobile checkout** — specifically payment friction, starting with one-click Apple Pay/Google Pay and address auto-fill.
* **Audit manufacturing on *Mr. Fuzzy*** and bundle mid-tier products with accessories to close the value gap before the cart.

Put a number on it: reallocating the wasted mobile ad spend, closing the mid-tier gap, and fixing the *Mr. Fuzzy* defect rate are three separate levers, each tied to a specific line item in this dashboard — not three guesses.

<img width="1624" height="815" alt="Data analysing step " src="https://github.com/user-attachments/assets/e969cc5a-69d6-4661-a92c-0c935d3e97d8" />

---

### Step 5 — The Dashboard

Three tabs, each answering one question on its own instead of one dense grid trying to do everything at once.

#### Tab 1 — Conversion Funnel Diagnostics
Shows the store is healthy overall and isolates the leak to the product page — while ruling out UX bugs along the way.
<img width="1890" height="886" alt="Funnel Diagnostics" src="https://github.com/user-attachments/assets/d5bcb9c0-f840-44ba-a8c8-89d6c7e05e9c" />
* KPI cards: `Overall Store CR`, entry-page drop-off (`1 − CR Entry to Products`), cart drop-off (`1 − CR Details to Cart`), mid-tier traffic share.
* Charts: sequential funnel progression; traffic share vs. conversion rate by price point (where the mid-tier collapse shows up); conversion consistency across devices and channels.

#### Tab 2 — Refunds & Margin Preservation
Confirms store-wide returns are under control and traces the actual loss to specific SKUs, not to cross-selling.

<img width="1890" height="886" alt="Refunds   Margin" src="https://github.com/user-attachments/assets/5612f1aa-49ee-4c48-a116-cab48eed1332" />
* KPI cards: store refund rate, add-on refund rate, mid-tier return rate, top refunded SKU's share of all refunds.
* Charts: Pareto of refund volume/rate by product; primary vs. add-on sales share against return share; day 1–15 refund timeline.

#### Tab 3 — Acquisition & ROI
Strips revenue down to real net profit after COGS and refunds, and shows exactly which campaigns to cut and which to scale.

<img width="1890" height="886" alt="Acquisition   ROI" src="https://github.com/user-attachments/assets/e7eec4a3-6ab5-43ab-8708-b76cf7d7d8b6" />

* KPI cards: store net margin %, desktop's share of total profit, the mobile-vs-desktop conversion gap.
* Charts: full campaign profitability matrix (net profit, refunds, COGS); scale-vs-cut quadrant map by session volume and margin; side-by-side desktop vs. mobile funnel comparison; sales-vs-returns share by SKU, echoed from Tab 2 to keep product risk visible alongside the ROI view.

---
