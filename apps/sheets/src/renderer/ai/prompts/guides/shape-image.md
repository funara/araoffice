# Shapes & images guide (shape-image)

## Default visual guidance

Prefer native shapes/text boxes, charts, and sparklines for visuals. Do not fetch remote images or author external `IMAGE("https://...")` formulas. Use `add_image` only when the user explicitly supplies a local image path; never guess or search for one.

## Inserting a shape/text box (supported)

`{op:"add_shape", sheetId, shapeType, anchorCell:"E2", fillColor?, text?}`

- shapeType: any OOXML prstGeom name from the insert gallery — rectangles ("rect", "roundRect", snip/round corner variants), basic shapes ("ellipse", "triangle", "diamond", "hexagon", "heart", "cloud", "donut", "smileyFace", …), block arrows ("rightArrow", "quadArrow", "chevron", …), stars/banners ("star5", "star12", "ribbon", "wave", …), flowchart nodes ("flowChartProcess", "flowChartDecision", "flowChartDocument", …), callouts ("wedgeRectCallout", "cloudCallout", …), or "textbox"
- textbox defaults to a white background with "Text"; other shapes default to a light-blue fill. text can carry the annotation directly.
- Layout-class: can share a batch with content/format, not with structural operations; only works on imported xlsx, written with ⌘S.

### Common uses

- **Callout/annotation boxes**: put a `textbox` next to key data with the interpretation as text (e.g. "+23% YoY, all-time high").
- **Flow/relationship diagrams**: combine `rect`/`roundRect` + `rightArrow`, staggering anchorCells for layout.
- **Emphasis blocks**: a `roundRect` with a translucent fill framing a group of key metrics.

## In-cell sparklines (SPARKLINE, supported)

No chart object needed — draw a mini trend chart inside a cell with a formula:

- Line sparkline: `=SPARKLINE(B2:M2)`
- Column sparkline: `=SPARKLINE(B2:M2, {"charttype","column"})`
- Win/loss: `=SPARKLINE(B2:M2, {"charttype","winloss"})`

Great for a "Trend" column at the end of each summary row — the trend is visible at a glance and is lighter than inserting many charts.

```json
{
  "op": "set_formula",
  "sheetId": "s1",
  "address": "N2",
  "formula": "=SPARKLINE(B2:M2,{\"charttype\",\"column\"})"
}
```

## Images and links

- Local image: `{op:"add_image", sheetId:"s1", path:"~/logo.png", anchorCell:"B2"}` — use only for a local path explicitly supplied by the user (PNG/JPEG/GIF, ≤20MB).
- Remote image fetching is prohibited. Do not author `=IMAGE("https://...")` or any other agent-authored external image formula.
- Legitimate user-requested hyperlinks remain allowed; prefer `set_hyperlink` (see the data guide), or use `=HYPERLINK("https://...", "Click here")` when the user specifically requests a formula.

## Rules

- Shapes/text boxes, charts, and sparklines are the default visual tools; use images only when the user supplies the local path.
- Shapes/images are for **supporting expression** — don't let them upstage the data; correctness first.
- For lots of trend displays prefer SPARKLINE (light); for a single focused analysis use add_chart (heavy).
