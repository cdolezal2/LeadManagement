# Project Instructions

## Customer Websites

When building websites for customers, never use emojis anywhere in the HTML — not in icons, labels, buttons, cards, or any other UI element.

Use inline SVG icons instead. All SVGs should be:
- Stroke-based (fill="none", stroke="currentColor")
- 24x24 viewBox
- Consistent stroke-width (1.8 for regular icons, 2.5 for emphasis like checkmarks)
- stroke-linecap="round" stroke-linejoin="round"
- Sized and colored via CSS on the parent element or directly on the svg tag

## SVG Logo Text

When placing text in an inline SVG logo, always use `text-anchor="middle"` with `x` set to the horizontal center of the viewBox so the text is truly centered regardless of font rendering. Example for a 220px-wide viewBox:

```html
<text x="110" y="48" text-anchor="middle" ...>M&amp;G</text>
```

Never position SVG text at `x="0"` or a small offset — it will appear left-aligned even if the SVG container is centered via CSS.

## Contact / Info Cards

For two-column layouts with an info card and a form card, use `align-items: stretch` on the grid so both tiles match height automatically.

Make the info card a flex column with `justify-content: space-between` so content fills the card naturally:

```css
.contact-info-card {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}
```

For contact info lists inside a dark card:
- Use a subtle `border-top` divider to separate the intro text from the list items
- Use `display: flex; flex-direction: column; gap: Npx` on the list for consistent spacing
- Icon tiles: 46×46px with a semi-transparent background and border
- Label: small uppercase (12px, 0.06em letter-spacing) above the value
- Value text: 17px

## Hero Right Panel (Dark Panel)

- Remove all background grid patterns and concentric rings — keep it clean
- Center all content with `align-items: center; justify-content: center; text-align: center`
- Center flex children like map-pin notes with `justify-content: center`
- Logo slot should use `display: flex; justify-content: center` to center the SVG inside it

## Tone & Copy

- Small business / family-owned sites: write copy that feels personal and local, not corporate
- Avoid italicized `<em>` wording in headings — it reads as excessive
- About sections should reference specific founding story details, direct owner relationships, and local pride

## Footer Attribution & Legal

Every customer website footer must include:
1. A subtle "Website created by [Bow Lane Leads](https://bowlaneleads.com/)" attribution line
2. Links to the site's Privacy Policy and Terms of Service pages

Footer attribution markup pattern:
```html
<div style="margin-top:16px;padding-top:16px;border-top:1px solid rgba(255,255,255,0.06);font-size:12px;">
  <a href="./[demo-slug]-privacy.html">Privacy Policy</a> &middot; <a href="./[demo-slug]-terms.html">Terms of Service</a>
</div>
<div style="margin-top:8px;font-size:11px;color:rgba(255,255,255,0.2);">
  Website created by <a href="https://bowlaneleads.com/" target="_blank" rel="noopener">Bow Lane Leads</a>
</div>
```

Each customer site needs its own `[demo-slug]-privacy.html` and `[demo-slug]-terms.html` pages styled to match the site's theme.
