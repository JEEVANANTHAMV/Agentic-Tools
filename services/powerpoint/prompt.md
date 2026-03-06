# Guidelines for Generating Content for PowerPoint Presentation API

## Tool Name - presentation_creator

## Core Objective
Create **premium, visually stunning, and modern** PowerPoint presentations that "WOW" the audience. Use high-contrast color palettes, absolute positioning for balanced layouts, and rich typography.

## Premium Design Strategy

1.  **Rich Aesthetics**: Avoid basic white backgrounds. Use a curated color palette:
    *   **Primary Blue**: `#4287f5`
    *   **Secondary Indigo**: `#667eea`
    *   **Accent Cyan**: `#00b6d6`
    *   **Alert/Metric Red**: `#ff4d4d`
    *   **Dark Mode BG**: `#121416`
2.  **Layout Precision**: 
    *   Use `position: absolute;` with percentages (e.g., `left: 5%; top: 10%; width: 90%;`) for every element. 
    *   This removes the "boring bullet list" look and creates custom, professional layouts.
3.  **Visual Hierarchy**:
    *   Larger font sizes for key metrics (e.g., `font-size: 32px;` or `48px;`).
    *   Use `<strong>` tags with vibrant colors to make keywords pop inside paragraphs.
4.  **Icons and SVGs**:
    *   **Emojis**: Use standard emojis for quick icons (🚀, 📈, ⚖️).
    *   **Images**: Use the `<img>` tag for high-quality photos or brand logos.
    *   **SVG-style Shapes**: Create "Cards" or "Containers" using `<div>` with `background-color`, `border-radius: 15px;`, and `border-color`.
5.  **Data Visualization**:
    *   Use `<table>` for structured data.
    *   Use colored boxes with large text inside for "KPI dashboards."

## Supported HTML Elements & Styles

| Element | Supported Styles / Attributes |
| :--- | :--- |
| **Slide** | `<div class="slide" style="background-color: #HEX;">` |
| **Containers (Divs)** | `position: absolute`, `left`, `top`, `width`, `height`, `background-color`, `border-radius`, `border-color`, `border-width` |
| **Headings (H1-H6)** | `color`, `font-size`, `text-align`, `left`, `top`, `width` |
| **Paragraphs (P)** | `color`, `font-size`, `text-align`, `left`, `top`, `width` |
| **Lists (UL, OL)** | `color`, `left`, `top`, `width` (Supports nested `<li>`) |
| **Inline Formatting** | `<strong>` (Bold), `<em>` (Italic) — both support inline `color` and `font-size` styles. |
| **Images (IMG)** | `src` (URL or path), `width`, `height`, `left`, `top` |
| **Tables (TABLE)** | `left`, `top`, `width`, `background-color` (Header `<th>` is auto-styled blue). |

## Design Principles
*   **The 6x6 Rule**: Max 6 lines of text, 6 words per line. Let the design speak.
*   **Contrast**: Always ensure high contrast between text and background.
*   **Alignment**: Align elements to a grid (e.g., left margins always at 5% or 10%).

## Premium Example: High-Tech Pitch Deck

```html
<!-- Slide 1: Impactful Title -->
<div class="slide" style="background-color: #121416;">
  <h1 style="color: #ffffff; font-size: 54px; text-align: center; top: 25%;">QUANTUM NETWORKS 🚀</h1>
  <p style="color: #00b6d6; font-size: 24px; text-align: center; top: 40%;">The Future of Unhackable Communication</p>
  <div style="position: absolute; top: 70%; width: 100%; text-align: center;">
    <p style="color: #667eea; font-size: 14px;">CONFIDENTIAL | SERIES A PITCH 2025</p>
  </div>
</div>

<!-- Slide 2: The Problem (Dashboard Style) -->
<div class="slide" style="background-color: #ffffff;">
  <h1 style="color: #4287f5; left: 5%; top: 5%;">Cybersecurity Crisis 📈</h1>
  
  <div style="position: absolute; left: 5%; top: 20%; width: 45%;">
    <p style="color: #333; font-size: 20px;">Traditional encryption is <strong style="color: #ff4d4d;">failing</strong>.</p>
    <ul>
      <li>Quantum computers can crack RSA in seconds</li>
      <li>Global cyber-theft costs <strong style="color: #ff4d4d;">$10.5T</strong> annually</li>
      <li>Trust in digital infrastructure is at an all-time low</li>
    </ul>
  </div>

  <!-- Stat Card -->
  <div style="position: absolute; left: 55%; top: 20%; width: 40%; height: 50%; background-color: #f0f4ff; border-radius: 20px; border-color: #4287f5;">
    <h3 style="color: #000; text-align: center; top: 10%;">The Impact</h3>
    <div style="text-align: center; top: 30%;">
      <p style="font-size: 42px;"><strong style="color: #ff4d4d;">60%</strong></p>
      <p>of breaches are due to <br/>ancient cryptography</p>
    </div>
  </div>
</div>

<!-- Slide 3: Conclusion & Next Steps -->
<div class=\"slide\" style=\"background-color: #4287f5;\">
  <h1 style=\"color: #ffffff; text-align: center; top: 30%;\">Ready to Scale?</h1>
  <div style=\"position: absolute; left: 20%; top: 50%; width: 60%; text-align: center; background-color: rgba(255,255,255,0.1); border-radius: 50px;\">
    <p style=\"color: #ffffff; padding: 20px;\">Contact us at investors@quantum.net</p>
  </div>
</div>
```

## API Call Format
`POST /api/v1/generate-presentation`

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/generate-presentation' \
-H 'Content-Type: application/json' \
-d '{
  "content": "SLIDE_HTML_HERE",
  "filename": "quantum_pitch.pptx"
}'
```
