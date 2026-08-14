# NX-913 — Campus Event & Hackathon OS · Pitch Deck

A god-level, animated presentation deck for **NX-913**, the Next-Generation Campus Event & Hackathon Operating System.

## Files

| File | Description |
|---|---|
| `index.html` | **The deck** — self-contained 16-slide animated presentation. Open in any browser. |
| `NX-913_Pitch_Deck.pptx` | Native PowerPoint export (16:9, 16 slides, presenter scripts as slide notes). |
| `assets/site/` | Live screenshots of [nx-913.com](https://nx-913.com) captured for the PPTX. |
| `tools/pptx/` | Generator scripts that build the PPTX from the deck content. |

## Presenting the deck

Open `index.html` in a browser and press:

- `←` / `→` — navigate slides
- `N` — presenter notes (per-slide speaking script)
- `O` — overview grid
- `F` — fullscreen
- `⬇` — download the PPTX right from the browser

## Regenerating the PPTX

```bash
cd tools/pptx && npm install && node export-pptx.mjs
```

The generator parses each slide's title/subtitle/notes directly from `index.html` so content stays in sync, then writes `NX-913_Pitch_Deck.pptx` to the project root.

## Design

Brand identity matched to [nx-913.com](https://nx-913.com): `#050505` black background, brand red `#dc2626` / `#ef4444` accents, gold highlights, **Chakra Petch** (headings) + **Rajdhani** (body) typography.
