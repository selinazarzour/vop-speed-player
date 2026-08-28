# ⏩ VOP Speed Player

**🔗 Live site: [selinazarzour.github.io/vop-speed-player](https://selinazarzour.github.io/vop-speed-player/)**

Watch narrated (voice-over) PowerPoints **at any speed**, right in your browser.

Professors love recording lectures as PowerPoint voice-overs — but PowerPoint only
plays them back at 1x. This tool fixes that:

1. **Drop your `.pptx` or `.ppsx`** onto the page
2. It extracts the narration that "Record Slide Show" embeds on each slide
3. **Play at 0.75x – 3x** with the slides, auto-advancing when narration ends

## ✨ Features

- 🔒 **100% private** — the file is unpacked with JSZip *inside your browser*; nothing is uploaded anywhere
- 🎙️ Auto-detects per-slide narration audio (and video "cameo" recordings)
- 🔴 Replays the presenter's **laser pointer** movements, synced to the narration at any speed
- ⚡ Speed chips: 0.75x → 3x, pitch-corrected
- ⏭ Auto-advance to the next slide when narration ends
- ⌨️ Keyboard shortcuts: `space` play/pause, `←`/`→` slides, `↑`/`↓` speed

## 🚀 Run it

It's a fully static site — no build step, no dependencies to install.

- Open the [live site](https://selinazarzour.github.io/vop-speed-player/), or
- Clone and serve locally: `python3 -m http.server` then open http://localhost:8000

## 🧠 How it works

A `.pptx` (or `.ppsx` slideshow — same format, different extension) is just a zip file. `ppt/presentation.xml` lists the slides in order, and each
slide's `_rels/slideN.xml.rels` points at the media PowerPoint embedded when you recorded
narration (usually an `.m4a` in `ppt/media/`). We pull those out as blob URLs, render the
slides with [pptx-preview](https://www.npmjs.com/package/pptx-preview), and play the
narration through a media element with `playbackRate` control.

## Tech

Vanilla HTML/CSS/JS + [JSZip](https://stuk.github.io/jszip/) + [pptx-preview](https://www.npmjs.com/package/pptx-preview), loaded from CDN. No bundler, so it deploys anywhere that serves static files.
