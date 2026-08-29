/* ============================================================
   VOP Speed Player
   Everything runs client-side:
   1. JSZip unpacks the .pptx (it's just a zip archive)
   2. We read presentation.xml to get the slide order, and each
      slide's .rels file to find the narration audio/video that
      PowerPoint's "Record Slide Show" embeds per slide
   3. pptx-preview renders the slides visually
   4. A single media element plays the narration with an
      adjustable playbackRate
   ============================================================ */

(() => {
  "use strict";

  const SPEEDS = [0.75, 1, 1.25, 1.5, 1.75, 2, 2.5, 3];
  const NO_AUDIO_ADVANCE_MS = 4000;

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const uploadView = $("uploadView");
  const loadingView = $("loadingView");
  const loadingText = $("loadingText");
  const playerView = $("playerView");
  const dropzone = $("dropzone");
  const fileInput = $("fileInput");
  const uploadError = $("uploadError");
  const renderRoot = $("renderRoot");
  const fallbackSlide = $("fallbackSlide");
  const cameoVideo = $("cameoVideo");
  const noAudioTag = $("noAudioTag");
  const fileNameEl = $("fileName");
  const narrationBadge = $("narrationBadge");
  const playBtn = $("playBtn");
  const prevBtn = $("prevBtn");
  const nextBtn = $("nextBtn");
  const backBtn = $("backBtn");
  const slideNowEl = $("slideNow");
  const slideTotalEl = $("slideTotal");
  const timeNow = $("timeNow");
  const timeTotal = $("timeTotal");
  const progressBar = $("progressBar");
  const progressFill = $("progressFill");
  const speedRow = $("speedRow");
  const autoplayToggle = $("autoplayToggle");

  // ---------- state ----------
  let slides = [];            // [{ mediaUrl, mediaKind, text }]
  let slideEls = null;        // rendered slide DOM nodes (from pptx-preview)
  let current = 0;
  let speed = 1.5;
  let isPlaying = false;
  let noAudioTimer = null;
  let objectUrls = [];
  const audio = new Audio();

  // ---------- helpers ----------
  const MIME = {
    m4a: "audio/mp4", mp3: "audio/mpeg", wav: "audio/wav", aac: "audio/aac",
    ogg: "audio/ogg", mp4: "video/mp4", m4v: "video/mp4", webm: "video/webm",
    mov: "video/quicktime", wma: "audio/x-ms-wma", wmv: "video/x-ms-wmv",
  };
  const ext = (p) => p.split(".").pop().toLowerCase();
  const fmtTime = (s) => {
    if (!isFinite(s)) return "0:00";
    const m = Math.floor(s / 60), sec = Math.floor(s % 60);
    return `${m}:${String(sec).padStart(2, "0")}`;
  };
  const show = (view) => {
    for (const v of [uploadView, loadingView, playerView]) v.hidden = v !== view;
  };
  const parseXml = (str) => new DOMParser().parseFromString(str, "application/xml");

  // Resolve "../media/media1.m4a" relative to "ppt/slides/" -> "ppt/media/media1.m4a"
  function resolvePath(base, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = base.split("/").concat(target.split("/"));
    const out = [];
    for (const p of parts) {
      if (p === "..") out.pop();
      else if (p !== "." && p !== "") out.push(p);
    }
    return out.join("/");
  }

  // ---------- pptx parsing ----------
  async function parsePptx(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);

    const presXmlFile = zip.file("ppt/presentation.xml");
    const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    if (!presXmlFile || !presRelsFile) {
      throw new Error("This doesn't look like a valid .pptx file.");
    }

    // rId -> slide path
    const relsDoc = parseXml(await presRelsFile.async("string"));
    const relMap = {};
    for (const rel of relsDoc.getElementsByTagName("Relationship")) {
      relMap[rel.getAttribute("Id")] = resolvePath("ppt", rel.getAttribute("Target"));
    }

    // ordered slide list from <p:sldIdLst>
    const presDoc = parseXml(await presXmlFile.async("string"));

    // real slide dimensions (EMUs) — decks can be 4:3, 16:9, or custom
    let aspect = 9 / 16, slideCx = 12192000, slideCy = 6858000;
    const sldSz = presDoc.getElementsByTagNameNS("*", "sldSz")[0];
    if (sldSz) {
      const cx = Number(sldSz.getAttribute("cx")), cy = Number(sldSz.getAttribute("cy"));
      if (cx > 0 && cy > 0) { aspect = cy / cx; slideCx = cx; slideCy = cy; }
    }

    const slidePaths = [];
    const sldIds = presDoc.getElementsByTagNameNS("*", "sldId");
    for (const sld of sldIds) {
      const rId = sld.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
        || sld.getAttribute("r:id");
      if (rId && relMap[rId]) slidePaths.push(relMap[rId]);
    }
    if (slidePaths.length === 0) throw new Error("No slides found in this file.");

    const parsed = [];
    for (const slidePath of slidePaths) {
      const entry = { mediaUrl: null, mediaKind: null, text: "", laser: null };

      // extract visible text (used for fallback rendering)
      const slideFile = zip.file(slidePath);
      if (slideFile) {
        const doc = parseXml(await slideFile.async("string"));
        const texts = [];
        for (const t of doc.getElementsByTagNameNS("*", "t")) {
          if (t.textContent.trim()) texts.push(t.textContent.trim());
        }
        entry.text = texts.join(" · ").slice(0, 300);

        // laser pointer traces recorded with "Record Slide Show"
        // (p14:laserTraceLst -> tracePtLst -> tracePt t="ms" x/y in EMU)
        const traces = [];
        for (const lst of doc.getElementsByTagNameNS("*", "laserTraceLst")) {
          for (const ptLst of lst.getElementsByTagNameNS("*", "tracePtLst")) {
            const pts = [];
            for (const pt of ptLst.getElementsByTagNameNS("*", "tracePt")) {
              const t = Number(pt.getAttribute("t")), x = Number(pt.getAttribute("x")), y = Number(pt.getAttribute("y"));
              if (isFinite(t) && isFinite(x) && isFinite(y)) {
                pts.push({ t: t / 1000, x: x / slideCx, y: y / slideCy }); // seconds + fractions
              }
            }
            if (pts.length) { pts.sort((a, b) => a.t - b.t); traces.push(pts); }
          }
        }
        if (traces.length) entry.laser = traces;
      }

      // find narration media in the slide's rels
      const dir = slidePath.substring(0, slidePath.lastIndexOf("/"));
      const name = slidePath.substring(slidePath.lastIndexOf("/") + 1);
      const relsFile = zip.file(`${dir}/_rels/${name}.rels`);
      if (relsFile) {
        const doc = parseXml(await relsFile.async("string"));
        let audioTarget = null, videoTarget = null;
        for (const rel of doc.getElementsByTagName("Relationship")) {
          const type = rel.getAttribute("Type") || "";
          const target = rel.getAttribute("Target") || "";
          if (rel.getAttribute("TargetMode") === "External") continue;
          if (/\/(audio|media)$/.test(type) && /\.(m4a|mp3|wav|aac|ogg|wma)$/i.test(target)) {
            audioTarget = audioTarget || target;
          } else if (/\/video$/.test(type) || /\.(mp4|m4v|webm|mov|wmv)$/i.test(target)) {
            videoTarget = videoTarget || target;
          }
        }
        const chosen = audioTarget || videoTarget;
        if (chosen) {
          const mediaPath = resolvePath(dir, chosen);
          const mediaFile = zip.file(mediaPath);
          if (mediaFile) {
            const blob = new Blob([await mediaFile.async("arraybuffer")], { type: MIME[ext(mediaPath)] || "application/octet-stream" });
            const url = URL.createObjectURL(blob);
            objectUrls.push(url);
            entry.mediaUrl = url;
            entry.mediaKind = audioTarget ? "audio" : "video";
          }
        }
      }
      parsed.push(entry);
    }
    return { slides: parsed, aspect };
  }

  // ---------- slide rendering ----------
  function renderSlides(arrayBuffer, count, aspect) {
    renderRoot.innerHTML = "";
    slideEls = null;
    // fit the deck's true aspect ratio inside the stage width AND the viewport height
    // measure from <main> — renderRoot is inside the still-hidden player view
    const mainEl = document.querySelector("main");
    const avail = mainEl ? mainEl.clientWidth - 42 : 0; // minus main padding + stage border
    let width = Math.min(avail > 300 ? avail : 960, 1040);
    let height = Math.round(width * aspect);
    const maxHeight = Math.max(320, Math.round(window.innerHeight * 0.62));
    if (height > maxHeight) {
      height = maxHeight;
      width = Math.round(height / aspect);
    }
    return new Promise((resolve) => {
      let settled = false;
      const done = () => { if (!settled) { settled = true; resolve(); } };
      try {
        const previewer = pptxPreview.init(renderRoot, { width, height });
        Promise.resolve(previewer.preview(arrayBuffer))
          .then(() => {
            // pptx-preview stacks all slides; find the container whose
            // child count equals the slide count so we can page them
            slideEls = findSlideElements(renderRoot, count);
            done();
          })
          .catch(done);
        setTimeout(done, 15000); // safety net
      } catch (e) {
        console.warn("pptx-preview failed, using text fallback", e);
        done();
      }
    });
  }

  function findSlideElements(root, count) {
    const queue = [root];
    while (queue.length) {
      const el = queue.shift();
      if (el !== root && el.children.length === count && count > 0) {
        return Array.from(el.children);
      }
      queue.push(...el.children);
    }
    if (root.children.length === 1 && count === 1) return Array.from(root.children);
    return null;
  }

  // PowerPoint auto-shrinks text that exceeds its box ("autofit"), but
  // pptx-preview renders fixed-size shape boxes and lets text overflow.
  // Mimic autofit: scale down a shape's inline font sizes until the text
  // fits its box. Runs lazily the first time a slide becomes visible
  // (elements must be visible to measure).
  function autofitSlideText(slideEl) {
    if (!slideEl || slideEl.dataset.autofitDone) return;
    slideEl.dataset.autofitDone = "1";
    for (const shape of slideEl.querySelectorAll(".shape-wrapper")) {
      const tw = shape.querySelector(".text-wrapper");
      if (!tw) continue;
      const boxH = shape.clientHeight;
      if (!boxH) continue;
      const runs = [];
      for (const n of tw.querySelectorAll("*")) {
        const fs = parseFloat(n.style && n.style.fontSize);
        if (fs) runs.push({ n, fs, ls: parseFloat(n.style.letterSpacing) || 0 });
      }
      if (!runs.length) continue;
      let scale = 1;
      const MIN_SCALE = 0.5;
      while (tw.scrollHeight > boxH + 2 && scale > MIN_SCALE) {
        scale = Math.max(MIN_SCALE, scale - 0.05);
        for (const r of runs) {
          r.n.style.fontSize = `${r.fs * scale}px`;
          if (r.ls) r.n.style.letterSpacing = `${r.ls * scale}px`;
        }
      }
    }
  }

  function showSlide(i) {
    current = Math.max(0, Math.min(i, slides.length - 1));
    slideNowEl.textContent = current + 1;
    prevBtn.disabled = current === 0;
    nextBtn.disabled = current === slides.length - 1;
    clearTimeout(noAudioTimer);

    // visual
    if (slideEls) {
      slideEls.forEach((el, idx) => { el.style.display = idx === current ? "" : "none"; });
      renderRoot.hidden = false;
      fallbackSlide.hidden = true;
      autofitSlideText(slideEls[current]);
    } else {
      renderRoot.hidden = true;
      fallbackSlide.hidden = false;
      fallbackSlide.innerHTML = `<h2>Slide ${current + 1}</h2><p>${escapeHtml(slides[current].text) || "(no text on this slide)"}</p><p style="font-size:.8rem;opacity:.6">Slide preview unavailable — narration playback still works.</p>`;
    }

    // narration
    const s = slides[current];
    audio.pause();
    cameoVideo.pause();
    cameoVideo.hidden = true;
    noAudioTag.hidden = !!s.mediaUrl;
    playBtn.disabled = !s.mediaUrl;
    progressFill.style.width = "0%";
    timeNow.textContent = "0:00";
    timeTotal.textContent = "0:00";

    if (s.mediaUrl) {
      if (s.mediaKind === "video") {
        cameoVideo.src = s.mediaUrl;
        cameoVideo.hidden = false;
      } else {
        audio.src = s.mediaUrl;
      }
      setSpeedOnMedia();
      if (isPlaying) playMedia();
    } else if (isPlaying && autoplayToggle.checked) {
      // no narration on this slide: auto-advance after a short pause
      noAudioTimer = setTimeout(() => {
        if (current < slides.length - 1) showSlide(current + 1);
        else stopPlayback();
      }, NO_AUDIO_ADVANCE_MS);
    }
    updatePlayIcon();
  }

  function escapeHtml(s) {
    return (s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // ---------- laser pointer replay ----------
  // Timestamps in the traces are relative to the slide's narration start,
  // so reading currentTime keeps the dot in sync at ANY playback speed.
  const stage = $("stage");
  const laserDot = $("laserDot");
  const LASER_LINGER_S = 0.4;

  function laserPosAt(traces, t) {
    for (const pts of traces) {
      if (t < pts[0].t || t > pts[pts.length - 1].t + LASER_LINGER_S) continue;
      if (t >= pts[pts.length - 1].t) return pts[pts.length - 1];
      for (let i = 1; i < pts.length; i++) {
        if (t <= pts[i].t) {
          const a = pts[i - 1], b = pts[i];
          const f = (t - a.t) / Math.max(b.t - a.t, 1e-6);
          return { x: a.x + (b.x - a.x) * f, y: a.y + (b.y - a.y) * f };
        }
      }
    }
    return null;
  }

  function laserLoop() {
    requestAnimationFrame(laserLoop);
    const s = slides[current];
    if (playerView.hidden || !s || !s.laser || !s.mediaUrl) { laserDot.hidden = true; return; }
    const pos = laserPosAt(s.laser, activeMedia().currentTime);
    if (!pos) { laserDot.hidden = true; return; }
    const host = (slideEls && slideEls[current]) || renderRoot;
    const hostRect = host.getBoundingClientRect();
    const stageRect = stage.getBoundingClientRect();
    laserDot.style.left = `${hostRect.left - stageRect.left + pos.x * hostRect.width}px`;
    laserDot.style.top = `${hostRect.top - stageRect.top + pos.y * hostRect.height}px`;
    laserDot.hidden = false;
  }
  requestAnimationFrame(laserLoop);

  // ---------- playback ----------
  const activeMedia = () => (slides[current] && slides[current].mediaKind === "video" ? cameoVideo : audio);

  function setSpeedOnMedia() {
    audio.playbackRate = speed;
    audio.preservesPitch = true;
    cameoVideo.playbackRate = speed;
    cameoVideo.preservesPitch = true;
  }

  function playMedia() {
    setSpeedOnMedia();
    activeMedia().play().catch(() => { /* autoplay policy: user will press play */ });
  }

  function togglePlay() {
    if (!slides[current] || !slides[current].mediaUrl) return;
    const m = activeMedia();
    if (m.paused) { isPlaying = true; playMedia(); }
    else { isPlaying = false; m.pause(); }
    updatePlayIcon();
  }

  function stopPlayback() {
    isPlaying = false;
    audio.pause();
    cameoVideo.pause();
    clearTimeout(noAudioTimer);
    updatePlayIcon();
  }

  function updatePlayIcon() {
    const m = activeMedia();
    playBtn.textContent = m && !m.paused ? "❚❚" : "▶";
  }

  function onMediaEnded() {
    if (autoplayToggle.checked && current < slides.length - 1) {
      showSlide(current + 1);
    } else {
      stopPlayback();
    }
  }

  function onTimeUpdate() {
    const m = activeMedia();
    if (!m.duration) return;
    progressFill.style.width = `${(m.currentTime / m.duration) * 100}%`;
    timeNow.textContent = fmtTime(m.currentTime);
    timeTotal.textContent = fmtTime(m.duration);
  }

  for (const m of [audio, cameoVideo]) {
    m.addEventListener("ended", onMediaEnded);
    m.addEventListener("timeupdate", onTimeUpdate);
    m.addEventListener("loadedmetadata", onTimeUpdate);
    m.addEventListener("play", updatePlayIcon);
    m.addEventListener("pause", updatePlayIcon);
    m.addEventListener("ratechange", () => { /* keep pitch natural */ m.preservesPitch = true; });
  }

  progressBar.addEventListener("click", (e) => {
    const m = activeMedia();
    if (!m.duration) return;
    const rect = progressBar.getBoundingClientRect();
    m.currentTime = ((e.clientX - rect.left) / rect.width) * m.duration;
  });

  // ---------- speed chips ----------
  function buildSpeedChips() {
    for (const s of SPEEDS) {
      const b = document.createElement("button");
      b.className = "speed-chip" + (s === speed ? " active" : "");
      b.textContent = `${s}x`;
      b.dataset.speed = s;
      b.addEventListener("click", () => setSpeed(s));
      speedRow.appendChild(b);
    }
  }
  function setSpeed(s) {
    speed = s;
    setSpeedOnMedia();
    for (const chip of speedRow.querySelectorAll(".speed-chip")) {
      chip.classList.toggle("active", Number(chip.dataset.speed) === s);
    }
  }

  // ---------- file handling ----------
  async function handleFile(file) {
    uploadError.hidden = true;
    if (!file || !/\.(pptx|ppsx)$/i.test(file.name)) {
      uploadError.textContent = "Please choose a .pptx or .ppsx file (PowerPoint). Older .ppt files aren't supported — re-save as .pptx first.";
      uploadError.hidden = false;
      return;
    }
    show(loadingView);
    loadingText.textContent = "Unpacking your presentation…";
    try {
      const buf = await file.arrayBuffer();
      cleanup();
      const parsed = await parsePptx(buf);
      slides = parsed.slides;
      loadingText.textContent = "Rendering slides…";
      await renderSlides(buf, slides.length, parsed.aspect);

      fileNameEl.textContent = file.name;
      slideTotalEl.textContent = slides.length;
      const withAudio = slides.filter((s) => s.mediaUrl).length;
      narrationBadge.hidden = withAudio === 0;
      narrationBadge.textContent = `🎙️ narration on ${withAudio}/${slides.length} slides`;

      isPlaying = false;
      show(playerView);
      showSlide(0);
      if (withAudio === 0) {
        noAudioTag.textContent = "no narration found in this file";
      } else {
        noAudioTag.textContent = "no narration on this slide";
      }
    } catch (err) {
      console.error(err);
      show(uploadView);
      uploadError.textContent = `Couldn't open that file: ${err.message}`;
      uploadError.hidden = false;
    }
  }

  function cleanup() {
    stopPlayback();
    for (const u of objectUrls) URL.revokeObjectURL(u);
    objectUrls = [];
    slides = [];
    slideEls = null;
    audio.removeAttribute("src");
    cameoVideo.removeAttribute("src");
  }

  // ---------- events ----------
  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") fileInput.click(); });
  fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));

  for (const ev of ["dragenter", "dragover"]) {
    dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("dragover"); });
  }
  for (const ev of ["dragleave", "drop"]) {
    dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove("dragover"); });
  }
  dropzone.addEventListener("drop", (e) => handleFile(e.dataTransfer.files[0]));

  // allow dropping anywhere on the page too
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    if (!playerView.hidden || !uploadView.hidden) {
      if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    }
  });

  playBtn.addEventListener("click", togglePlay);
  prevBtn.addEventListener("click", () => showSlide(current - 1));
  nextBtn.addEventListener("click", () => showSlide(current + 1));
  backBtn.addEventListener("click", () => { cleanup(); fileInput.value = ""; show(uploadView); });

  document.addEventListener("keydown", (e) => {
    if (playerView.hidden) return;
    if (e.target.tagName === "INPUT") return;
    switch (e.key) {
      case " ": e.preventDefault(); togglePlay(); break;
      case "ArrowLeft": showSlide(current - 1); break;
      case "ArrowRight": showSlide(current + 1); break;
      case "ArrowUp": e.preventDefault(); bumpSpeed(1); break;
      case "ArrowDown": e.preventDefault(); bumpSpeed(-1); break;
    }
  });
  function bumpSpeed(dir) {
    const i = SPEEDS.indexOf(speed);
    const next = SPEEDS[Math.max(0, Math.min(SPEEDS.length - 1, i + dir))];
    setSpeed(next);
  }

  // ---------- init ----------
  buildSpeedChips();
  show(uploadView);
})();
