/* Noisy Colouring Book – web version matching VB6 screen structure
 *
 * VB6 flow:
 *   START.FRM  – splash background + horizontal row of category buttons (DirButtons)
 *   CREATE.FRM – colouring canvas 15/16 + palette 1/16
 *                clicking DirButton goes DIRECTLY here, loads first WMF
 *   LOADPIC.FRM – modal dialog opened from CREATE's "Load Picture" tool button
 *                 one WMF preview at a time, prev/OK/next
 */
(function () {
  'use strict';

  let data = null;
  let currentCategory = null;
  let currentScene = null;
  let selectedColour = 0;
  let svgEl = null;
  let undoState = null;
  let originalFills = null;
  let playSounds = true;
  let blending = false;
  let structuralShapes = new Set();
  const background = '#FFFFFF';

  // Colour cursors – palette index → matching-colour cursor file.
  // Palette: RED(0), GREEN(1), BLUE(2), YELLOW(3), MAGENTA(4),
  //          CYAN(5), WHITE(6), BLACK(7→clock)
  const cursorFiles = [
    'cursors/BIGRED.png',     // 0 red
    'cursors/BIGGREEN.png',   // 1 green
    'cursors/BIGBLUE.png',    // 2 blue
    'cursors/BIGYELLO.png',   // 3 yellow
    'cursors/BIGMAGENTA.png', // 4 magenta
    'cursors/BIGCYAN.png',    // 5 cyan
    'cursors/BIGWHITE.png',   // 6 white
    'cursors/CLOCK.png'       // 7 black → clock cursor
  ];
  // Hotspot as fraction of cursor size [x, y].
  // Arrows: tip at top-left (0,0). Clock: centre (0.5,0.5).
  const cursorHotspots = [
    [0, 0], [0, 0], [0, 0], [0, 0],
    [0, 0], [0, 0], [0, 0], [0.5, 0.5]
  ];

  // ── Scaled cursors — size proportional to window ───────────────────
  const cursorImageCache = {};   // url → Image element
  const scaledCursorCache = {};  // "url_size" → data-URL

  function cursorSize() {
    const minDim = Math.min(window.innerWidth, window.innerHeight);
    return Math.round(Math.max(12, Math.min(48, minDim / 30)));
  }

  function loadCursorImage(url) {
    if (cursorImageCache[url]) return Promise.resolve(cursorImageCache[url]);
    return new Promise(resolve => {
      const img = new Image();
      img.onload = () => { cursorImageCache[url] = img; resolve(img); };
      img.onerror = () => resolve(null);
      img.src = url;
    });
  }

  async function scaledCursorUrl(url, size) {
    const key = url + '_' + size;
    if (scaledCursorCache[key]) return scaledCursorCache[key];
    const img = await loadCursorImage(url);
    if (!img) return null;
    const c = document.createElement('canvas');
    c.width = size; c.height = size;
    c.getContext('2d').drawImage(img, 0, 0, size, size);
    const dataUrl = c.toDataURL('image/png');
    scaledCursorCache[key] = dataUrl;
    return dataUrl;
  }

  async function updateCursor() {
    if (!svgEl) return;
    const file = cursorFiles[selectedColour] || cursorFiles[0];
    const size = cursorSize();
    const hotspot = cursorHotspots[selectedColour] || [0, 0];
    const hx = Math.round(hotspot[0] * size);
    const hy = Math.round(hotspot[1] * size);
    const scaled = await scaledCursorUrl(file, size);
    if (scaled) {
      svgEl.style.cursor = `url(${scaled}) ${hx} ${hy}, pointer`;
    } else {
      svgEl.style.cursor = `url(${file}) 0 0, pointer`;
    }
  }

  // Re-scale cursors when window resizes (debounced)
  let resizeTimer = null;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      // Clear scaled cache so new sizes are generated
      Object.keys(scaledCursorCache).forEach(k => delete scaledCursorCache[k]);
      updateCursor();
    }, 200);
  });

  // Categories – all displayed at once (only categories with icons are included)

  // Scene sequential navigation – VB6 LOADPIC shows one at a time
  let sceneIndex = 0;

  // ── IndexedDB persistence ──────────────────────────────────────────
  // Save/restore fill colours for each scene so colouring is preserved.

  const DB_NAME = 'NoisyColouringBook';
  const DB_VERSION = 11;  // bumped: reload updated default-fills.json
  const STORE_NAME = 'sceneFills';
  let db = null;

  function openDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const d = req.result;
        if (d.objectStoreNames.contains(STORE_NAME)) {
          d.deleteObjectStore(STORE_NAME);
        }
        d.createObjectStore(STORE_NAME);
      };
      req.onsuccess = () => { db = req.result; resolve(db); };
      req.onerror = () => { console.warn('IndexedDB unavailable'); resolve(null); };
    });
  }

  function saveFills() {
    if (!db || !svgEl || !currentScene) return;
    const fills = {};
    svgEl.querySelectorAll('[data-shape]').forEach(el => {
      fills[el.getAttribute('data-shape')] = el.getAttribute('fill');
    });
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(fills, currentScene.wmf);
  }

  function loadFills(wmfPath) {
    if (!db) return Promise.resolve(null);
    return new Promise(resolve => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(wmfPath);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  }

  // First-run initialisation: load pre-computed fills from default-fills.json
  // (exported from the WMF editor with manual corrections applied).
  async function initializeAllScenes() {
    if (!db || !data) return;
    // Check if store already has entries (not first run)
    const count = await new Promise(resolve => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).count();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(0);
    });
    if (count > 0) return;

    try {
      const resp = await fetch('default-fills.json');
      const allFills = await resp.json();
      for (const [wmfPath, fills] of Object.entries(allFills)) {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        tx.objectStore(STORE_NAME).put(fills, wmfPath);
      }
    } catch (e) {
      console.warn('Failed to load default fills:', e);
    }
  }

  // Audio
  const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  const audioCache = {};

  async function loadSound(url) {
    if (audioCache[url]) return audioCache[url];
    try {
      const resp = await fetch(url);
      if (!resp.ok) return null;
      const buf = await resp.arrayBuffer();
      const decoded = await audioCtx.decodeAudioData(buf);
      audioCache[url] = decoded;
      return decoded;
    } catch { return null; }
  }

  function playAudio(url) {
    if (!url || !playSounds) return;
    loadSound(url).then(buf => {
      if (!buf) return;
      const src = audioCtx.createBufferSource();
      src.buffer = buf;
      src.connect(audioCtx.destination);
      src.start();
    });
  }

  function showScreen(id) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    if (audioCtx.state === 'suspended') audioCtx.resume();
  }

  // ── START SCREEN (VB6 START.FRM) ──────────────────────────────────────

  // Categories whose icons should be dynamically rendered from their first WMF
  const dynamicIconSlugs = new Set(['animals', 'fantasy', 'young', 'zapkids']);
  const dynamicIconBg = { animals: '#00ff00', fantasy: '#0000ff', young: '#ffff00', zapkids: '#00ffff' };

  async function buildCategories() {
    const row = document.getElementById('cat-row');
    row.innerHTML = '';
    // Vivid fill colours for dynamic icons (skip white & black)
    const iconPalette = ['#ff0000','#00ff00','#0000ff','#ffff00','#ff00ff','#00ffff'];
    for (const cat of data.categories) {
      const btn = document.createElement('button');
      btn.setAttribute('aria-label', cat.name);

      if (dynamicIconSlugs.has(cat.slug)) {
        // Render first scene as a coloured SVG thumbnail → convert to img
        try {
          const scene = cat.scenes[0];
          const svgText = await wmfToSvg(scene.wmf);
          const tmp = document.createElement('div');
          tmp.innerHTML = svgText;
          const svg = tmp.querySelector('svg');
          if (svg) {
            svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
            // Add white background rect
            const vb = svg.getAttribute('viewBox');
            if (vb) {
              const [vx, vy, vw, vh] = vb.split(/\s+/).map(Number);
              const bgRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
              bgRect.setAttribute('x', vx); bgRect.setAttribute('y', vy);
              bgRect.setAttribute('width', vw); bgRect.setAttribute('height', vh);
              bgRect.setAttribute('fill', dynamicIconBg[cat.slug] || '#ffffff');
              svg.insertBefore(bgRect, svg.firstChild);
            }
            let ci = 0;
            svg.querySelectorAll('[data-shape]').forEach(el => {
              el.setAttribute('fill', iconPalette[ci % iconPalette.length]);
              ci++;
            });
            // Convert to data URL and use as <img> like static icons
            const svgBlob = new Blob([tmp.innerHTML], {type: 'image/svg+xml'});
            const imgUrl = URL.createObjectURL(svgBlob);
            const img = document.createElement('img');
            img.src = imgUrl;
            img.alt = cat.name;
            btn.appendChild(img);
          }
        } catch (e) {
          // Fallback to static icon PNG
          if (cat.icon) {
            const img = document.createElement('img');
            img.src = cat.icon;
            img.alt = cat.name;
            btn.appendChild(img);
          }
        }
      } else if (cat.icon) {
        const img = document.createElement('img');
        img.src = cat.icon;
        img.alt = cat.name;
        btn.appendChild(img);
      }

      btn.addEventListener('click', () => openCategory(cat));
      row.appendChild(btn);
    }
  }

  // VB6: Form_Load plays start.mid automatically.
  // Browsers require a user gesture for audio, so we try on init and
  // fall back to playing on the first interaction anywhere on the page.
  let startPlayed = false;
  function playStartMusic() {
    if (startPlayed) return;
    startPlayed = true;
    if (audioCtx.state === 'suspended') audioCtx.resume();
    playAudio('snd/start.wav');
  }
  document.addEventListener('click', playStartMusic, { once: true });
  document.addEventListener('touchstart', playStartMusic, { once: true });

  // ── CATEGORY → COLOUR SCREEN (VB6 flow) ─────────────────────────────

  function openCategory(cat) {
    currentCategory = cat;
    sceneIndex = 0;
    // VB6: goes directly to CREATE.FRM, loads first WMF from directory
    openScene(cat.scenes[0]);
  }

  // ── SCENE PICKER MODAL (VB6 LOADPIC.FRM) ────────────────────────────
  // In VB6, LOADPIC is opened from SSCommand1(1) "Load Picture" tool button

  function showScenePicker() {
    // Save current colouring so the preview reflects latest changes
    saveFills();
    sceneIndex = currentCategory.scenes.indexOf(
      currentCategory.scenes.find(s => s === currentCategory.scenes[sceneIndex])
    );
    document.getElementById('scene-overlay').classList.add('active');
    showScenePreview();
  }

  function hideScenePicker() {
    document.getElementById('scene-overlay').classList.remove('active');
  }

  async function showScenePreview() {
    const preview = document.getElementById('scene-preview');
    preview.innerHTML = '';
    const scene = currentCategory.scenes[sceneIndex];
    if (!scene) return;

    try {
      const svgText = await wmfToSvg(scene.wmf);
      preview.innerHTML = svgText;
      const svg = preview.querySelector('svg');
      if (svg) {
        svg.removeAttribute('width');
        svg.removeAttribute('height');
        svg.setAttribute('preserveAspectRatio', 'none');

        // Apply saved fills FIRST so structural detection uses editor-corrected colours
        const allShapes = svg.querySelectorAll('[data-shape]');
        const saved = await loadFills(scene.wmf);
        if (saved) {
          allShapes.forEach(el => {
            const idx = el.getAttribute('data-shape');
            if (saved[idx]) el.setAttribute('fill', saved[idx]);
          });
        }

        // Detect structural shapes for preview (same logic as openScene)
        const previewStructural = new Set();
        allShapes.forEach(el => {
          const f = (el.getAttribute('fill') || '').toLowerCase();
          const idx = el.getAttribute('data-shape');
          if (f === '#000000' || f === '#010101') {
            previewStructural.add(idx);
          }
        });

        // Clear non-structural shapes to white, then apply saved colouring
        allShapes.forEach(el => {
          const idx = el.getAttribute('data-shape');
          if (!previewStructural.has(idx)) {
            el.setAttribute('fill', background);
          }
        });
        // Saved fills already applied above; re-apply to restore user colouring
        if (saved) {
          allShapes.forEach(el => {
            const idx = el.getAttribute('data-shape');
            if (saved[idx]) el.setAttribute('fill', saved[idx]);
          });
        }
      }
    } catch (e) {
      preview.textContent = scene.name;
      console.warn('Failed to load WMF preview:', e);
    }

    // VB6 LOADPIC: plays scene sound when navigating
    if (currentCategory.sounds && currentCategory.sounds['0']) {
      playAudio(currentCategory.sounds['0']);
    }
  }

  // VB6 LOADPIC: SSCommand1(0) = prev, SSCommand1(1) = next, wraps around
  document.getElementById('scene-prev').addEventListener('click', () => {
    if (sceneIndex > 0) sceneIndex--;
    else sceneIndex = currentCategory.scenes.length - 1;
    showScenePreview();
  });
  document.getElementById('scene-next').addEventListener('click', () => {
    if (sceneIndex < currentCategory.scenes.length - 1) sceneIndex++;
    else sceneIndex = 0;
    showScenePreview();
  });

  // VB6 LOADPIC: SSCommand2 = OK, loads selected scene into colour screen
  document.getElementById('scene-ok').addEventListener('click', () => {
    hideScenePicker();
    openScene(currentCategory.scenes[sceneIndex]);
  });

  // ── COLOURING SCREEN (VB6 CREATE.FRM) ─────────────────────────────────

  async function openScene(scene) {
    // Save current scene's fills before switching
    saveFills();

    undoState = null;
    originalFills = null;
    svgEl = null;
    currentScene = scene;

    const container = document.getElementById('svg-container');
    container.innerHTML = '';
    showScreen('colour-screen');

    try {
      const svgText = await wmfToSvg(scene.wmf);
      container.innerHTML = svgText;
      svgEl = container.querySelector('svg');
      if (!svgEl) return;

      svgEl.removeAttribute('width');
      svgEl.removeAttribute('height');
      svgEl.setAttribute('preserveAspectRatio', 'none');
      updateCursor();

      // Ensure a background rectangle exists as the first colourable shape.
      // Some WMFs (especially new categories) lack one.
      const vb = svgEl.getAttribute('viewBox');
      if (vb) {
        const hasBackgroundRect = (() => {
          const first = svgEl.querySelector('[data-shape="0"]');
          if (!first) return false;
          // Check if shape 0 is a rect or polygon covering most of the viewBox
          const [, , vw, vh] = vb.split(/\s+/).map(Number);
          const bbox = first.getBBox ? first.getBBox() : null;
          if (!bbox) return false;
          return bbox.width >= vw * 0.9 && bbox.height >= vh * 0.9;
        })();
        if (!hasBackgroundRect) {
          const [vx, vy, vw, vh] = vb.split(/\s+/).map(Number);
          const bgRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          bgRect.setAttribute('x', vx);
          bgRect.setAttribute('y', vy);
          bgRect.setAttribute('width', vw);
          bgRect.setAttribute('height', vh);
          bgRect.setAttribute('fill', background);
          bgRect.setAttribute('data-shape', 'bg');
          svgEl.insertBefore(bgRect, svgEl.firstChild);
        }
      }

      // Apply saved fills from IndexedDB FIRST (so structural detection
      // uses the editor-corrected colours, not raw WMF colours).
      const allShapes = svgEl.querySelectorAll('[data-shape]');
      const saved = await loadFills(scene.wmf);
      if (saved) {
        allShapes.forEach(el => {
          const idx = el.getAttribute('data-shape');
          if (saved[idx]) el.setAttribute('fill', saved[idx]);
        });
      }

      // Detect "structural" shapes — only black (#000000 / #010101) shapes
      // are outlines. Everything else is paintable.
      structuralShapes = new Set();
      allShapes.forEach(el => {
        const f = (el.getAttribute('fill') || '').toLowerCase();
        const idx = el.getAttribute('data-shape');
        if (idx === 'bg') return;
        if (f === '#000000' || f === '#010101') {
          structuralShapes.add(idx);
        }
      });

      // VB6: MDraw.FillColor = background on load
      // Clear non-structural shapes to white, then restore any saved user colouring
      originalFills = new Map();
      allShapes.forEach(el => {
        const idx = el.getAttribute('data-shape');
        originalFills.set(idx, el.getAttribute('fill'));
        if (!structuralShapes.has(idx)) {
          if (saved && saved[idx]) {
            el.setAttribute('fill', saved[idx]);
          } else {
            el.setAttribute('fill', background);
          }
        }
      });

      svgEl.addEventListener('click', onSvgClick);
      svgEl.addEventListener('touchstart', onSvgTouch, { passive: false });
    } catch (err) {
      console.error('Failed to load WMF:', err);
    }
  }

  // Find the shape element from a click/touch event.
  // If the click lands on a stroke-only outline (no data-shape),
  // use elementsFromPoint to find the filled shape underneath.
  function findShapeAt(x, y) {
    const els = document.elementsFromPoint(x, y);
    for (const el of els) {
      if (!svgEl.contains(el)) continue;
      let target = el;
      while (target && target !== svgEl) {
        if (target.hasAttribute && target.hasAttribute('data-shape')) return target;
        target = target.parentElement;
      }
    }
    return null;
  }

  function getShapeFromEvent(e) {
    // First try the direct target (fast path)
    let target = e.target;
    while (target && target !== svgEl) {
      if (target.hasAttribute && target.hasAttribute('data-shape')) return target;
      target = target.parentElement;
    }
    // Clicked on outline/stroke — probe all elements at that point
    return findShapeAt(e.clientX, e.clientY);
  }

  function onSvgTouch(e) {
    e.preventDefault();
    const touch = e.touches[0];
    const shape = findShapeAt(touch.clientX, touch.clientY);
    if (shape) fillShape(shape);
  }

  function onSvgClick(e) {
    const shape = getShapeFromEvent(e);
    if (shape) fillShape(shape);
  }

  // ── Colour blending (VB6 MDraw_MouseDown) ─────────────────────────────

  function hexToRgb(hex) {
    return {
      r: parseInt(hex.slice(1, 3), 16),
      g: parseInt(hex.slice(3, 5), 16),
      b: parseInt(hex.slice(5, 7), 16)
    };
  }

  function rgbToHex(r, g, b) {
    return '#' + [r, g, b].map(c =>
      Math.max(0, Math.min(255, Math.round(c))).toString(16).padStart(2, '0')
    ).join('');
  }

  function blendColour(targetHex, currentHex) {
    const t = hexToRgb(targetHex), c = hexToRgb(currentHex);
    return rgbToHex(
      (t.r + 3 * c.r) / 4,
      (t.g + 3 * c.g) / 4,
      (t.b + 3 * c.b) / 4
    );
  }

  function normalizeColour(c) {
    if (!c) return background;
    c = c.trim().toLowerCase();
    if (c.length === 4 && c[0] === '#')
      return '#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3];
    if (/^rgb/.test(c)) {
      const m = c.match(/(\d+)/g);
      if (m && m.length >= 3) return rgbToHex(+m[0], +m[1], +m[2]);
    }
    return c;
  }

  // VB6 MDrawPro groups shapes into MDCR regions ({...}).
  // Clicking any shape in a region fills ALL shapes in that region.
  // The WMF converter wraps individual shapes in their own mdcr-region,
  // so we walk up through groups until we find one with multiple shapes,
  // stopping if we reach the root (outermost) region.
  function getGroupSiblings(shape) {
    // Walk up to the innermost g.mdcr-region containing this shape
    let group = shape.parentElement;
    while (group && group !== svgEl) {
      if (group.matches && group.matches('g.mdcr-region')) break;
      group = group.parentElement;
    }
    if (!group || group === svgEl) return [shape];

    // Total shapes in the SVG — used to cap overly-broad groups
    const totalShapes = svgEl.querySelectorAll('[data-shape]').length;

    // Walk up through mdcr-region groups looking for a meaningful region.
    // A per-shape wrapper has only 1 descendant shape; keep going up.
    // Stop at the root region (no parent mdcr-region) — fill individually.
    while (group && group !== svgEl) {
      // Check if root region (no parent mdcr-region)
      let isNested = false;
      let p = group.parentElement;
      while (p && p !== svgEl) {
        if (p.matches && p.matches('g.mdcr-region')) { isNested = true; break; }
        p = p.parentElement;
      }
      if (!isNested) return [shape]; // root — fill individually

      // Count all descendant shapes in this group
      const descendants = group.querySelectorAll('[data-shape]');
      // If the group covers most of the image, treat it as root-like
      if (descendants.length > totalShapes * 0.5) return [shape];
      if (descendants.length > 1) return [...descendants];

      // Only 1 shape — this is a per-shape wrapper, try parent group
      group = group.parentElement;
      while (group && group !== svgEl) {
        if (group.matches && group.matches('g.mdcr-region')) break;
        group = group.parentElement;
      }
    }
    return [shape];
  }

  function fillShape(el) {
    // Don't fill structural shapes (they form outlines)
    const elIdx = el.getAttribute('data-shape');
    if (structuralShapes.has(elIdx)) return;

    const palette = data.palette;
    const targetColour = palette[selectedColour];

    // Save undo state – VB6: Image1.picture = MDraw.picture
    undoState = new Map();
    svgEl.querySelectorAll('[data-shape]').forEach(s => {
      undoState.set(s.getAttribute('data-shape'), s.getAttribute('fill'));
    });

    // Fill all shapes in the same MDCR region group (skip structural shapes)
    const groupShapes = getGroupSiblings(el);
    groupShapes.forEach(shape => {
      const idx = shape.getAttribute('data-shape');
      if (structuralShapes.has(idx)) return;
      const currentFill = normalizeColour(shape.getAttribute('fill'));
      let newColour;
      if (!blending || currentFill === background) {
        newColour = targetColour;
      } else {
        newColour = blendColour(targetColour, currentFill);
      }
      shape.setAttribute('fill', newColour);
    });

    if (currentCategory && currentCategory.sounds[String(selectedColour)]) {
      playAudio(currentCategory.sounds[String(selectedColour)]);
    }
  }

  // ── Palette ───────────────────────────────────────────────────────────

  function buildPalette() {
    const container = document.getElementById('colour-buttons');
    container.innerHTML = '';
    data.palette.forEach((colour, i) => {
      const btn = document.createElement('button');
      btn.className = 'colour-btn' + (i === 0 ? ' selected' : '');
      btn.style.backgroundColor = colour;
      btn.setAttribute('aria-label', 'Colour ' + (i + 1));
      btn.addEventListener('click', () => selectColour(i));
      // VB6: SSRibbon3_MouseMove plays sound on hover
      btn.addEventListener('mouseenter', () => {
        if (currentCategory && currentCategory.sounds[String(i)])
          playAudio(currentCategory.sounds[String(i)]);
      });
      container.appendChild(btn);
    });
    document.getElementById('current-colour').style.backgroundColor = data.palette[0];
  }

  function selectColour(i) {
    selectedColour = i;
    document.querySelectorAll('.colour-btn').forEach(b => b.classList.remove('selected'));
    document.querySelectorAll('.colour-btn')[i].classList.add('selected');
    document.getElementById('current-colour').style.backgroundColor = data.palette[i];
    updateCursor();
    playAudio('snd/tools/tool0.wav');
  }

  // ── Tools ─────────────────────────────────────────────────────────────

  // VB6 SSCommand1(0) = Undo
  document.getElementById('btn-undo').addEventListener('click', doUndo);
  function doUndo() {
    if (!undoState || !svgEl) return;
    svgEl.querySelectorAll('[data-shape]').forEach(el => {
      const idx = el.getAttribute('data-shape');
      if (undoState.has(idx)) el.setAttribute('fill', undoState.get(idx));
    });
    undoState = null;
    playAudio('snd/tools/tool1.wav');
  }

  // VB6 SSCommand1(1) = Load Picture (opens LOADPIC modal)
  document.getElementById('btn-load').addEventListener('click', () => {
    if (!currentCategory) return;
    showScenePicker();
  });

  // VB6 SSCommand1(2) = Clear
  document.getElementById('btn-clear').addEventListener('click', doClear);
  function doClear() {
    if (!svgEl) return;
    undoState = new Map();
    svgEl.querySelectorAll('[data-shape]').forEach(el => {
      undoState.set(el.getAttribute('data-shape'), el.getAttribute('fill'));
    });
    svgEl.querySelectorAll('[data-shape]').forEach(el => {
      const idx = el.getAttribute('data-shape');
      // Preserve structural shapes (outlines formed by filled background areas)
      if (!structuralShapes.has(idx)) {
        el.setAttribute('fill', background);
      }
    });
    playAudio('snd/tools/tool2.wav');
    saveFills();
  }

  // VB6: SSCommand1_MouseMove plays tool sounds on hover
  // VB6 buttons: Undo=tool0, Load=tool1, Clear=tool2
  // Web additions (Sound, Home) reuse tool sounds for consistency
  document.getElementById('btn-undo').addEventListener('mouseenter', () => playAudio('snd/tools/tool0.wav'));
  document.getElementById('btn-load').addEventListener('mouseenter', () => playAudio('snd/tools/tool1.wav'));
  document.getElementById('btn-clear').addEventListener('mouseenter', () => playAudio('snd/tools/tool2.wav'));
  document.getElementById('btn-sound').addEventListener('mouseenter', () => playAudio('snd/tools/tool1.wav'));
  document.getElementById('btn-back-colour').addEventListener('mouseenter', () => playAudio('snd/tools/tool0.wav'));

  // VB6: Escape unloads frmCreate, back to START
  document.getElementById('btn-back-colour').addEventListener('click', () => {
    saveFills();
    hideScenePicker();
    showScreen('start-screen');
  });

  // Sound toggle (VB6: F2)
  document.getElementById('btn-sound').addEventListener('click', toggleSound);
  function toggleSound() {
    playSounds = !playSounds;
    const btn = document.getElementById('btn-sound');
    btn.classList.toggle('muted', !playSounds);
    btn.title = playSounds ? 'Sound ON' : 'Sound OFF';
  }

  // ── Keyboard shortcuts (VB6: Form_KeyPress + MDraw_KeyDown) ───────────

  document.addEventListener('keydown', (e) => {
    const colourScreen = document.getElementById('colour-screen');
    if (!colourScreen.classList.contains('active')) return;

    // If scene picker is open, Escape closes it
    if (document.getElementById('scene-overlay').classList.contains('active')) {
      if (e.key === 'Escape') { hideScenePicker(); e.preventDefault(); }
      return;
    }

    const key = e.key;
    if (key >= '0' && key <= '7') { selectColour(parseInt(key)); e.preventDefault(); return; }
    if (key === 'F2') { toggleSound(); e.preventDefault(); return; }
    if (key === 'Delete' || key === 'Backspace' || (e.ctrlKey && key === 'z') || (e.metaKey && key === 'z')) {
      doUndo(); e.preventDefault(); return;
    }
    // VB6: '+' key = load picture
    if (key === '+' || key === '=') { showScenePicker(); e.preventDefault(); return; }
    if (key === 'Enter') { doClear(); e.preventDefault(); return; }
    if (key === 'Escape') { saveFills(); showScreen('start-screen'); e.preventDefault(); return; }
  });

  // Help
  document.getElementById('btn-help').addEventListener('click', () => {
    window.location.href = 'help.html';
  });

  // Save fills when the page unloads (tab close / navigate away)
  window.addEventListener('beforeunload', () => saveFills());

  // ── Randomise floating decorations ─────────────────────────────────────
  // Pick random emojis for each slot; guarantee palette, rainbow and notes
  // are always present (in random positions).
  const requiredEmojis = ['🎨', '🎨', '🌈', '🎵', '🎶'];
  const optionalPool = [
    '⭐','🌟','🖌️','✨','☺','🎀','🎈','💖','🦋',
    '🌸','🍭','🎪','🎉','🌻','🐣','🐾','💫','🧸','🎠',
    '🪁','🫧','🍬','❤️','🌺','🎊','🐱','🐶','🦄','🐠',
    '🌼','🍀','🔮','💎','🪅','🎯','♪','🦋','🌟'
  ];
  function shuffle(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  }
  function randomiseDecos() {
    const decos = document.querySelectorAll('.floating-deco');
    const count = decos.length;
    const picks = [...requiredEmojis];
    const shuffled = shuffle([...optionalPool]);
    // Track which emojis are already used (count occurrences for duplicates)
    const usedCounts = {};
    for (const e of picks) usedCounts[e] = (usedCounts[e] || 0) + 1;
    for (const e of shuffled) {
      if (picks.length >= count) break;
      if (!usedCounts[e]) { picks.push(e); usedCounts[e] = 1; }
    }
    shuffle(picks);
    decos.forEach((el, i) => { el.textContent = picks[i]; });
  }
  randomiseDecos();

  // ── Colouring-book background ──────────────────────────────────────────
  // Load a random pattern WMF, render as SVG outlines (white fills),
  // placed behind the animated rainbow gradient overlay.
  const bgPatterns = [
    'wmf/wierd/PATTERN1.WMF',
    'wmf/wierd/PATTERN2.WMF',
    'wmf/wierd/PATTERN3.WMF',
    'wmf/tessellation/TESS2.WMF',
    'wmf/tessellation/TESS3.WMF',
    'wmf/tessellation/TESS4.WMF',
    'wmf/tessellation/TESS5.WMF',
    'wmf/tessellation/TESS6.WMF',
    'wmf/tessellation/TESS7.WMF',
    'wmf/tessellation/TESS8.WMF',
    'wmf/tessellation/TESS9.WMF'
  ];

  async function loadStartBackground() {
    const container = document.getElementById('start-bg-pattern');
    if (!container) return;
    // Try patterns until we find one that renders visible outlines
    const shuffled = shuffle([...bgPatterns]);
    for (const wmfPath of shuffled) {
      try {
        const svgText = await wmfToSvg(wmfPath);
        container.innerHTML = svgText;
        const svg = container.querySelector('svg');
        if (!svg) continue;
        svg.removeAttribute('width');
        svg.removeAttribute('height');
        svg.setAttribute('preserveAspectRatio', 'none');
        // Set all shapes to white so only outlines (strokes) show
        svg.querySelectorAll('[data-shape]').forEach(el => {
          el.setAttribute('fill', '#ffffff');
        });
        // Check the SVG has visible stroke elements (not just filled shapes)
        const hasStrokes = svg.querySelector('polyline, line, [stroke]:not([stroke="none"])');
        if (hasStrokes) return; // good pattern, keep it
        // No visible outlines — try another pattern
      } catch (e) {
        console.warn('Failed to load start background pattern:', wmfPath, e);
      }
    }
  }

  // ── Init ──────────────────────────────────────────────────────────────

  async function init() {
    try {
      await openDB();
      const resp = await fetch('activities.json?v=2');
      data = await resp.json();
      await buildCategories();
      buildPalette();
      // Load colouring-book pattern background
      await loadStartBackground();
      // First run: save all-white fills for every scene so they start clean
      await initializeAllScenes();
      // VB6: Form_Load plays start.mid — try autoplay (may be blocked by browser)
      playStartMusic();
    } catch (e) {
      console.error('Failed to load activities:', e);
    }
  }

  init();
})();
