// --- CONFIGURATION ---
// Card is exactly 6" × 4" = 152.4 mm × 101.6 mm
const CARD_W_MM = 152.4;
const CARD_H_MM = 101.6;

// Print at 300 DPI for crisp text
const PRINT_DPI  = 300;
const PRINT_W_PX = Math.round((CARD_W_MM / 25.4) * PRINT_DPI); // 1800
const PRINT_H_PX = Math.round((CARD_H_MM / 25.4) * PRINT_DPI); // 1200

// ── POLYGON SPACING — one value controls both screen preview AND PDF ─────────
// Fraction of the map-window height used as padding around the territory polygon.
//   0.05 → tight (polygon nearly fills the map area)
//   0.10 → default balanced look
//   0.20 → zoomed-out with lots of surrounding context
const MAP_PAD_FRAC = 0.10;

// Zoom level used when displaying a Point feature — far enough out to show
// surrounding streets, road names and landmarks for navigation context.
// 15 = good urban detail,  14 = more context / rural roads
const POINT_ZOOM = 17;

const MAPTILER_KEY = '4USK1756KHWJylQrYYep';

let loadedFiles      = {};
let activeFeature    = null;
let checkedNames     = new Set();
let congregationName = '';   // set once per session via the cong-name dialogue

// ─── 1. MAPLIBRE SETUP ──────────────────────────────────────────────────────
const initialStyle = document.getElementById('styleSelect').value;

const map = new maplibregl.Map({
    container: 'map',
    style: `https://api.maptiler.com/maps/${initialStyle}/style.json?key=${MAPTILER_KEY}`,
    center: [0, 0],
    zoom: 1,
    preserveDrawingBuffer: true,
    attributionControl: false,
    fadeDuration: 0,
    interactive: false,   // view is automated — zoom/pan disabled on screen map
});

// Attribution shown via card-attr div in buildCardFrame — scales with card

map.on('load', () => {
    placeTheKh();
    showWelcomeDialogue();
});

// FIX: Handle missing images (stops console errors for 'mall', 'transparent-icon', etc.)
map.on('styleimagemissing', (e) => {
    // Add a 1x1 transparent placeholder synchronously — prevents repeated firing
    if (!map.hasImage(e.id)) {
        const blank = new Uint8Array(4);  // RGBA all zeros = transparent
        map.addImage(e.id, { width: 1, height: 1, data: blank });
    }
});

document.getElementById('styleSelect').addEventListener('change', e => {
    const styleId = e.target.value;
    // Remove active layers/source before the style swap so MapLibre's diff
    // engine doesn't race against the placement engine mid-transition
    ['poly-outline','mask-layer','terr-line','terr-point','terr-point-ring']
        .forEach(id => { try { if (map.getLayer(id)) map.removeLayer(id); } catch(e){} });
    try { if (map.getSource('active-data')) map.removeSource('active-data'); } catch(e){}

    map.setStyle(`https://api.maptiler.com/maps/${styleId}/style.json?key=${MAPTILER_KEY}`);
    map.once('idle', () => {
        placeTheKh();
        if (activeFeature) showTerritory(activeFeature);
    });
});

// ─── 2. CARD FRAME ────────────────────────────────────────────────────────────────────────────
// The card is a real DOM element in the flow (#territory-card-frame inside
// #card-stage). CSS keeps it 152.4:101.6 and centred. JS only injects the
// header/footer and tells MapLibre to resize into #card-window.

function buildCardFrame() {
    const frame = document.getElementById('territory-card-frame');
    if (!frame) return;

    // Measure card as laid out by CSS
    const W     = frame.offsetWidth;
    const H     = frame.offsetHeight;
    const scale = W / 874;                   // 874 = design reference width
    const headH = Math.round(110 * scale);
    const footH = Math.round(110 * scale);
    const fs    = n => Math.round(n * scale) + 'px';
    const pd    = n => Math.round(n * scale) + 'px';

    // Remove old header/footer/attr if rebuilding on resize
    ['card-top', 'card-bottom', 'card-attr'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.remove();
    });

    // ── Header ───────────────────────────────────────────────────────────────────────────────────
    const header = document.createElement('div');
    header.id = 'card-top';
    header.style.cssText = [
        'background:#fff;flex-shrink:0;height:' + headH + 'px;',
        'display:flex;flex-direction:column;justify-content:flex-end;',
        'border:1px solid #000;border-bottom:none;'
    ].join('');
    header.innerHTML =
        '<div style="text-align:center;font-weight:bold;font-size:' + fs(32) + ';margin-bottom:' + pd(8) + ';letter-spacing:0.5px;">' +
            'Territory Map Card' +
        '</div>' +
        '<div style="display:flex;align-items:flex-end;padding:0 ' + pd(25) + ' ' + pd(10) + ';font-size:' + fs(20) + ';font-weight:bold;">' +
            '<span style="white-space:nowrap;margin-right:5px;">Locality</span>' +
            '<span id="card-locality" style="flex:1;border-bottom:2px dotted #000;text-align:left;font-weight:bold;white-space:nowrap;overflow:hidden;padding-left:5px;font-size:' + fs(25) + ';">&nbsp;</span>' +
            '<span style="white-space:nowrap;margin-left:15px;margin-right:5px;">Terr. No.</span>' +
            '<span id="card-terrno" style="width:' + pd(100) + ';border-bottom:2px dotted #000;text-align:center;font-size:' + fs(25) + ';">&nbsp;</span>' +
        '</div>';

    // ── Attribution overlay on map window ───────────────────────────────────────────────────────────
    const attr = document.createElement('div');
    attr.id = 'card-attr';
    attr.style.cssText = [
        'position:absolute;bottom:2px;right:1px;z-index:20;',
        'pointer-events:none;background:rgba(255,255,255,0.82);',
        'padding:1px 5px;font-family:sans-serif;',
        'font-size:' + fs(7) + ';line-height:1.4;color:#444;'
    ].join('');
    attr.textContent = '© OpenStreetMap contributors © MapTiler';

    // ── Footer ───────────────────────────────────────────────────────────────────────────────────
    const footer = document.createElement('div');
    footer.id = 'card-bottom';
    footer.style.cssText = [
        'background:#fff;flex-shrink:0;height:' + footH + 'px;',
        'position:relative;border:1px solid #000;border-top:none;'
    ].join('');
    footer.innerHTML =
        '<div style="text-align:center;font-weight:bold;font-size:' + fs(14) + ';padding-top:' + pd(5) + ';"> ' +
            '(Paste map above or draw in territory)' +
        '</div>' +
        '<div style="font-size:' + fs(21) + ';font-weight:bold;text-align:justify;text-align-last:justify;padding:' + pd(5) + ' ' + pd(25) + ' 0;line-height:1.2;">' +
            'Please keep this card in the envelope. Do not soil, mark, or bend it. Each time the territory is covered, please inform the brother who cares for the territory files.' +
        '</div>' +
        '<div style="position:absolute;bottom:' + pd(8) + ';left:' + pd(25) + ';font-size:' + fs(12) + ';font-weight:normal;">' +
            'S-12-E &nbsp; 6/72' +
        '</div>';

    // Insert: header before card-window, attr inside it, footer after
    const cardWindow = document.getElementById('card-window');
    frame.insertBefore(header, cardWindow);
    cardWindow.appendChild(attr);
    frame.appendChild(footer);

    // Tell MapLibre the container dimensions changed
    if (typeof map !== 'undefined') map.resize();

    // Simple flat padding — fraction of map-window height, same as PDF
    const mapWinH = H - headH - footH;
    window._mapPadding = Math.round(mapWinH * MAP_PAD_FRAC);
}

function cleanNameAndSplit(rawName) {
    if (!rawName) return { number: '', locality: '' };
    let clean = rawName.replace(/^[>}\])\.,;\-\s]+/, '');
    const match = clean.match(/^(\d+)[>}\])\.,;\-\s]*(.*)/);
    return match
        ? { number: match[1], locality: match[2].trim() }
        : { number: '',       locality: clean };
}

function updateCardFields(locality, terrNo) {
    const elL = document.getElementById('card-locality');
    const elT = document.getElementById('card-terrno');
    if (elL) elL.textContent = locality || '\u00a0';
    if (elT) elT.textContent = terrNo   || '\u00a0';
}

// ─── 3. SHOW TERRITORY ────────────────────────────────────────────────────────────────────────
// Geometry-aware layer renderer. Works on any MapLibre instance so it can
// be reused for both the live screen map and the offscreen PDF render map.
function addTerritoryLayers(targetMap, feature, lineWidth) {
    if (!lineWidth) lineWidth = 4;
    const gtype = feature.geometry.type;

    // Clean up any existing layers/source
    ['terr-point', 'terr-point-ring', 'terr-line', 'poly-outline', 'mask-layer']
        .forEach(id => { try { if (targetMap.getLayer(id)) targetMap.removeLayer(id); } catch(e){} });
    try { if (targetMap.getSource('active-data')) targetMap.removeSource('active-data'); } catch(e){}

    const isPolygon  = gtype === 'Polygon'    || gtype === 'MultiPolygon';
    const isLine     = gtype === 'LineString' || gtype === 'MultiLineString';
    const isPoint    = gtype === 'Point'      || gtype === 'MultiPoint';

    // Build mask polygon for polygon features (dims area outside territory)
    const features = [feature];
    if (isPolygon) {
        let coords = gtype === 'Polygon'
            ? feature.geometry.coordinates
            : feature.geometry.coordinates.reduce((a, p) => a.concat(p), []);
        features.unshift({
            type: 'Feature',
            geometry: { type: 'Polygon', coordinates: [
                [[-180,90],[180,90],[180,-90],[-180,-90],[-180,90]],
                ...coords
            ]},
            properties: { type: 'mask' }
        });
    }

    targetMap.addSource('active-data', {
        type: 'geojson',
        data: { type: 'FeatureCollection', features }
    });

    if (isPolygon) {
        targetMap.addLayer({
            id: 'mask-layer', type: 'fill', source: 'active-data',
            filter: ['==', ['get', 'type'], 'mask'],
            paint: { 'fill-color': '#000000', 'fill-opacity': 0.1 }
        });
        targetMap.addLayer({
            id: 'poly-outline', type: 'line', source: 'active-data',
            filter: ['!=', ['get', 'type'], 'mask'],
            layout: { 'line-join': 'round', 'line-cap': 'round' },
            paint: { 'line-color': '#1a6fc4', 'line-width': lineWidth }
        });
    } else if (isLine) {
        // Dashed orange line for routes / boundary lines
        targetMap.addLayer({
            id: 'terr-line', type: 'line', source: 'active-data',
            layout: { 'line-join': 'round', 'line-cap': 'round' },
            paint: { 'line-color': '#e06c00', 'line-width': lineWidth, 'line-dasharray': [3, 2] }
        });
    } else if (isPoint) {
        // White halo ring
        targetMap.addLayer({
            id: 'terr-point-ring', type: 'circle', source: 'active-data',
            paint: {
                'circle-radius':       lineWidth * 4,
                'circle-color':        '#ffffff',
                'circle-stroke-width': lineWidth,
                'circle-stroke-color': '#1a6fc4'
            }
        });
        // Filled centre dot
        targetMap.addLayer({
            id: 'terr-point', type: 'circle', source: 'active-data',
            paint: {
                'circle-radius': lineWidth * 1.5,
                'circle-color':  '#1a6fc4'
            }
        });
    }
}

function showTerritory(feature) {
    activeFeature = feature;
    addTerritoryLayers(map, feature, 4);

    const gtype = feature.geometry.type;
    if (gtype === 'Point' || gtype === 'MultiPoint') {
        // For points fitBounds gives a zero-size box and zooms to max.
        // Instead centre on the point at a readable street-level zoom.
        const coords = gtype === 'Point'
            ? feature.geometry.coordinates
            : feature.geometry.coordinates[0];  // first point of MultiPoint
        map.jumpTo({ center: [coords[0], coords[1]], zoom: POINT_ZOOM, bearing: 0 });
    } else {
        const bbox     = turf.bbox(feature);
        const rotation = calculateBestRotation(feature);
        map.fitBounds(bbox, {
            padding: window._mapPadding || 60,
            bearing: rotation,
            animate: false
        });
    }
    // Update KH sticker and apply colourised film effect after camera settles
    requestAnimationFrame(() => {
        updateKhStickerOnScreen();
        setTimeout(() => applyColorizeEffect(map, feature), 200);
    });
}

// ─── 3b. ROTATION HELPER ────────────────────────────────────────────────────
function calculateBestRotation(feature) {
    const cardAspect = CARD_W_MM / CARD_H_MM;

    const bbox = turf.bbox(feature);
    const w = turf.distance(
        [bbox[0], (bbox[1] + bbox[3]) / 2],
        [bbox[2], (bbox[1] + bbox[3]) / 2],
        { units: 'meters' }
    );
    const h = turf.distance(
        [(bbox[0] + bbox[2]) / 2, bbox[1]],
        [(bbox[0] + bbox[2]) / 2, bbox[3]],
        { units: 'meters' }
    );

    let bestBearing = 0;
    let bestScore   = -Infinity;

    for (let angle = 0; angle < 180; angle += 5) {
        const rad = (angle * Math.PI) / 180;
        const cos = Math.abs(Math.cos(rad));
        const sin = Math.abs(Math.sin(rad));

        const rotW = w * cos + h * sin;
        const rotH = w * sin + h * cos;

        // Penalise angles far from 0/90 to avoid excessive diagonal overhang
        const overhang = Math.sin(2 * rad); // peaks at 45°/135°
        const ratio = rotW / rotH;
        const score = -Math.abs(ratio - cardAspect) - (overhang * 0.3);

        if (score > bestScore) {
            bestScore   = score;
            bestBearing = angle;
        }
    }

    return bestBearing;
}

// ─── 4. KH ICON — hide default star, show custom JW.ORG icon ────────────────
// Accepts any MapLibre instance so it works for both the live map and the
// offscreen PDF render map.
function placeTheKh(targetMap, iconSize) {
    if (iconSize === undefined) iconSize = 0.7;
    if (!targetMap) targetMap = map;
    const style = targetMap.getStyle();
    if (!style) return;

    // 1. Draw a dark blue square icon with "JW." on top and "ORG" below
    if (!targetMap.hasImage('jw-org-icon')) {
        const size = 36;
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = size;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#003580';
        ctx.fillRect(0, 0, size, size);
        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 11px sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('JW.',  size / 2, size / 2 - 6);
        ctx.fillText('ORG',  size / 2, size / 2 + 7);
        const imageData = ctx.getImageData(0, 0, size, size);
        targetMap.addImage('jw-org-icon', { width: size, height: size, data: imageData.data });
    }

    // Ensure we have a transparent icon for the fallback
    if (!targetMap.hasImage('transparent-icon')) {
        const tCanvas = document.createElement('canvas');
        tCanvas.width = 1; tCanvas.height = 1;
        const tCtx = tCanvas.getContext('2d');
        targetMap.addImage('transparent-icon', { width: 1, height: 1, data: tCtx.getImageData(0,0,1,1).data });
    }

    // 2. Hide the default star/icon for Kingdom Hall / Jehovah POIs
    style.layers.forEach(layer => {
        if (layer['source-layer'] !== 'poi') return;
        if (layer.layout && layer.layout['icon-image']) {
            const currentIcon = targetMap.getLayoutProperty(layer.id, 'icon-image');
            // Don't double-patch layers we already patched
            if (Array.isArray(currentIcon) && currentIcon[0] === 'case' &&
                JSON.stringify(currentIcon).includes('Kingdom Hall')) return;
            try {
                targetMap.setLayoutProperty(layer.id, 'icon-image', [
                    'case',
                    ['any',
                        ['in', 'Kingdom Hall', ['get', 'name']],
                        ['in', 'Jehovah',      ['get', 'name']]
                    ],
                    'transparent-icon',
                    currentIcon
                ]);
            } catch (e) {}
        }
    });

    // 3. Add custom JW.ORG icon layer on top
    const sourceId = Object.keys(style.sources).find(k => style.sources[k].type === 'vector');
    if (targetMap.getLayer('jw-org-text')) targetMap.removeLayer('jw-org-text');
    if (sourceId) {
        targetMap.addLayer({
            id: 'jw-org-text',
            type: 'symbol',
            source: sourceId,
            'source-layer': 'poi',
            minzoom: 13,
            filter: [
                'all',
                ['==', ['get', 'class'], 'place_of_worship'],
                ['any',
                    ['in', 'Kingdom Hall', ['get', 'name']],
                    ['in', 'Jehovah',      ['get', 'name']]
                ]
            ],
            layout: {
                'icon-image': 'jw-org-icon',
                'icon-size': iconSize,
                'icon-allow-overlap': true,
                'icon-ignore-placement': true,
                'icon-offset': [0, -8]
            }
        });
    }
}

// ─── 5. DATA LOADING ────────────────────────────────────────────────────────
document.getElementById('fileInput').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;

    let kmlText = '';
    if (file.name.endsWith('.kmz')) {
        const zip = await JSZip.loadAsync(file);
        const kmlFile = Object.keys(zip.files).find(n => n.endsWith('.kml'));
        if (kmlFile) kmlText = await zip.file(kmlFile).async('string');
    } else {
        kmlText = await file.text();
    }
    if (!kmlText) return;

    const dom     = new DOMParser().parseFromString(kmlText, 'text/xml');
    const geojson = toGeoJSON.kml(dom);

    geojson.features.sort((a, b) =>
        (a.properties.name || '').localeCompare(b.properties.name || '', undefined, { numeric: true, sensitivity: 'base' })
    );

    loadedFiles  = { [file.name]: geojson };
    checkedNames.clear();
    renderSidebar();

    // ── KMZ loaded indicator ──────────────────────────────────────────────

    const badge = document.getElementById('kmzFilenameBadge');
    if (badge) {
        badge.textContent = '📁 ' + file.name;
        badge.style.display = 'flex';
    }

    document.getElementById('printAllBtn').disabled = false;
    // Reset toggle button label for fresh file
    const tBtn = document.getElementById('togglePanelBtn');
    if (tBtn) { const l = tBtn.querySelector('.btn-label'); if (l) l.textContent = 'See Maps'; }
    // Zoom to the outer boundary of all loaded features so the user can
    // see the full territory area before the KH placement prompt appears
    const allBbox = turf.bbox({ type: 'FeatureCollection', features: geojson.features });
    map.fitBounds(allBbox, { padding: 40, animate: true, duration: 2200 });
    // Ask about KH sticker once the fly-in has settled
    setTimeout(askKhStickerPlacement, 5200);

    e.target.value = '';
});

// ─── 6. SIDEBAR ─────────────────────────────────────────────────────────────
function renderSidebar() {
    const container = document.getElementById('sidePanelInner');
    container.innerHTML = '';

    const keys = Object.keys(loadedFiles);
    if (!keys.length) return;

    const features = loadedFiles[keys[0]].features;
    features.forEach(f => {
        const info  = cleanNameAndSplit(f.properties.name);
        const label = (info.number ? info.number + ' ' : '') + info.locality;

        const div = document.createElement('div');
        div.className = 'side-item';
        div.dataset.name = f.properties.name;

        const chk = document.createElement('input');
        chk.type    = 'checkbox';
        chk.checked = checkedNames.has(f.properties.name);
        chk.addEventListener('change', () => {
            if (chk.checked) checkedNames.add(f.properties.name);
            else             checkedNames.delete(f.properties.name);
            updatePrintBar();
            updateSelectAllChk();
        });

        const span = document.createElement('span');
        span.className   = 'side-item-label';
        span.textContent = label;
        span.onclick = () => selectTerritory(f, div);

        div.appendChild(chk);
        div.appendChild(span);
        container.appendChild(div);
    });

    updatePrintBar();
}

function selectTerritory(feature, divEl) {
    document.querySelectorAll('.side-item').forEach(d => d.classList.remove('active'));
    if (divEl) divEl.classList.add('active');
    const info = cleanNameAndSplit(feature.properties.name);
    updateCardFields(info.locality, info.number);
    showTerritory(feature);
    // Show the single-card print button whenever a territory is actively displayed
    const oneBtn = document.getElementById('printOneBtn');
    if (oneBtn) oneBtn.style.display = 'flex';
    // Auto-close the side panel so the map is immediately visible
    const panel = document.getElementById('sidePanel');
    const toggleBtn = document.getElementById('togglePanelBtn');
    if (panel) panel.classList.remove('open');
    if (toggleBtn) {
        toggleBtn.classList.remove('active');
        // Update label to hint the user can change their selection
        const lbl = toggleBtn.querySelector('.btn-label');
        if (lbl) lbl.textContent = 'Change Map';
    }
}

function updatePrintBar() {
    const bar = document.getElementById('panelPrintBar');
    const cnt = document.getElementById('selCount');
    const n   = checkedNames.size;
    if (n > 0) {
        bar.classList.add('visible');
        cnt.textContent = `${n} selected`;
    } else {
        bar.classList.remove('visible');
    }
}

function updateSelectAllChk() {
    const keys = Object.keys(loadedFiles);
    if (!keys.length) return;
    const total = loadedFiles[keys[0]].features.length;
    const chk   = document.getElementById('selectAllChk');
    chk.indeterminate = checkedNames.size > 0 && checkedNames.size < total;
    chk.checked       = checkedNames.size === total;
}

function toggleSelectAll() {
    const keys = Object.keys(loadedFiles);
    if (!keys.length) return;
    const features   = loadedFiles[keys[0]].features;
    const allChecked = checkedNames.size === features.length;
    checkedNames.clear();
    if (!allChecked) features.forEach(f => checkedNames.add(f.properties.name));

    document.querySelectorAll('.side-item').forEach(div => {
        const chk = div.querySelector('input[type=checkbox]');
        if (chk) chk.checked = !allChecked;
    });
    updateSelectAllChk();
    updatePrintBar();
}

// ─── 7. PRINT HELPERS ───────────────────────────────────────────────────────
// Renders the card entirely on an offscreen canvas at full print resolution.
// ─── 7. PRINT HELPERS (Updated with Specific Line Break) ────────────────────
async function renderCardToCanvas(feature) {

    // ── Output canvas at full print resolution ──────────────────────────────
    const out  = document.createElement('canvas');
    out.width  = PRINT_W_PX;
    out.height = PRINT_H_PX;
    const ctx  = out.getContext('2d');

    // ── Proportions ─────────────────────────────────────────────────────────
    const BASE_W  = PRINT_W_PX;
    const BASE_H  = PRINT_H_PX;
    const HEAD_H  = Math.round(BASE_H * (110 / 555));  
    const FOOT_H  = Math.round(BASE_H * (110 / 555));
    const MAP_Y   = HEAD_H;
    const MAP_H   = BASE_H - HEAD_H - FOOT_H;

    const FONT    = "'Times New Roman', Times, serif";
    const SANS    = "sans-serif";

    // ── 1. White header ─────────────────────────────────────────────────────
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, BASE_W, HEAD_H);

    // border around header
    ctx.strokeStyle = '#000';
    ctx.lineWidth   = 2;
    ctx.strokeRect(1, 1, BASE_W - 2, HEAD_H - 1);

    // "Territory Map Card" title
    ctx.fillStyle  = '#000';
    ctx.font       = 'bold ' + Math.round(BASE_H * 32 / 555) + 'px ' + FONT;
    ctx.textAlign  = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Territory Map Card', BASE_W / 2, Math.round(HEAD_H * 0.36));

    // Locality / Terr No row
    const rowY    = Math.round(HEAD_H * 0.78);
    const padX    = Math.round(BASE_W * 25 / 874);
    const fldSize = Math.round(BASE_H * 25 / 555);
    const lblSize = Math.round(BASE_H * 20 / 555);
    const dotY    = rowY + Math.round(fldSize * 0.3);

    ctx.font      = 'bold ' + lblSize + 'px ' + FONT;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    ctx.fillText('Locality', padX, rowY);

    const lblW    = ctx.measureText('Locality').width + 6;
    const terrLbl = 'Terr. No.';
    ctx.font      = 'bold ' + lblSize + 'px ' + FONT;
    const terrLblW = ctx.measureText(terrLbl).width;
    const terrBoxW = Math.round(BASE_W * 100 / 874);
    const terrLblX = BASE_W - padX - terrBoxW - terrLblW - 10;

    // dotted underline for locality
    ctx.setLineDash([3, 4]);
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(padX + lblW, dotY);
    ctx.lineTo(terrLblX - 8, dotY);
    ctx.stroke();
    ctx.setLineDash([]);

    // locality text — clamp with ellipsis to fit available space
    const elL = document.getElementById('card-locality');
    const localityText = (elL && elL.textContent.trim() !== ' ') ? elL.textContent.trim() : '';
    ctx.font      = fldSize + 'px ' + FONT;
    ctx.textAlign = 'left';
    const localityMaxW = terrLblX - 8 - (padX + lblW + 4);
    let displayLocality = localityText;
    if (ctx.measureText(displayLocality).width > localityMaxW) {
        while (displayLocality.length > 0 &&
               ctx.measureText(displayLocality + '…').width > localityMaxW) {
            displayLocality = displayLocality.slice(0, -1);
        }
        displayLocality += '…';
    }
    ctx.fillText(displayLocality, padX + lblW + 4, rowY);

    // "Terr. No." label
    ctx.font      = 'bold ' + lblSize + 'px ' + FONT;
    ctx.textAlign = 'left';
    ctx.fillText(terrLbl, terrLblX, rowY);

    // dotted underline for terr no
    ctx.setLineDash([3, 4]);
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(terrLblX + terrLblW + 6, dotY);
    ctx.lineTo(BASE_W - padX, dotY);
    ctx.stroke();
    ctx.setLineDash([]);

    // terr no text
    const elT = document.getElementById('card-terrno');
    const terrNoText = (elT && elT.textContent.trim() !== ' ') ? elT.textContent.trim() : '';
    ctx.font      = fldSize + 'px ' + FONT;
    ctx.textAlign = 'center';
    ctx.fillText(terrNoText, terrLblX + terrLblW + 6 + terrBoxW / 2, rowY);

    // ── 2. Map area — offscreen MapLibre at full print resolution ───────────
    const dX = 0, dY = MAP_Y, dW = BASE_W, dH = MAP_H;

    await new Promise((resolve) => {

        // Create a hidden container sized exactly to the print map area
        const offContainer = document.createElement('div');
        offContainer.style.cssText = [
            'position:fixed;left:-9999px;top:-9999px;',
            'width:' + dW + 'px;height:' + dH + 'px;',
            'pointer-events:none;opacity:0.001;'  // visible to GPU but invisible to user
        ].join('');
        document.body.appendChild(offContainer);

        const styleId = document.getElementById('styleSelect').value;
        const bbox    = turf.bbox(feature);
        const bearing = calculateBestRotation(feature);

        const offMap = new maplibregl.Map({
            container:             offContainer,
            style:                 'https://api.maptiler.com/maps/' + styleId + '/style.json?key=' + MAPTILER_KEY,
            preserveDrawingBuffer: true,
            attributionControl:    false,
            fadeDuration:          0,
            interactive:           false,
            pixelRatio:            2,       // pin to 1 — identical output on any screen/DPI
        });

        offMap.once('load', () => {

            // Suppress missing image errors — synchronous so it never fires twice
            offMap.on('styleimagemissing', (e) => {
                if (!offMap.hasImage(e.id)) {
                    const blank = new Uint8Array(4);
                    offMap.addImage(e.id, { width: 1, height: 1, data: blank });
                }
            });

            // Apply KH / JW.ORG icon to the offscreen map exactly as on screen
            placeTheKh(offMap, 1.4);  // larger icon for print resolution

            const offGtype = feature.geometry.type;
            if (offGtype === 'Point' || offGtype === 'MultiPoint') {
                const offCoords = offGtype === 'Point'
                    ? feature.geometry.coordinates
                    : feature.geometry.coordinates[0];
                offMap.jumpTo({ center: [offCoords[0], offCoords[1]], zoom: POINT_ZOOM, bearing: 0 });
            } else {
                offMap.fitBounds(bbox, {
                    padding: Math.round((BASE_H - HEAD_H - FOOT_H) * MAP_PAD_FRAC),
                    bearing: bearing,
                    animate: false,
                });
            }

            // Geometry-aware territory layers (polygon / line / point)
            addTerritoryLayers(offMap, feature, 6);

            // Wait for tiles to finish, then rAF to ensure GPU frame is committed
            offMap.once('idle', () => {
    setTimeout(() => {
        requestAnimationFrame(() => {
            try {
                ctx.drawImage(offMap.getCanvas(), dX, dY, dW, dH);
            } catch(e) {
                ctx.fillStyle = '#e0e0e0';
                ctx.fillRect(dX, dY, dW, dH);
            }
            applyColorizeOnCanvas(ctx, offMap, feature, dX, dY, dW, dH);
            compositeKhStickerOnCanvas(ctx, offMap);
            offMap.remove();
            offContainer.remove();
            resolve();
        });
    }, 150);
});
        });

        // Safety timeout 15s
        setTimeout(() => {
            try { offMap.remove(); } catch(e) {}
            try { offContainer.remove(); } catch(e) {}
            resolve();
        }, 15000);
    });

    // Left/right borders on map area
    ctx.strokeStyle = '#000';
    ctx.lineWidth   = 2;
    ctx.setLineDash([]);
    ctx.beginPath();
    ctx.moveTo(1,         MAP_Y); ctx.lineTo(1,         MAP_Y + MAP_H);
    ctx.moveTo(BASE_W-1,  MAP_Y); ctx.lineTo(BASE_W-1,  MAP_Y + MAP_H);
    ctx.stroke();

    // ── 3. Attribution strip ────────────────────────────────────────────────
    const attrText  = '© OpenStreetMap contributors © MapTiler';
    const attrSize  = Math.round(BASE_H * 8 / 555);
    ctx.font        = attrSize + 'px ' + SANS;
    ctx.textAlign   = 'right';
    ctx.textBaseline = 'bottom';
    const attrMetrics = ctx.measureText(attrText);
    const attrPadX  = 5, attrPadY = 4;
    const attrBgW   = attrMetrics.width + attrPadX * 2;
    const attrBgH   = attrSize + attrPadY * 2;
    const attrBgX   = BASE_W - attrBgW - 2;
    const attrBgY   = MAP_Y + MAP_H - attrBgH - 2;
    ctx.fillStyle   = 'rgba(255,255,255,0.82)';
    ctx.fillRect(attrBgX, attrBgY, attrBgW, attrBgH);
    ctx.fillStyle   = '#444';
    ctx.fillText(attrText, BASE_W - 2 - attrPadX, MAP_Y + MAP_H - 2 - attrPadY);

    // ── 4. White footer ─────────────────────────────────────────────────────
    const footY = MAP_Y + MAP_H;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, footY, BASE_W, FOOT_H);

    // border around footer
    ctx.strokeStyle = '#000';
    ctx.lineWidth   = 2;
    ctx.strokeRect(1, footY + 1, BASE_W - 2, FOOT_H - 2);

    // "Paste map" hint
    const hintSize = Math.round(BASE_H * 14 / 555);
    ctx.fillStyle  = '#000';
    ctx.font       = 'bold ' + hintSize + 'px ' + FONT;
    ctx.textAlign  = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText('(Paste map above or draw in territory)', BASE_W / 2, footY + Math.round(FOOT_H * 0.06));

    // ── FIXED JUSTIFIED TEXT LOGIC ──────────────────────────────────────────
    // Hard-coded split to ensure "territory" starts the second line.
    
    const bodySize = Math.round(BASE_H * 20 / 555);
    ctx.font       = 'bold ' + bodySize + 'px ' + FONT;
    ctx.textBaseline = 'top';
    
    // Define the specific words for each line
    const lines = [
        ['Please', 'keep', 'this', 'card', 'in', 'the', 'envelope.', 'Do', 'not', 'soil,', 'mark,', 'or', 'bend', 'it.', 'Each', 'time', 'the'],
        ['territory', 'is', 'covered,', 'please', 'inform', 'the', 'brother', 'who', 'cares', 'for', 'the', 'territory', 'files.']
    ];

    const maxW    = BASE_W - padX * 2;
    const lineH   = Math.round(bodySize * 1.25);
    const textTop = footY + Math.round(FOOT_H * 0.28);

    lines.forEach((lineWords, i) => {
        const lineY = textTop + i * lineH;

        // Measure total width of words (without spaces)
        let totalWordsW = 0;
        lineWords.forEach(w => totalWordsW += ctx.measureText(w).width);

        // Calculate total empty space available for gaps
        const emptySpace = maxW - totalWordsW;

        // Calculate gap size. 
        // If it's a single word (unlikely here), gap is 0.
        // Otherwise, divide space by number of gaps (words - 1).
        let gap = 0;
        if (lineWords.length > 1) {
            gap = emptySpace / (lineWords.length - 1);
        }

        // Draw the words
        let cursor = padX;
        lineWords.forEach((w) => {
            ctx.textAlign = 'left';
            ctx.fillText(w, cursor, lineY);
            cursor += ctx.measureText(w).width + gap;
        });
    });
    // ────────────────────────────────────────────────────────────────────────

    // Form code bottom-left
    const codeSize = Math.round(BASE_H * 12 / 555);
    ctx.font       = codeSize + 'px ' + FONT;
    ctx.textAlign  = 'left';
    ctx.textBaseline = 'bottom';
    ctx.fillText('S-12-E   6/72', padX, footY + FOOT_H - Math.round(FOOT_H * 0.08));

    return out;
}

// ─── IMPROVED UI: Info Line Ribbon ──────────────────────────────────────────
function showOverlay(msg) {
    let ov = document.getElementById('printOverlay');
    if (!ov) {
        ov = document.createElement('div');
        ov.id = 'printOverlay';
        document.body.appendChild(ov);
    }
    
    // Positioned directly under the 52px toolbar, full width
    // Replaces the KMZ filename badge space temporarily
ov.style.cssText = `
    position: fixed !important;
    bottom: 0 !important;
    top: auto !important;
    left: 0;
    right: 0;
    height: 40px;
    background: #fff;
    border-top: 1px solid #ccc;
    display: flex;
    align-items: center;
    justify-content: center;
    font-family: sans-serif;
    font-size: 12px;
    color: #333;
    z-index: 999999;
    box-shadow: 0 -4px 6px -2px rgba(0,0,0,0.1);
    transition: opacity 0.3s;
`;

    // Hide the KMZ filename badge during printing
    const badge = document.getElementById('kmzFilenameBadge');
    if (badge) badge.style.display = 'none';

    // Static layout to prevent jumping
    ov.innerHTML = `
        <div style="margin-right:20px; font-weight:bold; color:#1a6fc4;">Generating PDF</div>
        <div style="width:250px; height:8px; background:#eee; border-radius:4px; overflow:hidden; margin-right:20px;">
            <div id="progressFill" style="width:0%; height:100%; background:#1a6fc4; transition:width 0.2s;"></div>
        </div>
        <div id="printStatus" style="font-variant-numeric: tabular-nums; width:140px;">Initializing...</div>
    `;
    
    if (msg) document.getElementById('printStatus').textContent = msg;
}

function setProgress(frac, msg) {
    const statusEl = document.getElementById('printStatus');
    const fillEl   = document.getElementById('progressFill');
    if (statusEl && msg) statusEl.textContent = msg;
    if (fillEl) fillEl.style.width = (frac * 100) + '%';
}

function hideOverlay() {
    const ov = document.getElementById('printOverlay');
    
    // Restore the KMZ filename badge
    const badge = document.getElementById('kmzFilenameBadge');
    if (badge && badge.textContent) badge.style.display = 'flex';

    if (ov) {
        ov.style.opacity = '0';
        setTimeout(() => { ov.style.display = 'none'; }, 350);
    }
}

async function buildPDF(features, filename, congName) {
    var jsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: [CARD_W_MM, CARD_H_MM] });

    // Save current screen state so we can restore it after the batch loop
    const savedFeature  = activeFeature;
    const savedLocality = document.getElementById('card-locality') ? document.getElementById('card-locality').textContent : '';
    const savedTerrNo   = document.getElementById('card-terrno')   ? document.getElementById('card-terrno').textContent   : '';

    for (let i = 0; i < features.length; i++) {
        const f    = features[i];
        const info = cleanNameAndSplit(f.properties.name);

        setProgress(i / features.length, `Card ${i + 1} of ${features.length}`);

        // ── Live map display (visual feedback for the worker) ──────────────────
        // showTerritory updates the live map BEFORE we await the offscreen render,
        // so the two MapLibre instances never touch the same source simultaneously.
        const displayLocality = (congName && congName.length)
            ? info.locality + ' [' + congName + ']'
            : info.locality;
        updateCardFields(displayLocality, info.number);
        showTerritory(f);   // draw on screen — purely cosmetic, PDF is offscreen
        // Brief yield so the browser paints the map before the offscreen render blocks
        await new Promise(r => requestAnimationFrame(r));

        const canvas  = await renderCardToCanvas(f);
        const imgData = canvas.toDataURL('image/png');

        if (i > 0) doc.addPage([CARD_W_MM, CARD_H_MM], 'landscape');
        doc.addImage(imgData, 'PNG', 0, 0, CARD_W_MM, CARD_H_MM);

        // Cooldown every 20 pages to avoid memory pressure
        if (i > 0 && i % 20 === 0) {
            setProgress(i / features.length, 'Cooling down…');
            await new Promise(r => setTimeout(r, 1000));
        }
    }

    setProgress(1, 'Saving PDF…');
    doc.save(filename);

    // Restore the map to whatever was showing before the batch started
    if (savedFeature) {
        // Re-highlight the correct sidebar item
        document.querySelectorAll('.side-item').forEach(d => d.classList.remove('active'));
        const savedDiv = document.querySelector(`.side-item[data-name="${savedFeature.properties.name}"]`);
        if (savedDiv) savedDiv.classList.add('active');
        // Redraw the map and restore card fields directly — avoids selectTerritory
        // overwriting the saved locality/terrNo with the plain unstyled name
        showTerritory(savedFeature);
        updateCardFields(savedLocality, savedTerrNo);
    }

    // Always finish with a full world view to signal the batch is complete
    // and clear any territory-specific layers
    ['poly-outline','mask-layer','terr-line','terr-point','terr-point-ring']
        .forEach(id => { try { if (map.getLayer(id)) map.removeLayer(id); } catch(e){} });
    try { if (map.getSource('active-data')) map.removeSource('active-data'); } catch(e){}
    map.jumpTo({ center: [0, 0], zoom: 1, bearing: 0 });
    activeFeature = null;
    updateCardFields(' ', ' ');
    document.querySelectorAll('.side-item').forEach(d => d.classList.remove('active'));
    // Reset toggle label now no territory is active
    const tBtn = document.getElementById('togglePanelBtn');
    if (tBtn) { const l = tBtn.querySelector('.btn-label'); if (l) l.textContent = 'See Maps'; }
    hideOverlay();
}

// ─── 8. CONGREGATION NAME DIALOGUE ────────────────────────────────────────────
// Resolves with a string (the congregation name) or null (user chose No).
function askCongregationName() {
    return new Promise(resolve => {

        const overlay = document.createElement('div');
        overlay.style.cssText = [
            'position:fixed;inset:0;z-index:10000;',
            'background:rgba(0,0,0,0.6);',
            'display:flex;align-items:center;justify-content:center;'
        ].join('');

        function makeBox(innerHTML) {
            const box = document.createElement('div');
            box.style.cssText = [
                'background:#fff;border-radius:10px;padding:32px 28px 24px;',
                'width:340px;max-width:90vw;',
                'box-shadow:0 8px 32px rgba(0,0,0,0.35);',
                'font-family:sans-serif;text-align:center;'
            ].join('');
            box.innerHTML = innerHTML;
            overlay.innerHTML = '';
            overlay.appendChild(box);
        }

        function btnStyle(primary) {
            return primary
                ? 'padding:10px 28px;background:#1a6fc4;color:#fff;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;'
                : 'padding:10px 28px;background:#e0e0e0;color:#333;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;';
        }

        function showStep1() {
            makeBox(
                '<div style="font-size:17px;font-weight:600;margin-bottom:10px;">Add congregation name?</div>' +
                '<div style="font-size:13px;color:#555;margin-bottom:24px;">Do you wish to add the congregation name to the cards?</div>' +
                '<div style="display:flex;gap:12px;justify-content:center;">' +
                  '<button id="congYesBtn" style="' + btnStyle(true)  + '">Yes</button>' +
                  '<button id="congNoBtn"  style="' + btnStyle(false) + '">No</button>'  +
                '</div>'
            );
            overlay.querySelector('#congNoBtn').onclick  = function() { overlay.remove(); resolve(null); };
            overlay.querySelector('#congYesBtn').onclick = showStep2;
        }

        function showStep2() {
            var keys      = Object.keys(loadedFiles);
            var raw       = keys.length ? keys[0] : '';
            var suggested = raw.replace(/\.(kmz|kml)$/i, '').replace(/[_-]/g, ' ').trim();

            makeBox(
                '<div style="font-size:17px;font-weight:600;margin-bottom:10px;">Congregation name</div>' +
                '<div style="font-size:13px;color:#555;margin-bottom:16px;">' +
                  'This will appear on every card as:<br>' +
                  '<em style="color:#1a6fc4;">Territory Name [Congregation Name]</em>' +
                '</div>' +
                '<input id="congNameInput" type="text"' +
                  ' style="width:100%;box-sizing:border-box;padding:9px 12px;font-size:14px;border:1px solid #ccc;border-radius:6px;margin-bottom:20px;"' +
                  ' placeholder="e.g. Westside Congregation" />' +
                '<div style="display:flex;gap:12px;justify-content:center;">' +
                  '<button id="congOkBtn"   style="' + btnStyle(true)  + '">OK</button>'   +
                  '<button id="congBackBtn" style="' + btnStyle(false) + '">Back</button>' +
                '</div>'
            );

            var input = overlay.querySelector('#congNameInput');
            input.value = suggested;
            setTimeout(function() { input.focus(); input.select(); }, 50);
            input.addEventListener('keydown', function(ev) {
                if (ev.key === 'Enter') overlay.querySelector('#congOkBtn').click();
            });
            overlay.querySelector('#congBackBtn').onclick = showStep1;
            overlay.querySelector('#congOkBtn').onclick   = function() {
                var name = input.value.trim();
                overlay.remove();
                resolve(name);
            };
        }

        showStep1();
        document.body.appendChild(overlay);
    });
}

// ─── 9. PRINT ACTIONS ─────────────────────────────────────────────────────────
async function startBatchPDF() {
    var keys = Object.keys(loadedFiles);
    if (!keys.length) return;

    // Show congregation name dialogue FIRST — before any other checks
    var congName = await askCongregationName();

    var lib = window.jspdf || window.jsPDF;
    if (!lib) return alert('jsPDF library missing.');
    var baseName = (congName && congName.length) ? congName : keys[0].replace(/\.(kmz|kml)$/i, '');
    showOverlay('Starting batch print…');
    await buildPDF(loadedFiles[keys[0]].features, baseName + '_Territory_Cards.pdf', congName);
}

async function printSelectedCards() {
    var keys = Object.keys(loadedFiles);
    if (!keys.length || !checkedNames.size) return;

    // Show congregation name dialogue FIRST — before any other checks
    var congName = await askCongregationName();

    var lib = window.jspdf || window.jsPDF;
    if (!lib) return alert('jsPDF library missing.');
    var baseName = (congName && congName.length) ? congName : keys[0].replace(/\.(kmz|kml)$/i, '');
    var features = loadedFiles[keys[0]].features.filter(function(f) { return checkedNames.has(f.properties.name); });
    showOverlay('Printing ' + features.length + ' selected card(s)…');
    await buildPDF(features, baseName + '_Territory_Cards_Selected.pdf', congName);
}

async function printActiveCard() {
    if (!activeFeature) return;
    var congName = await askCongregationName();
    var lib = window.jspdf || window.jsPDF;
    if (!lib) return alert('jsPDF library missing.');
    var keys     = Object.keys(loadedFiles);
    var baseName = (congName && congName.length) ? congName : (keys.length ? keys[0].replace(/\.(kmz|kml)$/i, '') : 'territory');
    var info     = cleanNameAndSplit(activeFeature.properties.name);
    var safeName = (info.number ? info.number + '_' : '') + (info.locality || 'card').replace(/[^a-z0-9_\-]/gi, '_');
    showOverlay('Printing card…');
    await buildPDF([activeFeature], baseName + '_' + safeName + '.pdf', congName);
}

// ─── INIT ────────────────────────────────────────────────────────────────────
// Wait for CSS layout before first build
requestAnimationFrame(() => buildCardFrame());

// On resize, wait for reflow then rebuild and re-fit
window.addEventListener('resize', () => {
    requestAnimationFrame(() => {
        buildCardFrame();
        if (activeFeature) showTerritory(activeFeature);
    });
});

// ─── 9b. COLOURISED FILM EFFECT ──────────────────────────────────────────────
// Pixels OUTSIDE the polygon → greyscale.
// Pixels INSIDE the polygon  → full colour.
// Applied as a canvas overlay on screen, and composited into the PDF canvas.

function applyColorizeEffect(targetMap, feature) {
    if (!feature) return;
    const gtype = feature.geometry.type;
    if (gtype !== 'Polygon' && gtype !== 'MultiPolygon') return;

    const old = document.getElementById('colorize-overlay');
    if (old) old.remove();

    const mapCanvas = targetMap.getCanvas();
    const W = mapCanvas.width;
    const H = mapCanvas.height;

    const overlay = document.createElement('canvas');
    overlay.id = 'colorize-overlay';
    overlay.width  = W;
    overlay.height = H;
    overlay.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;pointer-events:none;z-index:10;';
    const ctx = overlay.getContext('2d');

    // Read map pixels via drawImage (works with preserveDrawingBuffer)
    const readCanvas = document.createElement('canvas');
    readCanvas.width = W; readCanvas.height = H;
    const readCtx = readCanvas.getContext('2d');
    readCtx.drawImage(mapCanvas, 0, 0);
    const imageData = readCtx.getImageData(0, 0, W, H);
    const data = imageData.data;

    // Scale from CSS pixels to device pixels
    const container = targetMap.getContainer();
    const scaleX = W / container.offsetWidth;
    const scaleY = H / container.offsetHeight;

    // Build polygon hit-test canvas
    const hitCanvas = document.createElement('canvas');
    hitCanvas.width = W; hitCanvas.height = H;
    const hitCtx = hitCanvas.getContext('2d');
    hitCtx.fillStyle = '#fff';

    const rings = gtype === 'Polygon'
        ? feature.geometry.coordinates
        : feature.geometry.coordinates.reduce((a, p) => a.concat(p), []);

    rings.forEach(ring => {
        hitCtx.beginPath();
        ring.forEach((coord, i) => {
            const pt = targetMap.project([coord[0], coord[1]]);
            i === 0 ? hitCtx.moveTo(pt.x * scaleX, pt.y * scaleY)
                    : hitCtx.lineTo(pt.x * scaleX, pt.y * scaleY);
        });
        hitCtx.closePath();
        hitCtx.fill('evenodd');
    });

    const hitData = hitCtx.getImageData(0, 0, W, H).data;

    for (let i = 0; i < data.length; i += 4) {
        if (hitData[i] < 128) {  // outside polygon → greyscale
            const grey = Math.round(0.299 * data[i] + 0.587 * data[i+1] + 0.114 * data[i+2]);
            data[i] = data[i+1] = data[i+2] = grey;
        }
    }

    ctx.putImageData(imageData, 0, 0);

    const cardWindow = document.getElementById('card-window');
    if (cardWindow) cardWindow.appendChild(overlay);
}

// PDF version — operates directly on the output canvas region
function applyColorizeOnCanvas(ctx, offMap, feature, dX, dY, dW, dH) {
    if (!feature) return;
    const gtype = feature.geometry.type;
    if (gtype !== 'Polygon' && gtype !== 'MultiPolygon') return;

    const imageData = ctx.getImageData(dX, dY, dW, dH);
    const data = imageData.data;

    const hitCanvas = document.createElement('canvas');
    hitCanvas.width = dW; hitCanvas.height = dH;
    const hitCtx = hitCanvas.getContext('2d');
    hitCtx.fillStyle = '#fff';

    const rings = gtype === 'Polygon'
        ? feature.geometry.coordinates
        : feature.geometry.coordinates.reduce((a, p) => a.concat(p), []);

    rings.forEach(ring => {
        hitCtx.beginPath();
        ring.forEach((coord, i) => {
            const pt = offMap.project([coord[0], coord[1]]);
            i === 0 ? hitCtx.moveTo(pt.x, pt.y) : hitCtx.lineTo(pt.x, pt.y);
        });
        hitCtx.closePath();
        hitCtx.fill('evenodd');
    });

    const hitData = hitCtx.getImageData(0, 0, dW, dH).data;

    for (let i = 0; i < data.length; i += 4) {
        if (hitData[i] < 128) {
            const grey = Math.round(0.299 * data[i] + 0.587 * data[i+1] + 0.114 * data[i+2]);
            data[i] = data[i+1] = data[i+2] = grey;
        }
    }

    ctx.putImageData(imageData, dX, dY);
}

// ─── 10. KINGDOM HALL STICKER ────────────────────────────────────────────────
//
// APPROACH: The sticker lives at a real lat/lng chosen by the user.
// It appears on any territory card whose rendered map bounds contain that point.
// Position is projected to pixel coords each time a map is shown — no dragging
// on the card, no feature-name matching.
//
// FLOW:
//   1. On page load (after map ready) → ask "Place KH sticker?" with instructions
//   2. Yes → map goes interactive, user pans/zooms to find the KH, clicks/taps
//            to drop a crosshair → lat/lng stored → map goes non-interactive again
//            → 🔄 Restart button revealed
//   3. No  → nothing changes, app works exactly as before
//   4. When any territory is shown → if stored lat/lng is within the rendered
//            bounds, project it to card-window pixels, show sticker div there
//   5. 🔄 Restart → prints the current card with sticker composited at the
//            projected pixel position on the offscreen PDF canvas

window._khLngLat    = null;  // { lng, lat } — set when user clicks on the map
window._khPlacing   = false; // true while in placement mode

// ── 1. Startup prompt ────────────────────────────────────────────────────────
// Called once the map has loaded and the UI is ready.
function askKhStickerPlacement() {
    const overlay = document.createElement('div');
    overlay.id = 'kh-prompt-overlay';
    overlay.style.cssText = [
        'position:fixed;inset:0;background:rgba(0,0,0,0.6);',
        'z-index:9000;display:flex;align-items:center;justify-content:center;',
        'font-family:sans-serif;'
    ].join('');

    overlay.innerHTML = `
        <div style="background:#fff;border-radius:12px;padding:28px 32px;max-width:380px;
                    width:90vw;box-shadow:0 8px 32px rgba(0,0,0,0.35);text-align:center;">
            <div style="font-size:32px;margin-bottom:10px;">🏛️</div>
            <div style="font-size:17px;font-weight:700;color:#1a2236;margin-bottom:12px;">
                Place Kingdom Hall Sticker?
            </div>
            <div style="font-size:13px;color:#444;line-height:1.6;margin-bottom:22px;text-align:left;"
			
                <strong>YES</strong> — the map becomes interactive. Pan and zoom to your
                Kingdom Hall, then <b>tap or click its exact location</b> to pin
                the JW.ORG sticker there.<br><br>
                The sticker will automatically appear on every territory card
                whose map includes that location — just use
                <b>Print All Maps</b>.<br><br>
                <b>NO</b> — everything works exactly as normal.<br><br>
			
			<strong>NOTE:</strong> If your Kingdom Hall is placed on source data as "Kingdom Hall" 
(<a href="https://www.openstreetmap.org" target="_blank" rel="noopener">OpenStreetMap</a>) 
AND is registered as a Christian place of worship a JW.ORG sticker will automatically replace 
the cross that would otherwise be seen beside it.<br><br>
            </div>
            <div style="display:flex;gap:12px;justify-content:center;">
                <button id="khPromptNo"  style="flex:1;height:40px;border-radius:8px;border:1.5px solid #ccc;
                        background:#f5f5f5;font-size:14px;font-weight:600;cursor:pointer;">
                    No thanks
                </button>
                <button id="khPromptYes" style="flex:1;height:40px;border-radius:8px;border:none;
                        background:#1a6fc4;color:#fff;font-size:14px;font-weight:600;cursor:pointer;">
                    Yes, place it
                </button>
            </div>
        </div>`;

    document.body.appendChild(overlay);

    document.getElementById('khPromptNo').onclick = () => {
        overlay.remove();
    };

    document.getElementById('khPromptYes').onclick = () => {
        overlay.remove();
        enterKhPlacementMode();
    };
}

// ── 2. Placement mode ────────────────────────────────────────────────────────
function enterKhPlacementMode() {
    window._khPlacing = true;

    // Make map interactive so user can pan/zoom
    map.dragPan.enable();
    map.scrollZoom.enable();
    map.touchZoomRotate.enable();
    map.doubleClickZoom.enable();

    // Zoom to a useful starting level if still at world view
    if (map.getZoom() < 5) {
        map.jumpTo({ zoom: 14 });
    }

    // Show placement banner
    const banner = document.createElement('div');
    banner.id = 'kh-place-banner';
    banner.style.cssText = [
        'position:fixed;top:60px;left:50%;transform:translateX(-50%);',
        'background:#003580;color:#fff;',
        'padding:10px 20px;border-radius:8px;',
        'font-family:sans-serif;font-size:13px;font-weight:600;',
        'z-index:8000;text-align:center;',
        'box-shadow:0 4px 16px rgba(0,0,0,0.4);',
        'pointer-events:none;'
    ].join('');
    banner.textContent = '📍 Pan & zoom to your Kingdom Hall — tap or click to place the sticker';
    document.body.appendChild(banner);

    // Change cursor to crosshair over the map
    document.getElementById('map').style.cursor = 'crosshair';

    // Single click/tap handler
    function onMapClick(e) {
        window._khLngLat  = { lng: e.lngLat.lng, lat: e.lngLat.lat };
        window._khPlacing = false;

        // Restore non-interactive
        map.dragPan.disable();
        map.scrollZoom.disable();
        map.touchZoomRotate.disable();
        map.doubleClickZoom.disable();
        document.getElementById('map').style.cursor = '';

        // Remove banner
        const b = document.getElementById('kh-place-banner');
        if (b) b.remove();

        // Remove the click listener
        map.off('click', onMapClick);

        // Show sticker immediately at the clicked location on the live map
        updateKhStickerOnScreen();

        // Confirm
        showKhConfirmToast();
    }

    map.on('click', onMapClick);
}

// ── 3. Toast confirmation ─────────────────────────────────────────────────────
function showKhConfirmToast() {
    const toast = document.createElement('div');
    toast.style.cssText = [
        'position:fixed;top:70px;left:50%;transform:translateX(-50%);',
        'background:#1e8c3a;color:#fff;',
        'padding:12px 24px;border-radius:8px;',
        'font-family:sans-serif;font-size:13px;font-weight:600;',
        'z-index:8000;box-shadow:0 4px 16px rgba(0,0,0,0.35);',
        'pointer-events:none;opacity:1;transition:opacity 0.5s;',
        'text-align:center;max-width:90vw;'
    ].join('');
    toast.textContent = '✅ KH sticker placed! Browse to any territory — use 🔄 Restart to print it with the sticker.';
    document.body.appendChild(toast);
    setTimeout(() => { toast.style.opacity = '0'; }, 3000);
    setTimeout(() => { toast.remove(); }, 3600);
}

// ── 4. Show/hide sticker on screen when territory changes ────────────────────
// Called by showTerritory (hooked below). Projects _khLngLat to card-window
// pixel coordinates. If outside the current map bounds, removes the sticker.
function updateKhStickerOnScreen() {
    removeKhSticker();
    if (!window._khLngLat) return;

    const cardWindow = document.getElementById('card-window');
    if (!cardWindow) return;

    // Project the stored lng/lat to pixel coords in the map container
    const pt = map.project([window._khLngLat.lng, window._khLngLat.lat]);

    // map container = card-window (they're the same element via #map inset:0)
    const cw = cardWindow.offsetWidth;
    const ch = cardWindow.offsetHeight;

    // If placing (no territory shown yet), still show sticker if point is visible
    // If a territory is shown, hide sticker when KH is outside its bounds
    if (pt.x < 0 || pt.x > cw || pt.y < 0 || pt.y > ch) return;

    const xFrac = pt.x / cw;
    const yFrac = pt.y / ch;

    // Store fracs for PDF composite
    window._khStickerPos = { xFrac, yFrac };

    const sticker = document.createElement('div');
    sticker.id = 'kh-sticker';
    sticker.style.cssText = [
        'position:absolute;',
        'left:' + (xFrac * 100) + '%;',
        'top:'  + (yFrac * 100) + '%;',
        'transform:translate(-50%,-50%);',
        'z-index:50;pointer-events:none;user-select:none;',
        'background:#003580;color:#fff;',
        'font-family:sans-serif;font-weight:bold;font-size:13px;line-height:1.2;',
        'text-align:center;padding:5px 10px;border-radius:4px;',
        'box-shadow:0 2px 8px rgba(0,0,0,0.45);'
    ].join('');
    sticker.innerHTML = 'JW.<br>ORG';
    cardWindow.appendChild(sticker);
}

function removeKhSticker() {
    const old = document.getElementById('kh-sticker');
    if (old) old.remove();
    window._khStickerPos = null;
}

// ── 5. PDF composite ─────────────────────────────────────────────────────────
// Called from renderCardToCanvas after the offscreen map image is drawn.
// Projects _khLngLat through the offscreen map to canvas pixel coords.
function compositeKhStickerOnCanvas(ctx, offMap) {
    if (!window._khLngLat) return;

    // Derive print dimensions from globals (these match the local consts in renderCardToCanvas)
    const _BASE_W = PRINT_W_PX;
    const _BASE_H = PRINT_H_PX;
    const _HEAD_H = Math.round(_BASE_H * (110 / 555));
    const _FOOT_H = Math.round(_BASE_H * (110 / 555));
    const _MAP_Y  = _HEAD_H;
    const _MAP_H  = _BASE_H - _HEAD_H - _FOOT_H;

    const pt = offMap.project([window._khLngLat.lng, window._khLngLat.lat]);

    // offMap container is PRINT_W_PX × MAP_H at print resolution
    if (pt.x < 0 || pt.x > _BASE_W || pt.y < 0 || pt.y > _MAP_H) return;

    // Canvas coords — offset by _MAP_Y because the map area starts below the header
    const cx = pt.x;
    const cy = _MAP_Y + pt.y;

    const sW = Math.round(_BASE_W * 0.07);
    const sH = Math.round(sW * 0.7);
    const sx = cx - sW / 2;
    const sy = cy - sH / 2;
    const r  = Math.round(sH * 0.12);

    // Rounded rect background
    ctx.fillStyle = '#003580';
    ctx.beginPath();
    ctx.moveTo(sx + r, sy);
    ctx.lineTo(sx + sW - r, sy);
    ctx.quadraticCurveTo(sx + sW, sy,       sx + sW, sy + r);
    ctx.lineTo(sx + sW, sy + sH - r);
    ctx.quadraticCurveTo(sx + sW, sy + sH,  sx + sW - r, sy + sH);
    ctx.lineTo(sx + r,  sy + sH);
    ctx.quadraticCurveTo(sx, sy + sH,        sx, sy + sH - r);
    ctx.lineTo(sx, sy + r);
    ctx.quadraticCurveTo(sx, sy,             sx + r, sy);
    ctx.closePath();
    ctx.save();
    ctx.shadowColor = 'rgba(0,0,0,0.45)';
    ctx.shadowBlur  = Math.round(sH * 0.15);
    ctx.shadowOffsetY = Math.round(sH * 0.06);
    ctx.fill();
    ctx.restore();

    const fontSize = Math.round(sH * 0.36);
    ctx.fillStyle    = '#ffffff';
    ctx.font         = 'bold ' + fontSize + 'px sans-serif';
    ctx.textAlign    = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('JW.',  cx, cy - fontSize * 0.55);
    ctx.fillText('ORG',  cx, cy + fontSize * 0.55);
}

// khRestart removed — sticker appears automatically on any card that contains the KH location

// ── 7. Hook into showTerritory & renderCardToCanvas ──────────────────────────
// ── showTerritory and renderCardToCanvas hooks are patched inline above ──

// ─── 11. WELCOME DIALOGUE ─────────────────────────────────────────────────────
// Multi-step onboarding shown once on map load.
// Final step triggers the file picker so the flow is seamless.

function showWelcomeDialogue() {

    const overlay = document.createElement('div');
    overlay.id = 'welcome-overlay';
    overlay.style.cssText = [
        'position:fixed;inset:0;z-index:10500;',
        'background:rgba(10,18,40,0.82);',
        'display:flex;align-items:center;justify-content:center;',
        'font-family:sans-serif;padding:12px;'
    ].join('');
    document.body.appendChild(overlay);

    const STEPS = [
        {
            icon: '🗺️',
            title: 'Territory Card Printer',
            body: `This tool turns your congregation's KMZ territory file into
                   print-ready <b>6″ × 4″ territory cards</b> — one per territory,
                   laid out and sized exactly to the S-12 card standard.<br><br>
                   Each card shows the territory map, the locality name, territory
                   number.<br><br>
                   All cards are exported as a single PDF ready to send straight
                   to your printer.`,
            back: null,
            next: 'Next →'
        },
        {
            icon: '📂',
            title: 'What is a KMZ file?',
            body: `A <b>KMZ file</b> is the standard format used by
                   Google Earth to store boundaries.<br><br>
                   Your congregation's territory servant may have this file — or maybe you just created it.
                   It should contain all your territories as named polygons or points drawn on a map.<br><br>
                   KML files are also supported.`,
            back: '← Back',
            next: 'Next →'
        },
        {
            icon: '🔄',
            title: 'Auto-rotation',
            body: `Each territory is automatically <b>rotated to best fill the card</b>.`,
            back: '← Back',
            next: 'Next →'
        },
        {
            icon: '🖨️',
            title: 'Printing on A4',
            body: `The cards are sized at exactly <b>6″ × 4″ (152 × 102 mm)</b>
                   at 300 DPI — the S-12 standard.<br><br>
                   When your PDF opens, set your printer to:<br>
                   <div style="margin:10px 0 0 10px;line-height:2;">
                     • Paper size: <b>A4</b><br>
                     • Scaling: <b>Actual size</b> (not "fit to page")<br>
                     • Orientation: <b>Portrait</b><br>
                   </div><br>
                   Two cards fit neatly on one A4 sheet.`,
            back: '← Back',
            next: 'Choose KMZ →'
        }
    ];

    let step = 0;

    function render() {
        const s = STEPS[step];
        const isLast = step === STEPS.length - 1;

        // Dot indicators
        const dots = STEPS.map((_, i) =>
            `<div style="width:8px;height:8px;border-radius:50%;margin:0 3px;
                         background:${i === step ? '#1a6fc4' : '#ccc'};
                         transition:background 0.2s;"></div>`
        ).join('');

        overlay.innerHTML = `
            <div style="background:#fff;border-radius:14px;
                        padding:30px 28px 24px;
                        width:100%;max-width:420px;max-height:90dvh;
                        overflow-y:auto;
                        box-shadow:0 12px 40px rgba(0,0,0,0.45);">

                <!-- Progress dots -->
                <div style="display:flex;justify-content:center;margin-bottom:18px;">
                    ${dots}
                </div>

                <!-- Icon -->
                <div style="font-size:36px;text-align:center;margin-bottom:10px;">
                    ${s.icon}
                </div>

                <!-- Title -->
                <div style="font-size:18px;font-weight:700;color:#1a2236;
                            text-align:center;margin-bottom:14px;">
                    ${s.title}
                </div>

                <!-- Body -->
                <div style="font-size:13px;color:#444;line-height:1.7;
                            margin-bottom:26px;">
                    ${s.body}
                </div>

                <!-- Buttons -->
                <div style="display:flex;gap:10px;">
                    ${s.back
                        ? `<button id="wlc-back"
                               style="flex:1;height:42px;border-radius:8px;
                                      border:1.5px solid #ccc;background:#f5f5f5;
                                      font-size:14px;font-weight:600;cursor:pointer;
                                      color:#333;">
                               ${s.back}
                           </button>`
                        : '<div style="flex:1"></div>'
                    }
                    <button id="wlc-next"
                            style="flex:2;height:42px;border-radius:8px;border:none;
                                   background:#1a6fc4;color:#fff;
                                   font-size:14px;font-weight:600;cursor:pointer;">
                        ${s.next}
                    </button>
                </div>

                <!-- Skip -->
                <div style="text-align:center;margin-top:14px;">
                    <span id="wlc-skip"
                          style="font-size:12px;color:#aaa;cursor:pointer;
                                 text-decoration:underline;">
                        Skip introduction
                    </span>
                </div>
            </div>`;

        overlay.querySelector('#wlc-next').onclick = () => {
            if (isLast) {
                overlay.remove();
                document.getElementById('fileInput').click();
            } else {
                step++;
                render();
            }
        };

        const backBtn = overlay.querySelector('#wlc-back');
        if (backBtn) backBtn.onclick = () => { step--; render(); };

        overlay.querySelector('#wlc-skip').onclick = () => overlay.remove();
    }

    render();
}
