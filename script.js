// --- CONFIGURATION ---
// Card is exactly 6" × 4" = 152.4 mm × 101.6 mm
const CARD_W_MM = 152.4;
const CARD_H_MM = 101.6;

// Print at 300 DPI for crisp text
const PRINT_DPI  = 300;
const PRINT_W_PX = Math.round((CARD_W_MM / 25.4) * PRINT_DPI); // 1800
const PRINT_H_PX = Math.round((CARD_H_MM / 25.4) * PRINT_DPI); // 1200

const MAP_PAD_FRAC = 0.10;
const POINT_ZOOM = 17;

// ── CITY TILE FILES ───────────────────────────────────────────────────────────
const CITY_TILES = {
    // Coast
    mombasa:   { label: 'Mombasa',    file: 'mombasa.pmtiles',   centre: [39.67, -4.05],  zoom: 13 },
    kilifi:    { label: 'Kilifi',     file: 'kilifi.pmtiles',    centre: [39.97, -3.51],  zoom: 13 },
    malindi:   { label: 'Malindi',    file: 'malindi.pmtiles',   centre: [40.12, -3.22],  zoom: 13 },
    watamu:    { label: 'Watamu',     file: 'watamu.pmtiles',    centre: [40.02, -3.35],  zoom: 14 },
    voi:       { label: 'Voi',        file: 'voi.pmtiles',       centre: [38.57, -3.40],  zoom: 13 },
    // Nairobi Region
    nairobi:   { label: 'Nairobi',    file: 'nairobi.pmtiles',   centre: [36.82, -1.29],  zoom: 12 },
    thika:     { label: 'Thika',      file: 'thika.pmtiles',     centre: [37.09, -1.03],  zoom: 13 },
    kiambu:    { label: 'Kiambu',     file: 'kiambu.pmtiles',    centre: [36.84, -1.17],  zoom: 13 },
    ruiru:     { label: 'Ruiru',      file: 'ruiru.pmtiles',     centre: [36.96, -1.14],  zoom: 13 },
    limuru:    { label: 'Limuru',     file: 'limuru.pmtiles',    centre: [36.64, -1.11],  zoom: 13 },
    machakos:  { label: 'Machakos',   file: 'machakos.pmtiles',  centre: [37.26, -1.52],  zoom: 13 },
    athiriver: { label: 'Athi River', file: 'athiriver.pmtiles', centre: [36.98, -1.45],  zoom: 13 },
    // Central
    nyeri:     { label: 'Nyeri',      file: 'nyeri.pmtiles',     centre: [36.95, -0.42],  zoom: 13 },
    muranga:   { label: 'Muranga',    file: 'muranga.pmtiles',   centre: [37.15, -0.72],  zoom: 13 },
    nanyuki:   { label: 'Nanyuki',    file: 'nanyuki.pmtiles',   centre: [37.07,  0.01],  zoom: 13 },
    embu:      { label: 'Embu',       file: 'embu.pmtiles',      centre: [37.45, -0.53],  zoom: 13 },
    kerugoya:  { label: 'Kerugoya',   file: 'kerugoya.pmtiles',  centre: [37.28, -0.50],  zoom: 13 },
    // Rift Valley
    nakuru:    { label: 'Nakuru',     file: 'nakuru.pmtiles',    centre: [36.07, -0.30],  zoom: 13 },
    naivasha:  { label: 'Naivasha',   file: 'naivasha.pmtiles',  centre: [36.43, -0.72],  zoom: 13 },
    eldoret:   { label: 'Eldoret',    file: 'eldoret.pmtiles',   centre: [35.27,  0.52],  zoom: 13 },
    kericho:   { label: 'Kericho',    file: 'kericho.pmtiles',   centre: [35.28, -0.37],  zoom: 13 },
    kitale:    { label: 'Kitale',     file: 'kitale.pmtiles',    centre: [34.99,  1.02],  zoom: 13 },
    narok:     { label: 'Narok',      file: 'narok.pmtiles',     centre: [35.87, -1.08],  zoom: 13 },
    bomet:     { label: 'Bomet',      file: 'bomet.pmtiles',     centre: [35.34, -0.78],  zoom: 13 },
    nyahururu: { label: 'Nyahururu',  file: 'nyahururu.pmtiles', centre: [36.36,  0.03],  zoom: 13 },
    // Western
    kisumu:    { label: 'Kisumu',     file: 'kisumu.pmtiles',    centre: [34.75,  0.10],  zoom: 13 },
    kakamega:  { label: 'Kakamega',   file: 'kakamega.pmtiles',  centre: [34.75,  0.28],  zoom: 13 },
    bungoma:   { label: 'Bungoma',    file: 'bungoma.pmtiles',   centre: [34.56,  0.57],  zoom: 13 },
    busia:     { label: 'Busia',      file: 'busia.pmtiles',     centre: [34.11,  0.46],  zoom: 13 },
    kisii:     { label: 'Kisii',      file: 'kisii.pmtiles',     centre: [34.76, -0.68],  zoom: 13 },
    homabay:   { label: 'Homa Bay',   file: 'homabay.pmtiles',   centre: [34.46, -0.52],  zoom: 13 },
    migori:    { label: 'Migori',     file: 'migori.pmtiles',    centre: [34.47, -1.06],  zoom: 13 },
    siaya:     { label: 'Siaya',      file: 'siaya.pmtiles',     centre: [34.29,  0.06],  zoom: 13 },
    // Eastern
    meru:      { label: 'Meru',       file: 'meru.pmtiles',      centre: [37.65,  0.05],  zoom: 13 },
    kitui:     { label: 'Kitui',      file: 'kitui.pmtiles',     centre: [38.01, -1.37],  zoom: 13 },
    chuka:     { label: 'Chuka',      file: 'chuka.pmtiles',     centre: [37.65, -0.33],  zoom: 13 },
};

const BASE_URL = 'https://msa-territory.github.io/territory/';
let currentCity = 'mombasa';

function getPmtilesUrl(cityKey) {
    return BASE_URL + CITY_TILES[cityKey].file;
}

// ── PROTOMAPS SETUP ───────────────────────────────────────────────────────────
const _pmtilesProtocol = new pmtiles.Protocol();
maplibregl.addProtocol('pmtiles', _pmtilesProtocol.tile);

function buildProtomapsStyle(flavour, cityKey) {
    if (!cityKey) cityKey = currentCity;
    return {
        version: 8,
        glyphs: 'https://protomaps.github.io/basemaps-assets/fonts/{fontstack}/{range}.pbf',
        sprite: 'https://protomaps.github.io/basemaps-assets/sprites/v4/' + flavour,
        sources: {
            protomaps: {
                type: 'vector',
                url: 'pmtiles://' + getPmtilesUrl(cityKey),
                attribution: '<a href="https://protomaps.com">Protomaps</a> © <a href="https://openstreetmap.org">OpenStreetMap</a>'
            }
        },
        layers: basemaps.layers('protomaps', basemaps.namedFlavor(flavour), { lang: 'en' })
    };
}

let loadedFiles      = {};
let activeFeature    = null;
let checkedNames     = new Set();
let congregationName = '';

// ─── 1. MAPLIBRE SETUP ──────────────────────────────────────────────────────
const initialFlavour = document.getElementById('styleSelect').value;
const cityData = CITY_TILES[currentCity];

const map = new maplibregl.Map({
    container: 'map',
    style: buildProtomapsStyle(initialFlavour),
    center: cityData.centre,
    zoom: cityData.zoom,
    preserveDrawingBuffer: true,
    attributionControl: false,
    fadeDuration: 0,
    interactive: false,
});

map.on('load', () => {
    placeTheKh();
    showWelcomeDialogue();
});

map.on('styleimagemissing', (e) => {
    if (!map.hasImage(e.id)) {
        const blank = new Uint8Array(4);
        map.addImage(e.id, { width: 1, height: 1, data: blank });
    }
});

// ── Style selector ────────────────────────────────────────────────────────────
document.getElementById('styleSelect').addEventListener('change', e => {
    const flavour = e.target.value;
    ['poly-outline','mask-layer','terr-line','terr-point','terr-point-ring']
        .forEach(id => { try { if (map.getLayer(id)) map.removeLayer(id); } catch(e){} });
    try { if (map.getSource('active-data')) map.removeSource('active-data'); } catch(e){}
    map.setStyle(buildProtomapsStyle(flavour));
    map.once('idle', () => {
        placeTheKh();
        if (activeFeature) showTerritory(activeFeature);
    });
});

// ── City selector ─────────────────────────────────────────────────────────────
document.getElementById('citySelect').addEventListener('change', e => {
    currentCity = e.target.value;
    const city = CITY_TILES[currentCity];
    const flavour = document.getElementById('styleSelect').value;

    // Clear active territory layers
    ['poly-outline','mask-layer','terr-line','terr-point','terr-point-ring']
        .forEach(id => { try { if (map.getLayer(id)) map.removeLayer(id); } catch(e){} });
    try { if (map.getSource('active-data')) map.removeSource('active-data'); } catch(e){}

    // Load new city tiles
    map.setStyle(buildProtomapsStyle(flavour, currentCity));
    map.once('idle', () => {
        map.jumpTo({ center: city.centre, zoom: city.zoom });
        placeTheKh();
    });
});

// ─── 2. CARD FRAME ────────────────────────────────────────────────────────────
function buildCardFrame() {
    const frame = document.getElementById('territory-card-frame');
    if (!frame) return;

    const W     = frame.offsetWidth;
    const H     = frame.offsetHeight;
    const scale = W / 874;
    const headH = Math.round(110 * scale);
    const footH = Math.round(110 * scale);
    const fs    = n => Math.round(n * scale) + 'px';
    const pd    = n => Math.round(n * scale) + 'px';

    ['card-top', 'card-bottom', 'card-attr'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.remove();
    });

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

    const attr = document.createElement('div');
    attr.id = 'card-attr';
    attr.style.cssText = [
        'position:absolute;bottom:2px;right:1px;z-index:20;',
        'pointer-events:none;background:rgba(255,255,255,0.82);',
        'padding:1px 5px;font-family:sans-serif;',
        'font-size:' + fs(7) + ';line-height:1.4;color:#444;'
    ].join('');
    attr.textContent = '© OpenStreetMap contributors © Protomaps';

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

    const cardWindow = document.getElementById('card-window');
    frame.insertBefore(header, cardWindow);
    cardWindow.appendChild(attr);
    frame.appendChild(footer);

    if (typeof map !== 'undefined') map.resize();

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

// ─── 3. SHOW TERRITORY ────────────────────────────────────────────────────────
function addTerritoryLayers(targetMap, feature, lineWidth) {
    if (!lineWidth) lineWidth = 4;
    const gtype = feature.geometry.type;

    ['terr-point', 'terr-point-ring', 'terr-line', 'poly-outline', 'mask-layer']
        .forEach(id => { try { if (targetMap.getLayer(id)) targetMap.removeLayer(id); } catch(e){} });
    try { if (targetMap.getSource('active-data')) targetMap.removeSource('active-data'); } catch(e){}

    const isPolygon  = gtype === 'Polygon'    || gtype === 'MultiPolygon';
    const isLine     = gtype === 'LineString' || gtype === 'MultiLineString';
    const isPoint    = gtype === 'Point'      || gtype === 'MultiPoint';

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
        targetMap.addLayer({
            id: 'terr-line', type: 'line', source: 'active-data',
            layout: { 'line-join': 'round', 'line-cap': 'round' },
            paint: { 'line-color': '#e06c00', 'line-width': lineWidth, 'line-dasharray': [3, 2] }
        });
    } else if (isPoint) {
        targetMap.addLayer({
            id: 'terr-point-ring', type: 'circle', source: 'active-data',
            paint: {
                'circle-radius':       lineWidth * 4,
                'circle-color':        '#ffffff',
                'circle-stroke-width': lineWidth,
                'circle-stroke-color': '#1a6fc4'
            }
        });
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
        const coords = gtype === 'Point'
            ? feature.geometry.coordinates
            : feature.geometry.coordinates[0];
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
    requestAnimationFrame(() => updateKhStickerOnScreen());
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
        const overhang = Math.sin(2 * rad);
        const ratio = rotW / rotH;
        const score = -Math.abs(ratio - cardAspect) - (overhang * 0.3);
        if (score > bestScore) {
            bestScore   = score;
            bestBearing = angle;
        }
    }
    return bestBearing;
}

// ─── 4. KH ICON ─────────────────────────────────────────────────────────────
function placeTheKh(targetMap, iconSize) {
    if (iconSize === undefined) iconSize = 0.7;
    if (!targetMap) targetMap = map;
    const style = targetMap.getStyle();
    if (!style) return;

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

    if (!targetMap.hasImage('transparent-icon')) {
        const tCanvas = document.createElement('canvas');
        tCanvas.width = 1; tCanvas.height = 1;
        const tCtx = tCanvas.getContext('2d');
        targetMap.addImage('transparent-icon', { width: 1, height: 1, data: tCtx.getImageData(0,0,1,1).data });
    }

    style.layers.forEach(layer => {
        if (layer['source-layer'] !== 'poi') return;
        if (layer.layout && layer.layout['icon-image']) {
            const currentIcon = targetMap.getLayoutProperty(layer.id, 'icon-image');
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

    const badge = document.getElementById('kmzFilenameBadge');
    if (badge) {
        badge.textContent = '📁 ' + file.name;
        badge.style.display = 'flex';
    }

    document.getElementById('printAllBtn').disabled = false;
    const tBtn = document.getElementById('togglePanelBtn');
    if (tBtn) { const l = tBtn.querySelector('.btn-label'); if (l) l.textContent = 'See Maps'; }

    const allBbox = turf.bbox({ type: 'FeatureCollection', features: geojson.features });
    map.fitBounds(allBbox, { padding: 40, animate: true, duration: 2200 });
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
    const oneBtn = document.getElementById('printOneBtn');
    if (oneBtn) oneBtn.style.display = 'flex';
    const panel = document.getElementById('sidePanel');
    const toggleBtn = document.getElementById('togglePanelBtn');
    if (panel) panel.classList.remove('open');
    if (toggleBtn) {
        toggleBtn.classList.remove('active');
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
async function renderCardToCanvas(feature) {
    const out  = document.createElement('canvas');
    out.width  = PRINT_W_PX;
    out.height = PRINT_H_PX;
    const ctx  = out.getContext('2d');

    const BASE_W  = PRINT_W_PX;
    const BASE_H  = PRINT_H_PX;
    const HEAD_H  = Math.round(BASE_H * (110 / 555));
    const FOOT_H  = Math.round(BASE_H * (110 / 555));
    const MAP_Y   = HEAD_H;
    const MAP_H   = BASE_H - HEAD_H - FOOT_H;

    const FONT    = "'Times New Roman', Times, serif";
    const SANS    = "sans-serif";

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, BASE_W, HEAD_H);
    ctx.strokeStyle = '#000';
    ctx.lineWidth   = 2;
    ctx.strokeRect(1, 1, BASE_W - 2, HEAD_H - 1);

    ctx.fillStyle  = '#000';
    ctx.font       = 'bold ' + Math.round(BASE_H * 32 / 555) + 'px ' + FONT;
    ctx.textAlign  = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Territory Map Card', BASE_W / 2, Math.round(HEAD_H * 0.36));

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

    ctx.setLineDash([3, 4]);
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(padX + lblW, dotY);
    ctx.lineTo(terrLblX - 8, dotY);
    ctx.stroke();
    ctx.setLineDash([]);

    const elL = document.getElementById('card-locality');
    const localityText = (elL && elL.textContent.trim() !== ' ') ? elL.textContent.trim() : '';
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

    ctx.font      = 'bold ' + lblSize + 'px ' + FONT;
    ctx.textAlign = 'left';
    ctx.fillText(terrLbl, terrLblX, rowY);

    ctx.setLineDash([3, 4]);
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(terrLblX + terrLblW + 6, dotY);
    ctx.lineTo(BASE_W - padX, dotY);
    ctx.stroke();
    ctx.setLineDash([]);

    const elT = document.getElementById('card-terrno');
    const terrNoText = (elT && elT.textContent.trim() !== ' ') ? elT.textContent.trim() : '';
    ctx.font      = fldSize + 'px ' + FONT;
    ctx.textAlign = 'center';
    ctx.fillText(terrNoText, terrLblX + terrLblW + 6 + terrBoxW / 2, rowY);

    // ── Map area ─────────────────────────────────────────────────────────────
    const dX = 0, dY = MAP_Y, dW = BASE_W, dH = MAP_H;

    await new Promise((resolve) => {
        const offContainer = document.createElement('div');
        offContainer.style.cssText = [
            'position:fixed;left:-9999px;top:-9999px;',
            'width:' + dW + 'px;height:' + dH + 'px;',
            'pointer-events:none;opacity:0.001;'
        ].join('');
        document.body.appendChild(offContainer);

        const flavour = document.getElementById('styleSelect').value;
        const bbox    = turf.bbox(feature);
        const bearing = calculateBestRotation(feature);

        const offMap = new maplibregl.Map({
            container:             offContainer,
            style:                 buildProtomapsStyle(flavour, currentCity),
            preserveDrawingBuffer: true,
            attributionControl:    false,
            fadeDuration:          0,
            interactive:           false,
            pixelRatio:            2,
        });

        offMap.once('load', () => {
            offMap.on('styleimagemissing', (e) => {
                if (!offMap.hasImage(e.id)) {
                    const blank = new Uint8Array(4);
                    offMap.addImage(e.id, { width: 1, height: 1, data: blank });
                }
            });

            placeTheKh(offMap, 1.4);

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

            addTerritoryLayers(offMap, feature, 6);

            offMap.once('idle', () => {
                setTimeout(() => {
                    requestAnimationFrame(() => {
                        try {
                            ctx.drawImage(offMap.getCanvas(), dX, dY, dW, dH);
                        } catch(e) {
                            ctx.fillStyle = '#e0e0e0';
                            ctx.fillRect(dX, dY, dW, dH);
                        }
                        compositeKhStickerOnCanvas(ctx, offMap);
                        offMap.remove();
                        offContainer.remove();
                        resolve();
                    });
                }, 150);
            });
        });

        setTimeout(() => {
            try { offMap.remove(); } catch(e) {}
            try { offContainer.remove(); } catch(e) {}
            resolve();
        }, 15000);
    });

    ctx.strokeStyle = '#000';
    ctx.lineWidth   = 2;
    ctx.setLineDash([]);
    ctx.beginPath();
    ctx.moveTo(1,         MAP_Y); ctx.lineTo(1,         MAP_Y + MAP_H);
    ctx.moveTo(BASE_W-1,  MAP_Y); ctx.lineTo(BASE_W-1,  MAP_Y + MAP_H);
    ctx.stroke();

    const attrText  = '© OpenStreetMap contributors © Protomaps';
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

    const footY = MAP_Y + MAP_H;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, footY, BASE_W, FOOT_H);
    ctx.strokeStyle = '#000';
    ctx.lineWidth   = 2;
    ctx.strokeRect(1, footY + 1, BASE_W - 2, FOOT_H - 2);

    const hintSize = Math.round(BASE_H * 14 / 555);
    ctx.fillStyle  = '#000';
    ctx.font       = 'bold ' + hintSize + 'px ' + FONT;
    ctx.textAlign  = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText('(Paste map above or draw in territory)', BASE_W / 2, footY + Math.round(FOOT_H * 0.06));

    const bodySize = Math.round(BASE_H * 20 / 555);
    ctx.font       = 'bold ' + bodySize + 'px ' + FONT;
    ctx.textBaseline = 'top';

    const lines = [
        ['Please', 'keep', 'this', 'card', 'in', 'the', 'envelope.', 'Do', 'not', 'soil,', 'mark,', 'or', 'bend', 'it.', 'Each', 'time', 'the'],
        ['territory', 'is', 'covered,', 'please', 'inform', 'the', 'brother', 'who', 'cares', 'for', 'the', 'territory', 'files.']
    ];

    const maxW    = BASE_W - padX * 2;
    const lineH   = Math.round(bodySize * 1.25);
    const textTop = footY + Math.round(FOOT_H * 0.28);

    lines.forEach((lineWords, i) => {
        const lineY = textTop + i * lineH;
        let totalWordsW = 0;
        lineWords.forEach(w => totalWordsW += ctx.measureText(w).width);
        const emptySpace = maxW - totalWordsW;
        let gap = 0;
        if (lineWords.length > 1) gap = emptySpace / (lineWords.length - 1);
        let cursor = padX;
        lineWords.forEach((w) => {
            ctx.textAlign = 'left';
            ctx.fillText(w, cursor, lineY);
            cursor += ctx.measureText(w).width + gap;
        });
    });

    const codeSize = Math.round(BASE_H * 12 / 555);
    ctx.font       = codeSize + 'px ' + FONT;
    ctx.textAlign  = 'left';
    ctx.textBaseline = 'bottom';
    ctx.fillText('S-12-E   6/72', padX, footY + FOOT_H - Math.round(FOOT_H * 0.08));

    return out;
}

// ─── OVERLAY / PROGRESS ──────────────────────────────────────────────────────
function showOverlay(msg) {
    let ov = document.getElementById('printOverlay');
    if (!ov) {
        ov = document.createElement('div');
        ov.id = 'printOverlay';
        document.body.appendChild(ov);
    }
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
    const badge = document.getElementById('kmzFilenameBadge');
    if (badge) badge.style.display = 'none';
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

    const savedFeature  = activeFeature;
    const savedLocality = document.getElementById('card-locality') ? document.getElementById('card-locality').textContent : '';
    const savedTerrNo   = document.getElementById('card-terrno')   ? document.getElementById('card-terrno').textContent   : '';

    for (let i = 0; i < features.length; i++) {
        const f    = features[i];
        const info = cleanNameAndSplit(f.properties.name);
        setProgress(i / features.length, `Card ${i + 1} of ${features.length}`);
        const displayLocality = (congName && congName.length)
            ? info.locality + ' [' + congName + ']'
            : info.locality;
        updateCardFields(displayLocality, info.number);
        showTerritory(f);
        await new Promise(r => requestAnimationFrame(r));
        const canvas  = await renderCardToCanvas(f);
        const imgData = canvas.toDataURL('image/png');
        if (i > 0) doc.addPage([CARD_W_MM, CARD_H_MM], 'landscape');
        doc.addImage(imgData, 'PNG', 0, 0, CARD_W_MM, CARD_H_MM);
        if (i > 0 && i % 20 === 0) {
            setProgress(i / features.length, 'Cooling down…');
            await new Promise(r => setTimeout(r, 1000));
        }
    }

    setProgress(1, 'Saving PDF…');
    doc.save(filename);

    if (savedFeature) {
        document.querySelectorAll('.side-item').forEach(d => d.classList.remove('active'));
        const savedDiv = document.querySelector(`.side-item[data-name="${savedFeature.properties.name}"]`);
        if (savedDiv) savedDiv.classList.add('active');
        showTerritory(savedFeature);
        updateCardFields(savedLocality, savedTerrNo);
    }

    ['poly-outline','mask-layer','terr-line','terr-point','terr-point-ring']
        .forEach(id => { try { if (map.getLayer(id)) map.removeLayer(id); } catch(e){} });
    try { if (map.getSource('active-data')) map.removeSource('active-data'); } catch(e){}
    map.jumpTo({ center: CITY_TILES[currentCity].centre, zoom: CITY_TILES[currentCity].zoom, bearing: 0 });
    activeFeature = null;
    updateCardFields(' ', ' ');
    document.querySelectorAll('.side-item').forEach(d => d.classList.remove('active'));
    const tBtn = document.getElementById('togglePanelBtn');
    if (tBtn) { const l = tBtn.querySelector('.btn-label'); if (l) l.textContent = 'See Maps'; }
    hideOverlay();
}

// ─── 8. GROUP NAME DIALOGUE ──────────────────────────────────────────────────
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
                '<div style="font-size:17px;font-weight:600;margin-bottom:10px;">Add group name?</div>' +
                '<div style="font-size:13px;color:#555;margin-bottom:24px;">Do you wish to add a group name to the cards?</div>' +
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
                '<div style="font-size:17px;font-weight:600;margin-bottom:10px;">Group name</div>' +
                '<div style="font-size:13px;color:#555;margin-bottom:16px;">' +
                  'This will appear on every card as:<br>' +
                  '<em style="color:#1a6fc4;">Territory Name [Group Name]</em>' +
                '</div>' +
                '<input id="congNameInput" type="text"' +
                  ' style="width:100%;box-sizing:border-box;padding:9px 12px;font-size:14px;border:1px solid #ccc;border-radius:6px;margin-bottom:20px;"' +
                  ' placeholder="e.g. Westside" />' +
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
requestAnimationFrame(() => buildCardFrame());
window.addEventListener('resize', () => {
    requestAnimationFrame(() => {
        buildCardFrame();
        if (activeFeature) showTerritory(activeFeature);
    });
});

// ─── 10. KINGDOM HALL STICKER ────────────────────────────────────────────────
window._khLngLat    = null;
window._khPlacing   = false;

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
            <div style="font-size:13px;color:#444;line-height:1.6;margin-bottom:22px;text-align:left;">
                <strong>YES</strong> — the map becomes interactive. Pan and zoom to your
                Kingdom Hall, then <b>tap or click its exact location</b> to pin
                the JW.ORG sticker there.<br><br>
                The sticker will automatically appear on every territory card
                whose map includes that location.<br><br>
                <b>NO</b> — everything works exactly as normal.<br><br>
                <strong>NOTE:</strong> If your Kingdom Hall is listed on
                <a href="https://www.openstreetmap.org" target="_blank" rel="noopener">OpenStreetMap</a>
                as a Christian place of worship, a JW.ORG sticker will automatically appear beside it.
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
    document.getElementById('khPromptNo').onclick  = () => overlay.remove();
    document.getElementById('khPromptYes').onclick = () => { overlay.remove(); enterKhPlacementMode(); };
}

function enterKhPlacementMode() {
    window._khPlacing = true;
    map.dragPan.enable();
    map.scrollZoom.enable();
    map.touchZoomRotate.enable();
    map.doubleClickZoom.enable();
    if (map.getZoom() < 5) map.jumpTo({ zoom: 14 });

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
    document.getElementById('map').style.cursor = 'crosshair';

    function onMapClick(e) {
        window._khLngLat  = { lng: e.lngLat.lng, lat: e.lngLat.lat };
        window._khPlacing = false;
        map.dragPan.disable();
        map.scrollZoom.disable();
        map.touchZoomRotate.disable();
        map.doubleClickZoom.disable();
        document.getElementById('map').style.cursor = '';
        const b = document.getElementById('kh-place-banner');
        if (b) b.remove();
        map.off('click', onMapClick);
        updateKhStickerOnScreen();
        showKhConfirmToast();
    }
    map.on('click', onMapClick);
}

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
    toast.textContent = '✅ KH sticker placed! Browse to any territory to see it.';
    document.body.appendChild(toast);
    setTimeout(() => { toast.style.opacity = '0'; }, 3000);
    setTimeout(() => { toast.remove(); }, 3600);
}

function updateKhStickerOnScreen() {
    removeKhSticker();
    if (!window._khLngLat) return;
    const cardWindow = document.getElementById('card-window');
    if (!cardWindow) return;
    const pt = map.project([window._khLngLat.lng, window._khLngLat.lat]);
    const cw = cardWindow.offsetWidth;
    const ch = cardWindow.offsetHeight;
    if (pt.x < 0 || pt.x > cw || pt.y < 0 || pt.y > ch) return;
    const xFrac = pt.x / cw;
    const yFrac = pt.y / ch;
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

function compositeKhStickerOnCanvas(ctx, offMap) {
    if (!window._khLngLat) return;
    const _BASE_W = PRINT_W_PX;
    const _BASE_H = PRINT_H_PX;
    const _HEAD_H = Math.round(_BASE_H * (110 / 555));
    const _FOOT_H = Math.round(_BASE_H * (110 / 555));
    const _MAP_Y  = _HEAD_H;
    const _MAP_H  = _BASE_H - _HEAD_H - _FOOT_H;
    const pt = offMap.project([window._khLngLat.lng, window._khLngLat.lat]);
    if (pt.x < 0 || pt.x > _BASE_W || pt.y < 0 || pt.y > _MAP_H) return;
    const cx = pt.x;
    const cy = _MAP_Y + pt.y;
    const sW = Math.round(_BASE_W * 0.07);
    const sH = Math.round(sW * 0.7);
    const sx = cx - sW / 2;
    const sy = cy - sH / 2;
    const r  = Math.round(sH * 0.12);
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

// ─── 11. WELCOME DIALOGUE ─────────────────────────────────────────────────────
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
                   Each card shows the territory map, the locality name, and territory number.<br><br>
                   All cards are exported as a single PDF ready to send straight to your printer.`,
            back: null,
            next: 'Next →'
        },
        {
            icon: '📂',
            title: 'What is a KMZ file?',
            body: `A <b>KMZ file</b> is the standard format used by Google Earth to store boundaries.<br><br>
                   Your territory servant may have this file. It should contain all your territories
                   as named polygons or points drawn on a map.<br><br>
                   KML files are also supported.`,
            back: '← Back',
            next: 'Next →'
        },
        {
            icon: '📍',
            title: 'Select your city first',
            body: `Use the <b>City</b> dropdown in the toolbar to select the city where your
                   territories are located.<br><br>
                   This loads the correct map tiles for that area. You can change city at any time.`,
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
                <div style="display:flex;justify-content:center;margin-bottom:18px;">${dots}</div>
                <div style="font-size:36px;text-align:center;margin-bottom:10px;">${s.icon}</div>
                <div style="font-size:18px;font-weight:700;color:#1a2236;text-align:center;margin-bottom:14px;">${s.title}</div>
                <div style="font-size:13px;color:#444;line-height:1.7;margin-bottom:26px;">${s.body}</div>
                <div style="display:flex;gap:10px;">
                    ${s.back
                        ? `<button id="wlc-back" style="flex:1;height:42px;border-radius:8px;border:1.5px solid #ccc;background:#f5f5f5;font-size:14px;font-weight:600;cursor:pointer;color:#333;">${s.back}</button>`
                        : '<div style="flex:1"></div>'
                    }
                    <button id="wlc-next" style="flex:2;height:42px;border-radius:8px;border:none;background:#1a6fc4;color:#fff;font-size:14px;font-weight:600;cursor:pointer;">${s.next}</button>
                </div>
                <div style="text-align:center;margin-top:14px;">
                    <span id="wlc-skip" style="font-size:12px;color:#aaa;cursor:pointer;text-decoration:underline;">Skip introduction</span>
                </div>
            </div>`;

        overlay.querySelector('#wlc-next').onclick = () => {
            if (isLast) { overlay.remove(); document.getElementById('fileInput').click(); }
            else { step++; render(); }
        };
        const backBtn = overlay.querySelector('#wlc-back');
        if (backBtn) backBtn.onclick = () => { step--; render(); };
        overlay.querySelector('#wlc-skip').onclick = () => overlay.remove();
    }

    render();
}
