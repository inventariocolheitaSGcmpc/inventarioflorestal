const farmFilter = document.getElementById("farmFilter");
const clearSelectionBtn = document.getElementById("clearSelectionBtn");
const exportBtn = document.getElementById("exportBtn");
const emailBtn = document.getElementById("emailBtn");
const selectionSummary = document.getElementById("selectionSummary");
const sequenceTableBody = document.querySelector("#sequenceTable tbody");
const sequenceCount = document.getElementById("sequenceCount");
const farmCount = document.getElementById("farmCount");
const plotCount = document.getElementById("plotCount");
const generatedAt = document.getElementById("generatedAt");
const emailModal = document.getElementById("emailModal");
const emailModalBackdrop = document.getElementById("emailModalBackdrop");
const closeEmailModalBtn = document.getElementById("closeEmailModalBtn");
const openOutlookBtn = document.getElementById("openOutlookBtn");
const openGmailBtn = document.getElementById("openGmailBtn");
const openMailAppBtn = document.getElementById("openMailAppBtn");
const copyEmailTextBtn = document.getElementById("copyEmailTextBtn");
const emailBodyPreview = document.getElementById("emailBodyPreview");
const downloadEmlBtn = document.getElementById("downloadEmlBtn");

const state = {
  allFeatures: [],
  currentFeatures: [],
  featureLayers: new Map(),
  selectedIds: [],
  pendingEmailDraft: null,
  pendingAttachments: null
};

const map = L.map("map", {
  minZoom: 7,
  maxZoom: 18,
  zoomControl: true
}).setView([-30.3, -54.0], 9);

L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  attribution: "&copy; OpenStreetMap contributors",
  crossOrigin: true
}).addTo(map);

L.tileLayer(
  "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
  {
    attribution: "Tiles &copy; Esri",
    crossOrigin: true
  }
).addTo(map);

const mapLayers = {
  polygons: L.layerGroup().addTo(map),
  arrows: L.layerGroup().addTo(map),
  farmLabels: L.layerGroup().addTo(map),
  sequenceLabels: L.layerGroup().addTo(map)
};

const legend = L.control({ position: "bottomright" });
legend.onAdd = function onAdd() {
  this._div = L.DomUtil.create("div", "info legend");
  this.update();
  return this._div;
};
legend.update = function update(min = 0, max = 0) {
  this._div.innerHTML = `
    <div style="padding:12px 14px;background:rgba(255,255,255,.92);border-radius:16px;border:1px solid rgba(0,0,0,.08);box-shadow:0 24px 60px rgba(34,53,29,.12);font:12px Manrope,sans-serif;">
      <strong style="display:block;margin-bottom:8px;">VCSC (m3/ha)</strong>
      <div style="display:flex;align-items:center;gap:8px;">
        <span>${formatNumber(min)}</span>
        <div style="width:140px;height:12px;border-radius:999px;background:linear-gradient(90deg,#d73027,#fee08b,#1a9850);"></div>
        <span>${formatNumber(max)}</span>
      </div>
    </div>
  `;
};
legend.addTo(map);

Promise.all([
  fetch("site-data/base_info.geojson").then((r) => r.json()),
  fetch("site-data/summary.json").then((r) => r.json())
]).then(([geojson, summary]) => {
  state.allFeatures = geojson.features.map((feature) => ({
    ...feature,
    properties: normalizeProperties(feature.properties)
  }));

  farmCount.textContent = summary.total_fazendas ?? "-";
  plotCount.textContent = summary.total_talhoes ?? "-";
  generatedAt.textContent = summary.generated_at ?? "-";

  populateFarmFilter();
  applyFilter();
}).catch((error) => {
  selectionSummary.innerHTML = `<p class="empty-state">Falha ao carregar os dados do site: ${error.message}</p>`;
});

farmFilter.addEventListener("change", () => {
  applyFilter();
});

clearSelectionBtn.addEventListener("click", () => {
  state.selectedIds = [];
  renderMap();
  syncSelectionUI();
});

exportBtn.addEventListener("click", () => {
  const rows = selectedRows();
  if (!rows.length) {
    alert("Selecione pelo menos um talhao.");
    return;
  }
  downloadXlsx(rows, currentFarmName());
});

emailBtn.addEventListener("click", async () => {
  const rows = selectedRows();
  if (!rows.length) {
    alert("Selecione pelo menos um talhao.");
    return;
  }

  const excelAttachment = createXlsxAttachment(rows, currentFarmName());
  downloadBlob(excelAttachment.blob, excelAttachment.filename);
  await wait(250);
  const screenshotAttachment = await createMapScreenshotAttachment(currentFarmName());
  if (screenshotAttachment?.blob) {
    downloadBlob(screenshotAttachment.blob, screenshotAttachment.filename);
  }
  await wait(250);

  const screenshotName = screenshotAttachment?.filename || `HF_${safeFarmName(currentFarmName())}_Mapa_Sequencia_Colheita.png`;
  state.pendingAttachments = {
    excel: excelAttachment,
    screenshot: screenshotAttachment
  };
  state.pendingEmailDraft = buildEmailDraft(rows, currentFarmName(), screenshotName);
  openEmailModal(state.pendingEmailDraft);
});

closeEmailModalBtn.addEventListener("click", closeEmailModal);
emailModalBackdrop.addEventListener("click", closeEmailModal);

openOutlookBtn.addEventListener("click", () => {
  if (!state.pendingEmailDraft) return;
  const draft = state.pendingEmailDraft;
  const url = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(draft.toComma)}&subject=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;
  window.open(url, "_blank", "noopener,noreferrer");
});

openGmailBtn.addEventListener("click", () => {
  if (!state.pendingEmailDraft) return;
  const draft = state.pendingEmailDraft;
  const url = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(draft.toComma)}&su=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;
  window.open(url, "_blank", "noopener,noreferrer");
});

openMailAppBtn.addEventListener("click", () => {
  if (!state.pendingEmailDraft) return;
  const draft = state.pendingEmailDraft;
  window.location.href = `mailto:${draft.toSemicolon}?subject=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;
});

downloadEmlBtn.addEventListener("click", async () => {
  if (!state.pendingEmailDraft || !state.pendingAttachments?.excel) {
    return;
  }
  try {
    const emlBlob = await buildEmlBlob(state.pendingEmailDraft, state.pendingAttachments);
    const safeFarm = safeFarmName(currentFarmName());
    downloadBlob(emlBlob, `HF_${safeFarm}_Rascunho_Email.eml`);
  } catch (error) {
    console.warn("Nao foi possivel gerar o arquivo .eml.", error);
  }
});

copyEmailTextBtn.addEventListener("click", async () => {
  if (!state.pendingEmailDraft) return;
  try {
    await navigator.clipboard.writeText(state.pendingEmailDraft.fullText);
    copyEmailTextBtn.textContent = "Texto copiado";
    window.setTimeout(() => {
      copyEmailTextBtn.textContent = "Copiar texto do e-mail";
    }, 1800);
  } catch (error) {
    console.warn("Nao foi possivel copiar o texto.", error);
  }
});

function normalizeProperties(properties) {
  return {
    id_projeto: properties.id_projeto,
    cd_talhao: String(properties.cd_talhao),
    projeto: properties.projeto,
    produtividade: Number(properties.produtividade ?? 0),
    ab: Number(properties.ab ?? 0),
    vlr_area_gis: Number(properties.vlr_area_gis ?? 0),
    idade_inteira: Number(properties.idade_inteira ?? 0),
    ht: Number(properties.ht ?? 0),
    dap: Number(properties.dap ?? 0),
    n: Number(properties.n ?? 0),
    mg: Number(properties.mg ?? 0),
    pro_tlh: properties.pro_tlh,
    vmi: Number(properties.vmi ?? 0),
    vcsc: Number(properties.vcsc ?? 0),
    ima: Number(properties.ima ?? 0)
  };
}

function populateFarmFilter() {
  const farms = [...new Set(state.allFeatures.map((f) => f.properties.projeto))].sort();
  farmFilter.innerHTML = ["Todas", ...farms]
    .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    .join("");
}

function applyFilter() {
  const farm = farmFilter.value || "Todas";
  state.currentFeatures = farm === "Todas"
    ? [...state.allFeatures]
    : state.allFeatures.filter((feature) => feature.properties.projeto === farm);

  state.selectedIds = state.selectedIds.filter((id) =>
    state.currentFeatures.some((feature) => feature.properties.pro_tlh === id)
  );

  renderMap();
  syncSelectionUI();
}

function renderMap() {
  mapLayers.polygons.clearLayers();
  mapLayers.farmLabels.clearLayers();
  mapLayers.sequenceLabels.clearLayers();
  mapLayers.arrows.clearLayers();
  state.featureLayers.clear();

  if (!state.currentFeatures.length) {
    return;
  }

  const values = state.currentFeatures.map((feature) => feature.properties.produtividade);
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  legend.update(minValue, maxValue);

  const geoJsonLayer = L.geoJSON(
    {
      type: "FeatureCollection",
      features: state.currentFeatures
    },
    {
      style: (feature) => ({
        color: "#ffffff",
        weight: 1,
        fillOpacity: 0.72,
        fillColor: getColor(feature.properties.produtividade, minValue, maxValue)
      }),
      onEachFeature: (feature, layer) => {
        const id = feature.properties.pro_tlh;
        state.featureLayers.set(id, layer);
        layer.bindTooltip(
          `<strong>Talhao:</strong> ${escapeHtml(feature.properties.cd_talhao)}`,
          { className: "custom-tooltip" }
        );
        layer.on("click", () => toggleSelection(id));
        layer.on("mouseover", () => {
          layer.setStyle({ weight: 3, color: "#ffff00", fillOpacity: 0.9 });
        });
        layer.on("mouseout", () => {
          if (!state.selectedIds.includes(id)) {
            geoJsonLayer.resetStyle(layer);
          }
        });
      }
    }
  );

  geoJsonLayer.addTo(mapLayers.polygons);

  if (farmFilter.value && farmFilter.value !== "Todas") {
    map.fitBounds(geoJsonLayer.getBounds(), { padding: [20, 20] });

    state.currentFeatures.forEach((feature) => {
      const center = turfCentroid(feature);
      if (!center) return;
      L.marker([center[1], center[0]], {
        interactive: false,
        icon: L.divIcon({
          className: "farm-label",
          html: `<div style="color:#d9534f;font-weight:800;font-size:11px;text-shadow:1px 1px 2px white;">${escapeHtml(feature.properties.cd_talhao)}</div>`
        })
      }).addTo(mapLayers.farmLabels);
    });
  } else {
    map.setView([-30.3, -54.0], 9);
  }
}

function toggleSelection(id) {
  if (state.selectedIds.includes(id)) {
    state.selectedIds = state.selectedIds.filter((value) => value !== id);
  } else {
    state.selectedIds.push(id);
  }
  syncSelectionUI();
}

function syncSelectionUI() {
  updateHighlightLayers();
  renderSelectionSummary();
  renderTable();
}

function updateHighlightLayers() {
  mapLayers.arrows.clearLayers();
  mapLayers.sequenceLabels.clearLayers();

  state.currentFeatures.forEach((feature) => {
    const id = feature.properties.pro_tlh;
    const layer = state.featureLayers.get(id);
    if (layer) {
      layer.setStyle({
        color: state.selectedIds.includes(id) ? "#1f1f1f" : "#ffffff",
        weight: state.selectedIds.includes(id) ? 4 : 1,
        fillOpacity: state.selectedIds.includes(id) ? 0.9 : 0.72
      });
    }
  });

  const selectedFeatures = selectedFeatureObjects();
  if (!selectedFeatures.length) {
    sequenceCount.textContent = "0 itens";
    return;
  }

  sequenceCount.textContent = `${selectedFeatures.length} ${selectedFeatures.length === 1 ? "item" : "itens"}`;

  const points = selectedFeatures
    .map((feature) => turfCentroid(feature))
    .filter(Boolean);

  selectedFeatures.forEach((feature, index) => {
    const center = turfCentroid(feature);
    if (!center) return;
    L.marker([center[1], center[0]], {
      interactive: false,
      icon: L.divIcon({
        className: "sequence-index",
        html: `<div class="sequence-label">${index + 1}</div>`
      })
    }).addTo(mapLayers.sequenceLabels);
  });

  if (points.length > 1) {
    L.polyline(points.map(([lng, lat]) => [lat, lng]), {
      color: "#ffff00",
      weight: 5,
      dashArray: "8, 12",
      opacity: 1
    }).addTo(mapLayers.arrows);
  }
}

function renderSelectionSummary() {
  const rows = selectedRows();
  if (!rows.length) {
    selectionSummary.innerHTML = `<p class="empty-state">Nenhum talhao selecionado.</p>`;
    return;
  }

  const totalArea = rows.reduce((acc, row) => acc + row.area, 0);
  const totalVcsc = rows.reduce((acc, row) => acc + row.vcsc, 0);

  selectionSummary.innerHTML = `
    <div class="summary-grid">
      <div class="summary-item">
        <span>Talhoes selecionados</span>
        <strong>${rows.length}</strong>
      </div>
      <div class="summary-item">
        <span>Area total</span>
        <strong>${formatNumber(totalArea)} ha</strong>
      </div>
      <div class="summary-item">
        <span>Volume Comercial (VCSC)</span>
        <strong>${formatNumber(totalVcsc)} m3</strong>
      </div>
    </div>
  `;
}

function renderTable() {
  const rows = selectedRows();
  sequenceTableBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${row.ordem}</td>
      <td>${escapeHtml(row.fazendaTalhao)}</td>
      <td>${formatNumber(row.area)}</td>
      <td>${formatNumber(row.vcsc)}</td>
      <td>${formatNumber(row.vmi)}</td>
    </tr>
  `).join("");
}

function selectedFeatureObjects() {
  return state.selectedIds
    .map((id) => state.currentFeatures.find((feature) => feature.properties.pro_tlh === id))
    .filter(Boolean);
}

function selectedRows() {
  return selectedFeatureObjects().map((feature, index) => ({
    ordem: index + 1,
    fazendaTalhao: feature.properties.pro_tlh,
    area: feature.properties.vlr_area_gis,
    vcsc: feature.properties.vcsc,
    vmi: feature.properties.vmi
  }));
}

function currentFarmName() {
  return farmFilter.value || "Geral";
}

function createXlsxAttachment(rows, farmName) {
  const worksheetRows = rows.map((row) => ({
    "Ordem sequencia": row.ordem,
    "Fazenda/Talhao": row.fazendaTalhao,
    "Area (ha)": row.area,
    "VCSC (m3)": row.vcsc,
    "VMI (m3)": row.vmi
  }));

  const safeFarm = safeFarmName(farmName);
  const filename = `HF_${safeFarm}_Sequencia_Talhonar.xlsx`;
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(worksheetRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sequencia");
  const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob(
    [buffer],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
  );
  return { filename, blob };
}

async function createMapScreenshotAttachment(farmName) {
  const safeFarm = safeFarmName(farmName);
  const filename = `HF_${safeFarm}_Mapa_Sequencia_Colheita.png`;

  try {
    const canvas = await html2canvas(document.getElementById("map"), {
      useCORS: true,
      backgroundColor: null,
      scale: window.devicePixelRatio > 1 ? 2 : 1,
      onclone: (clonedDoc) => normalizeLeafletClone(clonedDoc)
    });
    const blob = await new Promise((resolve) => {
      canvas.toBlob((blob) => {
        if (!blob) {
          resolve(null);
          return;
        }
        resolve(blob);
      }, "image/png");
    });
    if (!blob) {
      return { filename, blob: null };
    }
    return { filename, blob };
  } catch (error) {
    console.warn("Nao foi possivel gerar a imagem do mapa.", error);
  }

  return { filename, blob: null };
}

function safeFarmName(farmName) {
  return farmName === "Todas" ? "Geral" : farmName.replace(/[^a-zA-Z0-9_-]/g, "_");
}

function wait(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function downloadBlob(blob, filename) {
  if (!blob) {
    return;
  }
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function buildEmailDraft(rows, farmName, screenshotName) {
  const recipients = [
    "pietro.duran@cmpc.com",
    "matheus.roberto@cmpc.com"
  ];
  const safeFarm = safeFarmName(farmName);
  const excelName = `HF_${safeFarm}_Sequencia_Talhonar.xlsx`;
  const totalArea = rows.reduce((acc, row) => acc + row.area, 0);
  const totalVcsc = rows.reduce((acc, row) => acc + row.vcsc, 0);
  const orderLines = rows
    .map((row) => `${row.ordem}. ${row.fazendaTalhao} | Area: ${formatNumber(row.area)} ha | VCSC: ${formatNumber(row.vcsc)} m3 | VMI: ${formatNumber(row.vmi)} m3`)
    .join("\n");

  const body = [
    "Ola,",
    "",
    `Segue a sequencia planejada para a fazenda ${farmName}.`,
    "",
    "Resumo da selecao:",
    `- Talhoes selecionados: ${rows.length}`,
    `- Area total: ${formatNumber(totalArea)} ha`,
    `- Volume Comercial (VCSC): ${formatNumber(totalVcsc)} m3`,
    "",
    "Sequencia selecionada:",
    orderLines,
    "",
    "Arquivos baixados para anexar:",
    `- ${excelName}`,
    `- ${screenshotName}`,
    "",
    "Observacao: em um site hospedado no GitHub Pages, os anexos nao podem ser inseridos automaticamente no e-mail. Por isso, os arquivos foram baixados no seu computador para voce anexar antes de enviar."
  ].join("\n");

  return {
    toComma: recipients.join(","),
    toSemicolon: recipients.join(";"),
    subject: `Sequencia - ${farmName}`,
    body,
    fullText: `Para: ${recipients.join("; ")}\nAssunto: Sequencia - ${farmName}\n\n${body}`
  };
}

function openEmailModal(draft) {
  emailBodyPreview.value = draft.fullText;
  emailModal.classList.remove("hidden");
  emailModal.setAttribute("aria-hidden", "false");
}

function closeEmailModal() {
  emailModal.classList.add("hidden");
  emailModal.setAttribute("aria-hidden", "true");
}

function normalizeLeafletClone(clonedDoc) {
  const selectors = [
    ".leaflet-map-pane",
    ".leaflet-tile-pane",
    ".leaflet-overlay-pane",
    ".leaflet-shadow-pane",
    ".leaflet-marker-pane",
    ".leaflet-tooltip-pane",
    ".leaflet-popup-pane",
    ".leaflet-zoom-animated",
    ".leaflet-zoom-hide"
  ];

  clonedDoc.querySelectorAll(selectors.join(",")).forEach((node) => {
    const style = clonedDoc.defaultView.getComputedStyle(node);
    const parsed = parseTransform(style.transform);
    if (!parsed) {
      return;
    }
    node.style.transform = "none";
    node.style.left = `${parsed.x}px`;
    node.style.top = `${parsed.y}px`;
  });
}

function parseTransform(transformValue) {
  if (!transformValue || transformValue === "none") {
    return null;
  }

  const matrixMatch = transformValue.match(/^matrix\((.+)\)$/);
  if (matrixMatch) {
    const values = matrixMatch[1].split(",").map((value) => Number(value.trim()));
    return { x: values[4] || 0, y: values[5] || 0 };
  }

  const matrix3dMatch = transformValue.match(/^matrix3d\((.+)\)$/);
  if (matrix3dMatch) {
    const values = matrix3dMatch[1].split(",").map((value) => Number(value.trim()));
    return { x: values[12] || 0, y: values[13] || 0 };
  }

  return null;
}

async function buildEmlBlob(draft, attachments) {
  const boundary = `----=_InventarioFlorestal_${Date.now()}`;
  const chunks = [];

  chunks.push(`To: ${draft.toComma}`);
  chunks.push(`Subject: ${draft.subject}`);
  chunks.push("MIME-Version: 1.0");
  chunks.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
  chunks.push("");
  chunks.push(`--${boundary}`);
  chunks.push('Content-Type: text/plain; charset="UTF-8"');
  chunks.push("Content-Transfer-Encoding: 8bit");
  chunks.push("");
  chunks.push(draft.body);
  chunks.push("");

  const allAttachments = [attachments.excel, attachments.screenshot].filter((item) => item?.blob);
  for (const attachment of allAttachments) {
    const base64 = await blobToBase64(attachment.blob);
    chunks.push(`--${boundary}`);
    chunks.push(`Content-Type: ${attachment.blob.type || "application/octet-stream"}; name="${attachment.filename}"`);
    chunks.push("Content-Transfer-Encoding: base64");
    chunks.push(`Content-Disposition: attachment; filename="${attachment.filename}"`);
    chunks.push("");
    chunks.push(wrapBase64(base64));
    chunks.push("");
  }

  chunks.push(`--${boundary}--`);
  chunks.push("");

  return new Blob([chunks.join("\r\n")], { type: "message/rfc822" });
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = String(reader.result || "");
      const base64 = result.split(",")[1] || "";
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function wrapBase64(value) {
  return value.replace(/(.{76})/g, "$1\r\n");
}

function getColor(value, min, max) {
  if (max === min) {
    return "rgb(26,152,80)";
  }
  const ratio = (value - min) / (max - min);
  if (ratio <= 0.5) {
    return interpolateColor([215, 48, 39], [254, 224, 139], ratio / 0.5);
  }
  return interpolateColor([254, 224, 139], [26, 152, 80], (ratio - 0.5) / 0.5);
}

function interpolateColor(start, end, factor) {
  const color = start.map((component, index) =>
    Math.round(component + factor * (end[index] - component))
  );
  return `rgb(${color[0]}, ${color[1]}, ${color[2]})`;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function turfCentroid(feature) {
  try {
    const coordinates = [];
    collectCoordinates(feature.geometry.coordinates, coordinates);
    if (!coordinates.length) {
      return null;
    }
    const sums = coordinates.reduce(
      (acc, current) => [acc[0] + current[0], acc[1] + current[1]],
      [0, 0]
    );
    return [sums[0] / coordinates.length, sums[1] / coordinates.length];
  } catch {
    return null;
  }
}

function collectCoordinates(input, output) {
  if (!Array.isArray(input)) {
    return;
  }
  if (typeof input[0] === "number" && typeof input[1] === "number") {
    output.push([input[0], input[1]]);
    return;
  }
  input.forEach((item) => collectCoordinates(item, output));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#039;");
}
