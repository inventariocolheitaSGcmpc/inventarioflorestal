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
const exportStage = document.getElementById("exportStage");
const exportTitle = document.getElementById("exportTitle");

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
  const excelAttachment = createXlsxAttachment(rows, currentFarmName());
  downloadBlob(excelAttachment.blob, excelAttachment.filename);
});

emailBtn.addEventListener("click", async () => {
  const rows = selectedRows();
  if (!rows.length) {
    alert("Selecione pelo menos um talhao.");
    return;
  }

  const excelAttachment = createXlsxAttachment(rows, currentFarmName());
  const screenshotAttachment = await createMapScreenshotAttachment(currentFarmName());
  downloadBlob(excelAttachment.blob, excelAttachment.filename);

  if (screenshotAttachment?.blob) {
    downloadBlob(screenshotAttachment.blob, screenshotAttachment.filename);
  } else {
    alert("A sequencia em Excel foi baixada, mas nao foi possivel gerar o print nesta tentativa.");
  }

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
  const selectedFeatures = selectedFeatureObjects();

  if (!selectedFeatures.length) {
    return { filename, blob: null };
  }

  try {
    const canvas = renderSequenceCanvas(selectedFeatures, farmName);
    const blob = await canvasToBlob(canvas);

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

function canvasToBlob(canvas) {
  return new Promise((resolve) => {
    canvas.toBlob((blob) => resolve(blob || null), "image/png");
  });
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
    orderLines
  ].join("\n");

  return {
    toComma: recipients.join(","),
    toSemicolon: recipients.join(";"),
    subject: `Sequencia - ${farmName}`,
    body,
    fullText: body
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

async function buildExportMap(selectedFeatures) {
  const exportMapContainer = document.getElementById("exportMap");
  exportMapContainer.innerHTML = "";
  const canvasRenderer = L.canvas({ padding: 0.5 });

  const exportMap = L.map(exportMapContainer, {
    preferCanvas: true,
    zoomControl: false,
    attributionControl: true
  });

  const satelliteLayer = L.tileLayer(
    "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
    {
      attribution: "Tiles &copy; Esri",
      crossOrigin: true
    }
  ).addTo(exportMap);

  const values = state.currentFeatures.map((feature) => feature.properties.produtividade);
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);

  const geoJsonLayer = L.geoJSON(
    {
      type: "FeatureCollection",
      features: selectedFeatures
    },
    {
      style: (feature) => ({
        renderer: canvasRenderer,
        color: "#ffff00",
        weight: 4,
        fillOpacity: 0.42,
        fillColor: getColor(feature.properties.produtividade, minValue, maxValue)
      })
    }
  ).addTo(exportMap);

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
    }).addTo(exportMap);
  });

  if (points.length > 1) {
    L.polyline(points.map(([lng, lat]) => [lat, lng]), {
      renderer: canvasRenderer,
      color: "#ffff00",
      weight: 5,
      dashArray: "8, 12",
      opacity: 1
    }).addTo(exportMap);
  }

  exportMap.fitBounds(geoJsonLayer.getBounds(), { padding: [40, 40] });

  addExportLegend(exportMap, minValue, maxValue);

  await waitForLeafletTiles(exportMap, [satelliteLayer]);
  await wait(500);

  return exportMap;
}

function renderLeafletMapToCanvas(targetMap) {
  return new Promise((resolve, reject) => {
    if (typeof leafletImage !== "function") {
      reject(new Error("leafletImage nao esta disponivel."));
      return;
    }
    leafletImage(targetMap, (error, canvas) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(canvas);
    });
  });
}

function composeExportCanvas(mapCanvas, farmName) {
  const width = 1280;
  const height = 920;
  const headerHeight = 160;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#eef2e7";
  ctx.fillRect(0, 0, width, height);

  roundRect(ctx, 24, 24, width - 48, height - 48, 28);
  ctx.fillStyle = "rgba(255,255,255,0.98)";
  ctx.fill();
  ctx.strokeStyle = "rgba(22,33,23,0.12)";
  ctx.lineWidth = 1;
  ctx.stroke();

  ctx.fillStyle = "#2f6b3b";
  ctx.font = "800 16px Manrope, sans-serif";
  ctx.fillText("SEQUENCIA DE COLHEITA", 56, 74);

  ctx.fillStyle = "#162117";
  ctx.font = "800 42px Manrope, sans-serif";
  ctx.fillText(`Inventario Florestal - ${farmName}`, 56, 122);

  const logo = document.querySelector(".export-logo");
  if (logo && logo.complete) {
    ctx.drawImage(logo, width - 210, 48, 130, 72);
  }

  ctx.drawImage(mapCanvas, 40, headerHeight + 24, width - 80, 640);

  return canvas;
}

function roundRect(ctx, x, y, width, height, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function renderSequenceCanvas(selectedFeatures, farmName) {
  const width = 1280;
  const height = 920;
  const headerHeight = 150;
  const footerHeight = 70;
  const mapX = 40;
  const mapY = headerHeight + 20;
  const mapWidth = width - 80;
  const mapHeight = height - headerHeight - footerHeight - 40;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#eef2e7";
  ctx.fillRect(0, 0, width, height);

  roundRect(ctx, 24, 24, width - 48, height - 48, 28);
  ctx.fillStyle = "rgba(255,255,255,0.98)";
  ctx.fill();
  ctx.strokeStyle = "rgba(22,33,23,0.12)";
  ctx.lineWidth = 1;
  ctx.stroke();

  ctx.fillStyle = "#2f6b3b";
  ctx.font = "800 16px Manrope, sans-serif";
  ctx.fillText("SEQUENCIA DE COLHEITA", 56, 72);
  ctx.fillStyle = "#162117";
  ctx.font = "800 40px Manrope, sans-serif";
  ctx.fillText(`Inventario Florestal - ${farmName}`, 56, 118);

  ctx.fillStyle = "#f8fbf4";
  roundRect(ctx, mapX, mapY, mapWidth, mapHeight, 22);
  ctx.fill();
  ctx.strokeStyle = "rgba(22,33,23,0.10)";
  ctx.stroke();

  const bbox = getFeatureBounds(selectedFeatures);
  const projected = projectFeatures(selectedFeatures, bbox, mapWidth, mapHeight, mapX, mapY);
  const values = selectedFeatures.map((feature) => feature.properties.produtividade);
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);

  projected.forEach((feature) => {
    feature.paths.forEach((polygon) => {
      ctx.beginPath();
      polygon.forEach((ring) => {
        ring.forEach((point, index) => {
          if (index === 0) {
            ctx.moveTo(point.x, point.y);
          } else {
            ctx.lineTo(point.x, point.y);
          }
        });
        ctx.closePath();
      });
      ctx.fillStyle = getColor(feature.properties.produtividade, minValue, maxValue);
      ctx.globalAlpha = 0.34;
      ctx.fill("evenodd");
      ctx.globalAlpha = 1;
      ctx.strokeStyle = "#fff34f";
      ctx.lineWidth = 4;
      ctx.stroke();
      ctx.strokeStyle = "#222222";
      ctx.lineWidth = 1.5;
      ctx.stroke();
    });
  });

  const centers = projected.map((feature) => feature.center);
  if (centers.length > 1) {
    ctx.beginPath();
    centers.forEach((center, index) => {
      if (index === 0) {
        ctx.moveTo(center.x, center.y);
      } else {
        ctx.lineTo(center.x, center.y);
      }
    });
    ctx.strokeStyle = "#ffe600";
    ctx.lineWidth = 5;
    ctx.setLineDash([10, 12]);
    ctx.stroke();
    ctx.setLineDash([]);
  }

  projected.forEach((feature, index) => {
    const { x, y } = feature.center;
    ctx.beginPath();
    ctx.arc(x, y, 16, 0, Math.PI * 2);
    ctx.fillStyle = "#ffffff";
    ctx.fill();
    ctx.strokeStyle = "#222222";
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.fillStyle = "#111111";
    ctx.font = "800 18px Manrope, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(String(index + 1), x, y + 0.5);

    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    ctx.fillStyle = "#d9534f";
    ctx.font = "800 15px Manrope, sans-serif";
    ctx.fillText(feature.properties.cd_talhao, x + 18, y - 10);
  });

  drawExportLegend(ctx, minValue, maxValue, width - 260, height - 120);

  return canvas;
}

function getFeatureBounds(features) {
  const allPoints = [];
  features.forEach((feature) => collectCoordinates(feature.geometry.coordinates, allPoints));
  const xs = allPoints.map((point) => point[0]);
  const ys = allPoints.map((point) => point[1]);
  return {
    minX: Math.min(...xs),
    maxX: Math.max(...xs),
    minY: Math.min(...ys),
    maxY: Math.max(...ys)
  };
}

function projectFeatures(features, bbox, mapWidth, mapHeight, offsetX, offsetY) {
  const padding = 40;
  const usableWidth = mapWidth - padding * 2;
  const usableHeight = mapHeight - padding * 2;
  const widthSpan = Math.max(bbox.maxX - bbox.minX, 0.000001);
  const heightSpan = Math.max(bbox.maxY - bbox.minY, 0.000001);
  const scale = Math.min(usableWidth / widthSpan, usableHeight / heightSpan);

  const projectPoint = ([x, y]) => ({
    x: offsetX + padding + (x - bbox.minX) * scale,
    y: offsetY + mapHeight - padding - (y - bbox.minY) * scale
  });

  return features.map((feature) => {
    const paths = geometryToProjectedPaths(feature.geometry, projectPoint);
    const centroid = turfCentroid(feature) || [bbox.minX, bbox.minY];
    return {
      properties: feature.properties,
      paths,
      center: projectPoint(centroid)
    };
  });
}

function geometryToProjectedPaths(geometry, projectPoint) {
  if (geometry.type === "Polygon") {
    return [geometry.coordinates.map((ring) => ring.map(projectPoint))];
  }
  if (geometry.type === "MultiPolygon") {
    return geometry.coordinates.map((polygon) => polygon.map((ring) => ring.map(projectPoint)));
  }
  return [];
}

function drawExportLegend(ctx, minValue, maxValue, x, y) {
  const width = 200;
  const height = 72;
  roundRect(ctx, x, y, width, height, 18);
  ctx.fillStyle = "rgba(255,255,255,0.94)";
  ctx.fill();
  ctx.strokeStyle = "rgba(0,0,0,0.08)";
  ctx.stroke();

  ctx.fillStyle = "#162117";
  ctx.font = "700 13px Manrope, sans-serif";
  ctx.fillText("VCSC (m3/ha)", x + 14, y + 22);

  const gradient = ctx.createLinearGradient(x + 14, y + 42, x + 154, y + 42);
  gradient.addColorStop(0, "#d73027");
  gradient.addColorStop(0.5, "#fee08b");
  gradient.addColorStop(1, "#1a9850");
  ctx.fillStyle = gradient;
  roundRect(ctx, x + 14, y + 34, 140, 12, 6);
  ctx.fill();

  ctx.fillStyle = "#5f6f60";
  ctx.font = "600 11px Manrope, sans-serif";
  ctx.fillText(formatNumber(minValue), x + 14, y + 62);
  ctx.fillText(formatNumber(maxValue), x + 118, y + 62);
}

function addExportLegend(targetMap, minValue, maxValue) {
  const exportLegend = L.control({ position: "bottomright" });
  exportLegend.onAdd = function onAdd() {
    const div = L.DomUtil.create("div", "info legend");
    div.innerHTML = `
      <div style="padding:12px 14px;background:rgba(255,255,255,.92);border-radius:16px;border:1px solid rgba(0,0,0,.08);box-shadow:0 24px 60px rgba(34,53,29,.12);font:12px Manrope,sans-serif;">
        <strong style="display:block;margin-bottom:8px;">VCSC (m3/ha)</strong>
        <div style="display:flex;align-items:center;gap:8px;">
          <span>${formatNumber(minValue)}</span>
          <div style="width:140px;height:12px;border-radius:999px;background:linear-gradient(90deg,#d73027,#fee08b,#1a9850);"></div>
          <span>${formatNumber(maxValue)}</span>
        </div>
      </div>
    `;
    return div;
  };
  exportLegend.addTo(targetMap);
}

function waitForLeafletTiles(targetMap, layers) {
  return new Promise((resolve) => {
    let remaining = 0;
    let done = false;

    const finish = () => {
      if (!done) {
        done = true;
        resolve();
      }
    };

    const register = (layer) => {
      if (!layer || typeof layer.isLoading !== "function") {
        return;
      }
      if (layer.isLoading()) {
        remaining += 1;
        layer.once("load", () => {
          remaining -= 1;
          if (remaining <= 0) {
            window.setTimeout(finish, 250);
          }
        });
      }
    };

    layers.forEach(register);
    targetMap.whenReady(() => {
      if (remaining === 0) {
        window.setTimeout(finish, 500);
      }
    });
    window.setTimeout(finish, 3500);
  });
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
