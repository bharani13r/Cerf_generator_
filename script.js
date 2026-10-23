
const $ = id => document.getElementById(id);

let participants = [];
let generatedFiles = [];
let placementMode = false;
let generatedRegisterBlob = null;
let previewUrl = null;
let currentType = "winner";

const templates = {
  winner: { image: null, data: null, positions: {} },
  participant: { image: null, data: null, positions: {} }
};

const requiredColumns = [
  "Serial No.", "Name", "Event", "Date", "Email", "Position"
];

function setStatus(id, message, isError = false) {
  const el = $(id);
  if (!el) return;

  el.textContent = message;
  el.classList.toggle("error", isError);
}

function normalizeKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[c]);
}

function createCertificateId(serialNo, usedIds) {
  if (!crypto?.getRandomValues) {
    throw new Error(
      "Secure random generation is unavailable. Use HTTPS or localhost."
    );
  }

  let id;

  do {
    const bytes = new Uint8Array(6);
    crypto.getRandomValues(bytes);

    const random = Array.from(bytes)
      .map(b => b.toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();

    id = `CERT-${serialNo}-${random}`;
  } while (usedIds.has(id));

  usedIds.add(id);
  return id;
}

function renderParticipants() {
  const container = $("participantsTable");

  if (!participants.length) {
    container.innerHTML = "";
    return;
  }

  const rows = participants.map(p => `
    <tr>
      <td>${escapeHtml(p["Serial No."])}</td>
      <td>${escapeHtml(p.Name)}</td>
      <td>${escapeHtml(p.Event)}</td>
      <td>${escapeHtml(p.Date)}</td>
      <td>${escapeHtml(p.Email)}</td>
      <td>${escapeHtml(p.Position || "Participation")}</td>
      <td>${escapeHtml(p.certificateId)}</td>
      <td>${p.type === "winner" ? "Winner" : "Participation"}</td>
    </tr>
  `).join("");

  container.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Serial No.</th>
          <th>Name</th>
          <th>Event</th>
          <th>Date</th>
          <th>Email (local only)</th>
          <th>Position</th>
          <th>Certificate Number</th>
          <th>Template</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
}

async function readExcel() {
  const file = $("excelFile").files[0];

  if (!file) {
    setStatus("excelStatus", "Please select an Excel file.", true);
    return;
  }

  if (!window.XLSX) {
    setStatus(
      "excelStatus",
      "Excel library failed to load. Check your internet connection.",
      true
    );
    return;
  }

  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: true
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false
    });

    if (!rows.length) {
      throw new Error("The Excel sheet is empty.");
    }

    if (rows.length > 5) {
      throw new Error("Maximum 5 participants are supported.");
    }

    const headerMap = {};

    Object.keys(rows[0]).forEach(header => {
      headerMap[normalizeKey(header)] = header;
    });

    const missing = requiredColumns.filter(
      column => !headerMap[normalizeKey(column)]
    );

    if (missing.length) {
      throw new Error(
        `Missing required columns: ${missing.join(", ")}`
      );
    }

    const serials = rows.map(row =>
      String(
        row[headerMap[normalizeKey("Serial No.")]] ?? ""
      ).trim()
    );

    if (
      serials.some(
        serial =>
          !serial ||
          !Number.isInteger(Number(serial)) ||
          Number(serial) < 1
      )
    ) {
      throw new Error(
        "Every Serial No. must be a positive whole number."
      );
    }

    if (new Set(serials).size !== serials.length) {
      throw new Error("Serial No. values must be unique.");
    }

    const usedIds = new Set();

    const loaded = rows.map((row, index) => {
      const person = {};

      requiredColumns.forEach(column => {
        person[column] = String(
          row[headerMap[normalizeKey(column)]] ?? ""
        ).trim();
      });

      for (const column of ["Name", "Event", "Date", "Email"]) {
        if (!person[column]) {
          throw new Error(
            `Row ${index + 2}: ${column} is empty.`
          );
        }
      }

      if (
        !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(person.Email)
      ) {
        throw new Error(
          `Row ${index + 2}: Email address is invalid.`
        );
      }

      
const position = person.Position.trim().toLowerCase();

    person.type = [
  "1", "2", "3",
  "first", "second", "third",
  "1st", "2nd", "3rd"
  ].includes(position)
  ? "winner"
  : "participant";


      person.certificateId = createCertificateId(
        person["Serial No."],
        usedIds
      );

      return person;
    });

    participants = loaded;
    generatedFiles = [];
    generatedRegisterBlob = null;

    $("downloads").innerHTML = "";
    $("registerDownload").innerHTML = "";
    $("pdfPreview").innerHTML = "";
    $("zipBtn").disabled = true;
    $("registerBtn").disabled = true;

    renderParticipants();

    setStatus(
      "excelStatus",
      `${participants.length} participant(s) loaded. ` +
      "Positions 1–3 use the winner template; other or blank positions use the participation template."
    );
  } catch (error) {
    setStatus(
      "excelStatus",
      error.message || "Could not read the spreadsheet.",
      true
    );
  }
}

function updateTemplateType() {
  placementMode = false;
  $("placeModeBtn").textContent = "Click template to place field";

  currentType = $("templateType").value;
  $("templateFile").value = "";

  drawTemplate();

  setStatus(
    "templateStatus",
    templates[currentType].image
      ? "Previously uploaded template selected."
      : "Upload an image for the selected template."
  );
}

function uploadTemplate() {
  const file = $("templateFile").files[0];

  if (!file) return;

  if (!["image/jpeg", "image/png"].includes(file.type)) {
    setStatus(
      "templateStatus",
      "Please upload a JPG or PNG image.",
      true
    );

    $("templateFile").value = "";
    return;
  }

  const reader = new FileReader();

  reader.onload = () => {
    const image = new Image();

    image.onload = () => {
      templates[currentType].image = image;
      templates[currentType].data = reader.result;

      drawTemplate();

      setStatus(
        "templateStatus",
        "Template uploaded successfully."
      );
    };

    image.onerror = () => {
      setStatus(
        "templateStatus",
        "Could not load the image.",
        true
      );
    };

    image.src = reader.result;
  };

  reader.readAsDataURL(file);
}

function drawTemplate() {
  const canvas = $("templateCanvas");
  const empty = $("emptyMessage");
  const markers = $("markers");
  const image = templates[currentType].image;

  if (!image) {
    canvas.style.display = "none";
    empty.style.display = "block";
    markers.innerHTML = "";
    return;
  }

  // Canvas drawing dimensions match the original image pixels.
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;

  const ctx = canvas.getContext("2d");

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(image, 0, 0);

  canvas.style.display = "block";
  empty.style.display = "none";

  // Keep the wrapper proportional to the original template.
  $("canvasWrap").style.aspectRatio =
    `${image.naturalWidth} / ${image.naturalHeight}`;

  requestAnimationFrame(renderPositions);
}

function getCanvasPoint(event) {
  const canvas = $("templateCanvas");
  const rect = canvas.getBoundingClientRect();

  if (!rect.width || !rect.height) {
    return { x: 0, y: 0 };
  }

  return {
    x: Math.max(
      0,
      Math.min(1, (event.clientX - rect.left) / rect.width)
    ),
    y: Math.max(
      0,
      Math.min(1, (event.clientY - rect.top) / rect.height)
    )
  };
}

function placeField(event) {
  if (!placementMode) return;

  if (!templates[currentType].image) {
    setStatus(
      "editorStatus",
      "Upload a template image first.",
      true
    );
    return;
  }

  const field = $("fieldSelect").value;
  const point = getCanvasPoint(event);

  templates[currentType].positions[field] = {
    x: point.x,
    y: point.y,
    fontSize: Math.max(
      6,
      Math.min(72, Number($("fontSize").value) || 24)
    ),
    font: $("fontSelect").value,
    bold: $("boldText").checked,
    color: $("textColor").value
  };

  placementMode = false;
  $("placeModeBtn").textContent = "Click template to place field";

  renderPositions();

  setStatus(
    "editorStatus",
    `${field} position saved.`
  );
}

function renderPositions() {
  const positions = templates[currentType].positions;
  const list = $("positionList");
  const markers = $("markers");

  list.innerHTML = "";
  markers.innerHTML = "";

  Object.entries(positions).forEach(([field, pos]) => {
    const line = document.createElement("div");

    line.textContent =
      `${field}: x ${(pos.x * 100).toFixed(1)}%, ` +
      `y ${(pos.y * 100).toFixed(1)}%`;

    list.appendChild(line);

    const marker = document.createElement("div");

    marker.className = "marker";
    marker.style.left = `${pos.x * 100}%`;
    marker.style.top = `${pos.y * 100}%`;

    const label = document.createElement("span");
    label.textContent = field;

    marker.appendChild(label);
    markers.appendChild(marker);
  });
}

function clearSelectedField() {
  const field = $("fieldSelect").value;

  delete templates[currentType].positions[field];

  renderPositions();

  setStatus(
    "editorStatus",
    `${field} position cleared.`
  );
}

function clearAllPositions() {
  templates[currentType].positions = {};

  renderPositions();

  setStatus(
    "editorStatus",
    "All positions cleared for this template."
  );
}

function hexToRgb(hex) {
  const s = hex.replace("#", "");

  return {
    r: parseInt(s.slice(0, 2), 16),
    g: parseInt(s.slice(2, 4), 16),
    b: parseInt(s.slice(4, 6), 16)
  };
}

function getVerificationPageUrl() {
  const entered = $("verificationUrl").value.trim();
  let url;

  if (entered) {
    try {
      url = new URL(entered);
    } catch {
      throw new Error(
        "Enter a valid public verification-page URL."
      );
    }

    if (!["http:", "https:"].includes(url.protocol)) {
      throw new Error(
        "The verification URL must start with http:// or https://."
      );
    }

    url.search = "";
    url.hash = "";

    return url;
  }

  if (["http:", "https:"].includes(location.protocol)) {
    return new URL("verify.html", location.href);
  }

  throw new Error(
    "Enter your published GitHub Pages verify.html URL before generating PDFs."
  );
}

function makeQrImage(url) {
  return new Promise((resolve, reject) => {
    if (!window.QRCode) {
      reject(
        new Error(
          "QR library failed to load. Check your internet connection."
        )
      );
      return;
    }

    const holder = document.createElement("div");

    Object.assign(holder.style, {
      position: "fixed",
      left: "-10000px",
      top: "0"
    });

    document.body.appendChild(holder);

    try {
      new QRCode(holder, {
        text: url,
        width: 220,
        height: 220,
        correctLevel: QRCode.CorrectLevel.M
      });

      setTimeout(() => {
        try {
          const canvas = holder.querySelector("canvas");
          const img = holder.querySelector("img");

          if (canvas) {
            resolve(canvas.toDataURL("image/png"));
          } else if (img?.src) {
            resolve(img.src);
          } else {
            reject(new Error("Could not create QR image."));
          }
        } catch (error) {
          reject(error);
        } finally {
          holder.remove();
        }
      }, 150);
    } catch (error) {
      holder.remove();
      reject(error);
    }
  });
}

function fieldValue(person, field) {
  if (field === "Position" && !person.Position) {
    return "";
  }

  return String(person[field] ?? "");
}

async function makePDF(person, template) {
  if (!window.jspdf) {
    throw new Error("PDF library failed to load.");
  }

  if (!template.image || !template.data) {
    throw new Error(
      `Upload the ${person.type === "winner" ? "winner" : "participation"} template first.`
    );
  }

  const { jsPDF } = window.jspdf;
  const image = template.image;

  // PDF page dimensions follow the exact aspect ratio of the template.
  const pageW = 842;
  const pageH =
    pageW * image.naturalHeight / image.naturalWidth;

  const pdf = new jsPDF({
    orientation:
      image.naturalWidth >= image.naturalHeight
        ? "landscape"
        : "portrait",
    unit: "pt",
    format: [pageW, pageH],
    compress: true
  });

  // Draw the template without stretching it.
  pdf.addImage(
    template.data,
    template.data.startsWith("data:image/png")
      ? "PNG"
      : "JPEG",
    0,
    0,
    pageW,
    pageH
  );

  // Positions are normalized from the editor's top-left origin.
  Object.entries(template.positions).forEach(([field, pos]) => {
    const value = fieldValue(person, field);

    if (!value) return;
    
    const horizontalCorrection = 48;
    const x = Number(pos.x) * pageW;
    const topY = Number(pos.y) * pageH;
    const fontSize = Number(pos.fontSize) || 24;

    const rgb = hexToRgb(pos.color || "#111111");

    pdf.setTextColor(rgb.r, rgb.g, rgb.b);

    pdf.setFont(
      pos.font || "helvetica",
      pos.bold ? "bold" : "normal"
    );

    pdf.setFontSize(fontSize);

    // The editor's y coordinate marks the text's approximate top edge.
    // jsPDF's default alphabetic baseline is below that point.
    const baselineY = topY + fontSize * 0.35;

    pdf.text(value, x, baselineY, {
      align: "center",
      baseline: "alphabetic"
    });
  });

  // Generate QR code linked to the public verification page.
  const verifyUrl = getVerificationPageUrl();

  verifyUrl.searchParams.set("id", person.certificateId);

  const qrImage = await makeQrImage(verifyUrl.href);

  // Keep the QR and certificate ID in the bottom-right corner.
  const qrSize = Math.min(
    58,
    pageW * 0.09,
    pageH * 0.13
  );

  const margin = Math.min(
    20,
    pageW * 0.03,
    pageH * 0.03
  );

  const qrX = pageW - qrSize - margin;
  const qrY = pageH - qrSize - margin;

  pdf.addImage(
    qrImage,
    "PNG",
    qrX,
    qrY,
    qrSize,
    qrSize
  );

  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(7);
  pdf.setTextColor(0, 0, 0);

  pdf.text(
    person.certificateId,
    qrX + qrSize / 2,
    qrY - 5,
    { align: "center" }
  );

  return pdf.output("blob");
}

async function previewFirstParticipant() {
  if (!participants.length) {
    setStatus(
      "generateStatus",
      "Read the Excel spreadsheet first.",
      true
    );
    return;
  }

  try {
    setStatus(
      "generateStatus",
      "Creating preview..."
    );

    const person = participants[0];
    const template = templates[person.type];

    const blob = await makePDF(person, template);

    if (previewUrl) {
      URL.revokeObjectURL(previewUrl);
    }

    previewUrl = URL.createObjectURL(blob);

    $("pdfPreview").innerHTML = "";

    const iframe = document.createElement("iframe");

    iframe.title = "Certificate preview";
    iframe.src = previewUrl;

    $("pdfPreview").appendChild(iframe);

    setStatus(
      "generateStatus",
      "Preview ready."
    );
  } catch (error) {
    setStatus(
      "generateStatus",
      error.message || "Could not create preview.",
      true
    );
  }
}

function createPublicRegister() {
  return {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),

    certificates: participants.map(person => ({
      id: person.certificateId,
      serialNo: person["Serial No."],
      name: person.Name,
      event: person.Event,
      date: person.Date,
      position: person.Position || "",
      type: person.type
    }))
  };
}

async function generatePDFs() {
  if (!participants.length) {
    setStatus(
      "generateStatus",
      "Read the Excel spreadsheet first.",
      true
    );
    return;
  }

  try {
    getVerificationPageUrl();
  } catch (error) {
    setStatus(
      "generateStatus",
      error.message,
      true
    );
    return;
  }

  if (!window.jspdf || !window.QRCode || !window.JSZip) {
    setStatus(
      "generateStatus",
      "A required library failed to load. Check your internet connection.",
      true
    );
    return;
  }

  generatedFiles = [];
  generatedRegisterBlob = null;

  $("downloads").innerHTML = "";
  $("registerDownload").innerHTML = "";
  $("pdfPreview").innerHTML = "";

  $("zipBtn").disabled = true;
  $("registerBtn").disabled = true;
  $("generateBtn").disabled = true;

  try {
    for (let i = 0; i < participants.length; i++) {
      const person = participants[i];

      setStatus(
        "generateStatus",
        `Generating certificate ${i + 1} of ${participants.length}...`
      );

      const blob = await makePDF(
        person,
        templates[person.type]
      );

      generatedFiles.push({
        name: `Certificate_${person["Serial No."]}.pdf`,
        blob
      });
    }

    generatedFiles.forEach(file => {
      const url = URL.createObjectURL(file.blob);
      const link = document.createElement("a");

      link.className = "download-link";
      link.href = url;
      link.download = file.name;
      link.textContent = `Download ${file.name}`;

      $("downloads").appendChild(link);
    });

    generatedRegisterBlob = new Blob(
      [JSON.stringify(createPublicRegister(), null, 2)],
      { type: "application/json;charset=utf-8" }
    );

    const registerUrl = URL.createObjectURL(
      generatedRegisterBlob
    );

    const registerLink = document.createElement("a");

    registerLink.className = "download-link";
    registerLink.href = registerUrl;
    registerLink.download = "certificates.json";
    registerLink.textContent =
      "Download certificates.json (public verification register)";

    $("registerDownload").appendChild(registerLink);

    $("zipBtn").disabled = false;
    $("registerBtn").disabled = false;

    setStatus(
      "generateStatus",
      `Generated ${generatedFiles.length} PDF(s) and certificates.json. ` +
      "Upload the JSON file beside verify.html on GitHub Pages. Email addresses are excluded."
    );
  } catch (error) {
    generatedFiles = [];
    generatedRegisterBlob = null;

    $("zipBtn").disabled = true;
    $("registerBtn").disabled = true;

    setStatus(
      "generateStatus",
      error.message || "PDF generation failed.",
      true
    );
  } finally {
    $("generateBtn").disabled = false;
  }
}

async function downloadZip() {
  if (!generatedFiles.length) return;

  try {
    const zip = new JSZip();

    generatedFiles.forEach(file => {
      zip.file(file.name, file.blob);
    });

    const blob = await zip.generateAsync({
      type: "blob"
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "Certificates.zip";
    link.click();

    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    setStatus(
      "generateStatus",
      error.message || "Could not create ZIP file.",
      true
    );
  }
}

function downloadRegister() {
  if (!generatedRegisterBlob) {
    setStatus(
      "generateStatus",
      "Generate certificates first.",
      true
    );
    return;
  }

  const url = URL.createObjectURL(generatedRegisterBlob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "certificates.json";
  link.click();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Event listeners
$("readBtn").addEventListener("click", readExcel);

$("templateType").addEventListener(
  "change",
  updateTemplateType
);

$("templateFile").addEventListener(
  "change",
  uploadTemplate
);

$("placeModeBtn").addEventListener("click", () => {
  placementMode = true;

  $("placeModeBtn").textContent =
    "Now click the template";

  setStatus(
    "editorStatus",
    `Click the template to place ${
      $("fieldSelect").selectedOptions[0].textContent
    }.`
  );
});

$("templateCanvas").addEventListener(
  "click",
  placeField
);

$("clearFieldBtn").addEventListener(
  "click",
  clearSelectedField
);

$("clearAllBtn").addEventListener(
  "click",
  clearAllPositions
);

$("previewBtn").addEventListener(
  "click",
  previewFirstParticipant
);

$("generateBtn").addEventListener(
  "click",
  generatePDFs
);

$("zipBtn").addEventListener(
  "click",
  downloadZip
);

$("registerBtn").addEventListener(
  "click",
  downloadRegister
);

window.addEventListener("resize", () => {
  if (templates[currentType].image) {
    drawTemplate();
  }
});
