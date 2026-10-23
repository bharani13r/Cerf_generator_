const $ = id => document.getElementById(id);

let participants = [];
let generatedFiles = [];
let currentType = "winner";
let placementMode = false;
let generatedRegisterBlob = null;

const templates = {
  winner: { image: null, data: null, positions: {} },
  participant: { image: null, data: null, positions: {} }
};

const requiredColumns = [
  "Serial No.",
  "Name",
  "Event",
  "Date",
  "Email",
  "Position"
];

function setStatus(id, message, isError = false) {
  const element = $(id);
  element.textContent = message;
  element.classList.toggle("error", isError);
}

function normalizeKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[character]);
}

// Generates a unique certificate number using the serial number
// and a cryptographically secure random component.
function createCertificateId(serialNo, usedIds) {
  if (!window.crypto || !window.crypto.getRandomValues) {
    throw new Error(
      "Secure random number generation is unavailable. Use a modern browser on HTTPS or localhost."
    );
  }

  let id;

  do {
    const bytes = new Uint8Array(6);
    window.crypto.getRandomValues(bytes);

    const randomPart = Array.from(bytes)
      .map(byte => byte.toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();

    id = `CERT-${serialNo}-${randomPart}`;
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

  const rows = participants.map(person => `
    <tr>
      <td>${escapeHtml(person["Serial No."])}</td>
      <td>${escapeHtml(person.Name)}</td>
      <td>${escapeHtml(person.Event)}</td>
      <td>${escapeHtml(person.Date)}</td>
      <td>${escapeHtml(person.Email)}</td>
      <td>${escapeHtml(person.Position || "Participation")}</td>
      <td>${escapeHtml(person.certificateId)}</td>
      <td>${escapeHtml(person.type === "winner" ? "Winner" : "Participation")}</td>
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
          <th>Email (kept local)</th>
          <th>Position</th>
          <th>Certificate Number</th>
          <th>Template</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
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
      "The Excel library did not load. Check your internet connection and reload.",
      true
    );
    return;
  }

  try {
    const buffer = await file.arrayBuffer();

    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rawRows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false
    });

    if (!rawRows.length) {
      setStatus("excelStatus", "The Excel sheet is empty.", true);
      return;
    }

    const headers = Object.keys(rawRows[0]);
    const headerMap = {};

    headers.forEach(header => {
      headerMap[normalizeKey(header)] = header;
    });

    const missing = requiredColumns.filter(
      column => !headerMap[normalizeKey(column)]
    );

    if (missing.length) {
      setStatus(
        "excelStatus",
        `Missing required column headers: ${missing.join(", ")}`,
        true
      );
      return;
    }

    if (rawRows.length > 5) {
      setStatus(
        "excelStatus",
        "Maximum 5 participants are supported in this version.",
        true
      );
      return;
    }

    // Validate all serial numbers before creating participant records.
    const serialValues = rawRows.map(row =>
      String(row[headerMap[normalizeKey("Serial No.")]] ?? "").trim()
    );

    const numericSerials = serialValues.map(Number);

    if (serialValues.some((value, index) =>
      !value ||
      !Number.isInteger(numericSerials[index]) ||
      numericSerials[index] < 1
    )) {
      const badIndex = serialValues.findIndex((value, index) =>
        !value ||
        !Number.isInteger(numericSerials[index]) ||
        numericSerials[index] < 1
      );

      throw new Error(
        `Row ${badIndex + 2}: Serial No. must be a positive whole number.`
      );
    }

    if (new Set(serialValues).size !== serialValues.length) {
      setStatus(
        "excelStatus",
        "Serial No. values must be unique.",
        true
      );
      return;
    }

    const usedIds = new Set();

    const newParticipants = rawRows.map((row, index) => {
      const person = {};

      requiredColumns.forEach(column => {
        person[column] = String(
          row[headerMap[normalizeKey(column)]] ?? ""
        ).trim();
      });

      // Name, Event, Date and Email remain required.
      // Position may be blank for participation certificates.
      for (const column of ["Name", "Event", "Date", "Email"]) {
        if (!person[column]) {
          throw new Error(
            `Row ${index + 2}: ${column} is empty.`
          );
        }
      }

      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(person.Email)) {
        throw new Error(
          `Row ${index + 2}: Email address is invalid.`
        );
      }

      const position = person.Position.trim();

      // Only positions 1, 2 and 3 are winners.
      // Blank or other positions use the participation template.
      person.type = ["1", "2", "3"].includes(position)
        ? "winner"
        : "participant";

      person.certificateId = createCertificateId(
        person["Serial No."],
        usedIds
      );

      return person;
    });

    participants = newParticipants;
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
      `${participants.length} participant(s) loaded successfully. Certificate numbers created. Positions 1–3 use the winner template; blank or other positions use the participation template.`
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
  currentType = $("templateType").value;
  $("templateFile").value = "";

  drawTemplate();
  renderPositions();

  setStatus(
    "templateStatus",
    templates[currentType].image
      ? `${currentType === "winner" ? "Winner" : "Participation"} template selected.`
      : "Upload an image for the selected template."
  );
}

function uploadTemplate() {
  const file = $("templateFile").files[0];

  if (!file) return;

  const allowed = ["image/jpeg", "image/png"];

  if (!allowed.includes(file.type)) {
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
        `${currentType === "winner" ? "Winner" : "Participation"} template uploaded successfully.`
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
  const emptyMessage = $("emptyMessage");
  const image = templates[currentType].image;

  if (!image) {
    canvas.style.display = "none";
    emptyMessage.style.display = "block";
    $("markers").innerHTML = "";
    return;
  }

  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;

  const context = canvas.getContext("2d");
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0);

  canvas.style.display = "block";
  canvas.style.width = "100%";
  canvas.style.maxWidth = "850px";
  canvas.style.height = "auto";

  emptyMessage.style.display = "none";

  requestAnimationFrame(() => {
    const rect = canvas.getBoundingClientRect();
    const markers = $("markers");

    markers.style.width = `${rect.width}px`;
    markers.style.height = `${rect.height}px`;

    renderPositions();
  });
}

function getCanvasPoint(event) {
  const canvas = $("templateCanvas");
  const rect = canvas.getBoundingClientRect();

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

  const point = getCanvasPoint(event);
  const field = $("fieldSelect").value;

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
    `${field === "CertificateNumber" ? "Certificate number" : field} position saved for this template.`
  );
}

function renderPositions() {
  const positions = templates[currentType].positions;
  const list = $("positionList");
  const markers = $("markers");

  list.innerHTML = "";
  markers.innerHTML = "";

  const labels = {
    CertificateNumber: "Certificate number"
  };

  Object.entries(positions).forEach(([field, position]) => {
    const line = document.createElement("div");

    line.textContent =
      `${labels[field] || field}: ` +
      `x ${(position.x * 100).toFixed(1)}%, ` +
      `y ${(position.y * 100).toFixed(1)}%`;

    list.appendChild(line);

    const marker = document.createElement("div");
    marker.className = "marker";
    marker.style.left = `${position.x * 100}%`;
    marker.style.top = `${position.y * 100}%`;

    const label = document.createElement("span");
    label.textContent = labels[field] || field;

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
    `${field === "CertificateNumber" ? "Certificate number" : field} position cleared.`
  );
}

function clearAllPositions() {
  templates[currentType].positions = {};

  renderPositions();

  setStatus(
    "editorStatus",
    "All field positions cleared for this template."
  );
}

function hexToRgb(hex) {
  const clean = String(hex).replace("#", "");

  return {
    r: parseInt(clean.substring(0, 2), 16),
    g: parseInt(clean.substring(2, 4), 16),
    b: parseInt(clean.substring(4, 6), 16)
  };
}

// Determines the public verification URL used by the QR code.
function getVerificationPageUrl() {
  const entered = $("verificationUrl").value.trim();
  let base;

  if (entered) {
    try {
      base = new URL(entered);
    } catch {
      throw new Error(
        "Enter a valid public verification-page URL."
      );
    }

    if (!["http:", "https:"].includes(base.protocol)) {
      throw new Error(
        "The verification-page URL must start with http:// or https://."
      );
    }

    if (base.username || base.password) {
      throw new Error(
        "The verification URL must not contain login credentials."
      );
    }

    base.search = "";
    base.hash = "";

    return base;
  }

  if (["http:", "https:"].includes(window.location.protocol)) {
    return new URL("verify.html", window.location.href);
  }

  throw new Error(
    "Before generating PDFs on your computer, enter your published GitHub Pages verify.html URL so the QR codes work online."
  );
}

function makeQrImage(url) {
  return new Promise((resolve, reject) => {
    if (!window.QRCode) {
      reject(
        new Error(
          "QR code library failed to load. Check your internet connection and reload."
        )
      );
      return;
    }

    const holder = document.createElement("div");

    holder.style.position = "fixed";
    holder.style.left = "-10000px";
    holder.style.top = "0";

    document.body.appendChild(holder);

    try {
      new QRCode(holder, {
        text: url,
        width: 220,
        height: 220,
        correctLevel: QRCode.CorrectLevel.M
      });

      window.setTimeout(() => {
        try {
          const canvas = holder.querySelector("canvas");
          const img = holder.querySelector("img");

          if (canvas) {
            resolve(canvas.toDataURL("image/png"));
          } else if (img && img.src) {
            resolve(img.src);
          } else {
            reject(new Error("Could not create the QR image."));
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
  if (field === "CertificateNumber") {
    return person.certificateId;
  }

  if (field === "Position" && !person.Position) {
    return "";
  }

  return String(person[field] ?? "");
}

async function makePDF(person, template) {
  if (!window.jspdf) {
    throw new Error(
      "The PDF library did not load. Check your internet connection and reload."
    );
  }

  if (!template.image || !template.data) {
    throw new Error(
      `Upload the ${person.type} template before generating certificates.`
    );
  }

  const { jsPDF } = window.jspdf;
  const image = template.image;

  const orientation =
    image.naturalWidth >= image.naturalHeight
      ? "landscape"
      : "portrait";

  const pdf = new jsPDF({
    orientation,
    unit: "pt",
    format: "a4"
  });

  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();

  pdf.addImage(
    template.data,
    template.data.startsWith("data:image/png") ? "PNG" : "JPEG",
    0,
    0,
    pageW,
    pageH
  );

  // Print all configured fields, including the certificate number.
  Object.entries(template.positions).forEach(([field, position]) => {
    const value = fieldValue(person, field);

    if (value === "") return;

    const rgb = hexToRgb(position.color);

    pdf.setTextColor(rgb.r, rgb.g, rgb.b);
    pdf.setFont(
      position.font || "helvetica",
      position.bold ? "bold" : "normal"
    );
    pdf.setFontSize(position.fontSize || 24);

    pdf.text(
      String(value),
      position.x * pageW,
      position.y * pageH,
      { align: "center" }
    );
  });

  // Add QR code with this certificate's ID already in the URL.
  const verificationUrl = getVerificationPageUrl();
  verificationUrl.searchParams.set("id", person.certificateId);

  const qrImage = await makeQrImage(verificationUrl.href);

  const qrSize = 58;
  const margin = 20;
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

  const person = participants[0];
  const template = templates[person.type];

  try {
    setStatus("generateStatus", "Creating preview...");

    const blob = await makePDF(person, template);
    const url = URL.createObjectURL(blob);

    $("pdfPreview").innerHTML = "";

    const iframe = document.createElement("iframe");
    iframe.title = "Certificate preview";
    iframe.src = url;

    $("pdfPreview").appendChild(iframe);

    setStatus("generateStatus", "Preview ready.");

  } catch (error) {
    setStatus(
      "generateStatus",
      error.message || "Could not create preview.",
      true
    );
  }
}

// Email addresses are intentionally excluded from the public register.
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

  if (!window.jspdf || !window.QRCode || !window.JSZip) {
    setStatus(
      "generateStatus",
      "A required PDF, QR, or ZIP library did not load. Check your internet connection and reload.",
      true
    );
    return;
  }

  // Validate QR destination before starting generation.
  try {
    getVerificationPageUrl();
  } catch (error) {
    setStatus("generateStatus", error.message, true);
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

    const register = createPublicRegister();

    generatedRegisterBlob = new Blob(
      [JSON.stringify(register, null, 2)],
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
      `Successfully generated ${generatedFiles.length} PDF(s) and certificates.json. Upload certificates.json beside verify.html in your GitHub Pages repository. Email addresses are excluded from the public register.`
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

  if (!window.JSZip) {
    setStatus(
      "generateStatus",
      "The ZIP library did not load.",
      true
    );
    return;
  }

  try {
    const zip = new JSZip();

    generatedFiles.forEach(file => {
      zip.file(file.name, file.blob);
    });

    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = "Certificates.zip";
    link.click();

    window.setTimeout(
      () => URL.revokeObjectURL(url),
      1000
    );

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
      "Generate the PDFs before downloading the register.",
      true
    );
    return;
  }

  const url = URL.createObjectURL(generatedRegisterBlob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "certificates.json";
  link.click();

  window.setTimeout(
    () => URL.revokeObjectURL(url),
    1000
  );
}

// Connect buttons to their functions.
$("readBtn").addEventListener("click", readExcel);
$("templateType").addEventListener("change", updateTemplateType);
$("templateFile").addEventListener("change", uploadTemplate);

$("placeModeBtn").addEventListener("click", () => {
  placementMode = true;
  $("placeModeBtn").textContent = "Now click the template";

  setStatus(
    "editorStatus",
    `Click the template to place ${$("fieldSelect").selectedOptions[0].textContent}.`
  );
});

$("templateCanvas").addEventListener("click", placeField);
$("clearFieldBtn").addEventListener("click", clearSelectedField);
$("clearAllBtn").addEventListener("click", clearAllPositions);

$("previewBtn").addEventListener("click", previewFirstParticipant);
$("generateBtn").addEventListener("click", generatePDFs);
$("zipBtn").addEventListener("click", downloadZip);
$("registerBtn").addEventListener("click", downloadRegister);

window.addEventListener("resize", () => {
  if (templates[currentType].image) {
    drawTemplate();
  }
});
