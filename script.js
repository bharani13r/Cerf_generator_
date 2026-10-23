
const $ = id => document.getElementById(id);

let participants = [];
let generatedFiles = [];
let currentType = "winner";
let placementMode = false;

const templates = {
  winner: { image: null, positions: {} },
  participant: { image: null, positions: {} }
};

let currentTemplateImage = null;
let currentImageData = null;

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

function getFieldValue(person, field) {
  return String(person[field] ?? "");
}

function createCertificateId(serialNo) {
  const bytes = new Uint8Array(6);
  crypto.getRandomValues(bytes);

  const randomPart = Array.from(bytes)
    .map(byte => byte.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();

  return `CERT-${serialNo}-${randomPart}`;
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
      <td>${escapeHtml(person.Position)}</td>
      <td>${escapeHtml(person.certificateId)}</td>
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
          <th>Email</th>
          <th>Position</th>
          <th>Certificate ID</th>
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
        `Missing required columns: ${missing.join(", ")}`,
        true
      );
      return;
    }

    if (rawRows.length > 5) {
      setStatus(
        "excelStatus",
        "Maximum 5 participants are supported.",
        true
      );
      return;
    }

    const newParticipants = rawRows.map((row, index) => {
      const person = {};

      requiredColumns.forEach(column => {
        person[column] = String(
          row[headerMap[normalizeKey(column)]] ?? ""
        ).trim();
      });

      const serial = Number(person["Serial No."]);

      if (
        !person["Serial No."] ||
        !Number.isInteger(serial) ||
        serial < 1
      ) {
        throw new Error(
          `Row ${index + 2}: Serial No. must be a positive whole number.`
        );
      }

      for (const column of requiredColumns.filter(
        column => column !== "Email"
      )) {
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

      person.type = ["1", "2", "3"].includes(person.Position)
        ? "winner"
        : "participant";

      person.certificateId = createCertificateId(serial);

      return person;
    });

    const serials = newParticipants.map(
      person => person["Serial No."]
    );

    if (new Set(serials).size !== serials.length) {
      setStatus(
        "excelStatus",
        "Serial No. values must be unique.",
        true
      );
      return;
    }

    participants = newParticipants;
    generatedFiles = [];

    $("downloads").innerHTML = "";
    $("pdfPreview").innerHTML = "";
    $("zipBtn").disabled = true;
    $("registerBtn").disabled = true;

    renderParticipants();

    setStatus(
      "excelStatus",
      `${participants.length} participant(s) loaded successfully. Certificate IDs created.`
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
  currentTemplateImage = templates[currentType].image;
  currentImageData = currentTemplateImage
    ? currentTemplateImage.src
    : null;

  drawTemplate();
  renderPositions();

  setStatus(
    "templateStatus",
    currentTemplateImage
      ? `${currentType} template selected.`
      : "Upload the selected template image."
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
    return;
  }

  const reader = new FileReader();

  reader.onload = () => {
    const image = new Image();

    image.onload = () => {
      templates[currentType].image = image;
      templates[currentType].data = reader.result;

      currentTemplateImage = image;
      currentImageData = reader.result;

      drawTemplate();

      setStatus(
        "templateStatus",
        `${currentType} template uploaded successfully.`
      );
    };

    image.onerror = () => {
      setStatus("templateStatus", "Could not load the image.", true);
    };

    image.src = reader.result;
  };

  reader.readAsDataURL(file);
}

function drawTemplate() {
  const canvas = $("templateCanvas");
  const emptyMessage = $("emptyMessage");
  const wrap = $("canvasWrap");

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
  wrap.style.width = "100%";

  renderPositions();
}

function getCanvasPoint(event) {
  const canvas = $("templateCanvas");
  const rect = canvas.getBoundingClientRect();

  return {
    x: (event.clientX - rect.left) / rect.width,
    y: (event.clientY - rect.top) / rect.height
  };
}

function placeField(event) {
  if (!placementMode) return;

  const image = templates[currentType].image;

  if (!image) {
    setStatus("editorStatus", "Upload a template image first.", true);
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
    `${field} position saved for the ${currentType} template.`
  );
}

function renderPositions() {
  const positions = templates[currentType].positions;
  const list = $("positionList");
  const markers = $("markers");

  list.innerHTML = "";
  markers.innerHTML = "";

  Object.entries(positions).forEach(([field, position]) => {
    const line = document.createElement("div");
    line.textContent =
      `${field}: x ${(position.x * 100).toFixed(1)}%, ` +
      `y ${(position.y * 100).toFixed(1)}%`;
    list.appendChild(line);

    const marker = document.createElement("div");
    marker.className = "marker";
    marker.style.left = `${position.x * 100}%`;
    marker.style.top = `${position.y * 100}%`;

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
  setStatus("editorStatus", `${field} position cleared.`);
}

function clearAllPositions() {
  templates[currentType].positions = {};

  renderPositions();
  setStatus("editorStatus", "All field positions cleared.");
}

function hexToRgb(hex) {
  const clean = String(hex).replace("#", "");

  return {
    r: parseInt(clean.substring(0, 2), 16),
    g: parseInt(clean.substring(2, 4), 16),
    b: parseInt(clean.substring(4, 6), 16)
  };
}

function makeQrImage(url) {
  return new Promise((resolve, reject) => {
    if (!window.QRCode) {
      reject(new Error("QR code library failed to load."));
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

      setTimeout(() => {
        const canvas = holder.querySelector("canvas");
        const img = holder.querySelector("img");

        try {
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

async function makePDF(person, template) {
  const { jsPDF } = window.jspdf;

  const image = template.image;

  if (!image) {
    throw new Error(
      `Upload the ${person.type} template before generating certificates.`
    );
  }

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

  Object.entries(template.positions).forEach(([field, position]) => {
    const value = field === "Serial No."
      ? person["Serial No."]
      : person[field];

    if (value === undefined || value === "") return;

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

  // The QR opens the verification page; the certificate number
  // must then be entered manually on that page.
  const verificationUrl = new URL(
    "verify.html",
    window.location.href
  ).href;

  const qrImage = await makeQrImage(verificationUrl);

  const qrSize = 58;
  const margin = 20;
  const qrX = pageW - qrSize - margin;
  const qrY = pageH - qrSize - margin;

  pdf.addImage(qrImage, "PNG", qrX, qrY, qrSize, qrSize);

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
    setStatus("generateStatus", "Read the Excel spreadsheet first.", true);
    return;
  }

  const person = participants[0];
  const template = templates[person.type];

  if (!template.image) {
    setStatus(
      "generateStatus",
      `Upload the ${person.type} template first.`,
      true
    );
    return;
  }

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

async function generatePDFs() {
  if (!participants.length) {
    setStatus("generateStatus", "Read the Excel spreadsheet first.", true);
    return;
  }

  if (!window.jspdf || !window.QRCode) {
    setStatus(
      "generateStatus",
      "A required PDF or QR library did not load. Check your internet connection and reload.",
      true
    );
    return;
  }

  generatedFiles = [];
  $("downloads").innerHTML = "";
  $("pdfPreview").innerHTML = "";
  $("zipBtn").disabled = true;
  $("registerBtn").disabled = true;

  $("generateBtn").disabled = true;

  try {
    for (let i = 0; i < participants.length; i++) {
      const person = participants[i];
      const template = templates[person.type];

      setStatus(
        "generateStatus",
        `Generating certificate ${i + 1} of ${participants.length}...`
      );

      const blob = await makePDF(person, template);

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

    $("zipBtn").disabled = false;
    $("registerBtn").disabled = false;

    setStatus(
      "generateStatus",
      `Successfully generated ${generatedFiles.length} PDF(s). Download the certificate register and publish it to GitHub to enable verification.`
    );
  } catch (error) {
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

  URL.revokeObjectURL(url);
}

function downloadRegister() {
  if (!participants.length || !generatedFiles.length) {
    setStatus(
      "generateStatus",
      "Generate the PDFs before downloading the register.",
      true
    );
    return;
  }

  const register = {
    certificates: participants.map(person => ({
      id: person.certificateId,
      serialNo: person["Serial No."],
      name: person.Name,
      event: person.Event,
      date: person.Date,
      position: person.Position
    }))
  };

  const blob = new Blob(
    [JSON.stringify(register, null, 2)],
    { type: "application/json" }
  );

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "certificates.json";
  link.click();

  URL.revokeObjectURL(url);
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
    `Click the template to place ${$("fieldSelect").value}.`
  );
});

$("templateCanvas").addEventListener("click", placeField);
$("clearFieldBtn").addEventListener("click", clearSelectedField);
$("clearAllBtn").addEventListener("click", clearAllPositions);

$("previewBtn").addEventListener("click", previewFirstParticipant);
$("generateBtn").addEventListener("click", generatePDFs);
$("zipBtn").addEventListener("click", downloadZip);
$("registerBtn").addEventListener("click", downloadRegister);
