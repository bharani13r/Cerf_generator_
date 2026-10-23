"use strict";

const $ = id => document.getElementById(id);

let participants = [];
let generatedFiles = [];
let certificateRegister = [];
let currentType = "winner";
let placementMode = false;

// IMPORTANT: These are two independent template objects.
const templates = {
  winner: {
    image: null,
    data: null,
    positions: {}
  },
  participant: {
    image: null,
    data: null,
    positions: {}
  }
};

const canvas = $("templateCanvas");
const ctx = canvas.getContext("2d");

function setStatus(id, message, isError = false) {
  const element = $(id);
  element.textContent = message;
  element.classList.toggle("error", isError);
}

function selectedTemplate() {
  return templates[currentType];
}

function templateStatusId(type) {
  return type === "winner"
    ? "winnerTemplateStatus"
    : "participantTemplateStatus";
}

function cleanFileName(value) {
  return String(value || "certificate")
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 100);
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

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getValue(row, possibleNames) {
  const rowKeys = Object.keys(row);
  for (const name of possibleNames) {
    const foundKey = rowKeys.find(
      key => normalizeHeader(key) === normalizeHeader(name)
    );
    if (foundKey !== undefined && String(row[foundKey]).trim() !== "") {
      return String(row[foundKey]).trim();
    }
  }
  return "";
}


// ------------------------------
// EXCEL UPLOAD
// ------------------------------

$("readExcel").addEventListener("click", async () => {
  const file = $("excelFile").files[0];

  if (!file) {
    setStatus("excelStatus", "Please select an Excel or CSV file.", true);
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false
    });

    if (!rows.length) {
      participants = [];
      renderParticipants();
      setStatus("excelStatus", "The spreadsheet is empty.", true);
      return;
    }

    participants = rows.map((row, index) => {
      const name = getValue(row, [
        "Name", "Participant Name", "Full Name", "Student Name"
      ]);

      const email = getValue(row, [
        "Email", "Email ID", "Email Address", "Mail"
      ]);

      const typeValue = getValue(row, [
        "Type", "Category", "Certificate Type", "Award"
      ]).toLowerCase();

      const type = (
        typeValue.includes("winner") ||
        typeValue.includes("first") ||
        typeValue.includes("second") ||
        typeValue.includes("third") ||
        typeValue.includes("prize")
      ) ? "winner" : "participant";

      return {
        name: name || `Participant ${index + 1}`,
        email,
        type,
        serial: index + 1,
        original: row
      };
    });

    renderParticipants();

    setStatus(
      "excelStatus",
      `Loaded ${participants.length} participant(s).`
    );
  } catch (error) {
    console.error(error);
    setStatus(
      "excelStatus",
      "Could not read this file. Check that it is a valid Excel or CSV file.",
      true
    );
  }
});

function renderParticipants() {
  const container = $("participantsTable");

  if (!participants.length) {
    container.innerHTML = "";
    return;
  }

  const visible = participants.slice(0, 5);

  let html = `
    <p>Participants loaded: ${participants.length}</p>
    <div style="overflow-x:auto">
    <table>
      <thead>
        <tr>
          <th>Serial</th>
          <th>Name</th>
          <th>Email</th>
          <th>Template</th>
        </tr>
      </thead>
      <tbody>
  `;

  for (const person of visible) {
    html += `
      <tr>
        <td>${person.serial}</td>
        <td>${escapeHtml(person.name)}</td>
        <td>${escapeHtml(person.email || "Not provided")}</td>
        <td>${escapeHtml(person.type)}</td>
      </tr>
    `;
  }

  html += "</tbody></table></div>";

  if (participants.length > 5) {
    html += "<p class='status'>Only the first 5 participants will be processed.</p>";
  }

  container.innerHTML = html;
}


// ------------------------------
// SEPARATE TEMPLATE UPLOADS
// ------------------------------

function uploadTemplate(type, inputId) {
  const input = $(inputId);
  const file = input.files[0];
  const statusId = templateStatusId(type);

  if (!file) return;

  if (!["image/png", "image/jpeg"].includes(file.type)) {
    setStatus(statusId, "Choose a PNG, JPG, or JPEG image.", true);
    input.value = "";
    return;
  }

  const reader = new FileReader();

  reader.onerror = () => {
    setStatus(statusId, "Could not read the selected image.", true);
  };

  reader.onload = () => {
    const image = new Image();

    image.onerror = () => {
      setStatus(statusId, "This image could not be opened.", true);
    };

    image.onload = () => {
      // The uploaded image is stored ONLY in templates[type].
      // Uploading a participant image cannot change templates.winner.
      templates[type].image = image;
      templates[type].data = reader.result;

      if ($("templateType").value === type) {
        currentType = type;
        drawTemplate();
        renderPositions();
      }

      setStatus(statusId, `${type} certificate uploaded successfully.`);
      setStatus("templateStatus", `Loaded ${type} certificate.`);
    };

    image.src = reader.result;
  };

  reader.readAsDataURL(file);
}

$("winnerTemplateFile").addEventListener("change", () => {
  uploadTemplate("winner", "winnerTemplateFile");
});

$("participantTemplateFile").addEventListener("change", () => {
  uploadTemplate("participant", "participantTemplateFile");
});


// ------------------------------
// TEMPLATE PREVIEW
// ------------------------------

function updateTemplateType() {
  currentType = $("templateType").value;
  placementMode = false;

  $("placeModeBtn").textContent = "Click template to place field";

  drawTemplate();
  renderPositions();

  const template = selectedTemplate();

  setStatus(
    "templateStatus",
    template.image
      ? `Editing ${currentType} template.`
      : `Upload the ${currentType} certificate image first.`
  );
}

$("templateType").addEventListener("change", updateTemplateType);

function drawTemplate() {
  const template = selectedTemplate();
  const emptyMessage = $("emptyMessage");

  if (!template.image) {
    canvas.style.display = "none";
    emptyMessage.style.display = "block";
    $("markers").innerHTML = "";
    return;
  }

  canvas.width = template.image.naturalWidth;
  canvas.height = template.image.naturalHeight;
  canvas.style.display = "block";
  canvas.style.maxWidth = "100%";
  canvas.style.height = "auto";
  emptyMessage.style.display = "none";

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(template.image, 0, 0);

  renderPositions();
}

function renderPositions() {
  const template = selectedTemplate();
  const list = $("positionList");
  const markers = $("markers");

  list.innerHTML = "";
  markers.innerHTML = "";

  const names = Object.keys(template.positions);

  if (!names.length) {
    list.textContent = "No fields positioned yet.";
    return;
  }

  names.forEach(field => {
    const position = template.positions[field];

    const item = document.createElement("div");
    item.textContent =
      `${field}: X ${Math.round(position.x * 100)}%, ` +
      `Y ${Math.round(position.y * 100)}%`;
    list.appendChild(item);

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

$("placeModeBtn").addEventListener("click", () => {
  if (!selectedTemplate().image) {
    setStatus(
      "templateStatus",
      `Upload the ${currentType} template before placing fields.`,
      true
    );
    return;
  }

  placementMode = !placementMode;

  $("placeModeBtn").textContent = placementMode
    ? "Now click the certificate"
    : "Click template to place field";

  setStatus(
    "templateStatus",
    placementMode
      ? `Click the preview to position ${$("fieldType").value}.`
      : "Field placement cancelled."
  );
});

canvas.addEventListener("click", event => {
  if (!placementMode) return;

  const rect = canvas.getBoundingClientRect();

  if (!rect.width || !rect.height) return;

  const x = Math.max(
    0,
    Math.min(1, (event.clientX - rect.left) / rect.width)
  );

  const y = Math.max(
    0,
    Math.min(1, (event.clientY - rect.top) / rect.height)
  );

  const field = $("fieldType").value;

  // Positions are stored independently for the selected template.
  selectedTemplate().positions[field] = { x, y };

  placementMode = false;
  $("placeModeBtn").textContent = "Click template to place field";

  renderPositions();

  setStatus(
    "templateStatus",
    `${field} position saved for the ${currentType} template.`
  );
});

$("clearPositionsBtn").addEventListener("click", () => {
  selectedTemplate().positions = {};
  renderPositions();
  setStatus(
    "templateStatus",
    `Positions cleared for the ${currentType} template only.`
  );
});


// ------------------------------
// PDF GENERATION HELPERS
// ------------------------------

function getPdfFont(fontName) {
  const fonts = PDFLib.StandardFonts;

  const mapping = {
    Helvetica: fonts.Helvetica,
    "Helvetica-Bold": fonts.HelveticaBold,
    TimesRoman: fonts.TimesRoman,
    "TimesRoman-Bold": fonts.TimesRomanBold,
    Courier: fonts.Courier
  };

  return mapping[fontName] || fonts.Helvetica;
}

function getFieldText(field, person, certificateId, eventName, eventDate) {
  switch (field) {
    case "name":
      return person.name;
    case "event":
      return eventName;
    case "date":
      return eventDate;
    case "certificateId":
      return certificateId;
    case "serial":
      return String(person.serial);
    default:
      return "";
  }
}

function hexToRgb(hex) {
  const clean = String(hex || "#000000").replace("#", "");
  const valid = /^[0-9a-fA-F]{6}$/.test(clean)
    ? clean
    : "000000";

  return {
    r: parseInt(valid.slice(0, 2), 16) / 255,
    g: parseInt(valid.slice(2, 4), 16) / 255,
    b: parseInt(valid.slice(4, 6), 16) / 255
  };
}

function makeCertificateId(person, index) {
  const prefix = person.type === "winner" ? "WIN" : "PAR";
  const serial = String(index + 1).padStart(3, "0");
  const random = Math.random().toString(36).slice(2, 7).toUpperCase();

  return `${prefix}-${Date.now().toString(36).toUpperCase()}-${serial}-${random}`;
}

async function makePDF(person, template, certificateId, eventName, eventDate) {
  if (!template.image || !template.data) {
    throw new Error(`The ${person.type} certificate template is missing.`);
  }

  const pdfDoc = await PDFLib.PDFDocument.create();

  const bytes = await fetch(template.data).then(response => response.arrayBuffer())
    .catch(() => null);

  // Decode the data URL directly, avoiding a network request.
  const base64 = template.data.split(",")[1];
  const binary = atob(base64);
  const imageBytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    imageBytes[i] = binary.charCodeAt(i);
  }

  let embeddedImage;

  if (template.data.startsWith("data:image/png")) {
    embeddedImage = await pdfDoc.embedPng(imageBytes);
  } else {
    embeddedImage = await pdfDoc.embedJpg(imageBytes);
  }

  const pageWidth = embeddedImage.width;
  const pageHeight = embeddedImage.height;

  const page = pdfDoc.addPage([pageWidth, pageHeight]);

  page.drawImage(embeddedImage, {
    x: 0,
    y: 0,
    width: pageWidth,
    height: pageHeight
  });

  const fontName = $("fontFamily").value;
  const font = await pdfDoc.embedFont(getPdfFont(fontName));
  const fontSize = Math.max(
    6,
    Math.min(100, Number($("fontSize").value) || 24)
  );

  const color = hexToRgb($("fontColor").value);
  const positions = template.positions;

  for (const field of Object.keys(positions)) {
    const position = positions[field];
    const text = getFieldText(
      field, person, certificateId, eventName, eventDate
    );

    if (!text) continue;

    const x = position.x * pageWidth;
    const y = pageHeight - position.y * pageHeight;
    const textWidth = font.widthOfTextAtSize(text, fontSize);

    const centered = $("centerText").checked;
    const drawX = centered ? x - textWidth / 2 : x;

    page.drawText(text, {
      x: Math.max(0, drawX),
      y: Math.max(0, Math.min(pageHeight - fontSize, y - fontSize / 2)),
      size: fontSize,
      font,
      color: PDFLib.rgb(color.r, color.g, color.b),
      maxWidth: pageWidth
    });
  }

  // Generate a QR code pointing to the verification page.
  const verifyBaseUrl = $("verifyBaseUrl").value.trim();

  if (verifyBaseUrl) {
    try {
      const separator = verifyBaseUrl.includes("?") ? "&" : "?";
      const verifyUrl =
        `${verifyBaseUrl}${separator}id=${encodeURIComponent(certificateId)}`;

      const qrDataUrl = await QRCode.toDataURL(verifyUrl, {
        width: 180,
        margin: 1
      });

      const qrBase64 = qrDataUrl.split(",")[1];
      const qrBinary = atob(qrBase64);
      const qrBytes = new Uint8Array(qrBinary.length);

      for (let i = 0; i < qrBinary.length; i++) {
        qrBytes[i] = qrBinary.charCodeAt(i);
      }

      const qrImage = await pdfDoc.embedPng(qrBytes);
      const qrSize = Math.min(70, pageWidth * 0.12, pageHeight * 0.18);

      page.drawImage(qrImage, {
        x: Math.max(0, pageWidth - qrSize - 15),
        y: 15,
        width: qrSize,
        height: qrSize
      });
    } catch (error) {
      console.warn("QR code could not be added:", error);
    }
  }

  const pdfBytes = await pdfDoc.save();

  return new Blob([pdfBytes], { type: "application/pdf" });
}


// ------------------------------
// GENERATE CERTIFICATES
// ------------------------------

$("generateBtn").addEventListener("click", async () => {
  if (!participants.length) {
    setStatus("generateStatus", "Upload and read an Excel file first.", true);
    return;
  }

  const people = participants.slice(0, 5);

  const missing = [...new Set(
    people
      .filter(person => !templates[person.type].image)
      .map(person => person.type)
  )];

  if (missing.length) {
    setStatus(
      "generateStatus",
      `Upload these required templates first: ${missing.join(", ")}.`,
      true
    );
    return;
  }

  const eventName = $("eventName").value.trim();
  const eventDate = $("eventDate").value.trim();

  if (!eventName || !eventDate) {
    setStatus(
      "generateStatus",
      "Enter both the event name and event date.",
      true
    );
    return;
  }

  const button = $("generateBtn");
  button.disabled = true;
  $("downloadArea").innerHTML = "";
  generatedFiles = [];
  certificateRegister = [];

  try {
    for (let i = 0; i < people.length; i++) {
      const person = people[i];

      // Choose the correct, independent template for this participant.
      const template = templates[person.type];
      const certificateId = makeCertificateId(person, i);

      setStatus(
        "generateStatus",
        `Generating certificate ${i + 1} of ${people.length}...`
      );

      const blob = await makePDF(
        person,
        template,
        certificateId,
        eventName,
        eventDate
      );

      const filename =
        `${String(person.serial).padStart(3, "0")}_` +
        `${cleanFileName(person.name)}.pdf`;

      generatedFiles.push({ filename, blob });

      certificateRegister.push({
        certificateId,
        name: person.name,
        email: person.email,
        type: person.type,
        serial: person.serial,
        event: eventName,
        date: eventDate,
        filename
      });
    }

    showDownloads();

    $("downloadRegisterBtn").disabled = false;

    setStatus(
      "generateStatus",
      `Successfully generated ${generatedFiles.length} certificate(s).`
    );
  } catch (error) {
    console.error(error);
    setStatus(
      "generateStatus",
      `Generation failed: ${error.message || error}`,
      true
    );
  } finally {
    button.disabled = false;
  }
});

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

function showDownloads() {
  const area = $("downloadArea");
  area.innerHTML = "";

  generatedFiles.forEach(file => {
    const link = document.createElement("a");
    link.className = "download-link";
    link.href = URL.createObjectURL(file.blob);
    link.download = file.filename;
    link.textContent = `Download ${file.filename}`;
    area.appendChild(link);
  });

  const zipButton = document.createElement("button");
  zipButton.textContent = "Download All Certificates (ZIP)";
  zipButton.addEventListener("click", downloadZip);
  area.appendChild(zipButton);
}

async function downloadZip() {
  if (!generatedFiles.length) return;

  try {
    const zip = new JSZip();

    generatedFiles.forEach(file => {
      zip.file(file.filename, file.blob);
    });

    const blob = await zip.generateAsync({ type: "blob" });
    downloadBlob(blob, "certificates.zip");
  } catch (error) {
    console.error(error);
    setStatus("generateStatus", "Could not create ZIP file.", true);
  }
}

$("downloadRegisterBtn").addEventListener("click", () => {
  const data = {
    certificates: certificateRegister
  };

  const blob = new Blob(
    [JSON.stringify(data, null, 2)],
    { type: "application/json" }
  );

  downloadBlob(blob, "certificates.json");
});


// Initial screen
updateTemplateType();
