
const $ = id => document.getElementById(id);

const templates = {
  winner: { image: null, data: null, positions: {} },
  participant: { image: null, data: null, positions: {} }
};

let participants = [];
let generatedFiles = [];
let certificateRegister = [];
let currentType = "winner";
let placementMode = false;

$("readBtn").addEventListener("click", readExcel);
$("templateType").addEventListener("change", switchTemplate);
$("templateFile").addEventListener("change", loadTemplate);
$("placeModeBtn").addEventListener("click", beginPlacement);
$("templateCanvas").addEventListener("click", placeField);
$("fieldSelect").addEventListener("change", loadFieldStyle);
$("fontSize").addEventListener("change", saveFieldStyle);
$("fontSelect").addEventListener("change", saveFieldStyle);
$("boldText").addEventListener("change", saveFieldStyle);
$("textColor").addEventListener("change", saveFieldStyle);
$("clearFieldBtn").addEventListener("click", clearField);
$("clearAllBtn").addEventListener("click", clearAll);
$("previewBtn").addEventListener("click", previewFirst);
$("generateBtn").addEventListener("click", generatePDFs);
$("zipBtn").addEventListener("click", downloadZip);
$("registerBtn").addEventListener("click", downloadRegister);
window.addEventListener("resize", renderMarkers);

function status(id, message, error = false) {
  $(id).textContent = message;
  $(id).classList.toggle("error", error);
}

function currentTemplate() {
  return templates[currentType];
}

function switchTemplate() {
  currentType = $("templateType").value;
  placementMode = false;
  $("templateFile").value = "";
  const t = currentTemplate();

  if (t.image) {
    drawTemplate();
  } else {
    $("templateCanvas").style.display = "none";
    $("emptyMessage").style.display = "block";
    $("markers").replaceChildren();
  }

  status("templateStatus", `Selected ${currentType} template.`);
  renderPositionList();
  loadFieldStyle();
}

function normalizeHeader(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function getColumn(headers, aliases) {
  return headers.find(header =>
    aliases.includes(normalizeHeader(header))
  );
}

// Positions 1, 2, and 3 are winners.
// All other positions are participants.
function isWinner(position) {
  const p = String(position ?? "").trim().toLowerCase();

  if (/^(1|1st|first)(\s*(place|prize|position))?$/.test(p)) return true;
  if (/^(2|2nd|second)(\s*(place|prize|position))?$/.test(p)) return true;
  if (/^(3|3rd|third)(\s*(place|prize|position))?$/.test(p)) return true;

  // Also accept labels such as "Winner", "First Prize", "2nd Place".
  if (/\b(winner|1st|2nd|3rd|first|second|third)\b/.test(p)) return true;

  // Everything else, including "Participant", uses the participant template.
  return false;
}

function formatDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return `${String(value.getDate()).padStart(2, "0")}-${String(value.getMonth() + 1).padStart(2, "0")}-${value.getFullYear()}`;
  }
  return String(value ?? "").trim();
}

function makeCertificateNumber(serialNo) {
  // Keep the spreadsheet serial number and add a random suffix.
  const serial = String(serialNo).trim().replace(/\s+/g, "");
  const random = Math.random().toString(36).slice(2, 8).toUpperCase();
  return `CERT-${serial}-${random}`;
}

async function readExcel() {
  participants = [];
  generatedFiles = [];
  certificateRegister = [];

  $("participantsTable").replaceChildren();
  $("downloads").replaceChildren();
  $("pdfPreview").replaceChildren();
  $("zipBtn").disabled = true;
  $("registerBtn").disabled = true;

  const file = $("excelFile").files[0];
  if (!file) {
    return status("excelStatus", "Choose an Excel file first.", true);
  }

  try {
    const wb = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: true
    });

    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) throw new Error("The spreadsheet is empty.");
    if (rows.length > 5) throw new Error("Maximum 5 participants allowed.");

    const headers = Object.keys(rows[0]);

    const column = {
      serial: getColumn(headers, ["serialno", "serialnumber", "serial", "sno", "slno"]),
      name: getColumn(headers, ["name", "participantname", "fullname"]),
      event: getColumn(headers, ["event", "eventname"]),
      date: getColumn(headers, ["date", "eventdate"]),
      email: getColumn(headers, ["email", "emailid", "emailaddress"]),
      position: getColumn(headers, ["position", "rank", "place", "prize"])
    };

    const missing = Object.entries(column)
      .filter(([, value]) => !value)
      .map(([key]) => key);

    if (missing.length) {
      throw new Error(
        "Missing required columns: " + missing.join(", ") +
        ". Required: Serial No., Name, Event, Date, Email, Position."
      );
    }

    participants = rows.map((row, index) => {
      const serialNo = String(row[column.serial] ?? "").trim();
      const name = String(row[column.name] ?? "").trim();
      const event = String(row[column.event] ?? "").trim();
      const date = formatDate(row[column.date]);
      const email = String(row[column.email] ?? "").trim();
      const position = String(row[column.position] ?? "").trim();

      if (!serialNo || !name || !event || !date || !email || !position) {
        throw new Error(`Please fill all six columns in spreadsheet row ${index + 2}.`);
      }

      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        throw new Error(`Invalid email in spreadsheet row ${index + 2}.`);
      }

      const type = isWinner(position) ? "winner" : "participant";

      return {
        serialNo,
        name,
        event,
        date,
        email,
        position,
        type,
        certificateNumber: makeCertificateNumber(serialNo)
      };
    });

    renderParticipants();
    status("excelStatus", `Loaded ${participants.length} participant(s). Winner/participant types assigned from Position.`);
  } catch (error) {
    participants = [];
    status("excelStatus", error.message, true);
  }
}

function renderParticipants() {
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  ["Serial No.", "Name", "Event", "Date", "Email", "Position", "Template", "Certificate Number"]
    .forEach(label => {
      const th = document.createElement("th");
      th.textContent = label;
      headerRow.appendChild(th);
    });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  participants.forEach(person => {
    const tr = document.createElement("tr");
    [
      person.serialNo,
      person.name,
      person.event,
      person.date,
      person.email,
      person.position,
      person.type,
      person.certificateNumber
    ].forEach(value => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  $("participantsTable").appendChild(table);
}

function loadTemplate() {
  const file = $("templateFile").files[0];
  if (!file) return;

  if (!["image/jpeg", "image/png"].includes(file.type)) {
    return status("templateStatus", "Upload a JPG, JPEG, or PNG image.", true);
  }

  const targetType = currentType;
  const reader = new FileReader();

  reader.onload = () => {
    const img = new Image();

    img.onload = () => {
      // Save the image to the selected template only.
      templates[targetType].image = img;
      templates[targetType].data = reader.result;

      if (currentType === targetType) drawTemplate();

      status("templateStatus", `Loaded ${targetType} template (${img.naturalWidth} × ${img.naturalHeight}).`);
    };

    img.onerror = () => status("templateStatus", "Unable to load this image.", true);
    img.src = reader.result;
  };

  reader.readAsDataURL(file);
}

function drawTemplate() {
  const img = currentTemplate().image;
  if (!img) return;

  const canvas = $("templateCanvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  canvas.getContext("2d").drawImage(img, 0, 0);
  canvas.style.display = "block";
  $("emptyMessage").style.display = "none";
  renderMarkers();
}

function beginPlacement() {
  if (!currentTemplate().image) {
    return status("editorStatus", "Upload a template image first.", true);
  }
  placementMode = true;
  status("editorStatus", `Now click where ${$("fieldSelect").value} should be placed.`);
}

function placeField(event) {
  if (!placementMode) return;

  const canvas = $("templateCanvas");
  const rect = canvas.getBoundingClientRect();
  const field = $("fieldSelect").value;

  currentTemplate().positions[field] = {
    x: (event.clientX - rect.left) / rect.width,
    y: (event.clientY - rect.top) / rect.height,
    size: fontSize(),
    font: $("fontSelect").value,
    bold: $("boldText").checked,
    color: $("textColor").value
  };

  placementMode = false;
  status("editorStatus", `${field} position saved for ${currentType} template.`);
  renderMarkers();
  renderPositionList();
}

function fontSize() {
  return Math.max(6, Math.min(72, Number($("fontSize").value) || 24));
}

function saveFieldStyle() {
  const p = currentTemplate().positions[$("fieldSelect").value];
  if (!p) return;

  p.size = fontSize();
  p.font = $("fontSelect").value;
  p.bold = $("boldText").checked;
  p.color = $("textColor").value;

  renderPositionList();
}

function loadFieldStyle() {
  const p = currentTemplate().positions[$("fieldSelect").value];
  if (!p) return;

  $("fontSize").value = p.size;
  $("fontSelect").value = p.font;
  $("boldText").checked = p.bold;
  $("textColor").value = p.color;
}

function clearField() {
  delete currentTemplate().positions[$("fieldSelect").value];
  renderMarkers();
  renderPositionList();
  status("editorStatus", "Selected field position cleared.");
}

function clearAll() {
  currentTemplate().positions = {};
  renderMarkers();
  renderPositionList();
  status("editorStatus", `All positions cleared for ${currentType} only.`);
}

function renderMarkers() {
  const holder = $("markers");
  holder.replaceChildren();

  Object.entries(currentTemplate().positions).forEach(([field, p]) => {
    const marker = document.createElement("div");
    marker.className = "marker";
    marker.style.left = `${p.x * 100}%`;
    marker.style.top = `${p.y * 100}%`;

    const label = document.createElement("span");
    label.textContent = field;
    marker.appendChild(label);
    holder.appendChild(marker);
  });
}

function renderPositionList() {
  const holder = $("positionList");
  holder.replaceChildren();

  const title = document.createElement("strong");
  title.textContent = `${currentType} positions`;
  holder.appendChild(title);

  ["Name", "Event", "Date", "Position", "CertificateNumber"].forEach(field => {
    const line = document.createElement("div");
    const p = currentTemplate().positions[field];

    line.textContent = p
      ? `${field}: placed · ${p.size} pt${p.bold ? " · bold" : ""}`
      : `${field}: not placed`;

    holder.appendChild(line);
  });
}

function getVerifyUrl(certificateNumber) {
  const url = new URL("verify.html", window.location.href);
  url.searchParams.set("id", certificateNumber);
  return url.href;
}

async function makePDF(person, template) {
  if (!template.image || !template.data) {
    throw new Error(`Upload the ${person.type} template first.`);
  }

  const { jsPDF } = window.jspdf;
  const img = template.image;

  const pdf = new jsPDF({
    orientation: img.naturalWidth >= img.naturalHeight ? "landscape" : "portrait",
    unit: "pt",
    format: [img.naturalWidth, img.naturalHeight]
  });

  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();

  pdf.addImage(
    template.data,
    template.data.startsWith("data:image/png") ? "PNG" : "JPEG",
    0, 0, pageW, pageH
  );

  const values = {
    Name: person.name,
    Event: person.event,
    Date: person.date,
    Position: person.position,
    CertificateNumber: person.certificateNumber
  };

  Object.entries(template.positions).forEach(([field, pos]) => {
    if (values[field] === undefined) return;

    pdf.setFont(pos.font, pos.bold ? "bold" : "normal");
    pdf.setFontSize(pos.size);
    pdf.setTextColor(pos.color);

    const x = pos.x * pageW;
    const y = pos.y * pageH + pos.size * 0.35;
    pdf.text(String(values[field]), x, y);
  });

  // Add a QR code in the bottom-right corner of the certificate.
  const qrUrl = getVerifyUrl(person.certificateNumber);
  const qrDataUrl = await QRCode.toDataURL(qrUrl, {
    width: 220,
    margin: 1,
    errorCorrectionLevel: "M"
  });

  const qrSize = Math.min(75, pageW * 0.14, pageH * 0.18);
  pdf.addImage(qrDataUrl, "PNG", pageW - qrSize - 15, pageH - qrSize - 15, qrSize, qrSize);

  return pdf.output("blob");
}

async function previewFirst() {
  if (!participants.length) {
    return status("generateStatus", "Read the spreadsheet first.", true);
  }

  const person = participants[0];
  const template = templates[person.type];

  try {
    const blob = await makePDF(person, template);
    const frame = document.createElement("iframe");
    frame.src = URL.createObjectURL(blob);
    $("pdfPreview").replaceChildren(frame);
    status("generateStatus", `Previewing ${person.name} (${person.type}).`);
  } catch (error) {
    status("generateStatus", error.message, true);
  }
}

async function generatePDFs() {
  if (!participants.length) {
    return status("generateStatus", "Read the spreadsheet first.", true);
  }

  generatedFiles = [];
  certificateRegister = [];
  $("downloads").replaceChildren();
  $("zipBtn").disabled = true;
  $("registerBtn").disabled = true;

  try {
    for (const person of participants) {
      const template = templates[person.type];

      if (!template.image) {
        throw new Error(`Upload the ${person.type} template before generating certificates.`);
      }

      const required = ["Name", "Event", "Date", "Position", "CertificateNumber"];
      const missing = required.filter(field => !template.positions[field]);

      if (missing.length) {
        throw new Error(
          `${person.type} template is missing positioned fields: ${missing.join(", ")}.`
        );
      }

      status("generateStatus", `Generating ${person.name}'s certificate...`);

      const blob = await makePDF(person, template);
      const safeName = person.name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
      const filename = `${person.serialNo}_${safeName}_certificate.pdf`;

      generatedFiles.push({ blob, filename, person });

      certificateRegister.push({
        certificateNumber: person.certificateNumber,
        serialNo: person.serialNo,
        name: person.name,
        event: person.event,
        date: person.date,
        email: person.email,
        position: person.position,
        certificateType: person.type,
        filename
      });
    }

    generatedFiles.forEach(file => {
      const link = document.createElement("a");
      link.className = "download-link";
      link.href = URL.createObjectURL(file.blob);
      link.download = file.filename;
      link.textContent = `${file.person.name} — ${file.person.type} certificate · Download PDF`;
      $("downloads").appendChild(link);
    });

    $("zipBtn").disabled = false;
    $("registerBtn").disabled = false;
    status("generateStatus", `Successfully generated ${generatedFiles.length} certificate(s). Download the JSON register and publish it as certificates.json.`);
  } catch (error) {
    status("generateStatus", error.message, true);
  }
}

async function downloadZip() {
  if (!generatedFiles.length) return;

  $("zipBtn").disabled = true;

  try {
    const zip = new JSZip();

    generatedFiles.forEach(file => {
      zip.file(file.filename, file.blob);
    });

    const blob = await zip.generateAsync({ type: "blob" });
    downloadBlob(blob, "certificates.zip");
  } catch (error) {
    status("generateStatus", "Could not create ZIP: " + error.message, true);
  } finally {
    $("zipBtn").disabled = false;
  }
}

function downloadRegister() {
  if (!certificateRegister.length) return;

  const data = {
    certificates: certificateRegister
  };

  const blob = new Blob(
    [JSON.stringify(data, null, 2)],
    { type: "application/json" }
  );

  downloadBlob(blob, "certificates.json");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
