const PRESETS = {
  google: "Google-style: clean single page, ATS-friendly, no photo. Sections: Summary (2 lines), Skills (grouped by category), Experience (STAR format with metrics), Education, Projects.",
  amazon: "Amazon Leadership Principles focused: results-oriented, heavy on quantifiable metrics. STAR-format bullets showing Ownership and Deliver Results.",
  microsoft: "Microsoft-style: technical focus. Technical Skills (languages, tools, platforms), Work Experience with impact metrics, Projects with tech stack, Education, Certifications.",
  mckinsey: "McKinsey consulting style: 1 page strictly. Bullet points with strong action verbs, quantified impact. Education at top, Experience, Leadership & Activities. No objective/summary.",
  goldman: "Goldman Sachs finance: very clean and concise. Education first with GPA. Experience with deal/project values. Skills section. Max 1 page.",
  startup: "Modern startup: brief 2-line summary, skills with proficiency. Experience focused on growth and ownership. Side projects highlighted.",
  custom: "",
  template: "",
};

let uploadedFile = null;
let uploadedTemplateFile = null;
let selectedCompany = "google";
let downloadFilename = null;

// ─── Resume PDF/DOCX upload ──────────────────────────────────────────────────

document.getElementById("resumeInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

const dropZone = document.getElementById("dropZone");
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.style.borderColor = "var(--accent)"; });
dropZone.addEventListener("dragleave", () => { dropZone.style.borderColor = ""; });
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  const name = file?.name.toLowerCase();
  if (file && (name.endsWith(".pdf") || name.endsWith(".docx"))) handleFile(file);
  else alert("Please upload a PDF or DOCX file.");
});

function handleFile(file) {
  if (file.size > 10 * 1024 * 1024) { alert("File too large. Max 10MB."); return; }
  uploadedFile = file;
  dropZone.classList.add("has-file");
  document.getElementById("dropLabel").innerHTML = `✓ <strong>${file.name}</strong>`;
  document.getElementById("format-card").classList.remove("disabled");
  document.getElementById("convertBtn").disabled = false;
  setStep(2);
}

// ─── Company template DOCX upload ───────────────────────────────────────────

const templateDropZone = document.getElementById("templateDropZone");

document.getElementById("templateInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleTemplateFile(file);
});

templateDropZone.addEventListener("dragover", (e) => { e.preventDefault(); templateDropZone.style.borderColor = "var(--accent)"; });
templateDropZone.addEventListener("dragleave", () => { templateDropZone.style.borderColor = ""; });
templateDropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (file && file.name.endsWith(".docx")) handleTemplateFile(file);
  else alert("Please upload a DOCX file.");
});

function handleTemplateFile(file) {
  if (file.size > 10 * 1024 * 1024) { alert("File too large. Max 10MB."); return; }
  uploadedTemplateFile = file;
  templateDropZone.classList.add("has-file");
  document.getElementById("templateDropLabel").innerHTML = `✓ <strong>${file.name}</strong>`;
}

// ─── Chips ───────────────────────────────────────────────────────────────────

document.querySelectorAll(".chip").forEach((chip) => {
  chip.addEventListener("click", () => {
    document.querySelectorAll(".chip").forEach((c) => c.classList.remove("active"));
    chip.classList.add("active");
    selectedCompany = chip.dataset.value;

    const ta = document.getElementById("formatText");
    const templateZone = document.getElementById("templateUploadZone");
    const hint = document.getElementById("presetHint");

    if (selectedCompany === "template") {
      ta.style.display = "none";
      templateZone.style.display = "block";
      hint.textContent = "📎 Upload a sample company resume (DOCX) — its style & structure will be copied, not its content.";
    } else if (selectedCompany === "custom") {
      ta.style.display = "block";
      templateZone.style.display = "none";
      ta.value = "";
      ta.placeholder = "Describe the exact resume format your company needs...";
      hint.textContent = "✦ Custom format — describe it below.";
    } else {
      ta.style.display = "block";
      templateZone.style.display = "none";
      ta.value = PRESETS[selectedCompany] || "";
      hint.textContent = `✓ Using ${chip.textContent.trim()} preset format.`;
    }
  });
});

document.getElementById("formatText").value = PRESETS["google"];

// ─── Convert ─────────────────────────────────────────────────────────────────

async function startConversion() {
  if (!uploadedFile) return;

  if (selectedCompany === "template" && !uploadedTemplateFile) {
    alert("Please upload a company template DOCX first.");
    return;
  }

  const formatText = document.getElementById("formatText").value.trim();
  const customFormat = selectedCompany === "custom" ? formatText : "";

  document.getElementById("result-card").style.display = "block";
  document.getElementById("loadingState").style.display = "flex";
  document.getElementById("previewArea").style.display = "none";
  document.getElementById("errorState").style.display = "none";
  document.getElementById("convertBtn").disabled = true;
  setStep(3);

  const messages = [
    "Reading your resume...",
    "Analyzing template structure...",
    "Reformatting for company style...",
    "Building DOCX...",
  ];
  let mi = 0;
  const msgEl = document.getElementById("loadingMsg");
  const msgInterval = setInterval(() => {
    mi = (mi + 1) % messages.length;
    msgEl.textContent = messages[mi];
  }, 1800);

  const formData = new FormData();
  formData.append("resume", uploadedFile);
  formData.append(
    "company",
    (selectedCompany === "custom" || selectedCompany === "template") ? "google" : selectedCompany
  );
  formData.append("custom_format", customFormat);

  if (selectedCompany === "template" && uploadedTemplateFile) {
    formData.append("company_template", uploadedTemplateFile);
  }

  try {
    const res = await fetch("/convert", { method: "POST", body: formData });
    const data = await res.json();
    clearInterval(msgInterval);

    if (!res.ok || data.error) throw new Error(data.error || "Conversion failed");

    downloadFilename = data.filename;
    // Pass template_style to renderPreview so it can follow the same order
    renderPreview(data.preview, data.template_style || null);
    document.getElementById("loadingState").style.display = "none";
    document.getElementById("previewArea").style.display = "block";
    setStepDone(3);
  } catch (err) {
    clearInterval(msgInterval);
    document.getElementById("loadingState").style.display = "none";
    document.getElementById("errorState").style.display = "block";
    document.getElementById("errorState").textContent = "Error: " + err.message;
    document.getElementById("convertBtn").disabled = false;
  }
}

// ─── Preview renderer (dynamic section order) ────────────────────────────────

function renderPreview(data, templateStyle) {
  // Name
  const nameEl = document.getElementById("previewName");
  nameEl.textContent = data.name || "Resume";
  const nameAlign = templateStyle?.name_alignment === "left" ? "left" : "center";
  nameEl.style.cssText = `display:block; text-align:${nameAlign}; font-size:32px; font-weight:700; letter-spacing:0.5px; margin-bottom:6px; width:100%;`;

  let html = "";

  // Contact
  if (data.contact) {
    const c = typeof data.contact === "string" ? { email: data.contact } : data.contact;
    const parts = [c.email, c.phone, c.address, c.linkedin, c.github].filter(Boolean);
    if (parts.length) {
      html += `<div class="preview-section" style="text-align:${nameAlign}; border-bottom: 1px solid var(--border); padding-bottom:10px; margin-bottom:14px;">
        <p style="color:var(--text-1); font-size:14px; font-weight:500; margin:0;">${parts.join("  ·  ")}</p>
      </div>`;
    }
  }

  // Section renderers
  const sectionNames = templateStyle?.section_names || {};
  const bulletPrefix = templateStyle?.bullet_prefix || "";

  function getHeading(key, fallback) {
    return sectionNames[key] || fallback;
  }

  function makeBullet(text) {
    return bulletPrefix
      ? `<li style="list-style:none; padding-left:0;">${bulletPrefix}${text}</li>`
      : `<li>${text}</li>`;
  }

  function renderSummary() {
    if (!data.summary) return "";
    return `<div class="preview-section"><h3>${getHeading("summary", "Summary")}</h3><p>${data.summary}</p></div>`;
  }

  function renderSkills() {
    if (!data.skills?.length) return "";
    const rows = data.skills.map(s =>
      `<div class="skill-row"><span class="skill-cat">${s.category}:</span> ${s.items.join(", ")}</div>`
    ).join("");
    return `<div class="preview-section"><h3>${getHeading("skills", "Skills")}</h3>${rows}</div>`;
  }

  function renderExperience() {
    if (!data.experience?.length) return "";
    const exps = data.experience.map(e => `
      <div class="exp-item">
        <div class="exp-header">${e.title} — ${e.company}</div>
        <div class="exp-meta">${e.duration || ""}</div>
        <ul style="padding-left:${bulletPrefix ? "0" : "18px"}">${(e.bullets || []).map(b => makeBullet(b)).join("")}</ul>
      </div>`).join("");
    return `<div class="preview-section"><h3>${getHeading("experience", "Experience")}</h3>${exps}</div>`;
  }

  function renderProjects() {
    if (!data.projects?.length) return "";
    const projs = data.projects.map(p => `
      <div class="proj-item">
        <div class="exp-header">${p.name} <span style="font-weight:400;font-size:12px;color:var(--text-2)">${p.tech || ""}</span></div>
        <ul style="padding-left:${bulletPrefix ? "0" : "18px"}">${(p.bullets || []).map(b => makeBullet(b)).join("")}</ul>
      </div>`).join("");
    return `<div class="preview-section"><h3>${getHeading("projects", "Projects")}</h3>${projs}</div>`;
  }

  function renderEducation() {
    if (!data.education?.length) return "";
    const edus = data.education.map(e =>
      `<div class="exp-item">
        <div class="exp-header">${e.degree} — ${e.institution}</div>
        <div class="exp-meta">${e.year || ""} ${e.gpa ? "· GPA: " + e.gpa : ""}</div>
      </div>`).join("");
    return `<div class="preview-section"><h3>${getHeading("education", "Education")}</h3>${edus}</div>`;
  }

  function renderCertifications() {
    if (!data.certifications?.length) return "";
    return `<div class="preview-section"><h3>${getHeading("certifications", "Certifications")}</h3><ul style="padding-left:${bulletPrefix ? "0" : "18px"}">${data.certifications.map(c => makeBullet(c)).join("")}</ul></div>`;
  }

  const renderMap = {
    summary: renderSummary,
    skills: renderSkills,
    experience: renderExperience,
    projects: renderProjects,
    education: renderEducation,
    certifications: renderCertifications,
  };

  // Use template section order if available, else default
  const sectionOrder = templateStyle?.section_order ||
    ["summary", "skills", "experience", "projects", "education", "certifications"];

  const rendered = new Set();

  // Render in template order
  for (const key of sectionOrder) {
    if (renderMap[key]) {
      html += renderMap[key]();
      rendered.add(key);
    }
  }

  // Render remaining sections not in template order
  for (const [key, fn] of Object.entries(renderMap)) {
    if (!rendered.has(key)) {
      html += fn();
    }
  }

  document.getElementById("previewBody").innerHTML = html;
}

function downloadFile() {
  if (!downloadFilename) return;
  window.location.href = `/download/${downloadFilename}`;
}

function resetApp() {
  uploadedFile = null;
  uploadedTemplateFile = null;
  downloadFilename = null;
  document.getElementById("resumeInput").value = "";
  document.getElementById("templateInput").value = "";
  dropZone.classList.remove("has-file");
  templateDropZone.classList.remove("has-file");
  document.getElementById("dropLabel").innerHTML = 'Drop PDF or DOCX here or <span class="link">click to browse</span>';
  document.getElementById("templateDropLabel").innerHTML = 'Drop a sample company resume (DOCX) here or <span class="link">click to browse</span>';
  document.getElementById("format-card").classList.add("disabled");
  document.getElementById("convertBtn").disabled = true;
  document.getElementById("result-card").style.display = "none";
  setStep(1);
}

function setStep(n) {
  for (let i = 1; i <= 3; i++) {
    const el = document.getElementById("s" + i);
    el.className = "step" + (i < n ? " done" : i === n ? " active" : "");
    if (i < n) el.querySelector(".step-num").textContent = "✓";
    else el.querySelector(".step-num").textContent = i;
  }
}
function setStepDone(n) {
  const el = document.getElementById("s" + n);
  el.className = "step done";
  el.querySelector(".step-num").textContent = "✓";
}