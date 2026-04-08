import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import * as mammoth from "mammoth";

const SA_HOLIDAYS = [
  "2025-01-01",
  "2025-03-21",
  "2025-04-18",
  "2025-04-21",
  "2025-04-27",
  "2025-05-01",
  "2025-06-16",
  "2025-08-09",
  "2025-09-24",
  "2025-12-16",
  "2025-12-25",
  "2025-12-26",
  "2026-01-01",
  "2026-03-21",
  "2026-03-27",
  "2026-04-17",
  "2026-04-27",
  "2026-05-01",
  "2026-06-16",
  "2026-08-09",
  "2026-09-24",
  "2026-12-16",
  "2026-12-25",
  "2026-12-26",
];

const ORANGE = "#FF6B35";
const BLUE = "#004E89";
const LIGHT = "#EEF4FB";
const DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTH_NAMES = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
const WD = ["Mon", "Tue", "Wed", "Thu", "Fri"];

const REFLECT_QUESTIONS = [
  { key: "q1", num: 1, q: "What went well this cycle?", type: "text" },
  { key: "q2", num: 2, q: "What did not go well this cycle?", type: "text" },
  {
    key: "q3",
    num: 3,
    q: "How can this be improved next cycle?",
    type: "text",
  },
  {
    key: "q4",
    num: 4,
    q: "Did you cover all the work for the cycle?",
    type: "select",
    opts: ["", "Yes", "No", "Partially"],
  },
  {
    key: "q5",
    num: 5,
    q: "If not, how will you get back on track?",
    type: "text",
  },
  {
    key: "q6",
    num: 6,
    q: "Do you need to support some learners?",
    type: "select",
    opts: ["", "Yes", "No"],
  },
  {
    key: "q7",
    num: 7,
    q: "How will you do this?",
    type: "select",
    opts: [
      "",
      "One-on-one support",
      "Group remediation",
      "Peer tutoring",
      "After-school support",
      "Worksheet intervention",
      "Combination of strategies",
      "Not applicable",
      "Other",
    ],
  },
];

function fmt(d) {
  return (
    DAY_NAMES[d.getDay()] +
    " " +
    d.getDate() +
    " " +
    MONTH_NAMES[d.getMonth()] +
    " " +
    d.getFullYear()
  );
}
function toISO(d) {
  return d.toISOString().split("T")[0];
}
function addDays(d, n) {
  var r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}
function mon0(d) {
  return d.getDay() === 0 ? 6 : d.getDay() - 1;
}
function isSchoolDay(d, extra) {
  if (d.getDay() === 0 || d.getDay() === 6) return false;
  var iso = toISO(d);
  return !SA_HOLIDAYS.includes(iso) && !(extra || []).includes(iso);
}

function assignDates(topics, startDate, lessonDays, extraHols) {
  var sorted = lessonDays.slice().sort();
  var cur = new Date(startDate);
  cur = addDays(cur, -mon0(cur));
  var dates = [];
  var safety = 0;
  while (dates.length < topics.length && safety < 500) {
    sorted.forEach(function (wd) {
      var d = addDays(cur, wd);
      if (isSchoolDay(d, extraHols) && dates.length < topics.length)
        dates.push(new Date(d));
    });
    cur = addDays(cur, 7);
    safety++;
  }
  return topics.map(function (t, i) {
    return {
      date: dates[i] || new Date(),
      weekLabel: t.weekLabel,
      topic: t.topic,
      done: false,
      initials: "",
    };
  });
}

function parseSheetRows(sheetRows, termNum) {
  if (!sheetRows || sheetRows.length === 0) return null;
  var termStart = -1;
  var termEnd = sheetRows.length;
  var targetRe = new RegExp("\\bTERM\\s*" + termNum + "\\b", "i");
  var otherRe = /\bTERM\s*(\d)\b/i;
  for (var i = 0; i < sheetRows.length; i++) {
    var rowText = Object.values(sheetRows[i])
      .map(function (v) {
        return String(v || "");
      })
      .join(" ");
    if (termStart === -1 && targetRe.test(rowText)) {
      termStart = i;
      continue;
    }
    if (termStart !== -1 && i > termStart) {
      var m = otherRe.exec(rowText);
      if (m && parseInt(m[1]) !== termNum) {
        termEnd = i;
        break;
      }
    }
  }
  var working =
    termStart === -1 ? sheetRows : sheetRows.slice(termStart, termEnd);
  var weekRowIdx = -1;
  var weekCols = [];
  for (var i = 0; i < working.length; i++) {
    var found = [];
    Object.keys(working[i]).forEach(function (k) {
      if (/week\s*\d+/i.test(String(working[i][k] || "")))
        found.push({ key: k, label: String(working[i][k]).trim() });
    });
    if (found.length >= 2) {
      weekRowIdx = i;
      weekCols = found;
      break;
    }
  }
  if (weekRowIdx === -1) return null;
  var conceptsRow = null;
  for (var j = weekRowIdx + 1; j < working.length; j++) {
    var filled = weekCols.filter(function (wc) {
      return String(working[j][wc.key] || "").trim().length > 10;
    });
    if (filled.length >= Math.floor(weekCols.length * 0.5)) {
      conceptsRow = working[j];
      break;
    }
  }
  if (!conceptsRow) return null;
  var result = [];
  weekCols.forEach(function (wc) {
    var raw = String(conceptsRow[wc.key] || "").trim();
    if (raw.length > 0 && !/formal.*assess/i.test(raw))
      result.push({ weekLabel: wc.label, topic: raw });
  });
  return result.length > 0 ? result : null;
}

function parseDocxHTML(html, termNum) {
  var parser = new DOMParser();
  var doc = parser.parseFromString(html, "text/html");
  var allEls = Array.from(
    doc.body.querySelectorAll("p, h1, h2, h3, h4, table"),
  );
  var targetRe = new RegExp("\\bTERM\\s*" + termNum + "\\b", "i");
  var otherRe = /\bTERM\s*(\d)\b/i;
  var termStartIdx = -1;
  var termEndIdx = allEls.length;
  for (var i = 0; i < allEls.length; i++) {
    var txt = allEls[i].textContent || "";
    if (termStartIdx === -1 && targetRe.test(txt)) {
      termStartIdx = i;
      continue;
    }
    if (termStartIdx !== -1) {
      var m = otherRe.exec(txt);
      if (m && parseInt(m[1]) !== termNum) {
        termEndIdx = i;
        break;
      }
    }
  }
  var section =
    termStartIdx === -1 ? allEls : allEls.slice(termStartIdx, termEndIdx);
  var tables = section.filter(function (el) {
    return el.tagName === "TABLE";
  });
  if (tables.length === 0) tables = Array.from(doc.querySelectorAll("table"));
  for (var t = 0; t < tables.length; t++) {
    var rows = tables[t].querySelectorAll("tr");
    var weekRowIdx = -1;
    var weekCols = [];
    for (var r = 0; r < rows.length; r++) {
      var cells = rows[r].querySelectorAll("td, th");
      var found = [];
      cells.forEach(function (cell, ci) {
        if (/week\s*\d+/i.test(cell.textContent))
          found.push({ col: ci, label: cell.textContent.trim() });
      });
      if (found.length >= 2) {
        weekRowIdx = r;
        weekCols = found;
        break;
      }
    }
    if (weekRowIdx === -1) continue;
    var conceptsRow = null;
    for (var j = weekRowIdx + 1; j < rows.length; j++) {
      var cells2 = rows[j].querySelectorAll("td, th");
      var filled = weekCols.filter(function (wc) {
        return cells2[wc.col] && cells2[wc.col].textContent.trim().length > 10;
      });
      if (filled.length >= Math.floor(weekCols.length * 0.5)) {
        conceptsRow = cells2;
        break;
      }
    }
    if (!conceptsRow) continue;
    var result = [];
    weekCols.forEach(function (wc) {
      var cell = conceptsRow[wc.col];
      var raw = cell ? cell.textContent.trim() : "";
      if (raw.length > 0 && !/formal.*assess/i.test(raw))
        result.push({ weekLabel: wc.label, topic: raw });
    });
    if (result.length > 0) return result;
  }
  return null;
}

function xmlEsc(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function wTc(width, fill, bold, color, size, text) {
  return (
    "<w:tc><w:tcPr><w:tcW w:w='" +
    width +
    "' w:type='dxa'/><w:shd w:val='clear' w:color='auto' w:fill='" +
    fill +
    "'/></w:tcPr>" +
    "<w:p><w:r><w:rPr>" +
    (bold ? "<w:b/>" : "") +
    "<w:color w:val='" +
    color +
    "'/><w:sz w:val='" +
    size +
    "'/></w:rPr>" +
    "<w:t xml:space='preserve'>" +
    xmlEsc(text) +
    "</w:t></w:r></w:p></w:tc>"
  );
}

function buildDocxXML(info, rows, reflection) {
  var borders =
    "<w:tblBorders><w:top w:val='single' w:sz='4' w:space='0' w:color='004E89'/><w:left w:val='single' w:sz='4' w:space='0' w:color='004E89'/><w:bottom w:val='single' w:sz='4' w:space='0' w:color='004E89'/><w:right w:val='single' w:sz='4' w:space='0' w:color='004E89'/><w:insideH w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/><w:insideV w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/></w:tblBorders>";
  var hdr =
    "<w:tr><w:trPr><w:trHeight w:val='400'/></w:trPr>" +
    wTc("1200", "004E89", true, "FFFFFF", "18", "Week") +
    wTc("2000", "004E89", true, "FFFFFF", "18", "Date") +
    wTc(
      "6000",
      "004E89",
      true,
      "FFFFFF",
      "18",
      "Topic / Activity (CAPS Aligned)",
    ) +
    wTc("800", "004E89", true, "FFFFFF", "18", "Done") +
    wTc("1000", "004E89", true, "FFFFFF", "18", "Initials") +
    "</w:tr>";
  var prevWeek = "";
  var dataRows = rows
    .map(function (r, i) {
      var show = r.weekLabel !== prevWeek;
      if (show) prevWeek = r.weekLabel;
      var fill = i % 2 === 0 ? "EEF4FB" : "FFFFFF";
      return (
        "<w:tr>" +
        wTc(
          "1200",
          fill,
          show,
          show ? "004E89" : "CCCCCC",
          "18",
          show ? r.weekLabel : "",
        ) +
        wTc("2000", fill, false, "222222", "18", fmt(r.date)) +
        wTc("6000", fill, false, "222222", "18", r.topic || "") +
        wTc("800", fill, false, "222222", "18", "") +
        wTc("1000", fill, false, "222222", "18", "") +
        "</w:tr>"
      );
    })
    .join("");
  var rf = reflection || {};
  var rfRows = REFLECT_QUESTIONS.map(function (item) {
    return (
      "<w:tr>" +
      wTc("3500", "EEF4FB", true, "004E89", "18", item.num + ". " + item.q) +
      wTc("7500", "FFFFFF", false, "222222", "18", rf[item.key] || "") +
      "</w:tr>"
    );
  }).join("");
  var signCell = function (role, name) {
    return (
      "<w:tc><w:tcPr><w:tcW w:w='5500' w:type='dxa'/></w:tcPr>" +
      "<w:p><w:r><w:rPr><w:b/><w:color w:val='004E89'/><w:sz w:val='20'/></w:rPr><w:t>" +
      xmlEsc(role) +
      "</w:t></w:r></w:p>" +
      "<w:p><w:r><w:rPr><w:sz w:val='20'/></w:rPr><w:t>Name: " +
      xmlEsc(name) +
      "</w:t></w:r></w:p>" +
      "<w:p><w:r><w:rPr><w:sz w:val='20'/></w:rPr><w:t>Signature: _________________________________</w:t></w:r></w:p>" +
      "<w:p><w:r><w:rPr><w:sz w:val='20'/></w:rPr><w:t>Date: ______________________</w:t></w:r></w:p>" +
      "</w:tc>"
    );
  };
  var signBorders =
    "<w:tblBorders><w:top w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/><w:left w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/><w:bottom w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/><w:right w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/><w:insideV w:val='single' w:sz='4' w:space='0' w:color='C8D8EA'/></w:tblBorders>";
  return (
    "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><?mso-application progid='Word.Document'?>" +
    "<w:wordDocument xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:body>" +
    "<w:p><w:r><w:rPr><w:b/><w:color w:val='004E89'/><w:sz w:val='28'/></w:rPr><w:t>" +
    xmlEsc(
      info.subject +
        " | Grade " +
        info.grade +
        " | Term " +
        info.term +
        " | " +
        info.year,
    ) +
    "</w:t></w:r></w:p>" +
    "<w:p><w:r><w:rPr><w:sz w:val='20'/></w:rPr><w:t>" +
    xmlEsc(
      "School: " +
        info.schoolName +
        "   Teacher: " +
        info.teacherName +
        "   HoD: " +
        (info.supervisorName || "___"),
    ) +
    "</w:t></w:r></w:p>" +
    "<w:p><w:r><w:rPr><w:sz w:val='20'/></w:rPr><w:t>" +
    xmlEsc(
      "Periods/Week: " +
        info.periodsPerWeek +
        "   Days: " +
        info.dayLabel +
        "   Term Start: " +
        info.startDate,
    ) +
    "</w:t></w:r></w:p>" +
    "<w:p><w:r><w:t> </w:t></w:r></w:p>" +
    "<w:tbl><w:tblPr><w:tblW w:w='11000' w:type='dxa'/>" +
    borders +
    "</w:tblPr>" +
    hdr +
    dataRows +
    "</w:tbl>" +
    "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t> </w:t></w:r></w:p>" +
    "<w:p><w:r><w:rPr><w:b/><w:color w:val='004E89'/><w:sz w:val='24'/></w:rPr><w:t>" +
    xmlEsc("Cycle Reflection — Term " + info.term) +
    "</w:t></w:r></w:p>" +
    "<w:p><w:r><w:t> </w:t></w:r></w:p>" +
    "<w:tbl><w:tblPr><w:tblW w:w='11000' w:type='dxa'/>" +
    borders +
    "</w:tblPr>" +
    "<w:tr>" +
    wTc("3500", "004E89", true, "FFFFFF", "18", "Question") +
    wTc("7500", "004E89", true, "FFFFFF", "18", "Response") +
    "</w:tr>" +
    rfRows +
    "<w:tr>" +
    wTc("3500", "EEF4FB", true, "004E89", "18", "SMT Comment") +
    wTc("7500", "FFFFFF", false, "222222", "18", rf.smtComment || "") +
    "</w:tr>" +
    "<w:tr>" +
    wTc("3500", "EEF4FB", true, "004E89", "18", "SMT Name and Signature") +
    wTc("7500", "FFFFFF", false, "222222", "18", rf.smtName || "") +
    "</w:tr>" +
    "<w:tr>" +
    wTc("3500", "EEF4FB", true, "004E89", "18", "Date") +
    wTc("7500", "FFFFFF", false, "222222", "18", rf.smtDate || "") +
    "</w:tr>" +
    "</w:tbl>" +
    "<w:p><w:r><w:t> </w:t></w:r></w:p>" +
    "<w:tbl><w:tblPr><w:tblW w:w='11000' w:type='dxa'/>" +
    signBorders +
    "</w:tblPr>" +
    "<w:tr>" +
    signCell("Educator", info.teacherName) +
    signCell("HoD / Supervisor", info.supervisorName || "_______________") +
    "</w:tr>" +
    "</w:tbl></w:body></w:wordDocument>"
  );
}

function buildPrintHTML(info, rows, reflection) {
  var prevWeek = "";
  var tableRows = rows
    .map(function (r, i) {
      var show = r.weekLabel !== prevWeek;
      if (show) prevWeek = r.weekLabel;
      var bg = i % 2 === 0 ? "#EEF4FB" : "#fff";
      return (
        "<tr style='background:" +
        bg +
        "'>" +
        "<td style='padding:5px 7px;font-size:11px;border-bottom:1px solid #dde6f0;font-weight:" +
        (show ? "700" : "400") +
        ";color:" +
        (show ? "#004E89" : "#aaa") +
        "'>" +
        (show ? r.weekLabel : "") +
        "</td>" +
        "<td style='padding:5px 7px;font-size:11px;border-bottom:1px solid #dde6f0;white-space:nowrap'>" +
        fmt(r.date) +
        "</td>" +
        "<td style='padding:5px 7px;font-size:11px;border-bottom:1px solid #dde6f0'>" +
        (r.topic || "") +
        "</td>" +
        "<td style='padding:5px 7px;font-size:11px;border-bottom:1px solid #dde6f0;text-align:center'>[ ]</td>" +
        "<td style='padding:5px 7px;font-size:11px;border-bottom:1px solid #dde6f0;text-align:center'>________</td>" +
        "</tr>"
      );
    })
    .join("");
  var rf = reflection || {};
  var rfRows = REFLECT_QUESTIONS.map(function (item) {
    return (
      "<tr><td style='padding:8px;font-size:11px;border:1px solid #dde6f0;width:35%;font-weight:600;color:#004E89;vertical-align:top'>" +
      item.num +
      ". " +
      item.q +
      "</td>" +
      "<td style='padding:8px;font-size:11px;border:1px solid #dde6f0'>" +
      (rf[item.key] || "") +
      "</td></tr>"
    );
  }).join("");
  return (
    "<!DOCTYPE html><html><head><meta charset='utf-8'><title>ATP Tracker</title>" +
    "<style>body{font-family:Arial,sans-serif;margin:20px;font-size:12px}h2{color:#004E89;margin-bottom:8px}h3{color:#004E89;margin-bottom:10px}" +
    ".meta{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:14px}" +
    ".meta div{background:#EEF4FB;padding:5px 8px;border-radius:4px;font-size:11px}" +
    "table{width:100%;border-collapse:collapse;margin-bottom:20px}" +
    "th{background:#004E89;color:#fff;padding:7px 8px;text-align:left;font-size:11px}" +
    "td{vertical-align:top}" +
    ".sign{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-top:24px}" +
    ".sc{border:1px solid #ccc;border-radius:6px;padding:16px}" +
    ".ln{display:inline-block;border-bottom:1px solid #333;margin-left:6px}" +
    ".newpage{page-break-before:always}" +
    "@media print{@page{size:A4 portrait;margin:12mm}}</style></head><body>" +
    "<h2>" +
    info.subject +
    " | Grade " +
    info.grade +
    " | Term " +
    info.term +
    " | " +
    info.year +
    "</h2>" +
    "<div class='meta'><div><b>School:</b> " +
    info.schoolName +
    "</div><div><b>Teacher:</b> " +
    info.teacherName +
    "</div><div><b>HoD/Supervisor:</b> " +
    (info.supervisorName || "___") +
    "</div>" +
    "<div><b>Periods/Week:</b> " +
    info.periodsPerWeek +
    "</div><div><b>Lesson Days:</b> " +
    info.dayLabel +
    "</div><div><b>Term Start:</b> " +
    info.startDate +
    "</div></div>" +
    "<table><thead><tr><th style='width:70px'>Week</th><th style='width:140px'>Date</th><th>Topic / Activity (CAPS Aligned)</th><th style='width:55px;text-align:center'>Done</th><th style='width:80px;text-align:center'>Initials</th></tr></thead><tbody>" +
    tableRows +
    "</tbody></table>" +
    "<div class='newpage'>" +
    "<h3>Cycle Reflection — Term " +
    info.term +
    "</h3>" +
    "<table><thead><tr><th style='width:35%'>Question</th><th>Response</th></tr></thead><tbody>" +
    rfRows +
    "<tr><td style='padding:8px;font-size:11px;border:1px solid #dde6f0;font-weight:700;color:#004E89'>SMT Comment</td><td style='padding:8px;font-size:11px;border:1px solid #dde6f0;min-height:50px'>" +
    (rf.smtComment || "") +
    "</td></tr>" +
    "<tr><td style='padding:8px;font-size:11px;border:1px solid #dde6f0;font-weight:700;color:#004E89'>SMT Name and Signature</td><td style='padding:8px;font-size:11px;border:1px solid #dde6f0'>" +
    (rf.smtName || "") +
    "</td></tr>" +
    "<tr><td style='padding:8px;font-size:11px;border:1px solid #dde6f0;font-weight:700;color:#004E89'>Date</td><td style='padding:8px;font-size:11px;border:1px solid #dde6f0'>" +
    (rf.smtDate || "") +
    "</td></tr>" +
    "</tbody></table>" +
    "<div class='sign'>" +
    "<div class='sc'><b style='color:#004E89'>Educator</b><br><br>Name: <b>" +
    info.teacherName +
    "</b><br><br>Signature: <span class='ln' style='width:190px'>&nbsp;</span><br><br>Date: <span class='ln' style='width:130px'>&nbsp;</span></div>" +
    "<div class='sc'><b style='color:#004E89'>HoD / Supervisor</b><br><br>Name: <b>" +
    (info.supervisorName || "_______________") +
    "</b><br><br>Signature: <span class='ln' style='width:190px'>&nbsp;</span><br><br>Date: <span class='ln' style='width:130px'>&nbsp;</span></div>" +
    "</div></div></body></html>"
  );
}

export default function App() {
  var fileRef = useRef();
  var [step, setStep] = useState(1);
  var [method, setMethod] = useState("");
  var [rawFileData, setRawFileData] = useState(null);
  var [loading, setLoading] = useState(false);
  var [error, setError] = useState("");
  var [imagePreview, setImagePreview] = useState("");
  var [rows, setRows] = useState(null);
  var [reflection, setReflection] = useState({
    q1: "",
    q2: "",
    q3: "",
    q4: "",
    q5: "",
    q6: "",
    q7: "",
    smtComment: "",
    smtName: "",
    smtDate: "",
  });
  var setR = function (k, v) {
    setReflection(function (p) {
      return Object.assign({}, p, { [k]: v });
    });
  };

  var [schoolName, setSchoolName] = useState("");
  var [teacherName, setTeacherName] = useState("");
  var [supervisorName, setSupervisorName] = useState("");
  var [subject, setSubject] = useState("");
  var [grade, setGrade] = useState("");
  var [year, setYear] = useState(String(new Date().getFullYear()));
  var [term, setTerm] = useState("1");
  var [startDate, setStartDate] = useState("");
  var [lessonDays, setLessonDays] = useState([0, 2, 4]);
  var [extraHols, setExtraHols] = useState([]);
  var [newHol, setNewHol] = useState("");

  var toggleDay = function (i) {
    setLessonDays(function (p) {
      return p.includes(i)
        ? p.filter(function (d) {
            return d !== i;
          })
        : p.concat([i]).sort();
    });
  };

  var handleExcel = function (file) {
    setLoading(true);
    setError("");
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!json || json.length === 0) throw new Error("File appears empty.");
        setRawFileData({ type: "excel", data: json });
        setStep(3);
      } catch (err) {
        setError(err.message);
      }
      setLoading(false);
    };
    reader.onerror = function () {
      setError("Could not read file.");
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  var handleDocx = function (file) {
    setLoading(true);
    setError("");
    var reader = new FileReader();
    reader.onload = async function (e) {
      try {
        var result = await mammoth.convertToHtml({
          arrayBuffer: e.target.result,
        });
        if (!result.value || result.value.trim().length === 0)
          throw new Error("Document appears empty.");
        setRawFileData({ type: "docx", data: result.value });
        setStep(3);
      } catch (err) {
        setError("Could not read Word document: " + err.message);
      }
      setLoading(false);
    };
    reader.onerror = function () {
      setError("Could not read file.");
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  var handleImage = function (file) {
    setLoading(true);
    setError("");
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var dataUrl = e.target.result;
        setImagePreview(dataUrl);
        setRawFileData({
          type: "image",
          data: dataUrl,
          fileType: file.type || "image/jpeg",
        });
        setStep(3);
      } catch (err) {
        setError("Could not read image: " + err.message);
      }
      setLoading(false);
    };
    reader.onerror = function () {
      setError("Could not read image.");
      setLoading(false);
    };
    reader.readAsDataURL(file);
  };

  var handleFileDrop = function (file) {
    if (!file) return;
    var ext = file.name.split(".").pop().toLowerCase();
    if (ext === "xlsx" || ext === "xls" || ext === "csv") handleExcel(file);
    else if (ext === "docx" || ext === "doc") handleDocx(file);
    else setError("Please upload an Excel, CSV or Word file here.");
  };

  var buildTracker = async function () {
    setError("");
    if (!startDate) {
      setError("Please enter the term start date.");
      return;
    }
    if (lessonDays.length === 0) {
      setError("Please select at least one lesson day.");
      return;
    }
    if (!schoolName || !teacherName || !subject) {
      setError("Please fill in school name, teacher name and subject.");
      return;
    }
    if (!rawFileData) {
      setError("No file loaded. Please go back and upload your ATP.");
      return;
    }
    setLoading(true);
    var termNum = parseInt(term);
    var topics = null;
    try {
      if (rawFileData.type === "excel") {
        topics = parseSheetRows(rawFileData.data, termNum);
        if (!topics || topics.length === 0)
          throw new Error(
            "Could not find Term " +
              termNum +
              " content. Make sure the ATP has a TERM " +
              termNum +
              " heading and WEEK 1, WEEK 2 etc. as column headers.",
          );
      } else if (rawFileData.type === "docx") {
        topics = parseDocxHTML(rawFileData.data, termNum);
        if (!topics || topics.length === 0)
          throw new Error(
            "Could not find Term " +
              termNum +
              " content. Make sure the document has a TERM " +
              termNum +
              " heading with a week table below it.",
          );
      } else if (rawFileData.type === "image") {
        var base64 = rawFileData.data.split(",")[1];
        var prompt = [
          "You are a South African CAPS curriculum expert.",
          "This image shows an Annual Teaching Plan which may contain multiple terms.",
          "Extract ONLY the lesson topics for TERM " +
            termNum +
            ". Ignore all other terms.",
          "Skip any week labelled Formal Assessment or Test.",
          "Return ONLY valid JSON starting with { and ending with }. No markdown.",
          '{"weeks":[{"weekLabel":"Week 1","topic":"topic text"}]}',
        ].join("\n");
        var res = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            model: "claude-sonnet-4-20250514",
            max_tokens: 3000,
            messages: [
              {
                role: "user",
                content: [
                  {
                    type: "image",
                    source: {
                      type: "base64",
                      media_type: rawFileData.fileType,
                      data: base64,
                    },
                  },
                  { type: "text", text: prompt },
                ],
              },
            ],
          }),
        });
        if (!res.ok) throw new Error("API error " + res.status);
        var data = await res.json();
        var raw = data.content
          .map(function (b) {
            return b.text || "";
          })
          .join("")
          .trim();
        var s = raw.indexOf("{");
        var en = raw.lastIndexOf("}");
        if (s === -1 || en === -1)
          throw new Error("No JSON found in AI response.");
        var parsed = JSON.parse(raw.slice(s, en + 1));
        topics = parsed.weeks || [];
        if (topics.length === 0)
          throw new Error("No Term " + termNum + " weeks found in image.");
      }
      if (!topics || topics.length === 0)
        throw new Error("No topics found for Term " + termNum + ".");
      var lpw = lessonDays.length;
      var expanded = [];
      topics.forEach(function (t) {
        var lines = t.topic
          .split(/\n/)
          .map(function (l) {
            return l.trim();
          })
          .filter(function (l) {
            return l.length > 4;
          });
        if (lines.length === 0) lines = [t.topic];
        for (var i = 0; i < lpw; i++)
          expanded.push({
            weekLabel: t.weekLabel,
            topic: i < lines.length ? lines[i] : lines[lines.length - 1],
          });
      });
      setRows(assignDates(expanded, startDate, lessonDays, extraHols));
      setStep(4);
    } catch (err) {
      setError(err.message);
    }
    setLoading(false);
  };

  var updRow = function (i, k, v) {
    setRows(function (p) {
      return p.map(function (r, j) {
        return j === i ? Object.assign({}, r, { [k]: v }) : r;
      });
    });
  };

  var getInfo = function () {
    return {
      schoolName: schoolName,
      teacherName: teacherName,
      supervisorName: supervisorName,
      subject: subject,
      grade: grade,
      term: term,
      year: year,
      periodsPerWeek: lessonDays.length,
      dayLabel: lessonDays
        .map(function (d) {
          return WD[d];
        })
        .join(", "),
      startDate: startDate,
    };
  };

  var doPrint = function () {
    var w = window.open("", "_blank");
    w.document.write(buildPrintHTML(getInfo(), rows, reflection));
    w.document.close();
    setTimeout(function () {
      w.focus();
      w.print();
    }, 500);
  };

  var doDocx = function () {
    var xml = buildDocxXML(getInfo(), rows, reflection);
    var blob = new Blob([xml], { type: "application/msword" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "ATP_Tracker_T" + term + "_Gr" + grade + ".doc";
    a.click();
  };

  var doCSV = function () {
    var lines = [
      subject + " | Grade " + grade + " | Term " + term + " | " + year,
      "School: " + schoolName + " | Teacher: " + teacherName,
      "",
      "Week,Date,Topic / Activity (CAPS Aligned),Done,Initials",
    ];
    rows.forEach(function (r) {
      lines.push(
        r.weekLabel +
          "," +
          fmt(r.date) +
          ',"' +
          (r.topic || "").replace(/"/g, '""') +
          '",' +
          (r.done ? "Yes" : "") +
          "," +
          (r.initials || ""),
      );
    });
    lines.push("", "Cycle Reflection", "Question,Response");
    REFLECT_QUESTIONS.forEach(function (item) {
      lines.push('"' + item.q + '",' + (reflection[item.key] || ""));
    });
    lines.push(
      "SMT Comment," + (reflection.smtComment || ""),
      "Educator: " + teacherName + ",Signature:,,Date:",
      "HoD/Supervisor: " + (supervisorName || "") + ",Signature:,,Date:",
    );
    var blob = new Blob([lines.join("\n")], {
      type: "text/csv;charset=utf-8;",
    });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "ATP_Tracker_T" + term + "_Gr" + grade + ".csv";
    a.click();
  };

  var inp = {
    width: "100%",
    padding: "8px 11px",
    border: "1.5px solid #c8d8ea",
    borderRadius: 6,
    fontSize: 13,
    boxSizing: "border-box",
    marginBottom: 12,
    fontFamily: "inherit",
  };
  var lbl = {
    display: "block",
    fontWeight: 600,
    fontSize: 12,
    color: BLUE,
    marginBottom: 4,
  };
  var card = {
    background: "#fff",
    borderRadius: 10,
    padding: 22,
    marginBottom: 16,
    boxShadow: "0 2px 8px rgba(0,0,0,.07)",
  };
  var sec = {
    fontSize: 15,
    fontWeight: 700,
    color: BLUE,
    marginBottom: 14,
    borderLeft: "4px solid " + ORANGE,
    paddingLeft: 10,
  };
  var pill = function (a) {
    return {
      display: "inline-block",
      cursor: "pointer",
      margin: "3px 4px",
      padding: "7px 15px",
      borderRadius: 6,
      fontWeight: 700,
      fontSize: 13,
      userSelect: "none",
      background: a ? ORANGE : LIGHT,
      color: a ? "#fff" : BLUE,
      border: "1.5px solid " + (a ? ORANGE : "#b8cfe0"),
    };
  };
  var btn = function (c) {
    return {
      background: c || ORANGE,
      color: "#fff",
      border: "none",
      borderRadius: 6,
      padding: "10px 22px",
      fontWeight: 700,
      fontSize: 13,
      cursor: "pointer",
      fontFamily: "inherit",
    };
  };
  var abox = function (t) {
    var m = {
      error: ["#fef2f2", "#b91c1c", "#fca5a5"],
      success: ["#f0fdf4", "#15803d", "#86efac"],
      info: ["#eff6ff", "#1d4ed8", "#bfdbfe"],
      warn: ["#fffbeb", "#92400e", "#fcd34d"],
    };
    var v = m[t] || m.info;
    return {
      padding: "10px 14px",
      borderRadius: 6,
      marginBottom: 12,
      fontSize: 13,
      background: v[0],
      color: v[1],
      border: "1px solid " + v[2],
    };
  };
  var uploadBox = {
    border: "2px dashed " + ORANGE,
    borderRadius: 8,
    padding: 28,
    textAlign: "center",
    cursor: "pointer",
    background: LIGHT,
  };
  var mCard = function (a) {
    return {
      border: "2px solid " + (a ? ORANGE : "#c8d8ea"),
      borderRadius: 10,
      padding: 18,
      cursor: "pointer",
      background: a ? "#fff8f5" : "#fff",
      flex: 1,
      minWidth: 160,
    };
  };

  return (
    <div
      style={{
        fontFamily: "'Segoe UI', sans-serif",
        background: "#f0f4f9",
        minHeight: "100vh",
      }}
    >
      <div style={{ background: BLUE, padding: "14px 24px" }}>
        <div style={{ fontSize: 20, fontWeight: 700, color: ORANGE }}>
          Qwerzy ATP Tracker
        </div>
        <div
          style={{ fontSize: 11, color: "rgba(255,255,255,0.7)", marginTop: 2 }}
        >
          We Play To Learn — CAPS-Aligned Lesson Tracker Generator
        </div>
      </div>

      <div
        style={{
          background: "#fff",
          borderBottom: "3px solid " + ORANGE,
          display: "flex",
          padding: "0 20px",
          overflowX: "auto",
        }}
      >
        {["Get Started", "Load ATP", "Configure", "Your Tracker"].map(
          function (s, i) {
            var active = step === i + 1;
            var done = step > i + 1;
            return (
              <div
                key={s}
                onClick={function () {
                  if (done) setStep(i + 1);
                }}
                style={{
                  padding: "11px 14px",
                  fontSize: 13,
                  fontWeight: 600,
                  whiteSpace: "nowrap",
                  borderBottom: active
                    ? "3px solid " + ORANGE
                    : "3px solid transparent",
                  color: active ? ORANGE : done ? BLUE : "#aaa",
                  marginBottom: -3,
                  cursor: done ? "pointer" : "default",
                }}
              >
                {done ? "✓ " : i + 1 + ". "}
                {s}
              </div>
            );
          },
        )}
      </div>

      <div style={{ maxWidth: 820, margin: "0 auto", padding: "20px 16px" }}>
        {step === 1 && (
          <div style={card}>
            <div style={sec}>Welcome to the Qwerzy ATP Tracker</div>
            <p
              style={{
                fontSize: 14,
                color: "#333",
                lineHeight: 1.7,
                marginBottom: 16,
              }}
            >
              This tool turns your Annual Teaching Plan into a dated lesson
              tracker, assigning real calendar dates to each lesson and skipping
              weekends and SA public holidays automatically.
            </p>
            <div
              style={{
                background: LIGHT,
                borderRadius: 8,
                padding: 16,
                marginBottom: 20,
              }}
            >
              <div
                style={{
                  fontWeight: 700,
                  color: BLUE,
                  marginBottom: 12,
                  fontSize: 14,
                }}
              >
                How to prepare your ATP:
              </div>
              <div style={{ marginBottom: 12 }}>
                <div style={{ fontWeight: 700, color: ORANGE, fontSize: 13 }}>
                  Best — Excel or CSV
                </div>
                <div
                  style={{
                    fontSize: 13,
                    color: "#444",
                    marginTop: 4,
                    lineHeight: 1.6,
                  }}
                >
                  Open your ATP in Microsoft Excel. Go to File, Save As, CSV.
                  Upload the CSV file. Instant and perfectly accurate.
                </div>
              </div>
              <div style={{ marginBottom: 12 }}>
                <div style={{ fontWeight: 700, color: ORANGE, fontSize: 13 }}>
                  Also great — Word or DOCX
                </div>
                <div
                  style={{
                    fontSize: 13,
                    color: "#444",
                    marginTop: 4,
                    lineHeight: 1.6,
                  }}
                >
                  Upload directly if your ATP is in Word format. If it is a PDF,
                  open it in Microsoft Word or Google Docs, save as .docx, and
                  upload here.
                </div>
              </div>
              <div>
                <div style={{ fontWeight: 700, color: ORANGE, fontSize: 13 }}>
                  Also supported — Photo or Screenshot
                </div>
                <div
                  style={{
                    fontSize: 13,
                    color: "#444",
                    marginTop: 4,
                    lineHeight: 1.6,
                  }}
                >
                  Take a clear photo or screenshot of your ATP. Claude reads the
                  table using AI vision.
                </div>
              </div>
            </div>
            <div
              style={{
                background: "#fff0e8",
                borderRadius: 6,
                padding: "10px 14px",
                marginBottom: 20,
                fontSize: 13,
                color: "#92400e",
                border: "1px solid #fcd34d",
              }}
            >
              Only lesson topics are extracted. No assessment tasks, no tests.
              Your tracker stays clean and easy to read.
            </div>
            <button
              style={Object.assign({}, btn(), {
                fontSize: 14,
                padding: "12px 32px",
              })}
              onClick={function () {
                setStep(2);
              }}
            >
              Get Started
            </button>
          </div>
        )}

        {step === 2 && (
          <div style={card}>
            <div style={sec}>Load Your ATP</div>
            <div
              style={{
                display: "flex",
                gap: 12,
                marginBottom: 20,
                flexWrap: "wrap",
              }}
            >
              <div
                style={mCard(method === "excel")}
                onClick={function () {
                  setMethod("excel");
                  setError("");
                }}
              >
                <div style={{ fontSize: 28, marginBottom: 6 }}>📊</div>
                <div style={{ fontWeight: 700, color: BLUE }}>Excel or CSV</div>
                <div style={{ fontSize: 12, color: "#666", marginTop: 4 }}>
                  Recommended
                </div>
              </div>
              <div
                style={mCard(method === "word")}
                onClick={function () {
                  setMethod("word");
                  setError("");
                }}
              >
                <div style={{ fontSize: 28, marginBottom: 6 }}>📝</div>
                <div style={{ fontWeight: 700, color: BLUE }}>Word or DOCX</div>
                <div style={{ fontSize: 12, color: "#666", marginTop: 4 }}>
                  Great for converted PDFs
                </div>
              </div>
              <div
                style={mCard(method === "image")}
                onClick={function () {
                  setMethod("image");
                  setError("");
                }}
              >
                <div style={{ fontSize: 28, marginBottom: 6 }}>📷</div>
                <div style={{ fontWeight: 700, color: BLUE }}>
                  Photo or Screenshot
                </div>
                <div style={{ fontSize: 12, color: "#666", marginTop: 4 }}>
                  AI reads from image
                </div>
              </div>
            </div>
            {error && <div style={abox("error")}>{error}</div>}
            {loading && (
              <div style={abox("info")}>Reading file, please wait...</div>
            )}
            {method === "excel" && (
              <div
                style={uploadBox}
                onClick={function () {
                  fileRef.current.click();
                }}
                onDragOver={function (e) {
                  e.preventDefault();
                }}
                onDrop={function (e) {
                  e.preventDefault();
                  handleFileDrop(e.dataTransfer.files[0]);
                }}
              >
                <div style={{ fontSize: 36 }}>📊</div>
                <div style={{ fontWeight: 700, color: BLUE, marginTop: 8 }}>
                  Click or drag your Excel or CSV file here
                </div>
                <div style={{ fontSize: 12, color: "#888", marginTop: 4 }}>
                  In Excel: File, Save As, CSV — then upload that file
                </div>
                <input
                  ref={fileRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  style={{ display: "none" }}
                  onChange={function (e) {
                    handleFileDrop(e.target.files[0]);
                  }}
                />
              </div>
            )}
            {method === "word" && (
              <div>
                <div style={abox("info")}>
                  If your ATP is a PDF, open it in Microsoft Word or Google Docs
                  first and save as .docx.
                </div>
                <div
                  style={uploadBox}
                  onClick={function () {
                    fileRef.current.click();
                  }}
                  onDragOver={function (e) {
                    e.preventDefault();
                  }}
                  onDrop={function (e) {
                    e.preventDefault();
                    handleDocx(e.dataTransfer.files[0]);
                  }}
                >
                  <div style={{ fontSize: 36 }}>📝</div>
                  <div style={{ fontWeight: 700, color: BLUE, marginTop: 8 }}>
                    Click or drag your Word file here
                  </div>
                  <div style={{ fontSize: 12, color: "#888", marginTop: 4 }}>
                    Microsoft Word or Google Docs .docx format
                  </div>
                  <input
                    ref={fileRef}
                    type="file"
                    accept=".docx,.doc"
                    style={{ display: "none" }}
                    onChange={function (e) {
                      handleDocx(e.target.files[0]);
                    }}
                  />
                </div>
              </div>
            )}
            {method === "image" && (
              <div>
                <div style={abox("info")}>
                  Take a clear, well-lit photo or screenshot of your ATP table.
                  Claude will read it using AI vision.
                </div>
                <div
                  style={uploadBox}
                  onClick={function () {
                    fileRef.current.click();
                  }}
                  onDragOver={function (e) {
                    e.preventDefault();
                  }}
                  onDrop={function (e) {
                    e.preventDefault();
                    handleImage(e.dataTransfer.files[0]);
                  }}
                >
                  <div style={{ fontSize: 36 }}>📷</div>
                  <div style={{ fontWeight: 700, color: BLUE, marginTop: 8 }}>
                    Click or drag your image here
                  </div>
                  <div style={{ fontSize: 12, color: "#888", marginTop: 4 }}>
                    JPG, PNG, GIF, WEBP supported
                  </div>
                  <input
                    ref={fileRef}
                    type="file"
                    accept="image/*"
                    style={{ display: "none" }}
                    onChange={function (e) {
                      handleImage(e.target.files[0]);
                    }}
                  />
                </div>
                {imagePreview && (
                  <div style={{ marginTop: 12, textAlign: "center" }}>
                    <img
                      src={imagePreview}
                      alt="preview"
                      style={{
                        maxWidth: "100%",
                        maxHeight: 260,
                        borderRadius: 8,
                        border: "1px solid #c8d8ea",
                      }}
                    />
                  </div>
                )}
              </div>
            )}
            {!method && (
              <div
                style={{
                  textAlign: "center",
                  color: "#aaa",
                  fontSize: 13,
                  padding: "20px 0",
                }}
              >
                Select an input method above to continue.
              </div>
            )}
          </div>
        )}

        {step === 3 && (
          <div>
            {rawFileData && (
              <div style={abox("success")}>
                {
                  "ATP loaded. Select your term, fill in details and click Generate."
                }
              </div>
            )}
            {!rawFileData && (
              <div style={abox("warn")}>
                No file loaded. Please go back and upload your ATP.
              </div>
            )}
            <div style={card}>
              <div style={sec}>Teacher and School Details</div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: 12,
                }}
              >
                <div>
                  <label style={lbl}>School Name *</label>
                  <input
                    style={inp}
                    value={schoolName}
                    onChange={function (e) {
                      setSchoolName(e.target.value);
                    }}
                    placeholder="e.g. Buhlebethu Primary"
                  />
                </div>
                <div>
                  <label style={lbl}>Teacher Name *</label>
                  <input
                    style={inp}
                    value={teacherName}
                    onChange={function (e) {
                      setTeacherName(e.target.value);
                    }}
                    placeholder="e.g. Ms N. Dlamini"
                  />
                </div>
                <div>
                  <label style={lbl}>HoD / Supervisor</label>
                  <input
                    style={inp}
                    value={supervisorName}
                    onChange={function (e) {
                      setSupervisorName(e.target.value);
                    }}
                    placeholder="e.g. Mr T. Mthembu"
                  />
                </div>
                <div>
                  <label style={lbl}>Subject *</label>
                  <input
                    style={inp}
                    value={subject}
                    onChange={function (e) {
                      setSubject(e.target.value);
                    }}
                    placeholder="e.g. Life Skills: PSW"
                  />
                </div>
                <div>
                  <label style={lbl}>Grade</label>
                  <input
                    style={inp}
                    value={grade}
                    onChange={function (e) {
                      setGrade(e.target.value);
                    }}
                    placeholder="e.g. 5"
                  />
                </div>
                <div>
                  <label style={lbl}>Year</label>
                  <input
                    style={inp}
                    type="number"
                    value={year}
                    onChange={function (e) {
                      setYear(e.target.value);
                    }}
                  />
                </div>
                <div>
                  <label style={lbl}>Term *</label>
                  <select
                    style={inp}
                    value={term}
                    onChange={function (e) {
                      setTerm(e.target.value);
                    }}
                  >
                    <option value="1">Term 1</option>
                    <option value="2">Term 2</option>
                    <option value="3">Term 3</option>
                    <option value="4">Term 4</option>
                  </select>
                </div>
              </div>
            </div>
            <div style={card}>
              <div style={sec}>Timetable</div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: 12,
                }}
              >
                <div>
                  <label style={lbl}>Term Start Date *</label>
                  <input
                    style={inp}
                    type="date"
                    value={startDate}
                    onChange={function (e) {
                      setStartDate(e.target.value);
                    }}
                  />
                </div>
                <div>
                  <label style={lbl}>Which days does this subject run? *</label>
                  <div
                    style={{
                      display: "flex",
                      gap: 6,
                      marginTop: 4,
                      flexWrap: "wrap",
                    }}
                  >
                    {WD.map(function (day, i) {
                      return (
                        <span
                          key={day}
                          style={Object.assign(
                            {},
                            pill(lessonDays.includes(i)),
                            {
                              minWidth: 46,
                              textAlign: "center",
                              padding: "7px 10px",
                            },
                          )}
                          onClick={function () {
                            toggleDay(i);
                          }}
                        >
                          {day}
                        </span>
                      );
                    })}
                  </div>
                  {lessonDays.length > 0 && (
                    <div style={{ fontSize: 11, color: "#555", marginTop: 6 }}>
                      {lessonDays.length +
                        " period(s)/week on " +
                        lessonDays
                          .map(function (d) {
                            return WD[d];
                          })
                          .join(", ")}
                    </div>
                  )}
                </div>
              </div>
              <div style={{ marginTop: 8 }}>
                <label style={lbl}>
                  Additional school closures (SA public holidays auto-excluded)
                </label>
                <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                  <input
                    style={Object.assign({}, inp, { marginBottom: 0, flex: 1 })}
                    type="date"
                    value={newHol}
                    onChange={function (e) {
                      setNewHol(e.target.value);
                    }}
                  />
                  <button
                    style={btn()}
                    onClick={function () {
                      if (newHol && !extraHols.includes(newHol)) {
                        setExtraHols(function (p) {
                          return p.concat([newHol]);
                        });
                        setNewHol("");
                      }
                    }}
                  >
                    Add
                  </button>
                </div>
                {extraHols.length > 0 && (
                  <div>
                    {extraHols.map(function (h) {
                      return (
                        <span
                          key={h}
                          style={{
                            display: "inline-flex",
                            alignItems: "center",
                            gap: 5,
                            background: "#fff0e8",
                            color: ORANGE,
                            border: "1px solid " + ORANGE,
                            borderRadius: 4,
                            padding: "3px 10px",
                            fontSize: 12,
                            margin: "2px 4px 2px 0",
                          }}
                        >
                          {h}
                          <span
                            style={{ cursor: "pointer", fontWeight: 700 }}
                            onClick={function () {
                              setExtraHols(function (p) {
                                return p.filter(function (x) {
                                  return x !== h;
                                });
                              });
                            }}
                          >
                            x
                          </span>
                        </span>
                      );
                    })}{" "}
                  </div>
                )}
                {extraHols.length === 0 && (
                  <span style={{ fontSize: 12, color: "#aaa" }}>
                    No extra closures added.
                  </span>
                )}
              </div>
            </div>
            {error && <div style={abox("error")}>{error}</div>}
            {loading && (
              <div style={abox("info")}>
                {"Reading Term " +
                  term +
                  " content from your ATP, please wait..."}
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <button
                style={btn(BLUE)}
                onClick={function () {
                  setStep(2);
                }}
              >
                Back
              </button>
              <button
                style={Object.assign({}, btn(), { opacity: loading ? 0.5 : 1 })}
                disabled={loading}
                onClick={buildTracker}
              >
                {"Generate Term " + term + " Tracker"}
              </button>
            </div>
          </div>
        )}

        {step === 4 && rows && (
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                flexWrap: "wrap",
                gap: 10,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  fontSize: 15,
                  fontWeight: 700,
                  color: BLUE,
                  borderLeft: "4px solid " + ORANGE,
                  paddingLeft: 10,
                }}
              >
                {subject +
                  " — Grade " +
                  grade +
                  " — Term " +
                  term +
                  " — " +
                  year}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button style={btn(BLUE)} onClick={doCSV}>
                  Download CSV
                </button>
                <button style={btn("#1d6a3a")} onClick={doDocx}>
                  Download DOCX
                </button>
                <button style={btn("#16a34a")} onClick={doPrint}>
                  Print / PDF
                </button>
                <button
                  style={btn()}
                  onClick={function () {
                    setStep(3);
                  }}
                >
                  Edit
                </button>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(3,1fr)",
                gap: 8,
                marginBottom: 16,
              }}
            >
              {[
                ["School", schoolName],
                ["Teacher", teacherName],
                ["HoD/Supervisor", supervisorName || "-"],
                ["Subject", subject],
                ["Grade", grade],
                ["Term and Year", "Term " + term + " — " + year],
                [
                  "Periods/Week",
                  lessonDays.length +
                    " (" +
                    lessonDays
                      .map(function (d) {
                        return WD[d];
                      })
                      .join(", ") +
                    ")",
                ],
                ["Term Start", startDate],
                ["Total Lessons", rows.length],
              ].map(function (kv) {
                return (
                  <div
                    key={kv[0]}
                    style={{
                      background: LIGHT,
                      padding: "6px 10px",
                      borderRadius: 6,
                    }}
                  >
                    <div style={{ fontSize: 10, color: "#888" }}>{kv[0]}</div>
                    <div style={{ fontWeight: 700, color: BLUE, fontSize: 13 }}>
                      {kv[1] || "-"}
                    </div>
                  </div>
                );
              })}
            </div>

            <div
              style={{
                background: "#fff",
                borderRadius: 10,
                overflow: "hidden",
                boxShadow: "0 2px 8px rgba(0,0,0,.07)",
                marginBottom: 16,
              }}
            >
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: 13,
                }}
              >
                <thead>
                  <tr style={{ background: BLUE, color: "#fff" }}>
                    <th
                      style={{
                        padding: "9px 10px",
                        textAlign: "left",
                        width: 70,
                      }}
                    >
                      Week
                    </th>
                    <th
                      style={{
                        padding: "9px 10px",
                        textAlign: "left",
                        width: 160,
                      }}
                    >
                      Date
                    </th>
                    <th style={{ padding: "9px 10px", textAlign: "left" }}>
                      Topic / Activity (CAPS Aligned)
                    </th>
                    <th
                      style={{
                        padding: "9px 10px",
                        textAlign: "center",
                        width: 65,
                      }}
                    >
                      Done
                    </th>
                    <th
                      style={{
                        padding: "9px 10px",
                        textAlign: "center",
                        width: 80,
                      }}
                    >
                      Initials
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {(function () {
                    var prevWeek = "";
                    return rows.map(function (r, i) {
                      var show = r.weekLabel !== prevWeek;
                      if (show) prevWeek = r.weekLabel;
                      return (
                        <tr
                          key={i}
                          style={{ background: i % 2 === 0 ? LIGHT : "#fff" }}
                        >
                          <td
                            style={{
                              padding: "8px 10px",
                              borderBottom: "1px solid #dde6f0",
                              fontWeight: show ? 700 : 400,
                              color: show ? BLUE : "#ccc",
                              fontSize: 12,
                            }}
                          >
                            {show ? r.weekLabel : ""}
                          </td>
                          <td
                            style={{
                              padding: "8px 10px",
                              borderBottom: "1px solid #dde6f0",
                              fontSize: 12,
                              whiteSpace: "nowrap",
                            }}
                          >
                            {fmt(r.date)}
                          </td>
                          <td
                            style={{
                              padding: "8px 10px",
                              borderBottom: "1px solid #dde6f0",
                            }}
                          >
                            <input
                              style={{
                                border: "none",
                                background: "transparent",
                                width: "100%",
                                fontSize: 13,
                                fontFamily: "inherit",
                              }}
                              value={r.topic}
                              onChange={function (e) {
                                updRow(i, "topic", e.target.value);
                              }}
                            />
                          </td>
                          <td
                            style={{
                              padding: "8px 10px",
                              borderBottom: "1px solid #dde6f0",
                              textAlign: "center",
                            }}
                          >
                            <select
                              value={r.done ? "yes" : ""}
                              onChange={function (e) {
                                updRow(i, "done", e.target.value === "yes");
                              }}
                              style={{
                                border: "1px solid #c8d8ea",
                                borderRadius: 4,
                                padding: "2px 4px",
                                fontSize: 13,
                              }}
                            >
                              <option value="">-</option>
                              <option value="yes">✓</option>
                            </select>
                          </td>
                          <td
                            style={{
                              padding: "8px 10px",
                              borderBottom: "1px solid #dde6f0",
                              textAlign: "center",
                            }}
                          >
                            <input
                              style={{
                                border: "none",
                                borderBottom: "1px solid #aaa",
                                background: "transparent",
                                width: 60,
                                textAlign: "center",
                                fontSize: 13,
                                fontFamily: "inherit",
                              }}
                              value={r.initials}
                              onChange={function (e) {
                                updRow(i, "initials", e.target.value);
                              }}
                              placeholder="___"
                            />
                          </td>
                        </tr>
                      );
                    });
                  })()}
                </tbody>
              </table>
            </div>

            <div style={card}>
              <div style={sec}>{"Cycle Reflection — Term " + term}</div>
              {REFLECT_QUESTIONS.map(function (item) {
                return (
                  <div
                    key={item.key}
                    style={{
                      display: "grid",
                      gridTemplateColumns: "2fr 3fr",
                      gap: 12,
                      alignItems: "center",
                      borderBottom: "1px solid #eef4fb",
                      padding: "10px 0",
                    }}
                  >
                    <div
                      style={{ fontSize: 13, color: "#333", fontWeight: 500 }}
                    >
                      {item.num + ". " + item.q}
                    </div>
                    {item.type === "select" ? (
                      <select
                        value={reflection[item.key]}
                        onChange={function (e) {
                          setR(item.key, e.target.value);
                        }}
                        style={{
                          padding: "8px 10px",
                          border: "1.5px solid #c8d8ea",
                          borderRadius: 6,
                          fontSize: 13,
                          fontFamily: "inherit",
                          color: BLUE,
                          background: LIGHT,
                        }}
                      >
                        {item.opts.map(function (o) {
                          return (
                            <option key={o} value={o}>
                              {o || "-- Select --"}
                            </option>
                          );
                        })}
                      </select>
                    ) : (
                      <input
                        value={reflection[item.key]}
                        onChange={function (e) {
                          setR(item.key, e.target.value);
                        }}
                        placeholder="Type your response..."
                        style={{
                          padding: "8px 10px",
                          border: "1.5px solid #c8d8ea",
                          borderRadius: 6,
                          fontSize: 13,
                          fontFamily: "inherit",
                          width: "100%",
                          boxSizing: "border-box",
                        }}
                      />
                    )}
                  </div>
                );
              })}
              <div
                style={{
                  marginTop: 16,
                  background: LIGHT,
                  borderRadius: 8,
                  padding: 16,
                }}
              >
                <div
                  style={{
                    fontWeight: 700,
                    color: BLUE,
                    fontSize: 13,
                    marginBottom: 10,
                  }}
                >
                  SMT Comment
                </div>
                <textarea
                  value={reflection.smtComment}
                  onChange={function (e) {
                    setR("smtComment", e.target.value);
                  }}
                  placeholder="SMT comment here..."
                  style={{
                    width: "100%",
                    padding: "8px 10px",
                    border: "1.5px solid #c8d8ea",
                    borderRadius: 6,
                    fontSize: 13,
                    fontFamily: "inherit",
                    minHeight: 70,
                    resize: "vertical",
                    boxSizing: "border-box",
                    marginBottom: 12,
                  }}
                />
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr",
                    gap: 12,
                  }}
                >
                  <div>
                    <label style={lbl}>SMT Name and Signature</label>
                    <input
                      value={reflection.smtName}
                      onChange={function (e) {
                        setR("smtName", e.target.value);
                      }}
                      placeholder="SMT name..."
                      style={{
                        width: "100%",
                        padding: "8px 10px",
                        border: "1.5px solid #c8d8ea",
                        borderRadius: 6,
                        fontSize: 13,
                        fontFamily: "inherit",
                        boxSizing: "border-box",
                      }}
                    />
                  </div>
                  <div>
                    <label style={lbl}>Date</label>
                    <input
                      type="date"
                      value={reflection.smtDate}
                      onChange={function (e) {
                        setR("smtDate", e.target.value);
                      }}
                      style={{
                        width: "100%",
                        padding: "8px 10px",
                        border: "1.5px solid #c8d8ea",
                        borderRadius: 6,
                        fontSize: 13,
                        fontFamily: "inherit",
                        boxSizing: "border-box",
                      }}
                    />
                  </div>
                </div>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 20,
                marginBottom: 20,
              }}
            >
              {[
                ["Educator", teacherName],
                ["HoD / Supervisor", supervisorName],
              ].map(function (rv) {
                return (
                  <div
                    key={rv[0]}
                    style={{
                      border: "1px solid #c8d8ea",
                      borderRadius: 8,
                      padding: 20,
                      background: "#fff",
                    }}
                  >
                    <div
                      style={{
                        fontWeight: 700,
                        color: BLUE,
                        marginBottom: 14,
                        fontSize: 14,
                      }}
                    >
                      {rv[0]}
                    </div>
                    <div style={{ fontSize: 13, marginBottom: 14 }}>
                      Name: <strong>{rv[1] || "_______________"}</strong>
                    </div>
                    <div style={{ fontSize: 13, marginBottom: 18 }}>
                      Signature:{" "}
                      <span
                        style={{
                          display: "inline-block",
                          borderBottom: "1.5px solid #333",
                          width: 180,
                          marginLeft: 6,
                        }}
                      >
                        &nbsp;
                      </span>
                    </div>
                    <div style={{ fontSize: 13 }}>
                      Date:{" "}
                      <span
                        style={{
                          display: "inline-block",
                          borderBottom: "1.5px solid #333",
                          width: 130,
                          marginLeft: 6,
                        }}
                      >
                        &nbsp;
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>

            <div style={{ display: "flex", justifyContent: "flex-start" }}>
              <button
                style={btn(BLUE)}
                onClick={function () {
                  setStep(3);
                }}
              >
                Back to Configure
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
