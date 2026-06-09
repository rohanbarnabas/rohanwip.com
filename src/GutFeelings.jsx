import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const C = {
  bg: "#FFFFFF", surface: "#F5F5F5", surfaceHigh: "#EBEBEB",
  accent: "#000000", accentDim: "#F0F0F0", accentText: "#FFFFFF",
  text: "#000000", muted: "#8E8E93", border: "#E5E5EA",
  success: "#34C759", error: "#FF3B30", white: "#FFFFFF",
};

const FONT = `-apple-system, BlinkMacSystemFont, "Helvetica Neue", Helvetica, Arial, sans-serif`;

const FULL_DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];

const FOODS = {
  breakfast: ["Overnight Oats + Cosmix","Poha + Cosmix","Uttapam + Cosmix"],
  roti: ["Ragi roti","Bajra roti","Atta roti (1 roti only — max once/week)"],
  sabzi: ["Bhindi","Aloo-bhindi","Mixed veg (beans, carrots, peas)","Broccoli","Patta gobi","Palak","Beetroot","Paneer (max once/week)"],
  salad: ["Kheera + gajar + apple + chole + chaat masala","Kheera + gajar (plain)"],
  rice: ["Veg pulao","Veg biryani","Khichdi","Idli-Sambar"],
};

const BRISTOL = {
  1:{label:"Type 1 – Hard lumps",color:"#8B2020"},
  2:{label:"Type 2 – Lumpy sausage",color:"#B84040"},
  3:{label:"Type 3 – Cracked sausage",color:"#D4785A"},
  4:{label:"Type 4 – Smooth sausage",color:"#34C759"},
  5:{label:"Type 5 – Soft blobs",color:"#FFA726"},
  6:{label:"Type 6 – Mushy",color:"#FF7043"},
  7:{label:"Type 7 – Liquid",color:"#FF3B30"},
};

const HINDI_DAYS = {
  "Monday":"सोमवार","Tuesday":"मंगलवार","Wednesday":"बुधवार",
  "Thursday":"गुरुवार","Friday":"शुक्रवार","Saturday":"शनिवार","Sunday":"रविवार"
};

const H = {
  "Overnight Oats + Cosmix":"ओट्स (सब्जा सीड्स के साथ, पानी से गाढ़ा)",
  "Poha + Cosmix":"पोहा",
  "Uttapam + Cosmix":"उत्तपम",
  "Ragi roti":"रागी रोटी",
  "Bajra roti":"बाजरा रोटी",
  "Atta roti (1 only)":"आटा रोटी (सिर्फ 1 रोटी)",
  "Atta roti (1 roti only — max once/week)":"आटा रोटी (सिर्फ 1 रोटी)",
  "Usal (dry)":"उसल (सूखा बनाना है)",
  "Bhindi":"भिंडी",
  "Aloo-bhindi":"आलू-भिंडी",
  "Mixed veg":"मिक्स सब्जी (बींस, गाजर, मटर)",
  "Mixed veg (beans, carrots, peas)":"मिक्स सब्जी (बींस, गाजर, मटर)",
  "Broccoli":"ब्रोकली",
  "Patta gobi":"पत्ता गोभी",
  "Palak":"पालक",
  "Beetroot":"चुकंदर",
  "Paneer":"पनीर (हफ्ते में 1 बार)",
  "Paneer (max once/week)":"पनीर (हफ्ते में 1 बार)",
  "Kheera + gajar + apple + chole + chaat masala":"खीरा + गाजर + सेब + छोले + चाट मसाला",
  "Kheera + gajar + apple + chole":"खीरा + गाजर + सेब + छोले + चाट मसाला",
  "Kheera + gajar (plain)":"खीरा + गाजर",
  "Kheera + gajar":"खीरा + गाजर",
  "Veg pulao":"वेज पुलाव",
  "Veg biryani":"वेज बिरयानी",
  "Khichdi":"खिचड़ी (कटी सब्जियों के साथ)",
  "Idli-Sambar":"इडली-सांभर",
};

function h(str) { return H[str] || str; }

function cookText(plan) {
  if (!plan) return "";
  const lines = ["इस हफ्ते का खाना\n"];
  plan.forEach(day => {
    lines.push(`*${HINDI_DAYS[day.day] || day.day}*`);
    lines.push(`नाश्ता:  ${h(day.breakfast)}`);
    lines.push(`लंच:     ${h(day.lunch?.roti)} + उसल (सूखा) + ${h(day.lunch?.sabzi)}`);
    lines.push(`डिनर:    2 उबले अंडे + ${h(day.dinner?.sabzi)} + ${h(day.dinner?.rice)}`);
    lines.push(`सलाद:    ${h(day.dinner?.salad)} (बिना नमक)`);
    lines.push("");
  });
  return lines.join("\n");
}

function exportCookXLSX(plan) {
  const headers = ["दिन (Day)","नाश्ता","लंच - रोटी","लंच - प्रोटीन","लंच - सब्जी","डिनर - अंडे","डिनर - सब्जी","डिनर - राइस","डिनर - सलाद"];
  const rows = plan.map(day => [
    `${HINDI_DAYS[day.day] || day.day} (${day.day})`,
    h(day.breakfast),
    h(day.lunch?.roti),
    "उसल (सूखा बनाना है)",
    h(day.lunch?.sabzi),
    "2 उबले अंडे",
    h(day.dinner?.sabzi),
    h(day.dinner?.rice),
    `${h(day.dinner?.salad)} (बिना नमक)`,
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = headers.map(() => ({ wch: 28 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Cook Sheet");
  XLSX.writeFile(wb, `cook_sheet_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportDhruviXLSX(entries) {
  const headers = [
    "Date","Time","Raw Entry","Meals","Ate as Planned",
    "Acidity","Gut Discomfort","Gut Notes","Energy Level",
    "Stomach Cleared","Bristol Type","Slept Well",
    "Morning Routine","Bedtime Routine","Phone 1st 30min",
    "Movement","Appetite Satisfaction","Dopamine Sources","Mood","Notes"
  ];
  const rows = entries.map(e => {
    const p = e.parsed || {};
    return [
      new Date(e.date).toLocaleDateString("en-IN"),
      e.time,
      e.raw,
      p.meals?.map(m => `${m.item}${m.portion ? ` (${m.portion})` : ""}`).join(", ") || "",
      p.ate_as_planned || "",
      p.acidity || "",
      p.gut_discomfort || "",
      p.gut_notes || "",
      p.energy_level || "",
      p.stomach_cleared || "",
      p.bristol_type != null ? `Type ${p.bristol_type}` : "",
      p.slept_well || "",
      p.morning_routine || "",
      p.bedtime_routine || "",
      p.phone_first_30min || "",
      p.movement || "",
      p.appetite_satisfaction || "",
      p.dopamine_sources?.join(", ") || "",
      p.mood || "",
      p.notes || "",
    ];
  });
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = [
    {wch:12},{wch:8},{wch:40},{wch:30},{wch:14},
    {wch:10},{wch:14},{wch:25},{wch:12},
    {wch:14},{wch:12},{wch:10},
    {wch:14},{wch:14},{wch:16},
    {wch:20},{wch:25},{wch:25},{wch:12},{wch:25}
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dhruvi Report");
  XLSX.writeFile(wb, `dhruvi_report_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function suggestWeek() {
  let paneerUsed = false, attaUsed = false;
  const sabziPool = ["Bhindi","Aloo-bhindi","Mixed veg","Broccoli","Patta gobi","Palak","Beetroot"];
  return FULL_DAYS.map((day, i) => {
    const breakfast = i === 2 ? "Poha + Cosmix" : i === 5 ? "Uttapam + Cosmix" : "Overnight Oats + Cosmix";
    let roti = i % 2 === 0 ? "Ragi roti" : "Bajra roti";
    if (!attaUsed && i === 3) { roti = "Atta roti (1 only)"; attaUsed = true; }
    let lunchSabzi = sabziPool[i % sabziPool.length];
    if (!paneerUsed && i === 1) { lunchSabzi = "Paneer"; paneerUsed = true; }
    const dinnerSabzi = sabziPool[(i + 3) % sabziPool.length];
    return {
      day,
      breakfast,
      lunch: { roti, protein: "Usal (dry)", sabzi: lunchSabzi },
      eveningSnack: "Cosmix shake (1 scoop)",
      dinner: {
        salad: i % 2 === 0 ? "Kheera + gajar + apple + chole" : "Kheera + gajar",
        eggs: "2 boiled eggs",
        sabzi: dinnerSabzi,
        rice: FOODS.rice[i % FOODS.rice.length],
      },
    };
  });
}

async function callClaude(system, userMsg) {
  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 1000,
        system,
        messages: [{ role: "user", content: userMsg }],
      }),
    });
    const data = await res.json();
    return data.content?.[0]?.text || "";
  } catch { return ""; }
}

async function parseLog(text) {
  const sys = `Parse the user's casual health log message into a JSON object. Return ONLY valid JSON, no markdown, no explanation.
Schema:
{
  "meals": [{"item":"string","portion":"string or null","mealType":"breakfast|lunch|snack|dinner|unknown"}],
  "ate_as_planned": "yes|no|partial|unknown",
  "acidity": "none|mild|moderate|severe|unknown",
  "gut_discomfort": "none|mild|moderate|severe|unknown",
  "gut_notes": "string or null",
  "energy_level": "low|medium|high|unknown",
  "stomach_cleared": "yes|no|unknown",
  "bristol_type": null or integer 1-7,
  "slept_well": "yes|no|partially|unknown",
  "morning_routine": "done|skipped|partial|unknown",
  "bedtime_routine": "done|skipped|partial|unknown",
  "phone_first_30min": "yes|no|unknown",
  "movement": "string or null",
  "appetite_satisfaction": "string or null",
  "dopamine_sources": [],
  "mood": "string or null",
  "notes": "string or null"
}`;
  const raw = await callClaude(sys, text);
  try { return JSON.parse(raw.replace(/```json|```/g, "").trim()); }
  catch { return null; }
}

async function updatePlanAI(plan, msg) {
  const sys = `You manage a 7-day meal plan JSON array. User will describe changes their dietician recommended. Update the plan and return ONLY the updated JSON array. No markdown, no explanation. Preserve the structure exactly.`;
  const raw = await callClaude(sys, `Current plan: ${JSON.stringify(plan)}\n\nDietician changes: ${msg}`);
  try { return JSON.parse(raw.replace(/```json|```/g, "").trim()); }
  catch { return null; }
}

const SK = { plan: "gf_plan_v2", entries: "gf_entries_v2" };

async function sget(key, fb) {
  try {
    const val = localStorage.getItem(key);
    return val ? JSON.parse(val) : fb;
  } catch { return fb; }
}

async function sset(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); }
  catch {}
}

function badge(label, value) {
  const map = {
    none: C.muted, low: C.error, medium: "#FFA726", high: C.success,
    yes: C.success, no: C.error, partial: "#FFA726", unknown: null,
    done: C.success, skipped: C.error, mild: "#FFA726", moderate: C.error,
    severe: C.error, partially: "#FFA726",
  };
  if (!value || value === "unknown") return null;
  const col = map[value] || C.muted;
  return (
    <span key={label} style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: col + "22", color: col, fontWeight: 500, marginRight: 4, marginBottom: 4, display: "inline-block", whiteSpace: "nowrap" }}>
      {label}: {value}
    </span>
  );
}

function ParsedCard({ parsed, time }) {
  const [open, setOpen] = useState(false);
  if (!parsed) return <div style={{ fontSize: 12, color: C.muted, padding: "8px 12px" }}>Could not parse — still saved.</div>;
  const bt = parsed.bristol_type;
  return (
    <div style={{ background: C.surface, borderRadius: 12, padding: 12, borderLeft: `2px solid ${C.text}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <span style={{ fontSize: 11, color: C.muted }}>{time} · parsed</span>
        <button onClick={() => setOpen(o => !o)} style={{ background: "none", border: "none", color: C.accent, fontSize: 12, cursor: "pointer", padding: 0 }}>
          {open ? "less ▲" : "more ▼"}
        </button>
      </div>
      {parsed.meals?.length > 0 && (
        <div style={{ marginBottom: 8 }}>
          {parsed.meals.map((m, i) => (
            <div key={i} style={{ fontSize: 13, color: C.text, marginBottom: 3 }}>
              · <b>{m.item}</b>{m.portion ? ` — ${m.portion}` : ""} <span style={{ fontSize: 11, color: C.muted }}>{m.mealType !== "unknown" ? m.mealType : ""}</span>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "flex", flexWrap: "wrap", marginBottom: 4 }}>
        {badge("planned", parsed.ate_as_planned)}
        {badge("energy", parsed.energy_level)}
        {parsed.mood && <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: C.muted + "22", color: C.muted, marginRight: 4, marginBottom: 4, display: "inline-block" }}>mood: {parsed.mood}</span>}
      </div>
      {bt && (
        <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: BRISTOL[bt]?.color + "22", borderRadius: 8, padding: "3px 10px", marginBottom: 4 }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: BRISTOL[bt]?.color, display: "inline-block", flexShrink: 0 }} />
          <span style={{ fontSize: 12, color: BRISTOL[bt]?.color }}>{BRISTOL[bt]?.label}</span>
        </div>
      )}
      {open && (
        <div style={{ marginTop: 8, borderTop: `1px solid ${C.border}`, paddingTop: 8 }}>
          <div style={{ display: "flex", flexWrap: "wrap" }}>
            {badge("acidity", parsed.acidity)}
            {badge("gut", parsed.gut_discomfort)}
            {badge("cleared", parsed.stomach_cleared)}
            {badge("sleep", parsed.slept_well)}
            {badge("AM routine", parsed.morning_routine)}
            {badge("PM routine", parsed.bedtime_routine)}
            {badge("phone AM", parsed.phone_first_30min)}
          </div>
          {parsed.gut_notes && <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>gut: {parsed.gut_notes}</div>}
          {parsed.movement && <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>movement: {parsed.movement}</div>}
          {parsed.appetite_satisfaction && <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>appetite: {parsed.appetite_satisfaction}</div>}
          {parsed.dopamine_sources?.length > 0 && <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>dopamine: {parsed.dopamine_sources.join(", ")}</div>}
          {parsed.notes && <div style={{ fontSize: 12, color: C.muted, marginTop: 3, fontStyle: "italic" }}>{parsed.notes}</div>}
        </div>
      )}
    </div>
  );
}

function MealRow({ label, value, options, onChange }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ marginBottom: 3 }}>
      <button onClick={() => setOpen(o => !o)} style={{ width: "100%", background: open ? C.accentDim : C.bg, border: `1px solid ${open ? C.accent : C.border}`, color: C.text, borderRadius: 8, padding: "7px 10px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
        <span style={{ fontSize: 11, color: C.muted, minWidth: 52, flexShrink: 0 }}>{label}</span>
        <span style={{ fontSize: 12, color: C.text, flex: 1, marginLeft: 8 }}>{value}</span>
        <span style={{ fontSize: 11, color: C.accent, marginLeft: 6 }}>✎</span>
      </button>
      {open && (
        <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, marginTop: 2, overflow: "hidden" }}>
          {options.map((opt, i) => (
            <button key={i} onClick={() => { onChange(opt); setOpen(false); }}
              style={{ width: "100%", background: opt === value ? C.accentDim : "none", border: "none", color: opt === value ? C.accent : C.text, padding: "9px 12px", textAlign: "left", fontSize: 12, cursor: "pointer", borderBottom: i < options.length - 1 ? `1px solid ${C.border}` : "none", display: "block" }}>
              {opt === value ? "✓ " : "  "}{opt}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function CopyButton({ text }) {
  const [copied, setCopied] = useState(false);
  const handle = () => {
    navigator.clipboard.writeText(text).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };
  return (
    <button onClick={handle} style={{ background: copied ? C.success + "22" : C.accentDim, border: `1px solid ${copied ? C.success : C.border}`, color: copied ? C.success : C.text, borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 500, transition: "all 0.2s" }}>
      {copied ? "✓ Copied" : "Copy text"}
    </button>
  );
}

function ExportButton({ label, onClick }) {
  return (
    <button onClick={onClick} style={{ background: C.accent, border: "none", color: C.accentText, borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 500 }}>
      {label}
    </button>
  );
}

function CookSheetView({ plan, onBack }) {
  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 16 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: C.surface, border: "none", color: C.text, borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 13 }}>← back</button>
          <span style={{ fontSize: 16, fontWeight: 700, color: C.text }}>Cook Sheet</span>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <CopyButton text={cookText(plan)} />
          <ExportButton label="Export .xlsx" onClick={() => exportCookXLSX(plan)} />
        </div>
      </div>
      <div style={{ fontSize: 12, color: C.muted, marginBottom: 14 }}>Hindi. Copy for WhatsApp or export as Excel.</div>
      {plan?.map((day, i) => (
        <div key={i} style={{ background: C.surface, borderRadius: 10, marginBottom: 8, overflow: "hidden" }}>
          <div style={{ background: C.surfaceHigh, padding: "8px 14px", borderBottom: `1px solid ${C.border}` }}>
            <span style={{ fontWeight: 700, fontSize: 14, color: C.text }}>{HINDI_DAYS[day.day] || day.day}</span>
            <span style={{ fontSize: 12, color: C.muted, marginLeft: 8 }}>{day.day}</span>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <tbody>
              {[
                ["नाश्ता", h(day.breakfast)],
                ["लंच · रोटी", h(day.lunch?.roti)],
                ["लंच · प्रोटीन", "उसल (सूखा बनाना है)"],
                ["लंच · सब्जी", h(day.lunch?.sabzi)],
                ["डिनर · अंडे", "2 उबले अंडे"],
                ["डिनर · सब्जी", h(day.dinner?.sabzi)],
                ["डिनर · राइस", h(day.dinner?.rice)],
                ["डिनर · सलाद", `${h(day.dinner?.salad)} (बिना नमक)`],
              ].map(([label, value], ri) => (
                <tr key={ri} style={{ borderBottom: ri < 7 ? `1px solid ${C.border}` : "none" }}>
                  <td style={{ padding: "8px 14px", color: C.muted, width: "40%", verticalAlign: "top", lineHeight: 1.5 }}>{label}</td>
                  <td style={{ padding: "8px 14px", color: C.text, lineHeight: 1.5 }}>{value}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}

function PlanTab({ plan, setPlan }) {
  const [expanded, setExpanded] = useState(null);
  const [cookSheet, setCookSheet] = useState(false);
  const [dhruviOpen, setDhruviOpen] = useState(false);
  const [dhruviText, setDhruviText] = useState("");
  const [updating, setUpdating] = useState(false);

  const savePlan = (p) => { setPlan(p); sset(SK.plan, p); };
  const handleSuggest = () => savePlan(suggestWeek());

  const handleDhruvi = async () => {
    if (!dhruviText.trim() || !plan) return;
    setUpdating(true);
    const updated = await updatePlanAI(plan, dhruviText);
    if (updated) savePlan(updated);
    setDhruviText(""); setDhruviOpen(false); setUpdating(false);
  };

  const changeMeal = (dayIdx, field, subfield, value) => {
    const updated = plan.map((d, i) => {
      if (i !== dayIdx) return d;
      if (subfield) return { ...d, [field]: { ...d[field], [subfield]: value } };
      return { ...d, [field]: value };
    });
    savePlan(updated);
  };

  if (cookSheet) return <CookSheetView plan={plan} onBack={() => setCookSheet(false)} />;

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
        <div>
          <div style={{ fontSize: 20, fontWeight: 700, color: C.text, letterSpacing: "-0.5px" }}>This Week</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Tap any meal to swap it</div>
        </div>
        {plan && <button onClick={() => setCookSheet(true)} style={{ background: C.accentDim, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 500 }}>Cook Sheet</button>}
      </div>

      {!plan ? (
        <div style={{ textAlign: "center", padding: "48px 20px" }}>
          <div style={{ fontSize: 38, marginBottom: 10 }}>🥘</div>
          <div style={{ fontSize: 15, color: C.text, marginBottom: 6, fontWeight: 600 }}>No plan yet</div>
          <div style={{ fontSize: 13, color: C.muted, marginBottom: 20 }}>I'll suggest a week based on your approved foods. You can swap anything.</div>
          <button onClick={handleSuggest} style={{ background: C.accent, border: "none", color: C.accentText, borderRadius: 10, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Suggest this week</button>
        </div>
      ) : (
        <>
          {plan.map((day, dayIdx) => (
            <div key={dayIdx} style={{ background: C.surface, borderRadius: 10, marginBottom: 6, overflow: "hidden" }}>
              <button onClick={() => setExpanded(expanded === dayIdx ? null : dayIdx)}
                style={{ width: "100%", background: "none", border: "none", padding: "11px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
                  <span style={{ fontSize: 13, fontWeight: 700, color: C.text, flexShrink: 0 }}>{day.day}</span>
                  <span style={{ fontSize: 11, color: C.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{day.breakfast}</span>
                </div>
                <span style={{ color: C.muted, fontSize: 11, marginLeft: 8, flexShrink: 0 }}>{expanded === dayIdx ? "▲" : "▼"}</span>
              </button>
              {expanded === dayIdx && (
                <div style={{ padding: "0 12px 12px" }}>
                  <MealRow label="Breakfast" value={day.breakfast} options={FOODS.breakfast} onChange={v => changeMeal(dayIdx, "breakfast", null, v)} />
                  <div style={{ fontSize: 10, color: C.muted, margin: "10px 0 5px", textTransform: "uppercase", letterSpacing: 1, fontWeight: 600 }}>Lunch</div>
                  <MealRow label="Roti" value={day.lunch?.roti} options={FOODS.roti} onChange={v => changeMeal(dayIdx, "lunch", "roti", v)} />
                  <MealRow label="Sabzi" value={day.lunch?.sabzi} options={FOODS.sabzi} onChange={v => changeMeal(dayIdx, "lunch", "sabzi", v)} />
                  <div style={{ fontSize: 11, color: C.muted, padding: "4px 2px" }}>Protein: Usal (dry) · always</div>
                  <div style={{ borderTop: `1px solid ${C.border}`, margin: "8px 0", paddingTop: 8 }}>
                    <div style={{ fontSize: 11, color: C.muted, marginBottom: 6 }}>Evening: Cosmix shake</div>
                  </div>
                  <div style={{ fontSize: 10, color: C.muted, margin: "4px 0 5px", textTransform: "uppercase", letterSpacing: 1, fontWeight: 600 }}>Dinner</div>
                  <MealRow label="Rice" value={day.dinner?.rice} options={FOODS.rice} onChange={v => changeMeal(dayIdx, "dinner", "rice", v)} />
                  <MealRow label="Sabzi" value={day.dinner?.sabzi} options={FOODS.sabzi} onChange={v => changeMeal(dayIdx, "dinner", "sabzi", v)} />
                  <MealRow label="Salad" value={day.dinner?.salad} options={FOODS.salad} onChange={v => changeMeal(dayIdx, "dinner", "salad", v)} />
                  <div style={{ fontSize: 11, color: C.muted, padding: "4px 2px" }}>+ 2 boiled eggs · always</div>
                </div>
              )}
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <button onClick={handleSuggest} style={{ flex: 1, background: C.accentDim, border: `1px solid ${C.border}`, color: C.text, borderRadius: 10, padding: "10px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>Re-suggest</button>
            <button onClick={() => setDhruviOpen(true)} style={{ flex: 1, background: C.accentDim, border: `1px solid ${C.border}`, color: C.text, borderRadius: 10, padding: "10px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>Update from Dhruvi</button>
          </div>
        </>
      )}

      {dhruviOpen && (
        <div style={{ position: "fixed", inset: 0, background: "#000000BB", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 100 }}>
          <div style={{ background: C.bg, borderRadius: "18px 18px 0 0", padding: 20, width: "100%", maxWidth: 430 }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: C.text, marginBottom: 4 }}>Update from Dhruvi</div>
            <div style={{ fontSize: 13, color: C.muted, marginBottom: 12 }}>Tell me what she said. I'll update the plan.</div>
            <textarea value={dhruviText} onChange={e => setDhruviText(e.target.value)}
              placeholder="e.g. she said to add moong dal at lunch on weekdays, reduce rice to 4 days a week..."
              style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, color: C.text, borderRadius: 10, padding: 12, fontSize: 13, resize: "none", height: 110, boxSizing: "border-box", outline: "none", fontFamily: FONT, lineHeight: 1.5 }} />
            <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
              <button onClick={() => { setDhruviOpen(false); setDhruviText(""); }} style={{ flex: 1, background: C.surface, border: "none", color: C.muted, borderRadius: 10, padding: 12, cursor: "pointer", fontSize: 13 }}>Cancel</button>
              <button onClick={handleDhruvi} disabled={updating} style={{ flex: 2, background: C.accent, border: "none", color: C.accentText, borderRadius: 10, padding: 12, fontWeight: 700, cursor: "pointer", fontSize: 14, opacity: updating ? 0.7 : 1 }}>
                {updating ? "Updating..." : "Update plan"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function LogTab({ entries, setEntries }) {
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [messages, setMessages] = useState([]);
  const bottomRef = useRef(null);
  const initialized = useRef(false);

  useEffect(() => {
    if (!initialized.current && entries.length > 0) {
      const rebuilt = entries.flatMap(e => [
        { type: "user", text: e.raw, time: e.time },
        { type: "parsed", parsed: e.parsed, time: e.time },
      ]);
      setMessages(rebuilt);
      initialized.current = true;
    }
  }, [entries]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, loading]);

  const handleSend = async () => {
    if (!input.trim() || loading) return;
    const text = input.trim();
    const time = new Date().toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
    setInput("");
    setMessages(m => [...m, { type: "user", text, time }]);
    setLoading(true);
    const parsed = await parseLog(text);
    const entry = { raw: text, parsed, time, date: new Date().toISOString() };
    const updated = [...entries, entry];
    setEntries(updated);
    sset(SK.entries, updated);
    setMessages(m => [...m, { type: "parsed", parsed, time }]);
    setLoading(false);
  };

  const handleKey = e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); } };

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", minHeight: 0 }}>
      <div style={{ flex: 1, overflowY: "auto", padding: "12px 16px" }}>
        {messages.length === 0 && !loading && (
          <div style={{ textAlign: "center", padding: "52px 20px" }}>
            <div style={{ fontSize: 34, marginBottom: 10 }}>💬</div>
            <div style={{ fontSize: 14, color: C.text, fontWeight: 600, marginBottom: 8 }}>Just start typing</div>
            <div style={{ fontSize: 13, color: C.muted, lineHeight: 1.7, marginBottom: 12 }}>
              Tell me about your day. What you ate, how your gut felt, sleep, energy, stress. I'll extract everything Dhruvi needs.
            </div>
            <div style={{ fontSize: 12, color: C.muted, fontStyle: "italic", lineHeight: 1.6 }}>
              "had poha this morning, stomach was gassy after lunch, energy tanked around 3pm, slept badly"
            </div>
          </div>
        )}
        {messages.map((msg, i) => (
          <div key={i} style={{ marginBottom: msg.type === "user" ? 4 : 14, display: "flex", flexDirection: msg.type === "user" ? "row-reverse" : "row" }}>
            {msg.type === "user" ? (
              <div style={{ maxWidth: "80%", background: C.accent, color: C.accentText, borderRadius: "18px 18px 4px 18px", padding: "10px 14px", fontSize: 14, lineHeight: 1.5 }}>
                {msg.text}
                <div style={{ fontSize: 10, color: C.accentText + "AA", marginTop: 4, textAlign: "right" }}>{msg.time}</div>
              </div>
            ) : (
              <div style={{ maxWidth: "94%", width: "94%" }}>
                <ParsedCard parsed={msg.parsed} time={msg.time} />
              </div>
            )}
          </div>
        ))}
        {loading && (
          <div style={{ display: "flex", gap: 5, padding: "8px 12px" }}>
            {[0, 1, 2].map(i => (
              <div key={i} style={{ width: 7, height: 7, borderRadius: "50%", background: C.accent, animation: `gfpulse 1.2s ease-in-out ${i * 0.2}s infinite` }} />
            ))}
          </div>
        )}
        <div ref={bottomRef} />
      </div>
      <div style={{ padding: "10px 14px 12px", background: C.bg, borderTop: `1px solid ${C.border}`, flexShrink: 0 }}>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
          <textarea ref={useRef(null)} value={input} onChange={e => setInput(e.target.value)} onKeyDown={handleKey}
            placeholder="How was your day food-wise..."
            rows={1}
            style={{ flex: 1, background: C.surface, border: `1px solid ${C.border}`, color: C.text, borderRadius: 20, padding: "10px 16px", fontSize: 14, resize: "none", outline: "none", fontFamily: FONT, lineHeight: 1.5, maxHeight: 90 }} />
          <button onClick={handleSend} disabled={loading || !input.trim()}
            style={{ background: input.trim() ? C.accent : C.surface, border: "none", color: input.trim() ? C.accentText : C.muted, width: 40, height: 40, borderRadius: "50%", fontSize: 16, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "all 0.15s" }}>
            ↑
          </button>
        </div>
      </div>
    </div>
  );
}

function ReportsTab({ plan, entries }) {
  const [view, setView] = useState("dhruvi");
  const recent = [...entries].reverse().slice(0, 20);

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 16 }}>
      <div style={{ display: "flex", background: C.surface, borderRadius: 10, padding: 3, marginBottom: 14, gap: 3 }}>
        {[["dhruvi","Dhruvi Report"],["cook","Cook Sheet"]].map(([id, label]) => (
          <button key={id} onClick={() => setView(id)}
            style={{ flex: 1, background: view === id ? C.accent : "none", border: "none", color: view === id ? C.accentText : C.muted, borderRadius: 8, padding: "8px", fontSize: 13, fontWeight: view === id ? 600 : 400, cursor: "pointer", transition: "all 0.15s" }}>
            {label}
          </button>
        ))}
      </div>

      {view === "dhruvi" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <span style={{ fontSize: 12, color: C.muted }}>{recent.length} entries · most recent first</span>
            {recent.length > 0 && <ExportButton label="Export .xlsx" onClick={() => exportDhruviXLSX([...entries].reverse())} />}
          </div>
          {recent.length === 0 ? (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>Nothing logged yet. Start in the Log tab.</div>
          ) : recent.map((e, i) => (
            <div key={i} style={{ background: C.surface, borderRadius: 10, padding: 12, marginBottom: 8 }}>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 6 }}>
                {new Date(e.date).toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short" })} · {e.time}
              </div>
              <div style={{ fontSize: 12, color: C.muted, fontStyle: "italic", marginBottom: 8, lineHeight: 1.5 }}>
                "{e.raw.length > 90 ? e.raw.slice(0, 90) + "..." : e.raw}"
              </div>
              <ParsedCard parsed={e.parsed} time={e.time} />
            </div>
          ))}
        </>
      )}

      {view === "cook" && (
        <>
          {!plan ? (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>No plan yet. Create one in the Plan tab.</div>
          ) : (
            <>
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 10 }}>
                <CopyButton text={cookText(plan)} />
                <ExportButton label="Export .xlsx" onClick={() => exportCookXLSX(plan)} />
              </div>
              {plan.map((day, i) => (
                <div key={i} style={{ background: C.surface, borderRadius: 10, marginBottom: 8, overflow: "hidden" }}>
                  <div style={{ background: C.surfaceHigh, padding: "8px 14px", borderBottom: `1px solid ${C.border}` }}>
                    <span style={{ fontWeight: 700, fontSize: 13, color: C.text }}>{HINDI_DAYS[day.day] || day.day}</span>
                    <span style={{ fontSize: 11, color: C.muted, marginLeft: 8 }}>{day.day}</span>
                  </div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <tbody>
                      {[
                        ["नाश्ता", h(day.breakfast)],
                        ["लंच · रोटी", h(day.lunch?.roti)],
                        ["लंच · प्रोटीन", "उसल (सूखा बनाना है)"],
                        ["लंच · सब्जी", h(day.lunch?.sabzi)],
                        ["डिनर · अंडे", "2 उबले अंडे"],
                        ["डिनर · सब्जी", h(day.dinner?.sabzi)],
                        ["डिनर · राइस", h(day.dinner?.rice)],
                        ["डिनर · सलाद", `${h(day.dinner?.salad)} (बिना नमक)`],
                      ].map(([label, value], ri) => (
                        <tr key={ri} style={{ borderBottom: ri < 7 ? `1px solid ${C.border}` : "none" }}>
                          <td style={{ padding: "7px 14px", color: C.muted, width: "40%", verticalAlign: "top", lineHeight: 1.5 }}>{label}</td>
                          <td style={{ padding: "7px 14px", color: C.text, lineHeight: 1.5 }}>{value}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </>
          )}
        </>
      )}
    </div>
  );
}

export default function GutFeelings() {
  const [tab, setTab] = useState("plan");
  const [plan, setPlan] = useState(null);
  const [entries, setEntries] = useState([]);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `
      * { box-sizing: border-box; }
      body { margin: 0; background: #FFFFFF; font-family: ${FONT}; }
      @keyframes gfpulse { 0%,100%{opacity:.3;transform:scale(.8)} 50%{opacity:1;transform:scale(1.1)} }
      ::-webkit-scrollbar { width: 3px; }
      ::-webkit-scrollbar-track { background: transparent; }
      ::-webkit-scrollbar-thumb { background: #E5E5EA; border-radius: 2px; }
      textarea:focus { border-color: #000000 !important; outline: none; }
    `;
    document.head.appendChild(style);

    Promise.all([sget(SK.plan, null), sget(SK.entries, [])]).then(([p, e]) => {
      if (p) setPlan(p);
      if (e?.length) setEntries(e);
      setReady(true);
    });
  }, []);

  if (!ready) return (
    <div style={{ background: C.bg, height: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT }}>
      <div style={{ fontSize: 13, color: C.muted }}>loading...</div>
    </div>
  );

  const TABS = [
    { id: "plan", icon: "📅", label: "Plan" },
    { id: "log", icon: "💬", label: "Log" },
    { id: "reports", icon: "📊", label: "Reports" },
  ];

  return (
    <div style={{ fontFamily: FONT, background: C.bg, color: C.text, height: "100vh", maxWidth: 430, margin: "0 auto", display: "flex", flexDirection: "column" }}>
      <div style={{ padding: "14px 20px 10px", background: C.bg, borderBottom: `1px solid ${C.border}`, flexShrink: 0, display: "flex", alignItems: "baseline", gap: 10 }}>
        <span style={{ fontSize: 20, fontWeight: 700, color: C.text, letterSpacing: "-0.5px" }}>Gut Feelings</span>
        <span style={{ fontSize: 12, color: C.muted }}>
          {new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "short" })}
        </span>
      </div>
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minHeight: 0 }}>
        {tab === "plan" && <PlanTab plan={plan} setPlan={setPlan} />}
        {tab === "log" && <LogTab entries={entries} setEntries={setEntries} />}
        {tab === "reports" && <ReportsTab plan={plan} entries={entries} />}
      </div>
      <div style={{ display: "flex", background: C.bg, borderTop: `1px solid ${C.border}`, flexShrink: 0, paddingBottom: "env(safe-area-inset-bottom)" }}>
        {TABS.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)}
            style={{ flex: 1, background: "none", border: "none", padding: "9px 0 6px", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
            <span style={{ fontSize: 18 }}>{t.icon}</span>
            <span style={{ fontSize: 10, color: tab === t.id ? C.accent : C.muted, fontWeight: tab === t.id ? 700 : 400 }}>{t.label}</span>
            {tab === t.id && <div style={{ width: 18, height: 2, background: C.accent, borderRadius: 1, marginTop: 1 }} />}
          </button>
        ))}
      </div>
    </div>
  );
}
