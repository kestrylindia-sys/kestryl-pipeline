import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const CFG_BASE  = "appTVlVjKPvZ6VSAt";
const ANTHROPIC_KEY = typeof __ANTHROPIC_KEY__ !== "undefined" ? __ANTHROPIC_KEY__ : "";
const CFG_TABLE = "tbl6iZQbNEddYvfn3";
const CFG_FLD   = {
  mouserKey:"fldeEz6qCd7BRFnsu", e14Key:"fldJOYFsaxcrvMo6R",
  e14Store:"fldcCgfLjeiVlOZUb",  margin:"fldI777OsZQY3DFjv",
  bcd:"fldJjkGVg0MADwysy",       sws:"fldSltfhMeOxbksgt",
  igst:"fld3Vq8s7RHkne2Oh",      freight:"fld8rIkof2aIUiPSi",
};

const AT_BASES = [
  { id:"appqDBdhwiFxlRS6Y", name:"Kestryl India - Operations Hub" },
  { id:"appelQ6k1OmDGihcp", name:"Tender & Procurement Management" },
  { id:"appAsHtfvYnOftbSa", name:"Kestryl RFQ Tracker" },
  { id:"appgb3I0tRORlTYvX", name:"RFQ Management" },
];

const STEPS = ["Upload PDF","Extract BOM & T&C","Price Sourcing","Landed Cost & Bid","Airtable Log","Export & Submit"];

const S = {
  bg:"#0d1117", surface:"#161b22", surface2:"#21262d",
  border:"#30363d", accent:"#f97316",
  green:"#22c55e", red:"#ef4444", yellow:"#eab308", blue:"#3b82f6",
  text:"#e6edf3", muted:"#8b949e",
};

const fmt = n => n != null && !isNaN(n) ? "₹" + Number(n).toLocaleString("en-IN", { maximumFractionDigits:2 }) : "—";

function Tag({ c, children }) {
  const bg = { blue:"rgba(59,130,246,.15)", green:"rgba(34,197,94,.15)", yellow:"rgba(234,179,8,.15)", red:"rgba(239,68,68,.15)", orange:"rgba(249,115,22,.15)" };
  const tx = { blue:"#60a5fa", green:"#4ade80", yellow:"#facc15", red:"#f87171", orange:"#fb923c" };
  const col = c || "blue";
  return <span style={{ display:"inline-block", padding:"2px 8px", borderRadius:12, fontSize:10, fontWeight:700, background:bg[col]||bg.blue, color:tx[col]||tx.blue }}>{children}</span>;
}

function Btn({ onClick, disabled, children, v, sz }) {
  const variant = v || "primary";
  const size    = sz || "md";
  const base = { padding: size==="sm" ? "3px 10px" : "7px 16px", fontSize: size==="sm" ? 11 : 12, fontWeight:700, borderRadius:6, border:"none", cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1, display:"inline-flex", alignItems:"center", gap:6, transition:"opacity .15s" };
  const vs = { primary:{ background:S.accent, color:"#fff" }, secondary:{ background:S.surface2, color:S.text, border:`1px solid ${S.border}` }, green:{ background:S.green, color:"#000" } };
  return <button onClick={onClick} disabled={disabled} style={{ ...base, ...(vs[variant] || vs.primary) }}>{children}</button>;
}

function Panel({ title, badge, children, noPad }) {
  return (
    <div style={{ background:S.surface, border:`1px solid ${S.border}`, borderRadius:8, marginBottom:16, overflow:"hidden" }}>
      <div style={{ padding:"11px 16px", borderBottom:`1px solid ${S.border}`, display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
        <span style={{ fontWeight:700, fontSize:13 }}>{title}</span>
        {badge}
      </div>
      <div style={noPad ? {} : { padding:16 }}>{children}</div>
    </div>
  );
}

function Th({ children }) {
  return <th style={{ background:S.surface2, color:S.muted, fontWeight:600, padding:"7px 10px", textAlign:"left", borderBottom:`1px solid ${S.border}`, fontSize:11, whiteSpace:"nowrap" }}>{children}</th>;
}

function Td({ children, s }) {
  return <td style={{ padding:"7px 10px", borderBottom:`1px solid ${S.border}`, verticalAlign:"middle", fontSize:12, ...s }}>{children}</td>;
}

function Dot({ state }) {
  const colors = { idle:S.muted, loading:S.yellow, ok:S.green, err:S.red };
  return <span style={{ width:8, height:8, borderRadius:"50%", background: colors[state] || S.muted, display:"inline-block", flexShrink:0 }} />;
}

function Spinner() {
  return <span style={{ display:"inline-block", width:11, height:11, border:"2px solid rgba(255,255,255,.2)", borderTopColor:"#fff", borderRadius:"50%", animation:"spin .6s linear infinite" }} />;
}

function LogBox({ lines, logRef }) {
  const colorMap = { ok:S.green, err:S.red, info:S.blue, warn:S.yellow };
  return (
    <div ref={logRef} style={{ background:S.bg, border:`1px solid ${S.border}`, borderRadius:6, padding:"9px 13px", fontFamily:"monospace", fontSize:11, maxHeight:110, overflowY:"auto" }}>
      {lines.map(function(l, i) {
        const msg  = typeof l === "string" ? l : l.msg;
        const type = typeof l === "string" ? "info" : l.type;
        return <div key={i} style={{ color: colorMap[type] || S.blue, padding:"1px 0" }}>{msg}</div>;
      })}
    </div>
  );
}

function NumInput({ label, val, set, step }) {
  return (
    <div>
      <div style={{ fontSize:11, color:S.muted, marginBottom:3 }}>{label}</div>
      <input
        type="number" value={val} step={step || 0.5}
        onChange={function(e) { set(parseFloat(e.target.value) || 0); }}
        style={{ background:S.surface2, border:`1px solid ${S.border}`, color:S.text, padding:"5px 9px", borderRadius:6, fontSize:12, width:78, outline:"none" }}
      />
    </div>
  );
}

function StatCard({ value, label }) {
  const isDeadline = label === "Submission Deadline";
  return (
    <div style={{ background:S.surface, border:`1px solid ${S.border}`, borderRadius:8, padding:"13px 15px" }}>
      <div style={{ fontSize:20, fontWeight:700, color: isDeadline ? S.red : S.text }}>{value}</div>
      <div style={{ fontSize:11, color:S.muted, marginTop:2 }}>{label}</div>
    </div>
  );
}

export default function App() {
  const [step, setStep]               = useState(0);
  const [pdfFile, setPdfFile]         = useState(null);
  const [pdfB64, setPdfB64]           = useState(null);
  const [extractState, setExtractState] = useState("idle");
  const [extractLog, setExtractLog]   = useState([{ msg:"// Upload a PDF to begin extraction.", type:"info" }]);
  const [rfq, setRfq]                 = useState(null);
  const [bom, setBom]                 = useState([]);
  const [tc, setTc]                   = useState([]);
  const [compliance, setCompliance]   = useState([]);
  const [prices, setPrices]           = useState([]);
  const [fetchState, setFetchState]   = useState("idle");
  const [pricingLog, setPricingLog]   = useState([{ msg:"// Click 'Fetch Live Prices' after extraction.", type:"info" }]);
  const [bcd, setBcd]                 = useState(0);
  const [sws, setSws]                 = useState(10);
  const [igst, setIgst]               = useState(18);
  const [margin, setMargin]           = useState(20);
  const [freight, setFreight]         = useState(500);
  const [atBase, setAtBase]           = useState(AT_BASES[0].id);
  const [atTable, setAtTable]         = useState("RFQ Pipeline");
  const [atState, setAtState]         = useState("idle");
  const [atRecordId, setAtRecordId]   = useState(null);
  const [atLog, setAtLog]             = useState([{ msg:"// Airtable MCP ready.", type:"info" }]);
  const [mouserKey, setMouserKey]         = useState("");
  const [e14Key, setE14Key]               = useState("");
  const [e14Store, setE14Store]           = useState("in.element14.com");
  const [nexarClientId, setNexarClientId] = useState("");
  const [nexarSecret, setNexarSecret]     = useState("");
  const [workerUrl, setWorkerUrl]         = useState("");
  const [configState, setConfigState]     = useState("loading");

  const extractLogRef = useRef(null);
  const pricingLogRef = useRef(null);
  const atLogRef      = useRef(null);

  useEffect(function() {
    async function loadConfig() {
      try {
        // On Netlify: use env vars injected at build time
        const envWorker  = typeof __WORKER_URL__  !== "undefined" ? __WORKER_URL__  : "";
        const envAirtable= typeof __AIRTABLE_PAT__ !== "undefined" ? __AIRTABLE_PAT__ : "";
        if (envWorker)  setWorkerUrl(envWorker);
        if (envBcd    != null) setBcd(0);

        // Fetch config record from Airtable REST API directly
        if (!envAirtable) { setConfigState("ok"); return; }
        const resp = await fetch(
          "https://api.airtable.com/v0/" + CFG_BASE + "/" + CFG_TABLE + "?maxRecords=1",
          { headers: { "Authorization": "Bearer " + envAirtable } }
        );
        if (!resp.ok) { setConfigState("ok"); return; }
        const data    = await resp.json();
        const fields  = data.records?.[0]?.fields || {};
        // Map by field name (Netlify deployment uses field names not IDs)
        if (fields["Mouser API Key"])      setMouserKey(fields["Mouser API Key"]);
        if (fields["Element14 API Key"])   setE14Key(fields["Element14 API Key"]);
        if (fields["Element14 Store ID"])  setE14Store(fields["Element14 Store ID"]);
        if (fields["Nexar Client ID"])     setNexarClientId(fields["Nexar Client ID"]);
        if (fields["Nexar Client Secret"]) setNexarSecret(fields["Nexar Client Secret"]);
        if (fields["Worker URL"])          setWorkerUrl(fields["Worker URL"]);
        if (fields["Default Margin %"]    != null) setMargin(fields["Default Margin %"]);
        if (fields["Default BCD %"]       != null) setBcd(fields["Default BCD %"]);
        if (fields["Default SWS %"]       != null) setSws(fields["Default SWS %"]);
        if (fields["Default IGST %"]      != null) setIgst(fields["Default IGST %"]);
        if (fields["Default Freight INR"] != null) setFreight(fields["Default Freight INR"]);
        setConfigState("ok");
      } catch(e) {
        setConfigState("err");
      }
    }
    loadConfig();
  }, []);

  function scrollLog(ref) { setTimeout(function() { if (ref.current) ref.current.scrollTop = 9999; }, 50); }
  function addExtLog(msg, type) { setExtractLog(function(l) { return [...l, { msg:msg, type:type||"info" }]; }); scrollLog(extractLogRef); }
  function addPLog(msg, type)   { setPricingLog(function(l) { return [...l, { msg:msg, type:type||"info" }]; }); scrollLog(pricingLogRef); }
  function addAtLog(msg, type)  { setAtLog(function(l)      { return [...l, { msg:msg, type:type||"info" }]; }); scrollLog(atLogRef); }

  function getBest(i) {
    const p = prices[i];
    if (!p) return 0;
    if (p.mouserPrice && p.e14Price) return Math.min(p.mouserPrice, p.e14Price);
    return p.mouserPrice || p.e14Price || 0;
  }

  function calcLine(i) {
    const base = getBest(i);
    if (!base) return { base:0, duty:0, igstAmt:0, landed:0, bid:0, line:0 };
    const duty    = base * ((bcd + sws) / 100);
    const igstAmt = (base + duty) * (igst / 100);
    const landed  = base + duty + igstAmt + (bom[i] ? freight / bom[i].qty : 0);
    const bid     = landed * (1 + margin / 100);
    return { base:base, duty:duty, igstAmt:igstAmt, landed:landed, bid:bid, line: bid * (bom[i] ? bom[i].qty : 1) };
  }

  const totals   = bom.reduce(function(s, _, i) { return s + calcLine(i).line; }, 0);
  const totalGst = totals * 0.18;
  const liveKeys  = (mouserKey && !mouserKey.startsWith("PASTE")) || (e14Key && !e14Key.startsWith("PASTE"));
  const nexarReady = !!(nexarClientId && nexarSecret && !nexarClientId.startsWith("PASTE") && !nexarSecret.startsWith("PASTE"));

  function onFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setPdfFile(file);
    setExtractState("idle");
    setExtractLog([
      { msg:"// File: " + file.name + " (" + (file.size/1024).toFixed(1) + " KB)", type:"info" },
      { msg:"// Click 'Extract with Claude AI' to begin.", type:"info" }
    ]);
    const reader = new FileReader();
    reader.onload = function(ev) { setPdfB64(ev.target.result.split(",")[1]); };
    reader.readAsDataURL(file);
  }

  async function extractPDF() {
    if (!pdfB64) { addExtLog("No PDF loaded.", "err"); return; }
    setExtractState("loading");
    addExtLog("Sending PDF to Claude API…", "info");
    const prompt = "Extract all information from this ECIL/BEL/HAL/DRDO RFQ PDF and return ONLY valid JSON (no markdown) in this exact structure: {\"rfqNumber\":\"\",\"buyer\":\"\",\"dept\":\"\",\"scope\":\"\",\"deadline\":\"\",\"bidOpening\":\"\",\"delivery\":\"\",\"payment\":\"\",\"warranty\":\"\",\"ld\":\"\",\"coc\":\"\",\"emd\":\"\",\"bidValidity\":\"\",\"optionClause\":\"\",\"evaluation\":\"\",\"consignee\":\"\",\"bom\":[{\"id\":1,\"pn\":\"\",\"desc\":\"\",\"mfr\":\"\",\"cat\":\"\",\"qty\":0,\"delivery\":\"\"}],\"tc\":[{\"label\":\"\",\"val\":\"\"}],\"compliance\":[{\"item\":\"\",\"req\":\"\",\"status\":\"ok\"}]}";
    try {
      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{ "Content-Type":"application/json", "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({
          model:"claude-sonnet-4-6",
          max_tokens:4000,
          messages:[{ role:"user", content:[
            { type:"document", source:{ type:"base64", media_type:"application/pdf", data:pdfB64 } },
            { type:"text", text:prompt }
          ]}],
        }),
      });
      const data = await resp.json();
      if (data.error) { addExtLog("API error: " + data.error.message, "err"); setExtractState("err"); return; }
      const raw = (data.content || []).filter(function(b) { return b.type === "text"; }).map(function(b) { return b.text; }).join("");
      addExtLog("Parsing JSON…", "info");
      let parsed;
      try {
        parsed = JSON.parse(raw.replace(/```json|```/g, "").trim());
      } catch(pe) {
        const m = raw.match(/\{[\s\S]*\}/);
        if (m) { parsed = JSON.parse(m[0]); }
        else { addExtLog("Could not parse JSON from response.", "err"); setExtractState("err"); return; }
      }
      setRfq({
        rfqNumber:parsed.rfqNumber||"", buyer:parsed.buyer||"", dept:parsed.dept||"",
        scope:parsed.scope||"", deadline:parsed.deadline||"", bidOpening:parsed.bidOpening||"",
        delivery:parsed.delivery||"", payment:parsed.payment||"", warranty:parsed.warranty||"",
        ld:parsed.ld||"", coc:parsed.coc||"", emd:parsed.emd||"",
        bidValidity:parsed.bidValidity||"", optionClause:parsed.optionClause||"",
        evaluation:parsed.evaluation||"", consignee:parsed.consignee||"",
      });
      const bomData = (parsed.bom || []).map(function(r, i) {
        return { id:r.id||i+1, pn:r.pn||"", desc:r.desc||"", mfr:r.mfr||"", cat:r.cat||"Component", qty:Number(r.qty)||0, delivery:r.delivery||"" };
      });
      setBom(bomData);
      setPrices(bomData.map(function() { return { mouserPrice:null, e14Price:null, stock:"—" }; }));
      setFetchState("idle");
      setTc(parsed.tc || []);
      setCompliance(parsed.compliance || []);
      addExtLog("✓ " + bomData.length + " BOM line(s) · " + (parsed.tc||[]).length + " T&C clauses", "ok");
      addExtLog("✓ RFQ: " + (parsed.rfqNumber||"") + " · Deadline: " + (parsed.deadline||""), "ok");
      setExtractState("done");
    } catch(e) {
      addExtLog("✗ " + e.message, "err");
      setExtractState("err");
    }
  }

  const workerReady = !!(workerUrl && workerUrl.includes("workers.dev"));

  // ── Direct Worker call — works in real browser (Netlify), no CORS issues ──
  async function fetchPrices() {
    if (!bom.length) { addPLog("No BOM items.", "err"); return; }
    if (!workerReady) { addPLog("⚠ Worker URL not set in Airtable Config.", "warn"); return; }
    setFetchState("fetching");
    addPLog("✓ Worker: " + workerUrl, "ok");
    addPLog("Fetching live prices for " + bom.length + " line(s)…", "info");

    const newPrices = prices.slice();
    for (var i = 0; i < bom.length; i++) {
      var r = bom[i];
      addPLog("[" + (i+1) + "/" + bom.length + "] " + r.pn + " — " + r.mfr, "info");
      try {
        var resp = await fetch(workerUrl + "?action=price", {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
          body: JSON.stringify({ pn: r.pn, mfr: r.mfr, qty: r.qty }),
        });
        if (!resp.ok) throw new Error("Worker HTTP " + resp.status);
        var result = await resp.json();
        if (result.error) throw new Error(result.error);

        var allOffers = result.allOffers || [];
        var mp = result.mouserPrice || null;
        var ep = result.e14Price    || null;

        if (allOffers.length > 0 || mp || ep) {
          var best = allOffers.length > 0 ? allOffers[0].priceINR : (mp || ep);
          var bs   = allOffers.length > 0 ? allOffers[0].seller : (mp ? "Mouser" : "E14");
          addPLog("✓ " + allOffers.length + " distributor(s) · Best: " + fmt(best) + " (" + bs + ") · Mouser: " + (mp ? fmt(mp) : "—") + " · E14: " + (ep ? fmt(ep) : "—"), "ok");
          newPrices[i] = { mouserPrice: mp, e14Price: ep, stock: "Live", source: bs, allOffers: allOffers };
        } else {
          var dbg = result.debug ? JSON.stringify(result.debug) : "";
          addPLog("⚠ " + r.pn + " — no results. " + dbg, "warn");
          newPrices[i] = { mouserPrice: null, e14Price: null, stock: "Not found", source: "—", allOffers: [] };
        }
      } catch(e) {
        addPLog("✗ " + r.pn + ": " + e.message, "err");
        newPrices[i] = { mouserPrice: null, e14Price: null, stock: "Error", source: "—", allOffers: [] };
      }
    }
    setPrices(newPrices);
    addPLog("✓ All " + bom.length + " done.", "ok");
    setFetchState("done");
  }

  function setMouserPrice(i, val) {
    var newP = prices.slice();
    newP[i] = Object.assign({}, newP[i] || {}, { mouserPrice: val ? Math.round(parseFloat(val)) : null });
    setPrices(newP);
  }
  function setE14Price(i, val) {
    var newP = prices.slice();
    newP[i] = Object.assign({}, newP[i] || {}, { e14Price: val ? Math.round(parseFloat(val)) : null });
    setPrices(newP);
  }
  function allPricesEntered() {
    return bom.length > 0 && bom.every(function(_, i) {
      var p = prices[i] || {};
      return p.mouserPrice || p.e14Price || (p.allOffers && p.allOffers.length > 0);
    });
  }
  function openLink(url) { window.open(url, "_blank", "noopener"); }
  function sellerTagColor(name) {
    var n = (name || "").toLowerCase();
    if (n.indexOf("mouser") >= 0) return "blue";
    if (n.indexOf("element14") >= 0 || n.indexOf("newark") >= 0 || n.indexOf("farnell") >= 0) return "green";
    if (n.indexOf("arrow") >= 0) return "orange";
    return "yellow";
  }

  async function logToAirtable() {
    setAtState("loading");
    addAtLog("Connecting to Airtable · Base: " + atBase, "info");
    addAtLog("Table: \"" + atTable + "\"", "info");
    const mfrs = [...new Set(bom.map(function(r) { return r.mfr; }).filter(Boolean))].join(", ");
    const payload = {
      "RFQ Number": rfq ? rfq.rfqNumber : "",
      "Buyer": rfq ? rfq.buyer : "",
      "Scope": rfq ? rfq.scope : "",
      "Deadline": rfq ? rfq.deadline : "",
      "BOM Lines": bom.length,
      "Total Qty": bom.reduce(function(s, r) { return s + r.qty; }, 0),
      "Estimated Bid Value (ex-GST)": Math.round(totals).toString(),
      "EMD Amount": rfq ? rfq.emd : "",
      "Status": "Sourcing",
      "Delivery Terms": rfq ? rfq.delivery : "",
      "Manufacturers": mfrs,
      "Source": "Kestryl RFQ Pipeline v3",
    };
    try {
      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{ "Content-Type":"application/json", "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({
          model:"claude-sonnet-4-6",
          max_tokens:1000,
          system:"You have Airtable MCP access. Create a record in base \"" + atBase + "\" table \"" + atTable + "\" (use closest matching table if name not found) with fields: " + JSON.stringify(payload) + ". Use typecast:true. Reply with \"RECORD_CREATED: [recordId]\" on its own line.",
          messages:[{ role:"user", content:"Create the Airtable record now." }],
          mcp_servers:[{ type:"url", url:"https://mcp.airtable.com/mcp", name:"airtable-mcp" }],
        }),
      });
      const data = await resp.json();
      const textParts  = (data.content || []).filter(function(b) { return b.type === "text"; }).map(function(b) { return b.text; });
      const toolParts  = (data.content || []).filter(function(b) { return b.type === "mcp_tool_result"; }).map(function(b) { return b.content && b.content[0] ? b.content[0].text : ""; });
      const combined   = textParts.concat(toolParts).join("");
      const m = combined.match(/rec[A-Za-z0-9]{14}/);
      if (m) {
        addAtLog("✓ Record created: " + m[0], "ok");
        setAtRecordId(m[0]);
        setAtState("done");
      } else if (combined.toLowerCase().indexOf("creat") >= 0 || combined.toLowerCase().indexOf("success") >= 0) {
        addAtLog("✓ Record created successfully", "ok");
        setAtState("done");
      } else {
        addAtLog("⚠ " + combined.substring(0, 160), "warn");
        setAtState("err");
      }
    } catch(e) {
      addAtLog("✗ " + e.message, "err");
      setAtState("err");
    }
  }

  function exportXLSX() {
    if (!bom.length) { alert("No BOM data."); return; }
    const wb = XLSX.utils.book_new();

    const bomRows = bom.map(function(r, i) {
      const p = prices[i] || {};
      const c = calcLine(i);
      return [r.id, r.pn, r.desc, r.mfr, r.cat, r.qty, r.delivery, p.mouserPrice||"", p.e14Price||"", c.base, +c.landed.toFixed(2), +c.bid.toFixed(2), +c.line.toFixed(0)];
    });
    const ws1 = XLSX.utils.aoa_to_sheet([
      ["KESTRYL INDIA — BID WORKBOOK"],
      ["Tender: " + (rfq?rfq.rfqNumber:"") + "   |   Deadline: " + (rfq?rfq.deadline:"") + "   |   Generated: " + new Date().toLocaleString("en-IN")],
      [],
      ["Item No","Part Number","Description","Manufacturer","Category","Qty","Delivery","Mouser INR","E14 INR","Best Price","Landed/unit","Bid/unit","Line Total"],
    ].concat(bomRows).concat([
      [],
      ["","","","","", bom.reduce(function(s,r){return s+r.qty;},0), "","","","","","TOTAL (ex-GST):", +totals.toFixed(0)],
      ["","","","","","","","","","","","GST @18%:", +totalGst.toFixed(0)],
      ["","","","","","","","","","","","GRAND TOTAL:", +(totals+totalGst).toFixed(0)],
    ]));
    ws1["!cols"] = [{wch:8},{wch:18},{wch:26},{wch:22},{wch:12},{wch:6},{wch:12},{wch:14},{wch:14},{wch:16},{wch:18},{wch:18},{wch:16}];
    ws1["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:12} }, { s:{r:1,c:0}, e:{r:1,c:12} }];
    XLSX.utils.book_append_sheet(wb, ws1, "BOM + Pricing");

    const tcRows = tc.map(function(t) { return [t.label, t.val]; });
    const ws2 = XLSX.utils.aoa_to_sheet([["T&C — " + (rfq?rfq.rfqNumber:"")], [], ["Clause","Details"]].concat(tcRows));
    ws2["!cols"] = [{wch:28},{wch:90}];
    ws2["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:1} }];
    XLSX.utils.book_append_sheet(wb, ws2, "T&C Summary");

    const bidRows = bom.map(function(r, i) {
      const c = calcLine(i);
      const ext = +((c.bid||0) * r.qty).toFixed(0);
      const gstAmt = +(ext * igst / 100).toFixed(0);
      return [r.id, r.desc, r.pn, r.mfr, r.qty, r.delivery, +c.bid.toFixed(2), ext, igst, gstAmt, ext+gstAmt];
    });
    const bidTotal = bom.reduce(function(s,_,i){const c=calcLine(i);const ext=(c.bid||0)*bom[i].qty;return s+ext+(ext*igst/100);},0);
    const ws3 = XLSX.utils.aoa_to_sheet([
      ["PRICE BID FORMAT — Tender: " + (rfq?rfq.rfqNumber:"")], ["Annexure — BoQ Format"], [],
      ["Item No","Description","Part Number","Make","Qty","Delivery","Unit Price (ex-GST)","Extended Value","GST %","GST Amt","Total incl GST"],
    ].concat(bidRows).concat([[],[,"","","","","","","","","GRAND TOTAL:", +bidTotal.toFixed(0)]]));
    ws3["!cols"] = [{wch:8},{wch:26},{wch:18},{wch:22},{wch:6},{wch:14},{wch:20},{wch:18},{wch:8},{wch:16},{wch:20}];
    ws3["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:10} }, { s:{r:1,c:0}, e:{r:1,c:10} }];
    XLSX.utils.book_append_sheet(wb, ws3, "Price Bid Format");

    const cocRows = bom.map(function(r) { return [r.id,r.desc,r.pn,r.mfr,"OEM / Authorised Distributor COC","As per PO date","Pending","",""]; });
    const ws4 = XLSX.utils.aoa_to_sheet([
      ["COC CHECKLIST — " + (rfq?rfq.rfqNumber:"")],
      ["OEM/AD COC mandatory. Material rejected without COC."], [],
      ["Item No","Description","P/N","Manufacturer","COC Type","Date Code","Status","COC Ref","Remarks"],
    ].concat(cocRows));
    ws4["!cols"] = [{wch:8},{wch:26},{wch:18},{wch:22},{wch:32},{wch:18},{wch:12},{wch:20},{wch:24}];
    ws4["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:8} }, { s:{r:1,c:0}, e:{r:1,c:8} }];
    XLSX.utils.book_append_sheet(wb, ws4, "COC Checklist");

    const ws5 = XLSX.utils.aoa_to_sheet([
      ["OUTCOME TRACKING — Fill after bid opening"], [], ["Field","Value"],
      ["RFQ Number", rfq?rfq.rfqNumber:""], ["Submission Date",""],
      ["Our Bid Value (ex-GST)", +totals.toFixed(0)],
      ["Grand Total (incl GST)", +(totals+totalGst).toFixed(0)],
      ["Bid Opening Date", rfq?rfq.bidOpening:""],
      ["L1 Price",""], ["L2 Price",""], ["Our Rank",""], ["Win / Loss",""], ["Loss Reason",""], ["PO Number",""], ["PO Value",""], ["Notes",""],
    ]);
    ws5["!cols"] = [{wch:32},{wch:40}];
    ws5["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:1} }];
    XLSX.utils.book_append_sheet(wb, ws5, "Outcome Tracking");

    const fname = "Kestryl_Bid_" + (rfq ? rfq.rfqNumber : "RFQ").replace(/[\/\\:*?"<>|]/g, "_") + ".xlsx";
    XLSX.writeFile(wb, fname);
  }

  const configBadge = configState === "loading"
    ? <span style={{ background:S.surface2, border:`1px solid ${S.border}`, padding:"2px 8px", borderRadius:4, fontSize:10, color:S.muted }}>Loading config…</span>
    : workerReady
    ? <span style={{ background:"rgba(34,197,94,.1)", border:"1px solid rgba(34,197,94,.3)", padding:"2px 8px", borderRadius:4, fontSize:10, fontWeight:700, color:S.green }}>⚡ Worker Live</span>
    : <span style={{ background:"rgba(234,179,8,.1)", border:"1px solid rgba(234,179,8,.3)", padding:"2px 8px", borderRadius:4, fontSize:10, fontWeight:700, color:S.yellow }}>⚠ Worker URL missing</span>;

  const mfrs = [...new Set(bom.map(function(r){return r.mfr;}).filter(Boolean))];

  return (
    <div style={{ background:S.bg, color:S.text, fontFamily:"system-ui,sans-serif", minHeight:"100vh", fontSize:13 }}>
      <style>{"@keyframes spin{to{transform:rotate(360deg)}} input:focus,select:focus{outline:none!important;border-color:" + S.accent + "!important;} *{box-sizing:border-box;}"}</style>

      <div style={{ background:S.surface, borderBottom:`1px solid ${S.border}`, padding:"12px 20px", display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
        <span style={{ fontSize:15, fontWeight:800, color:S.accent }}>⚡ Kestryl RFQ Pipeline</span>
        <span style={{ background:S.surface2, border:`1px solid ${S.border}`, padding:"2px 8px", borderRadius:4, fontSize:10, color:S.muted, fontWeight:600 }}>v4 · Live Pricing</span>
        {configBadge}
        {rfq && rfq.rfqNumber && <span style={{ background:S.surface2, border:`1px solid ${S.border}`, padding:"3px 10px", borderRadius:20, fontSize:11, color:S.muted }}>{rfq.rfqNumber}</span>}
        {rfq && rfq.scope     && <span style={{ background:S.surface2, border:`1px solid ${S.border}`, padding:"3px 10px", borderRadius:20, fontSize:11, color:S.muted, maxWidth:240, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{rfq.scope}</span>}
        {rfq && rfq.deadline  && <span style={{ background:"rgba(239,68,68,.1)", border:"1px solid rgba(239,68,68,.3)", padding:"3px 10px", borderRadius:20, fontSize:11, color:S.red, marginLeft:"auto" }}>⏰ Due: {rfq.deadline}</span>}
      </div>

      <div style={{ display:"flex", background:S.surface, borderBottom:`1px solid ${S.border}`, overflowX:"auto" }}>
        {STEPS.map(function(s, i) {
          const isActive = i === step;
          const isDone   = i < step;
          const canClick = i === 0 || (rfq && i > 0);
          return (
            <div key={i} onClick={function() { if (canClick) setStep(i); }}
              style={{ padding:"9px 16px", fontSize:12, fontWeight:500, cursor: canClick ? "pointer" : "not-allowed", borderBottom:"2px solid " + (isActive ? S.accent : "transparent"), color: isDone ? S.green : isActive ? S.accent : S.muted, whiteSpace:"nowrap", display:"flex", alignItems:"center", gap:6, opacity: (!rfq && i > 0) ? 0.5 : 1 }}>
              <span style={{ width:18, height:18, borderRadius:"50%", border:"1.5px solid currentColor", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:700, background: isDone ? S.green : "transparent", color: isDone ? "#000" : "currentColor", flexShrink:0 }}>
                {isDone ? "✓" : i+1}
              </span>
              {s}
            </div>
          );
        })}
      </div>

      <div style={{ padding:20, maxWidth:1120, margin:"0 auto" }}>

        {step === 0 && (
          <div>
            <Panel
              title="Upload RFQ PDF"
              badge={
                <Tag c={extractState==="done"?"green":extractState==="loading"?"yellow":extractState==="err"?"red":"blue"}>
                  {extractState==="done"?"✓ Extracted":extractState==="loading"?"Extracting…":extractState==="err"?"Error":"Ready"}
                </Tag>
              }
            >
              <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                <label style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:10, border:"2px dashed " + (pdfFile ? S.accent : S.border), borderRadius:10, padding:"32px 20px", cursor:"pointer", background: pdfFile ? "rgba(249,115,22,.05)" : S.surface2 }}>
                  <span style={{ fontSize:32 }}>{pdfFile ? "📄" : "📂"}</span>
                  <span style={{ fontWeight:600, fontSize:13, color: pdfFile ? S.accent : S.muted }}>{pdfFile ? pdfFile.name : "Click to upload RFQ PDF"}</span>
                  {pdfFile && <span style={{ fontSize:11, color:S.muted }}>{(pdfFile.size/1024).toFixed(1)} KB</span>}
                  {!pdfFile && <span style={{ fontSize:11, color:S.muted }}>Supports any government tender PDF</span>}
                  <input type="file" accept=".pdf" onChange={onFileChange} style={{ display:"none" }}/>
                </label>
                <div style={{ display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
                  <Btn onClick={extractPDF} disabled={!pdfFile || extractState==="loading"}>
                    {extractState==="loading" ? <span><Spinner/> Extracting…</span> : extractState==="done" ? "✓ Re-extract" : "🤖 Extract with Claude AI"}
                  </Btn>
                  {pdfFile && (
                    <Btn v="secondary" onClick={function() { setPdfFile(null); setPdfB64(null); setExtractState("idle"); setExtractLog([{msg:"// Upload a PDF to begin.",type:"info"}]); }}>
                      ✕ Clear
                    </Btn>
                  )}
                  {extractState==="done" && <Btn v="green" onClick={function(){setStep(1);}}>View BOM & T&C →</Btn>}
                </div>
                <div>
                  <div style={{ fontSize:11, color:S.muted, marginBottom:5, fontWeight:600 }}>EXTRACTION LOG</div>
                  <LogBox lines={extractLog} logRef={extractLogRef}/>
                </div>
              </div>
            </Panel>
            {rfq && (
              <Panel title="Last Extracted RFQ" badge={<Tag c="green">✓ Loaded</Tag>}>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:10 }}>
                  {[
                    { l:"RFQ Number", v:rfq.rfqNumber },
                    { l:"Buyer", v:rfq.buyer },
                    { l:"Dept", v:rfq.dept },
                    { l:"Scope", v:rfq.scope },
                    { l:"Deadline", v:rfq.deadline },
                    { l:"BOM Lines", v:bom.length + " items" },
                  ].map(function(item) {
                    return (
                      <div key={item.l} style={{ background:S.surface2, border:`1px solid ${S.border}`, borderRadius:6, padding:"9px 12px" }}>
                        <div style={{ fontSize:10, fontWeight:700, color:S.muted, textTransform:"uppercase", letterSpacing:".5px", marginBottom:3 }}>{item.l}</div>
                        <div style={{ fontSize:12, color: item.l==="Deadline" ? S.red : item.l==="RFQ Number" ? S.accent : S.text, fontWeight: item.l==="RFQ Number" ? 700 : 400 }}>{item.v || "—"}</div>
                      </div>
                    );
                  })}
                </div>
              </Panel>
            )}
          </div>
        )}

        {step === 1 && (
          <div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:16 }}>
              <StatCard value={bom.length} label="BOM Line Items"/>
              <StatCard value={mfrs.length} label="Manufacturers"/>
              <StatCard value={bom.reduce(function(s,r){return s+r.qty;},0)} label="Total Qty (pcs)"/>
              <StatCard value={rfq ? rfq.deadline : "—"} label="Submission Deadline"/>
            </div>
            <Panel title="Bill of Quantity" badge={<Tag c="green">✓ {bom.length} item{bom.length !== 1 ? "s" : ""}</Tag>} noPad>
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><Th>#</Th><Th>Part Number</Th><Th>Description</Th><Th>Manufacturer</Th><Th>Category</Th><Th>Qty</Th><Th>Delivery</Th></tr></thead>
                  <tbody>
                    {bom.map(function(r) {
                      return (
                        <tr key={r.id}>
                          <Td s={{ color:S.muted }}>{r.id}</Td>
                          <Td s={{ color:S.accent, fontWeight:700 }}>{r.pn || "—"}</Td>
                          <Td>{r.desc}</Td>
                          <Td s={{ fontSize:11 }}>{r.mfr}</Td>
                          <Td><Tag c="blue">{r.cat}</Tag></Td>
                          <Td s={{ fontWeight:700 }}>{r.qty}</Td>
                          <Td s={{ fontSize:11, color:S.muted }}>{r.delivery}</Td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Panel>
            <Panel title="Terms & Conditions" badge={<Tag c="blue">{tc.length} clauses</Tag>}>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))", gap:9 }}>
                {tc.map(function(t, i) {
                  return (
                    <div key={i} style={{ background:S.surface2, border:`1px solid ${S.border}`, borderRadius:6, padding:"9px 12px" }}>
                      <div style={{ fontSize:10, fontWeight:700, color:S.muted, textTransform:"uppercase", marginBottom:3 }}>{t.label}</div>
                      <div style={{ fontSize:12, lineHeight:1.5 }}>{t.val}</div>
                    </div>
                  );
                })}
              </div>
            </Panel>
            {compliance.length > 0 && (
              <Panel title="Compliance Checklist" noPad>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><Th>Item</Th><Th>Requirement</Th><Th>Status</Th></tr></thead>
                  <tbody>
                    {compliance.map(function(c, i) {
                      return (
                        <tr key={i}>
                          <Td>{c.item}</Td>
                          <Td s={{ color:S.muted, fontSize:11 }}>{c.req}</Td>
                          <Td><Tag c={c.status==="ok"?"green":c.status==="action"?"red":"yellow"}>{c.status==="ok"?"✓ OK":c.status==="action"?"⚠ Action":"Verify"}</Tag></Td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </Panel>
            )}
            <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
              <Btn v="secondary" onClick={function(){setStep(0);}}>← Back</Btn>
              <Btn onClick={function(){setStep(2);}}>Next: Source Prices →</Btn>
            </div>
          </div>
        )}

        {step === 2 && (
          <div>
            <Panel title="Price Sourcing — Live via Cloudflare Worker"
              badge={
                <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                  <span style={{ display:"flex", alignItems:"center", gap:6, background:S.surface2, border:`1px solid ${S.border}`, padding:"4px 10px", borderRadius:6, fontSize:11 }}>
                    <Dot state={fetchState==="done"?"ok":fetchState==="fetching"?"loading":"idle"}/>
                    {fetchState==="done" ? bom.length+"/"+bom.length+" sourced" : fetchState==="fetching" ? "Fetching…" : workerReady ? "Worker ready" : "Worker not set"}
                  </span>
                  <Btn onClick={fetchPrices} disabled={fetchState==="fetching"||fetchState==="done"||!workerReady}>
                    {fetchState==="done" ? "✓ Prices fetched" : fetchState==="fetching" ? <span><Spinner/> Fetching…</span> : "⚡ Fetch Live Prices"}
                  </Btn>
                  {fetchState==="done" && <Btn v="secondary" sz="sm" onClick={function(){ setFetchState("idle"); setPricingLog([{msg:"// Ready.",type:"info"}]); setPrices(bom.map(function(){ return {mouserPrice:null,e14Price:null,stock:"—",allOffers:[]}; })); }}>Re-fetch</Btn>}
                </div>
              }
              noPad
            >
              {/* Worker status */}
              <div style={{ padding:"10px 16px", borderBottom:`1px solid ${S.border}`, display:"flex", alignItems:"center", gap:8, fontSize:11 }}>
                <Dot state={workerReady?"ok":"idle"}/>
                <span style={{ color: workerReady ? S.green : S.muted }}>
                  {workerReady ? "Worker: " + workerUrl + " · Mouser India · Element14 · Nexar" : "Worker URL not configured — check Airtable Config"}
                </span>
              </div>

              {bom.map(function(r, i) {
                var p         = prices[i] || {};
                var best      = getBest(i);
                var allOffers = p.allOffers || [];
                var hasPrice  = !!(p.mouserPrice || p.e14Price || allOffers.length > 0);

                return (
                  <div key={r.id} style={{ borderBottom:`1px solid ${S.border}` }}>
                    {/* Part header */}
                    <div style={{ padding:"12px 16px", display:"flex", alignItems:"center", gap:10, flexWrap:"wrap", background: hasPrice ? "rgba(34,197,94,0.03)" : "transparent" }}>
                      <span style={{ background:S.surface2, border:`1px solid ${S.border}`, borderRadius:4, padding:"2px 8px", fontSize:10, color:S.muted, fontWeight:700 }}>#{r.id}</span>
                      <span style={{ fontWeight:700, color:S.accent, fontSize:13 }}>{r.pn}</span>
                      <span style={{ fontSize:12, color:S.muted }}>{r.desc}</span>
                      <span style={{ fontSize:11, color:S.muted }}>· {r.mfr}</span>
                      <span style={{ fontSize:11, color:S.muted }}>· Qty: <strong style={{color:S.text}}>{r.qty}</strong></span>
                      <span style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:12 }}>
                        {best > 0 && <span style={{ fontSize:15, fontWeight:700, color:S.green }}>Best: {fmt(best)}</span>}
                        {p.mouserPrice && <span style={{ fontSize:11, color:"#60a5fa" }}>Mouser: {fmt(p.mouserPrice)}</span>}
                        {p.e14Price    && <span style={{ fontSize:11, color:"#4ade80" }}>E14: {fmt(p.e14Price)}</span>}
                        {fetchState==="done" && !hasPrice && <span style={{ fontSize:11, color:S.yellow }}>⚠ Not found</span>}
                      </span>
                    </div>

                    {/* Distributor table */}
                    {allOffers.length > 0 && (
                      <div style={{ overflowX:"auto" }}>
                        <table style={{ width:"100%", borderCollapse:"collapse" }}>
                          <thead><tr><Th>Distributor</Th><Th>Unit Price (INR)</Th><Th>MOQ</Th><Th>Stock</Th><Th>vs Best</Th><Th>Use</Th></tr></thead>
                          <tbody>
                            {allOffers.map(function(o, oi) {
                              var isBest  = oi === 0;
                              var diffPct = allOffers[0].priceINR > 0 ? Math.round(((o.priceINR - allOffers[0].priceINR) / allOffers[0].priceINR) * 100) : 0;
                              var isM = (o.seller||"").toLowerCase().indexOf("mouser") >= 0;
                              var isE = (o.seller||"").toLowerCase().indexOf("element14") >= 0 || (o.seller||"").toLowerCase().indexOf("newark") >= 0 || (o.seller||"").toLowerCase().indexOf("farnell") >= 0;
                              return (
                                <tr key={oi} style={{ background: isBest ? "rgba(34,197,94,0.05)" : "transparent" }}>
                                  <Td><Tag c={sellerTagColor(o.seller)}>{o.seller}</Tag></Td>
                                  <Td s={{ fontWeight:700, color: isBest ? S.green : S.text }}>{fmt(o.priceINR)}</Td>
                                  <Td s={{ color:S.muted, fontSize:11 }}>{o.moq||1}</Td>
                                  <Td s={{ fontSize:11 }}>{o.stock > 0 ? o.stock+" pcs" : <span style={{color:S.yellow}}>Check</span>}</Td>
                                  <Td>{isBest ? <span style={{fontSize:10,fontWeight:700,color:S.green}}>✓ Best</span> : <span style={{fontSize:11,color:S.yellow}}>+{diffPct}%</span>}</Td>
                                  <Td>
                                    <Btn v="secondary" sz="sm" onClick={function() {
                                      var newP = prices.slice();
                                      if (isM)      newP[i] = Object.assign({}, newP[i], { mouserPrice: o.priceINR });
                                      else if (isE) newP[i] = Object.assign({}, newP[i], { e14Price: o.priceINR });
                                      else          newP[i] = Object.assign({}, newP[i], { mouserPrice: o.priceINR });
                                      setPrices(newP);
                                    }}>Use</Btn>
                                  </Td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* Manual override + quick links */}
                    <div style={{ padding:"10px 16px", background:S.surface2, borderTop:`1px solid ${S.border}`, display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
                      <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                        {[
                          { label:"🔵 Mouser", url:"https://www.mouser.in/Search/Refine?Keyword="+encodeURIComponent(r.pn) },
                          { label:"🟢 E14",    url:"https://in.element14.com/search#st="+encodeURIComponent(r.pn) },
                          { label:"🔶 Nexar",  url:"https://octopart.com/search?q="+encodeURIComponent(r.pn)+"&currency=INR" },
                        ].map(function(lnk) {
                          return <button key={lnk.label} onClick={function(){ openLink(lnk.url); }}
                            style={{ padding:"3px 8px", borderRadius:4, fontSize:10, fontWeight:600, border:`1px solid ${S.border}`, background:S.surface, color:S.muted, cursor:"pointer" }}>{lnk.label} ↗</button>;
                        })}
                      </div>
                      <span style={{ fontSize:11, color:S.muted }}>Override:</span>
                      <input type="number" placeholder="Mouser ₹" value={p.mouserPrice||""} onChange={function(e){ setMouserPrice(i, e.target.value); }}
                        style={{ background:S.bg, border:`1px solid ${S.border}`, color:S.text, padding:"4px 8px", borderRadius:6, fontSize:12, width:100, outline:"none" }}/>
                      <input type="number" placeholder="E14 ₹" value={p.e14Price||""} onChange={function(e){ setE14Price(i, e.target.value); }}
                        style={{ background:S.bg, border:`1px solid ${S.border}`, color:S.text, padding:"4px 8px", borderRadius:6, fontSize:12, width:100, outline:"none" }}/>
                    </div>
                  </div>
                );
              })}

              <div style={{ padding:"11px 16px", borderTop:`1px solid ${S.border}` }}>
                <LogBox lines={pricingLog} logRef={pricingLogRef}/>
              </div>
            </Panel>

            <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
              <Btn v="secondary" onClick={function(){setStep(1);}}>← Back</Btn>
              <Btn onClick={function(){setStep(3);}}>Next: Landed Cost →</Btn>
            </div>
          </div>
        )}

        {step === 3 && (
          <div>
            <Panel title="Landed Cost Calculator & Bid Price Builder">
              <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginBottom:16, alignItems:"flex-end" }}>
                <NumInput label="BCD %" val={bcd} set={setBcd}/>
                <NumInput label="SWS %" val={sws} set={setSws}/>
                <NumInput label="IGST %" val={igst} set={setIgst} step={1}/>
                <NumInput label="Margin %" val={margin} set={setMargin} step={1}/>
                <NumInput label="Freight ₹" val={freight} set={setFreight} step={100}/>
                <div style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}>
                  <span style={{ fontSize:11, color:S.muted }}>Delivery:</span>
                  <span style={{ fontSize:12, color:S.green, fontWeight:600 }}>FOR, ECIL Hyderabad</span>
                </div>
              </div>
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><Th>#</Th><Th>Part Number</Th><Th>Qty</Th><Th>Unit Price</Th><Th>Import Duty</Th><Th>IGST</Th><Th>Landed/unit</Th><Th>Margin</Th><Th>Bid/unit</Th><Th>Line Total</Th></tr></thead>
                  <tbody>
                    {bom.map(function(r, i) {
                      const c = calcLine(i);
                      return (
                        <tr key={r.id}>
                          <Td s={{ color:S.muted }}>{r.id}</Td>
                          <Td s={{ color:S.accent, fontWeight:700 }}>{r.pn || "—"}</Td>
                          <Td s={{ fontWeight:600 }}>{r.qty}</Td>
                          <Td>{fmt(c.base)}</Td>
                          <Td s={{ color:S.muted }}>{fmt(c.duty.toFixed(0))}</Td>
                          <Td s={{ color:S.muted }}>{fmt(c.igstAmt.toFixed(0))}</Td>
                          <Td s={{ color:S.yellow, fontWeight:600 }}>{fmt(c.landed.toFixed(2))}</Td>
                          <Td s={{ color:S.muted }}>{margin}%</Td>
                          <Td s={{ color:S.accent, fontWeight:700 }}>{fmt(c.bid.toFixed(2))}</Td>
                          <Td s={{ color:S.green, fontWeight:700 }}>{fmt(c.line.toFixed(0))}</Td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div style={{ marginTop:12, padding:"12px 14px", background:S.surface2, borderRadius:6, display:"flex", gap:24, flexWrap:"wrap" }}>
                <div><span style={{ color:S.muted, fontSize:11 }}>Total Bid (ex-GST): </span><span style={{ color:S.accent, fontSize:16, fontWeight:700 }}>{fmt(totals.toFixed(0))}</span></div>
                <div><span style={{ color:S.muted, fontSize:11 }}>GST @{igst}%: </span><span style={{ color:S.text, fontSize:14, fontWeight:600 }}>{fmt(totalGst.toFixed(0))}</span></div>
                <div><span style={{ color:S.muted, fontSize:11 }}>Grand Total: </span><span style={{ color:S.green, fontSize:16, fontWeight:700 }}>{fmt((totals+totalGst).toFixed(0))}</span></div>
                {rfq && rfq.emd && <div><span style={{ color:S.muted, fontSize:11 }}>EMD: </span><span style={{ color:S.yellow, fontSize:14, fontWeight:600 }}>{rfq.emd}</span></div>}
              </div>
            </Panel>
            <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
              <Btn v="secondary" onClick={function(){setStep(2);}}>← Back</Btn>
              <Btn onClick={function(){setStep(4);}}>Next: Log to Airtable →</Btn>
            </div>
          </div>
        )}

        {step === 4 && (
          <div>
            <Panel title="Airtable — Log RFQ to Operations Hub"
              badge={
                <div style={{ display:"flex", alignItems:"center", gap:6, background:S.surface2, border:`1px solid ${S.border}`, padding:"4px 10px", borderRadius:6, fontSize:11 }}>
                  <Dot state={atState==="done"?"ok":atState==="loading"?"loading":atState==="err"?"err":"idle"}/>
                  {atState==="done"?"Record created":atState==="loading"?"Creating…":atState==="err"?"Check table name":"Connected"}
                </div>
              }
            >
              <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:14, alignItems:"flex-end" }}>
                <div>
                  <div style={{ fontSize:11, color:S.muted, marginBottom:3 }}>Target Base</div>
                  <select value={atBase} onChange={function(e){setAtBase(e.target.value);}} style={{ background:S.surface2, border:`1px solid ${S.border}`, color:S.text, padding:"6px 10px", borderRadius:6, fontSize:12 }}>
                    {AT_BASES.map(function(b){ return <option key={b.id} value={b.id}>{b.name}</option>; })}
                  </select>
                </div>
                <div>
                  <div style={{ fontSize:11, color:S.muted, marginBottom:3 }}>Table name</div>
                  <input value={atTable} onChange={function(e){setAtTable(e.target.value);}} style={{ background:S.surface2, border:`1px solid ${S.border}`, color:S.text, padding:"6px 10px", borderRadius:6, fontSize:12, width:190 }}/>
                </div>
              </div>
              <div style={{ background:S.surface2, borderRadius:6, padding:"11px 14px", marginBottom:14 }}>
                <div style={{ fontSize:10, fontWeight:700, color:S.muted, textTransform:"uppercase", marginBottom:8 }}>Record Preview</div>
                {[
                  { k:"RFQ Number", v: rfq ? rfq.rfqNumber : "", c:"accent" },
                  { k:"Buyer",      v: rfq ? rfq.buyer    : "", c:"" },
                  { k:"Scope",      v: rfq ? rfq.scope    : "", c:"" },
                  { k:"Deadline",   v: rfq ? rfq.deadline : "", c:"red" },
                  { k:"BOM Lines",  v: bom.length + " items", c:"" },
                  { k:"Estimated Bid", v: fmt(totals.toFixed(0)) + " ex-GST", c:"accent" },
                  { k:"Status",     v: "Sourcing", c:"" },
                  { k:"Manufacturers", v: mfrs.join(", "), c:"" },
                ].map(function(item) {
                  return (
                    <div key={item.k} style={{ display:"flex", gap:8, fontSize:12, lineHeight:1.9 }}>
                      <span style={{ color:S.muted, width:210, flexShrink:0 }}>{item.k}</span>
                      <span style={{ color: item.c==="red" ? S.red : item.c==="accent" ? S.accent : S.text }}>{item.v || "—"}</span>
                    </div>
                  );
                })}
              </div>
              <LogBox lines={atLog} logRef={atLogRef}/>
              <div style={{ display:"flex", alignItems:"center", gap:12, marginTop:12 }}>
                <Btn v="green" onClick={logToAirtable} disabled={atState==="loading"||atState==="done"}>
                  {atState==="done" ? "✓ Record Created" : atState==="loading" ? <span><Spinner/> Creating…</span> : "📋 Create Airtable Record"}
                </Btn>
                {atRecordId && <span style={{ fontSize:12, color:S.green }}>✓ ID: {atRecordId}</span>}
                {atState==="err" && <Btn v="secondary" sz="sm" onClick={function(){setAtState("idle");}}>Retry</Btn>}
              </div>
            </Panel>
            <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
              <Btn v="secondary" onClick={function(){setStep(3);}}>← Back</Btn>
              <Btn onClick={function(){setStep(5);}}>Next: Export →</Btn>
            </div>
          </div>
        )}

        {step === 5 && (
          <div>
            <Panel title="Export Bid Workbook (.xlsx)">
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:10, marginBottom:16 }}>
                {[
                  { num:"Sheet 1", title:"BOM + Pricing",       desc:"Items · Mouser/E14 · Landed · Bid price" },
                  { num:"Sheet 2", title:"T&C Summary",         desc:"All extracted clauses" },
                  { num:"Sheet 3", title:"Price Bid Format",    desc:"ECIL BoQ · GST column" },
                  { num:"Sheet 4", title:"COC Checklist",       desc:"OEM COC tracking per line item" },
                  { num:"Sheet 5", title:"Outcome Tracking",    desc:"Fill after bid opening → Phase 4 analytics" },
                ].map(function(s) {
                  return (
                    <div key={s.num} style={{ background:S.surface2, border:`1px solid ${S.border}`, borderRadius:6, padding:12 }}>
                      <div style={{ fontSize:10, fontWeight:700, color:S.accent, textTransform:"uppercase", marginBottom:4 }}>{s.num} — {s.title}</div>
                      <div style={{ fontSize:11, color:S.muted }}>{s.desc}</div>
                    </div>
                  );
                })}
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                <Btn onClick={exportXLSX}>⬇ Download .xlsx Workbook</Btn>
                <span style={{ fontSize:11, color:S.muted }}>5 sheets · Ready for submission reference</span>
              </div>
            </Panel>
            <Panel title="Submission Checklist" noPad>
              <table style={{ width:"100%", borderCollapse:"collapse" }}>
                <thead><tr><Th>Item</Th><Th>Status</Th><Th>Notes</Th></tr></thead>
                <tbody>
                  {[
                    { item:"Signed & stamped RFQ copy",          status:"action",  note:"Mandatory — bid rejected without it" },
                    { item:"Technical Bid upload",               status:"prepare", note:"Upload to etenders.ecil.co.in" },
                    { item:"Price Bid entry",                    status:"prepare", note:"Financial offer section of portal" },
                    { item:"OEM/AD COC for " + mfrs.join(", "),  status:"action",  note:"Mandatory along with supply" },
                    { item:"Digital Signature Certificate",      status:"verify",  note:"Required for e-portal submission" },
                    { item:"T&C Compliance sheet",               status:"prepare", note:"Signed & stamped, each clause stated" },
                    { item:"EMD",                                status: rfq && rfq.emd ? "action" : "ok", note: rfq && rfq.emd ? rfq.emd : "Check portal — may not be applicable" },
                    { item:"Delivery date LIVE before dispatch", status:"ok",      note:"Confirm on portal before shipping" },
                  ].map(function(row) {
                    return (
                      <tr key={row.item}>
                        <Td>{row.item}</Td>
                        <Td><Tag c={row.status==="action"?"red":row.status==="ok"?"green":"yellow"}>{row.status==="action"?"⚠ Action needed":row.status==="ok"?"✓ Noted":"Prepare"}</Tag></Td>
                        <Td s={{ color:S.muted, fontSize:11 }}>{row.note}</Td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </Panel>
            <Panel title="📊 Outcome Tracking — Phase 3D" badge={<Tag c="orange">Action after bid opening</Tag>}>
              <p style={{ fontSize:12, color:S.muted, lineHeight:1.7, marginBottom:12 }}>
                Bid opening: <strong style={{ color:S.text }}>{rfq ? rfq.bidOpening : "—"}</strong>. Log outcome within 30 minutes to build your analytics dataset.
              </p>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))", gap:9 }}>
                {[
                  { l:"Our Bid Value", v: fmt(totals.toFixed(0)) + " ex-GST" },
                  { l:"Bid Opening",   v: rfq ? rfq.bidOpening : "—" },
                  { l:"Sheet 5",       v: "Fill L1 Price, Win/Loss, Reason" },
                  { l:"Airtable",      v: "Update Status → Won/Lost" },
                ].map(function(item) {
                  return (
                    <div key={item.l} style={{ background:S.surface2, border:`1px solid ${S.border}`, borderRadius:6, padding:"9px 12px" }}>
                      <div style={{ fontSize:10, fontWeight:700, color:S.muted, textTransform:"uppercase", marginBottom:3 }}>{item.l}</div>
                      <div style={{ fontSize:12 }}>{item.v}</div>
                    </div>
                  );
                })}
              </div>
            </Panel>
            <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
              <Btn v="secondary" onClick={function(){setStep(4);}}>← Back</Btn>
              <Btn v="green" onClick={function(){ alert("Pipeline complete for " + (rfq?rfq.rfqNumber:"this tender") + "\nBOM: " + bom.length + " items | Bid: " + fmt(totals.toFixed(0)) + " ex-GST\nDeadline: " + (rfq?rfq.deadline:"—") + "\n\nRecord logged to Airtable. Download .xlsx for submission reference."); }}>
                ✅ Complete — Get Submission Plan
              </Btn>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
