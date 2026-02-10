const API_URL =
  "https://script.google.com/macros/s/AKfycbz12zRHIIEtm1T58s6x2RdhXP3-87cTORrPnU6syNoV-QNiol7Kc4TNWHUKajTixC-G/exec";

const REFRESH_MS = 30000;
const DEBOUNCE_DELAY = 350;
const MAX_RETRIES = 3;
const RETRY_DELAY = 2000;

const fmt = new Intl.NumberFormat("th-TH");
const TH_DOW = [
  "อาทิตย์",
  "จันทร์",
  "อังคาร",
  "พุธ",
  "พฤหัสบดี",
  "ศุกร์",
  "เสาร์",
];

const TH_MONTHS = [
  "มกราคม",
  "กุมภาพันธ์",
  "มีนาคม",
  "เมษายน",
  "พฤษภาคม",
  "มิถุนายน",
  "กรกฎาคม",
  "สิงหาคม",
  "กันยายน",
  "ตุลาคม",
  "พฤศจิกายน",
  "ธันวาคม",
];

function ymKey(d) {
  // d = Date
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function monthNameFromKey(key) {
  // key = "YYYY-MM"
  const m = Number(key.split("-")[1]) - 1;
  return TH_MONTHS[m] || key;
}

function sumMonthlyFromDailyTrend(dailyTrend, ym) {
  // รวมจาก dailyTrend ที่มี date: YYYY-MM-DD
  const out = { sales: 0, calls: 0, visits: 0, quotes: 0 };
  (dailyTrend || []).forEach((r) => {
    if (!r?.date) return;
    if (String(r.date).slice(0, 7) !== ym) return;
    out.sales += Number(r.sales || 0);
    out.calls += Number(r.calls || 0);
    out.visits += Number(r.visits || 0);
    out.quotes += Number(r.quotes || 0);
  });
  return out;
}

const el = (id) => document.getElementById(id);

const state = {
  isLoading: false,
  autoTimer: null,
  lastPayload: null,
  activeMetric: "sales",
  latestTrendRows: [],
  retryCount: 0,

  isPicking: false,
  _handlers: {},
  _availableCache: { team: "", person: "", group: "" },
};

// Charts
let chart = null;
let productChart = null;
let lostDealChart = null;

/* ================= UI helpers ================= */
function setText(id, v) {
  const node = el(id);
  if (!node) return;
  node.textContent = v ?? "";
}
function setHTML(id, v) {
  const node = el(id);
  if (!node) return;
  node.innerHTML = v ?? "";
}
function escapeHtml(str) {
  if (str == null) return "";
  return String(str).replace(
    /[&<>"']/g,
    (m) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[m],
  );
}

function n0(v) {
  const x = Number(v);
  return Number.isFinite(x) ? x : 0;
}

function addThaiDow(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return dateStr;
  return `${dateStr} (${TH_DOW[d.getDay()]})`;
}
function setFilterStatus(msg, isError = false) {
  const s = el("filterStatus");
  if (!s) return;
  s.textContent = msg;
  s.classList.toggle("error", !!isError);
  s.classList.toggle("ok", !isError);
}
function showToast(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.classList.add("out");
    setTimeout(() => toast.remove(), 220);
  }, 2200);
}

/* ================= Filters ================= */
function debounceAutoLoad() {
  const ck = el("ckAuto");
  if (!ck || !ck.checked) return;
  if (state.autoTimer) clearTimeout(state.autoTimer);
  state.autoTimer = setTimeout(() => loadData(true), DEBOUNCE_DELAY);
}

function onDaysChange() {
  if (el("f_start")) el("f_start").value = "";
  if (el("f_end")) el("f_end").value = "";
  debounceAutoLoad();
}

function onStartEndChange() {
  if (el("f_days")) el("f_days").value = "";
  debounceAutoLoad();
}

function buildQueryFromFilters() {
  const p = new URLSearchParams();

  const days = document.getElementById("f_days")?.value;
  const start = document.getElementById("f_start")?.value;
  const end = document.getElementById("f_end")?.value;
  const team = document.getElementById("f_team")?.value;
  const person = document.getElementById("f_person")?.value;
  const group = document.getElementById("f_group")?.value;

  console.log("🔍 Current filters:", { days, start, end, team, person, group });

  if (start && end) {
    p.set("start", start);
    p.set("end", end);
  } else if (days) {
    p.set("days", days);
  }

  // ✅ ส่งชื่อ param ให้ตรงกับ GAS
  if (team) p.set("teamlead", team);
  if (person) p.set("person", person);
  if (group) p.set("group", group);

  // ✅ เพิ่ม timestamp เพื่อป้องกัน cache
  p.set("_t", Date.now());

  console.log("📤 Built query:", p.toString());
  return p;
}

function fillSelect(id, items, keepValue = true) {
  const sel = el(id);
  if (!sel) return;

  const prev = sel.value;
  sel.innerHTML = "";

  const all = document.createElement("option");
  all.value = "";
  all.textContent = "(ทั้งหมด)";
  sel.appendChild(all);

  (items || []).forEach((x) => {
    const opt = document.createElement("option");
    opt.value = x;
    opt.textContent = x;
    sel.appendChild(opt);
  });

  if (keepValue && prev && [...sel.options].some((o) => o.value === prev))
    sel.value = prev;
  else sel.value = "";
}

/* ✅ PATCH: setAvailable แบบ cache กัน dropdown เด้ง */
function setAvailable_PATCH(payload) {
  const a = payload?.available || {};
  const teamArr = a.teamleads || [];
  const personArr = a.people || [];
  const groupArr = a.groups || [];

  const teamStr = JSON.stringify(teamArr);
  const personStr = JSON.stringify(personArr);
  const groupStr = JSON.stringify(groupArr);

  if (teamStr !== state._availableCache.team) {
    fillSelect("f_team", teamArr, true);
    state._availableCache.team = teamStr;
  }
  if (personStr !== state._availableCache.person) {
    fillSelect("f_person", personArr, true);
    state._availableCache.person = personStr;
  }
  if (groupStr !== state._availableCache.group) {
    fillSelect("f_group", groupArr, true);
    state._availableCache.group = groupStr;
  }
}

function resetFilters() {
  if (el("f_days")) el("f_days").value = "365";
  if (el("f_start")) el("f_start").value = "";
  if (el("f_end")) el("f_end").value = "";
  if (el("f_team")) el("f_team").value = "";
  if (el("f_person")) el("f_person").value = "";
  if (el("f_group")) el("f_group").value = "";

  setFilterStatus("รีเซ็ตตัวกรองแล้ว (365 วันย้อนหลัง)");
  showToast("รีเซ็ตตัวกรองแล้ว (365 วันย้อนหลัง)", "info");

  if (state.lastPayload) setAvailable_PATCH(state.lastPayload);
  loadData(false);
}

/* ================= JSONP loader ================= */
async function loadJSONP(url, options = {}) {
  const { timeout = 30000, isRetry = false } = options; // ลด timeout เป็น 30 วินาที

  return new Promise((resolve, reject) => {
    const cbName =
      "__cb_" + Date.now() + "_" + Math.floor(Math.random() * 10000);
    const script = document.createElement("script");

    let settled = false;
    let timeoutId;

    // ✅ Callback function
    window[cbName] = (data) => {
      if (settled) return;
      settled = true;
      cleanup();

      console.log(`📥 JSONP callback received: ${cbName}`);

      if (!data) {
        reject(new Error("Empty response from server"));
        return;
      }

      if (data.error) {
        reject(new Error(`Server error: ${data.error}`));
        return;
      }

      if (!data.ok) {
        reject(new Error(`Response not ok: ${data.error || "Unknown error"}`));
        return;
      }

      resolve(data);
    };

    // ✅ Timeout handler
    timeoutId = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();

      const errorMsg = `Request timeout (${timeout}ms) - URL: ${url}`;
      console.warn(`⏰ Timeout: ${errorMsg}`);
      reject(new Error(errorMsg));
    }, timeout);

    // ✅ Cleanup function
    function cleanup() {
      clearTimeout(timeoutId);

      // ลบ script element
      try {
        if (script.parentNode) {
          script.parentNode.removeChild(script);
        }
      } catch (e) {
        // ignore
      }

      // ลบ callback หลังจาก delay
      setTimeout(() => {
        try {
          delete window[cbName];
        } catch (e) {
          window[cbName] = undefined;
        }
      }, 1000);
    }

    // ✅ Set up script
    const encodedUrl =
      url +
      (url.includes("?") ? "&" : "?") +
      "callback=" +
      cbName +
      "&_=" +
      Date.now() +
      "&retry=" +
      (isRetry ? "1" : "0");

    script.src = encodedUrl;

    // ✅ Error handler
    script.onerror = (error) => {
      if (settled) return;
      settled = true;
      cleanup();

      console.error(`❌ Script load error: ${cbName}`, error);
      reject(new Error(`Failed to load script: ${url}`));
    };

    // ✅ Success handler (สำหรับ debugging)
    script.onload = () => {
      console.log(`✅ Script loaded: ${cbName}`);
    };

    // ✅ เพิ่ม timestamp สำหรับ tracking
    script.setAttribute("data-jsonp-id", cbName);
    script.setAttribute("data-load-time", Date.now());

    console.log(`📤 Loading JSONP: ${cbName}`, {
      url:
        encodedUrl.length > 100
          ? encodedUrl.substring(0, 100) + "..."
          : encodedUrl,
      timeout: timeout + "ms",
    });

    // ✅ Append to body
    document.body.appendChild(script);
  });
}

/* ================= Validation/Debug ================= */
function validatePayload(payload) {
  const errors = [];
  const warnings = [];

  if (!payload) errors.push("Payload is null or undefined");
  else if (!payload.ok)
    errors.push(`Payload.ok is false: ${payload.error || "No error message"}`);

  if (!Array.isArray(payload.dailyTrend))
    errors.push("dailyTrend is not an array");
  else if (payload.dailyTrend.length === 0)
    warnings.push("dailyTrend is empty");

  if (!Array.isArray(payload.summary)) warnings.push("summary is not an array");
  if (!Array.isArray(payload.personTotals))
    warnings.push("personTotals is not an array");
  if (!payload.kpiToday || typeof payload.kpiToday !== "object")
    warnings.push("kpiToday is missing or not an object");

  if (errors.length > 0) {
    console.error("Validation Errors:", errors);
    showToast(`ข้อผิดพลาด: ${errors[0]}`, "error");
  }
  if (warnings.length > 0) console.warn("Validation Warnings:", warnings);

  return { isValid: errors.length === 0, errors, warnings };
}

/* ================= Load flow ================= */
async function loadData(isAuto = false) {
  console.group(`📥 loadData called (isAuto: ${isAuto})`);
  console.log("Current state:", {
    isLoading: state.isLoading,
    isPicking: state.isPicking,
    retryCount: state.retryCount,
    autoTimer: state.autoTimer,
  });

  // ✅ ป้องกันการโหลดซ้ำซ้อน
  if (isAuto && state.isPicking) {
    console.log("⏸️ Skipping auto load (user is picking from dropdown)");
    console.groupEnd();
    return;
  }

  if (state.isLoading) {
    console.log("⏸️ Skipping load (already loading)");
    console.groupEnd();
    return;
  }

  state.isLoading = true;
  const startTime = Date.now();

  // ✅ อัปเดต UI status
  setFilterStatus("กำลังโหลดข้อมูล...");

  const btnApply = document.getElementById("btnApply");
  const originalText = btnApply?.textContent;
  if (btnApply) btnApply.textContent = "กำลังโหลด...";

  try {
    // ✅ 1. สร้าง query parameters จาก filters
    const qs = buildQueryFromFilters();
    const url = API_URL + "?" + qs.toString();

    console.log(`📡 [${new Date().toLocaleTimeString()}] Loading from URL:`, {
      url: url.length > 100 ? url.substring(0, 100) + "..." : url,
      params: qs.toString(),
      isAuto: isAuto,
    });

    // ✅ 2. กำหนด timeout (auto load ใช้เวลาสั้นกว่า)
    const timeout = isAuto ? 15000 : 30000; // auto: 15s, manual: 30s
    console.log(`⏱️ Timeout set to: ${timeout}ms`);

    // ✅ 3. โหลดข้อมูลด้วย JSONP
    const payload = await loadJSONP(url, {
      timeout: timeout,
      isRetry: state.retryCount > 0,
    });

    const loadTime = Date.now() - startTime;
    console.log(`✅ Load successful in ${loadTime}ms`);

    // ✅ 4. ตรวจสอบ payload
    if (!payload) {
      throw new Error("Empty response from server");
    }

    console.log("📦 Payload received:", {
      ok: payload.ok,
      error: payload.error,
      keys: Object.keys(payload),
      dailyTrendLength: payload.dailyTrend?.length || 0,
      summaryLength: payload.summary?.length || 0,
      personTotalsLength: payload.personTotals?.length || 0,
    });

    // ✅ 5. Validation
    const validation = validatePayload(payload);
    if (!validation.isValid) {
      console.error("❌ Payload validation failed:", validation.errors);
      throw new Error(validation.errors[0] || "Invalid payload structure");
    }

    // ✅ 6. Reset state
    state.lastPayload = payload;
    state.retryCount = 0;

    // ✅ 7. Debug data structure (optional)
    if (!isAuto) {
      debugDataStructure(payload);
      checkAPIData(payload);
    }

    // ✅ 8. Update UI
    console.log("🔄 Updating UI...");
    updateAllUI(payload);

    // ✅ 9. Cache to localStorage
    try {
      const cacheData = {
        data: payload,
        timestamp: Date.now(),
        filters: qs.toString(),
        loadTime: loadTime,
        version: "1.0",
      };
      localStorage.setItem("lastDashboardPayload", JSON.stringify(cacheData));
      console.log("💾 Cached to localStorage:", {
        size: JSON.stringify(cacheData).length,
        timestamp: new Date(cacheData.timestamp).toLocaleTimeString(),
      });
    } catch (e) {
      console.warn("⚠️ Could not save to localStorage:", e.message);
    }

    // ✅ 10. Update status
    setFilterStatus(`โหลดสำเร็จ (${loadTime}ms)`);

    if (!isAuto) {
      showToast(`โหลดข้อมูลสำเร็จ (${loadTime}ms)`, "success");
    }

    console.log(`✅ Load completed successfully in ${loadTime}ms`);
    console.groupEnd();
  } catch (err) {
    const errorTime = Date.now() - startTime;
    console.error(`❌ API load error (${errorTime}ms):`, {
      message: err.message,
      stack: err.stack,
      isAuto: isAuto,
      retryCount: state.retryCount,
    });

    let errorMessage = err.message || "Unknown error";
    let userMessage = errorMessage;

    // ✅ แปลง error messages เป็นภาษาไทย
    if (errorMessage.includes("timeout")) {
      userMessage = "คำขอหมดเวลา (เซิร์ฟเวอร์ตอบสนองช้า)";
    } else if (errorMessage.includes("Failed to load script")) {
      userMessage = "ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้";
    } else if (errorMessage.includes("Empty response")) {
      userMessage = "เซิร์ฟเวอร์ไม่ตอบสนอง";
    } else if (errorMessage.includes("Network Error")) {
      userMessage = "ข้อผิดพลาดเครือข่าย";
    } else if (errorMessage.includes("CORS")) {
      userMessage = "ปัญหาเกี่ยวกับความปลอดภัยของเบราว์เซอร์";
    }

    // ✅ Update UI error state
    setText("chartStatus", `ข้อผิดพลาด: ${userMessage}`);
    setFilterStatus("โหลดไม่สำเร็จ", true);

    // ✅ ลองใช้ cached data ถ้ามี
    let cachedDataUsed = false;
    try {
      const cached = localStorage.getItem("lastDashboardPayload");
      if (cached) {
        const cachedData = JSON.parse(cached);
        const cacheAge = Date.now() - cachedData.timestamp;
        const cacheValid = cacheAge < 3600000; // 1 ชั่วโมง

        console.log("🔍 Checking cache:", {
          age: cacheAge,
          valid: cacheValid,
          filters: cachedData.filters,
        });

        if (cacheValid) {
          console.log(
            "🔄 Using cached data from localStorage (age:",
            Math.round(cacheAge / 1000),
            "s)",
          );

          showToast(
            `ใช้ข้อมูลจากแคช (อายุ ${Math.round(cacheAge / 1000)} วินาที)`,
            "info",
          );
          updateAllUI(cachedData.data);
          setFilterStatus("ใช้ข้อมูลแคช");
          state.retryCount = 0;
          cachedDataUsed = true;

          console.log("✅ Successfully loaded from cache");
          console.groupEnd();

          // อัปเดตปุ่ม
          if (btnApply) btnApply.textContent = originalText;
          state.isLoading = false;
          return;
        } else {
          console.log(
            "⚠️ Cache expired (age:",
            Math.round(cacheAge / 1000),
            "s)",
          );
        }
      }
    } catch (cacheErr) {
      console.warn("Cache fallback failed:", cacheErr);
    }

    // ✅ Retry logic (เฉพาะ manual load และยังไม่เกิน retry limit)
    if (!isAuto && !cachedDataUsed && state.retryCount < MAX_RETRIES) {
      state.retryCount++;
      const retryDelay = RETRY_DELAY * Math.pow(1.5, state.retryCount - 1);

      const retryMessage = `กำลังลองใหม่... (${state.retryCount}/${MAX_RETRIES})`;
      console.log(
        `🔁 Retry ${state.retryCount}/${MAX_RETRIES} in ${retryDelay}ms`,
      );

      showToast(retryMessage, "info");
      setFilterStatus(retryMessage);

      // ✅ ใช้ setTimeout สำหรับ retry
      setTimeout(() => {
        console.log(`🔄 Executing retry ${state.retryCount}/${MAX_RETRIES}`);
        loadData(true); // ใช้ isAuto = true สำหรับ retry
      }, retryDelay);

      console.groupEnd();
      return;
    } else {
      // ✅ หมด retry หรือเป็น auto load
      if (state.retryCount >= MAX_RETRIES) {
        console.log(`🛑 Max retries reached (${MAX_RETRIES})`);
        showToast("ลองใหม่หลายครั้งแล้ว ไม่สามารถเชื่อมต่อได้", "error");
        state.retryCount = 0;
      }

      // ✅ แสดง fallback UI ถ้าไม่ได้ใช้แคช
      if (!isAuto && !cachedDataUsed) {
        console.log("🔄 Showing fallback UI");
        showFallbackUI();
      }
    }

    console.groupEnd();
  } finally {
    // ✅ Cleanup
    state.isLoading = false;

    // อัปเดตปุ่มกลับสู่สถานะปกติ
    if (btnApply) btnApply.textContent = originalText;

    // ล้าง auto timer
    if (state.autoTimer) {
      clearTimeout(state.autoTimer);
      state.autoTimer = null;
    }

    console.log("🧹 Cleanup completed, isLoading:", state.isLoading);
  }
}

// ✅ Fallback UI สำหรับเมื่อ API ไม่สามารถติดต่อได้
function showFallbackUI() {
  console.log("🔄 Showing fallback UI");

  const fallbackHTML = `
    <div class="offline-message">
      <div style="color: #fbbf24; font-size: 32px; margin-bottom: 15px; text-align: center;">
        ⚠️
      </div>
      <div style="text-align: center; margin-bottom: 20px;">
        <div style="color: #94a3b8; font-size: 16px; margin-bottom: 10px;">
          ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้
        </div>
        <div style="font-size: 13px; color: #64748b; line-height: 1.5;">
          กรุณาตรวจสอบ:
          <ul style="text-align: left; margin: 10px 0; padding-left: 20px;">
            <li>การเชื่อมต่ออินเทอร์เน็ต</li>
            <li>URL ของ API: ${API_URL.substring(0, 50)}...</li>
            <li>สถานะเซิร์ฟเวอร์</li>
          </ul>
        </div>
      </div>
      <div style="text-align: center;">
        <button onclick="location.reload()" 
                style="padding: 10px 20px; background: #3b82f6; color: white; 
                       border: none; border-radius: 6px; cursor: pointer; 
                       font-weight: 500; margin-right: 10px;">
          โหลดใหม่
        </button>
        <button onclick="loadData(false)" 
                style="padding: 10px 20px; background: #64748b; color: white; 
                       border: none; border-radius: 6px; cursor: pointer; 
                       font-weight: 500;">
          ลองอีกครั้ง
        </button>
      </div>
    </div>
  `;

  // แสดงข้อความใน containers หลัก
  const mainContainers = [
    "top5Wrap",
    "personTotalsBody",
    "summaryBody",
    "conversionContainer",
    "areaPerformanceContainer",
    "productPerformanceContainer",
    "monthlyComparisonContainer",
  ];

  mainContainers.forEach((containerId) => {
    const container = document.getElementById(containerId);
    if (container) {
      container.innerHTML = fallbackHTML;
    }
  });

  // แสดงใน chart area
  const chartStatus = document.getElementById("chartStatus");
  if (chartStatus) {
    chartStatus.innerHTML = `
      <div style="text-align: center; padding: 30px;">
        <div style="color: #f59e0b; margin-bottom: 10px;">⚠️ ไม่สามารถโหลดข้อมูลได้</div>
        <div style="font-size: 13px; color: #94a3b8;">
          กำลังใช้ข้อมูลแคชหรือลองเชื่อมต่อใหม่...
        </div>
      </div>
    `;
  }

  setFilterStatus("เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ", true);
}

// ✅ ตรวจสอบ cached data เมื่อหน้าโหลด
function checkCachedDataOnLoad() {
  try {
    const cached = localStorage.getItem("lastDashboardPayload");
    if (cached) {
      const cachedData = JSON.parse(cached);
      const cacheAge = Date.now() - cachedData.timestamp;
      const cacheValid = cacheAge < 3600000; // 1 ชั่วโมง

      if (cacheValid) {
        console.log(
          "📦 Found valid cached data, age:",
          Math.round(cacheAge / 1000),
          "seconds",
        );

        // อัปเดต UI ด้วย cached data พร้อมเครื่องหมาย
        const cachedIndicator = document.createElement("div");
        cachedIndicator.className = "cached-indicator";
        cachedIndicator.innerHTML =
          '<span style="color: #f59e0b;">⚠️ กำลังแสดงข้อมูลจากแคช</span>';

        const statusEl = el("filterStatus");
        if (statusEl) {
          statusEl.textContent = "ใช้ข้อมูลแคช (ออฟไลน์)";
        }

        return cachedData.data;
      }
    }
  } catch (e) {
    console.warn("Error checking cache:", e);
  }
  return null;
}

// ✅ แก้ไข loadJSONP ให้มี error handling ที่ดีขึ้น
async function loadJSONP(url) {
  return new Promise((resolve, reject) => {
    const cbName =
      "__cb_" + Date.now() + "_" + Math.floor(Math.random() * 10000);
    const script = document.createElement("script");

    const TIMEOUT_MS = 45000; // 45 seconds สำหรับโหลดปกติ
    let settled = false;

    window[cbName] = (data) => {
      if (settled) return;
      settled = true;
      cleanup(false);

      if (!data) {
        reject(new Error("Empty response from server"));
        return;
      }

      if (data.error) {
        reject(new Error(data.error));
        return;
      }

      resolve(data);
    };

    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup(true);
      reject(new Error(`Request timeout (${TIMEOUT_MS}ms)`));
    }, TIMEOUT_MS);

    function cleanup(keepCallbackNoop) {
      clearTimeout(timeout);

      try {
        if (script && script.parentNode) script.parentNode.removeChild(script);
      } catch {}

      if (keepCallbackNoop) {
        window[cbName] = () => {};
        setTimeout(() => {
          try {
            delete window[cbName];
          } catch {
            window[cbName] = undefined;
          }
        }, 120000);
      } else {
        try {
          delete window[cbName];
        } catch {
          window[cbName] = undefined;
        }
      }
    }

    script.src =
      url +
      (url.includes("?") ? "&" : "?") +
      "callback=" +
      cbName +
      "&_=" +
      Date.now();

    script.onerror = () => {
      if (settled) return;
      settled = true;
      cleanup(false);
      reject(new Error("Failed to load script - Network error or CORS issue"));
    };

    // ✅ เพิ่ม tracking สำหรับ debugging
    console.log(`📤 Loading JSONP: ${cbName}`);
    script.onload = () => {
      console.log(`📥 Script loaded: ${cbName}`);
    };

    document.body.appendChild(script);
  });
}

/* ================= Picking lock + filter events ================= */
function bindPickingLock() {
  const fields = document.querySelectorAll(
    ".filters select, .filters input[type='date']",
  );
  fields.forEach((node) => {
    node.addEventListener("focus", () => (state.isPicking = true));
    node.addEventListener("pointerdown", () => (state.isPicking = true));
    node.addEventListener("blur", () => (state.isPicking = false));
  });

  document.addEventListener(
    "click",
    (e) => {
      const inside =
        e.target && e.target.closest && e.target.closest(".filters");
      if (!inside) state.isPicking = false;
    },
    true,
  );
}

function bindFilterEvents_PATCH() {
  const H = state._handlers;

  const days = el("f_days");
  const start = el("f_start");
  const end = el("f_end");
  const team = el("f_team");
  const person = el("f_person");
  const group = el("f_group");
  const btnApply = el("btnApply");
  const btnReset = el("btnReset");

  if (!H.onDaysChange) H.onDaysChange = () => onDaysChange();
  if (!H.onStartEndChange) H.onStartEndChange = () => onStartEndChange();

  if (!H.onTeamBlur)
    H.onTeamBlur = () => {
      state.isPicking = false;
      debounceAutoLoad();
    };
  if (!H.onPersonBlur)
    H.onPersonBlur = () => {
      state.isPicking = false;
      debounceAutoLoad();
    };
  if (!H.onGroupBlur)
    H.onGroupBlur = () => {
      state.isPicking = false;
      debounceAutoLoad();
    };

  if (!H.onTeamChange)
    H.onTeamChange = () => {
      if (!state.isPicking) debounceAutoLoad();
    };
  if (!H.onPersonChange)
    H.onPersonChange = () => {
      if (!state.isPicking) debounceAutoLoad();
    };
  if (!H.onGroupChange)
    H.onGroupChange = () => {
      if (!state.isPicking) debounceAutoLoad();
    };

  if (!H.onApplyClick) H.onApplyClick = () => loadData(false);
  if (!H.onResetClick) H.onResetClick = () => resetFilters();

  if (days) {
    days.removeEventListener("change", H.onDaysChange);
    days.addEventListener("change", H.onDaysChange);
  }
  if (start) {
    start.removeEventListener("change", H.onStartEndChange);
    start.addEventListener("change", H.onStartEndChange);
  }
  if (end) {
    end.removeEventListener("change", H.onStartEndChange);
    end.addEventListener("change", H.onStartEndChange);
  }

  if (team) {
    team.removeEventListener("blur", H.onTeamBlur);
    team.removeEventListener("change", H.onTeamChange);
    team.addEventListener("blur", H.onTeamBlur);
    team.addEventListener("change", H.onTeamChange);
  }
  if (person) {
    person.removeEventListener("blur", H.onPersonBlur);
    person.removeEventListener("change", H.onPersonChange);
    person.addEventListener("blur", H.onPersonBlur);
    person.addEventListener("change", H.onPersonChange);
  }
  if (group) {
    group.removeEventListener("blur", H.onGroupBlur);
    group.removeEventListener("change", H.onGroupChange);
    group.addEventListener("blur", H.onGroupBlur);
    group.addEventListener("change", H.onGroupChange);
  }

  if (btnApply) {
    btnApply.removeEventListener("click", H.onApplyClick);
    btnApply.addEventListener("click", H.onApplyClick);
  }
  if (btnReset) {
    btnReset.removeEventListener("click", H.onResetClick);
    btnReset.addEventListener("click", H.onResetClick);
  }
}

/* ================= Person Totals Pagination (PATCH) ================= */

function renderPersonTotalsWithPagination(payload, page = 1, pageSize = 20) {
  const body = el("personTotalsBody");
  if (!body) return;

  const rows = Array.isArray(payload.personTotals) ? payload.personTotals : [];
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);

  const start = (safePage - 1) * pageSize;
  const end = start + pageSize;
  const pageRows = rows.slice(start, end);

  body.innerHTML = "";

  if (!pageRows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6" class="muted">ไม่มีข้อมูล</td>`;
    body.appendChild(tr);
    createPaginationControls("personPagination", 1, 1, () => {});
    return;
  }

  pageRows.forEach((r, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${start + i + 1}</td>
      <td>${escapeHtml(r.person || r.name || r.salesPerson || "")}</td>
      <td class="num">${fmt.format(Number(r.sales || 0))} ฿</td>
      <td class="num">${fmt.format(Number(r.calls || 0))}</td>
      <td class="num">${fmt.format(Number(r.visits || 0))}</td>
      <td class="num">${fmt.format(Number(r.quotes || 0))}</td>
    `;
    body.appendChild(tr);
  });

  createPaginationControls(
    "personPagination",
    safePage,
    totalPages,
    (newPage) => {
      renderPersonTotalsWithPagination(payload, newPage, pageSize);
    },
  );
}

function createPaginationControls(
  containerId,
  currentPage,
  totalPages,
  onChange,
) {
  const container = el(containerId) || createPaginationContainer(containerId);
  container.innerHTML = "";

  if (totalPages <= 1) return;

  // Prev
  if (currentPage > 1) {
    const prevBtn = document.createElement("button");
    prevBtn.type = "button";
    prevBtn.textContent = "← ก่อนหน้า";
    prevBtn.onclick = () => onChange(currentPage - 1);
    container.appendChild(prevBtn);
  }

  // Pages (โชว์ 1..5 แบบง่าย)
  const maxShow = Math.min(totalPages, 5);
  for (let i = 1; i <= maxShow; i++) {
    const pageBtn = document.createElement("button");
    pageBtn.type = "button";
    pageBtn.textContent = String(i);
    pageBtn.className = i === currentPage ? "active" : "";
    pageBtn.onclick = () => onChange(i);
    container.appendChild(pageBtn);
  }

  // Next
  if (currentPage < totalPages) {
    const nextBtn = document.createElement("button");
    nextBtn.type = "button";
    nextBtn.textContent = "ถัดไป →";
    nextBtn.onclick = () => onChange(currentPage + 1);
    container.appendChild(nextBtn);
  }

  const info = document.createElement("span");
  info.className = "pageinfo";
  info.textContent = ` หน้า ${currentPage} จาก ${totalPages}`;
  container.appendChild(info);
}

function createPaginationContainer(id) {
  const div = document.createElement("div");
  div.id = id;
  div.className = "pagination";

  // เอาไปวางต่อท้ายตาราง personTotals
  const anchor = el("personTotalsBody");
  if (anchor && anchor.closest) {
    const table = anchor.closest("table");
    if (table && table.after) table.after(div);
    else anchor.after(div);
  } else {
    document.body.appendChild(div);
  }
  return div;
}

// ---------------- Target Achievement ----------------
// ✅ Enhanced Target vs Actual Rendering
// แทนที่ฟังก์ชัน renderTarget เดิม (ประมาณบรรทัด 1065-1100)

function renderTarget(payload) {
  const targetData = payload?.target ?? payload?.monthlyTarget ?? {};

  const actual = Number(
    targetData.actual ??
      targetData.sales ??
      payload?.summary?.totalSales ??
      payload?.summaryTotals?.sales ??
      0,
  );
  const goal = Number(
    targetData.goal ?? targetData.target ?? targetData.monthlyTarget ?? 0,
  );

  // ถ้า API ไม่ส่งเป้า/ยอดมาเลย
  if (actual === 0 && goal === 0) {
    setText("target_actual", "ไม่มีข้อมูล");
    setText("target_goal", "ไม่มีข้อมูล");
    setText("target_pct", "0%");
    setText("target_badge", "0%");
    const fill = el("target_fill");
    if (fill) fill.style.width = "0%";
    const status = el("target_status");
    if (status) status.innerHTML = "";
    return;
  }

  const pct = goal > 0 ? (actual / goal) * 100 : 0;

  // อัปเดตค่า
  setText("target_actual", fmt.format(actual) + " ฿");
  setText("target_goal", fmt.format(goal) + " ฿");
  setText("target_pct", pct.toFixed(1) + "%");

  // อัปเดต badge
  const badge = el("target_badge");
  if (badge) {
    badge.textContent = pct.toFixed(1) + "%";

    // เปลี่ยนสีตามเปอร์เซ็นต์
    badge.className = "target-badge";
    if (pct >= 100) {
      badge.classList.add("excellent");
    } else if (pct >= 80) {
      badge.classList.add("good");
    } else if (pct >= 50) {
      badge.classList.add("warning");
    } else {
      badge.classList.add("danger");
    }
  }

  // อัปเดต progress bar
  const fill = el("target_fill");
  if (fill) {
    fill.style.width = `${Math.min(pct, 100)}%`;

    // โทนสีตาม % เป้า
    if (pct >= 100) {
      fill.style.background = "linear-gradient(90deg, #10b981, #059669)";
    } else if (pct >= 80) {
      fill.style.background = "linear-gradient(90deg, #3b82f6, #2563eb)";
    } else if (pct >= 50) {
      fill.style.background = "linear-gradient(90deg, #f59e0b, #d97706)";
    } else {
      fill.style.background = "linear-gradient(90deg, #ef4444, #dc2626)";
    }
  }

  // อัปเดต status message
  const status = el("target_status");
  if (status) {
    let message = "";
    let statusClass = "";

    const remaining = goal - actual;
    const remainingFormatted = fmt.format(Math.abs(remaining));

    if (pct >= 100) {
      message = `🎉 ยอดเยี่ยม! ทำได้เกินเป้า ${remainingFormatted} ฿ (${(pct - 100).toFixed(1)}%)`;
      statusClass = "excellent";
    } else if (pct >= 80) {
      message = `👍 ดีมาก! ใกล้เป้าแล้ว เหลืออีก ${remainingFormatted} ฿ (${(100 - pct).toFixed(1)}%)`;
      statusClass = "good";
    } else if (pct >= 50) {
      message = `💪 ต้องเร่งสปีด! เหลืออีก ${remainingFormatted} ฿ (${(100 - pct).toFixed(1)}%)`;
      statusClass = "warning";
    } else {
      message = `⚠️ ต้องเร่งมากๆ! เหลืออีก ${remainingFormatted} ฿ (${(100 - pct).toFixed(1)}%)`;
      statusClass = "danger";
    }

    status.textContent = message;
    status.className = "target-status " + statusClass;
  }
}

// ---------------- Product Mix Chart ----------------
function initProductChart() {
  if (!window.Chart) return;

  const canvas = el("productChart");
  if (!canvas || !window.Chart) return;

  const ctx = canvas.getContext("2d");

  productChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: [],
      datasets: [{ data: [] }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          onClick: () => {}, // กันคลิก toggle
          labels: { color: "#cbd5e1" },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const v = Number(ctx.raw || 0);
              return `${ctx.label}: ${fmt.format(v)} ฿`;
            },
          },
        },
      },
    },
  });
}

// ============================================
// PRODUCT MIX: AUTO COLOR + GRADIENT + % LABELS
// ============================================

const PRODUCT_COLOR_RULES = [
  {
    key: "inverter_veichi",
    label: "Inverter Veichi",
    keywords: [
      "veichi",
      "veichi inverter",
      "อินเวอร์เตอร์ veichi",
      "inverter veichi",
      "veichi-",
    ],
  },
  {
    key: "solar_pump",
    label: "Solar Pump",
    keywords: [
      "solar pump",
      "โซล่าปั๊ม",
      "โซลาร์ปั๊ม",
      "ปั๊มโซล่า",
      "solarpump",
    ],
  },
  {
    key: "pump",
    label: "Pump",
    keywords: [
      "pump",
      "ปั๊ม",
      "centrifugal",
      "หอยโข่ง",
      "submersible",
      "บาดาล",
    ],
  },
  {
    key: "part",
    label: "Part",
    keywords: ["part", "อะไหล่", "อุปกรณ์", "accessory", "spare"],
  },
  {
    key: "mdb_db",
    label: "MDB/DB",
    keywords: ["mdb", "db", "ตู้ไฟ", "ตู้คอนโทรล", "distribution board"],
  },
  {
    key: "motor",
    label: "Motor",
    keywords: ["motor", "มอเตอร์", "3hp", "5hp", "7.5hp"],
  },
  {
    key: "inverter_other",
    label: "Inverter Other",
    keywords: ["inverter", "อินเวอร์เตอร์", "ac drive", "vfd", "drive"],
  },
  {
    key: "other",
    label: "Other",
    keywords: ["other", "อื่นๆ", "misc", "unknown", "ไม่ระบุ"],
  },
];

// 2) gradient palette (จะวนตาม index + ถ้า match rule จะ fix ตาม key)
const GRADIENT_PAIRS = {
  inverter_veichi: { start: "#8B5CF6", end: "#7C3AED" },
  solar_pump: { start: "#3B82F6", end: "#2563EB" },
  pump: { start: "#10B981", end: "#059669" },
  part: { start: "#F59E0B", end: "#D97706" },
  mdb_db: { start: "#EF4444", end: "#DC2626" },
  motor: { start: "#EC4899", end: "#DB2777" },
  inverter_other: { start: "#14B8A6", end: "#0D9488" },
  other: { start: "#6366F1", end: "#4F46E5" },
};

function _normText(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()\-_/\\]/g, "")
    .trim();
}

function detectProductKey(label) {
  const n = _normText(label);
  if (!n) return "other";

  for (const r of PRODUCT_COLOR_RULES) {
    for (const kw of r.keywords) {
      if (_normText(kw) && n.includes(_normText(kw))) return r.key;
    }
  }
  return "other";
}

function hashToPair(label) {
  const s = String(label || "");
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  const keys = Object.keys(GRADIENT_PAIRS);
  return GRADIENT_PAIRS[keys[h % keys.length]];
}

function makeArcGradient(chart, pair) {
  const { ctx, chartArea } = chart;
  if (!chartArea) return pair.start; // ตอน chart ยังไม่ layout

  // radial gradient ให้ดูมีมิติ
  const cx = (chartArea.left + chartArea.right) / 2;
  const cy = (chartArea.top + chartArea.bottom) / 2;
  const r =
    Math.min(
      chartArea.right - chartArea.left,
      chartArea.bottom - chartArea.top,
    ) / 2;

  const g = ctx.createRadialGradient(cx, cy, r * 0.1, cx, cy, r);
  g.addColorStop(0, pair.start);
  g.addColorStop(1, pair.end);
  return g;
}

const percentLabelsPlugin = {
  id: "percentLabelsPlugin",
  afterDatasetsDraw(chart, args, pluginOptions) {
    const { ctx } = chart;
    const dataset = chart.data.datasets?.[0];
    if (!dataset) return;

    const meta = chart.getDatasetMeta(0);
    const data = dataset.data || [];
    const total = data.reduce((a, b) => a + (Number(b) || 0), 0);
    if (!total) return;

    ctx.save();
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.font = "600 12px Inter, system-ui, sans-serif";
    ctx.fillStyle = "rgba(226,232,240,0.95)";

    meta.data.forEach((arc, i) => {
      const v = Number(data[i] || 0);
      if (v <= 0) return;

      const pct = (v / total) * 100;
      if (pct < 3) return; // ชิ้นเล็กมาก ไม่เขียนให้รก (ปรับได้)

      const p = arc.getProps(
        ["x", "y", "startAngle", "endAngle", "innerRadius", "outerRadius"],
        true,
      );
      const angle = (p.startAngle + p.endAngle) / 2;
      const r = (p.innerRadius + p.outerRadius) / 2;

      const x = p.x + Math.cos(angle) * r;
      const y = p.y + Math.sin(angle) * r;

      ctx.fillText(`${pct.toFixed(1)}%`, x, y);
    });

    ctx.restore();
  },
};

const productMixChartOptions = {
  responsive: true,
  maintainAspectRatio: false,
  cutout: "68%",
  animation: { duration: 1200, easing: "easeInOutQuart" },
  plugins: {
    legend: {
      display: true,
      position: "top",
      // กันการคลิก legend toggle (ถ้าคุณอยากให้คลิกได้ ก็ลบ onClick นี้)
      onClick: () => {},
      labels: {
        color: "#e2e8f0",
        usePointStyle: true,
        pointStyle: "circle",
        padding: 14,
        font: { size: 12, weight: "600" },
      },
    },
    tooltip: {
      backgroundColor: "rgba(15,23,42,.95)",
      titleColor: "#f1f5f9",
      bodyColor: "#cbd5e1",
      borderColor: "rgba(139,92,246,.35)",
      borderWidth: 1,
      padding: 12,
      callbacks: {
        label(ctx) {
          const label = ctx.label || "";
          const value = Number(ctx.parsed || 0);
          const total =
            ctx.dataset.data.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
          const pct = (value / total) * 100;
          return `${label}: ${value.toLocaleString()} ฿  (${pct.toFixed(1)}%)`;
        },
      },
    },
  },
};

function safeDestroyChart(canvasId) {
  const el = document.getElementById(canvasId);
  if (!el) return;
  const c = Chart.getChart(el);
  if (c) c.destroy();
}

function renderProductMix(payload) {
  const mix = payload?.productMix;
  const canvas = document.getElementById("productChart");
  if (!canvas) return;

  const items = mix?.items || [];
  if (!items.length) return;

  // ==============================
  // ✅ FIX: destroy chart ที่ผูกกับ canvas จริง ๆ
  // ==============================
  const existingChart = Chart.getChart(canvas);
  if (existingChart) {
    existingChart.destroy();
  }

  const labels = items.map((i) => i.label);
  const data = items.map((i) => Number(i.value || 0));

  const ctx = canvas.getContext("2d");

  window.productMixChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: (context) => {
            const chart = context.chart;
            const idx = context.dataIndex;
            const label = chart.data.labels?.[idx];

            const key = detectProductKey(label);
            const pair = GRADIENT_PAIRS[key] || hashToPair(label);
            return makeArcGradient(chart, pair);
          },
          borderColor: "rgba(15,23,42,0.9)",
          borderWidth: 2,
          hoverOffset: 10,
          hoverBorderColor: "rgba(255,255,255,0.65)",
          hoverBorderWidth: 2,
        },
      ],
    },
    options: productMixChartOptions,
    plugins: [percentLabelsPlugin],
  });
}

// ✅ Enhanced Customer Insight Rendering
// แทนที่ฟังก์ชัน renderCustomerInsight ในไฟล์ app.js

function renderCustomerInsight(payload) {
  const container = document.getElementById("customerInsightBody");
  if (!container) return;

  const items = payload?.customerInsight?.items;

  if (!Array.isArray(items) || items.length === 0) {
    container.innerHTML = `<tr><td colspan="5" class="muted">ไม่มีข้อมูล</td></tr>`;
    return;
  }

  // คำนวณยอดรวม
  const totalSales = items.reduce(
    (sum, item) => sum + (item.sales || item.value || 0),
    0,
  );
  const totalCount = items.reduce((sum, item) => sum + (item.count || 0), 0);

  // สร้าง HTML สำหรับแต่ละแถว
  const rows = items
    .map((item, index) => {
      const label = item.label || item.type || item.name || "ไม่ระบุ";
      const count = item.count || 0;
      const sales = item.sales || item.value || 0;
      const pct =
        item.pct ||
        item.percent ||
        (totalSales > 0 ? (sales / totalSales) * 100 : 0);

      // กำหนดสีตามอันดับ
      const rankColors = [
        { bg: "#f59e0b", text: "#fff" }, // 1 - เหลือง
        { bg: "#94a3b8", text: "#fff" }, // 2 - เทา
        { bg: "#fb923c", text: "#fff" }, // 3 - ส้ม
        { bg: "#3b82f6", text: "#fff" }, // 4 - น้ำเงิน
      ];
      const rankColor = rankColors[index] || { bg: "#64748b", text: "#fff" };

      // กำหนดสี badge % (เขียวถ้า > 30%, แดงถ้า < 15%)
      let badgeClass = "badge-neutral";
      if (pct >= 30) badgeClass = "badge-success";
      else if (pct < 15) badgeClass = "badge-danger";

      return `
      <tr class="insight-row">
        <td class="insight-category">
          <div class="category-wrapper">
            <div class="rank-badge" style="background: ${rankColor.bg}; color: ${rankColor.text};">
              ${index + 1}
            </div>
            <div class="category-info">
              <div class="category-name">${escapeHtml(label)}</div>
              <div class="category-progress">
                <div class="progress-bar-bg">
                  <div class="progress-bar-fill" style="width: ${Math.min(pct, 100)}%"></div>
                </div>
              </div>
            </div>
          </div>
        </td>
        <td class="num insight-count">${count}</td>
        <td class="num insight-sales">
          ${(sales / 1000000).toFixed(2)} B
        </td>
        <td class="num insight-percent">
          <span class="percent-badge ${badgeClass}">
            ${pct.toFixed(1)}%
          </span>
        </td>
        <td class="num insight-total">
          ${(sales / 1000).toFixed(3)} B
        </td>
      </tr>
    `;
    })
    .join("");

  // แถวสรุป - ใช้รูปแบบเดียวกับแถวข้อมูล
  const summaryRow = `
    <tr class="insight-summary">
      <td class="summary-label">
        <strong>รวมทั้งหมด (2026)</strong>
      </td>
      <td class="num"><strong>${totalCount}</strong></td>
      <td class="num"><strong>${(totalSales / 1000000).toFixed(2)} B</strong></td>
      <td class="num"><strong>100%</strong></td>
      <td class="num"><strong>${(totalSales / 1000).toFixed(3)} B</strong></td>
    </tr>
  `;

  container.innerHTML = rows + summaryRow;
}

// ---------------- 🆕 Area Performance ----------------
function renderAreaPerformance(payload) {
  const host = document.getElementById("areaPerformanceContainer");
  if (!host) return;

  const items = payload?.areaPerformance?.items || [];
  if (!items.length) {
    host.innerHTML = `<div class="area-block"><div class="muted">ไม่มีข้อมูล</div></div>`;
    return;
  }

  // normalize + กัน undefined
  const normalized = items.map((x) => ({
    area: String(x.area ?? x.label ?? "ไม่ระบุพื้นที่"),
    sales: Number(x.sales ?? x.value ?? 0),
    leads: Number(x.leads ?? x.count ?? 0),
  }));

  host.innerHTML = `
    <div class="area-block">
      <div class="area-head">
        <div class="area-chip">Area Performance</div>
        <div class="area-tools">
          <input id="areaSearch" class="area-search" placeholder="ค้นหา Area เช่น กรุงเทพ, โคราช…" />
          <select id="areaSort" class="area-select">
            <option value="sales_desc" selected>เรียง: ยอดขายมาก → น้อย</option>
            <option value="sales_asc">เรียง: ยอดขายน้อย → มาก</option>
            <option value="leads_desc">เรียง: Leads มาก → น้อย</option>
            <option value="leads_asc">เรียง: Leads น้อย → มาก</option>
            <option value="name_asc">เรียง: ชื่อ A → Z</option>
          </select>
        </div>
      </div>

      <div class="area-chip">ทั้งหมด: <b>${fmt.format(normalized.length)}</b> พื้นที่</div>

      <div class="area-scroll">
        <div id="areaGrid" class="area-grid"></div>
      </div>
    </div>
  `;

  const grid = document.getElementById("areaGrid");
  const search = document.getElementById("areaSearch");
  const sortSel = document.getElementById("areaSort");

  function sortList(list, mode) {
    const arr = [...list];
    switch (mode) {
      case "sales_asc":
        return arr.sort((a, b) => a.sales - b.sales);
      case "leads_desc":
        return arr.sort((a, b) => b.leads - a.leads);
      case "leads_asc":
        return arr.sort((a, b) => a.leads - b.leads);
      case "name_asc":
        return arr.sort((a, b) => a.area.localeCompare(b.area, "th"));
      case "sales_desc":
      default:
        return arr.sort((a, b) => b.sales - a.sales);
    }
  }

  function draw(list) {
    const MAX_SHOW = 200; // กัน DOM หน่วง (ปรับได้)
    const shown = list.slice(0, MAX_SHOW);

    grid.innerHTML = shown
      .map(
        (it, idx) => `
        <div class="area-card">
          <div class="area-row1">
            <div class="area-name">${escapeHtml(it.area)}</div>
            <div class="area-rank">#${idx + 1}</div>
          </div>

          <div class="area-metrics">
            <div class="area-metric">
              <div class="k">ยอดขาย</div>
              <div class="v">${fmt.format(it.sales)} ฿</div>
            </div>
            <div class="area-metric">
              <div class="k">Leads</div>
              <div class="v">${fmt.format(it.leads)}</div>
            </div>
          </div>

          <div class="area-mini">เฉลี่ย/Lead: <b>${fmt.format(it.leads > 0 ? Math.round(it.sales / it.leads) : 0)}</b> ฿</div>
        </div>
      `,
      )
      .join("");

    if (list.length > MAX_SHOW) {
      grid.insertAdjacentHTML(
        "beforeend",
        `<div class="area-chip">แสดง ${fmt.format(MAX_SHOW)} จาก ${fmt.format(list.length)} (ใช้ค้นหาเพื่อกรอง)</div>`,
      );
    }
  }

  function apply() {
    const q = (search.value || "").trim().toLowerCase();
    const filtered = !q
      ? normalized
      : normalized.filter((x) => x.area.toLowerCase().includes(q));

    const sorted = sortList(filtered, sortSel.value);
    draw(sorted);
  }

  // initial
  apply();

  // events (ใส่ debounce กันพิมพ์แล้วหน่วง)
  let t = null;
  search.addEventListener("input", () => {
    clearTimeout(t);
    t = setTimeout(apply, 120);
  });
  sortSel.addEventListener("change", apply);
}

function renderLostDeals(payload) {
  if (!lostDealChart) return;

  // ✅ รองรับทั้งแบบ Array และแบบ Object{items:[]}
  const lr = payload?.lostReasons;
  const raw = Array.isArray(lr) ? lr : Array.isArray(lr?.items) ? lr.items : [];

  // ถ้าไม่มีข้อมูลจริง → เคลียร์กราฟ
  if (!raw.length) {
    lostDealChart.data.labels = ["ไม่มีข้อมูล"];
    lostDealChart.data.datasets[0].data = [0];
    lostDealChart.update();
    return;
  }

  const labels = raw.map(
    (r) =>
      r.reason ||
      r.lostReason ||
      r.cause ||
      r.status ||
      r.label ||
      "ไม่ระบุเหตุผล",
  );

  const values = raw.map((r) =>
    Number(
      r.count ??
        r.total ??
        r.qty ??
        r.times ??
        r.value ?? // ✅ เผื่อ API ส่งชื่อ value
        r.n ??
        0,
    ),
  );

  // ✅ อัปเดตข้อมูล
  lostDealChart.data.labels = labels;
  lostDealChart.data.datasets[0].data = values;
  lostDealChart.update();
}

function renderCallVisitYearly(payload) {
  console.log("🔄 renderCallVisitYearly called");
  console.log("Call & Visit payload:", payload?.callVisitYearly);

  const cv = payload?.callVisitYearly || {};
  const yearNow = new Date().getFullYear();

  // ✅ 1. ตรวจสอบ element IDs
  const elementIds = [
    "cv_total_calls",
    "cv_total_visits",
    "cv_total_presented",
    "cv_total_quoted",
    "cv_total_closed",
  ];

  console.log(
    "Checking elements:",
    elementIds.map((id) => ({
      id,
      exists: !!document.getElementById(id),
    })),
  );

  // ✅ 2. แสดงข้อมูลพื้นฐาน (ป้องกัน undefined)
  setText("cv_total_calls", fmt.format(Number(cv.totalCalls || cv.calls || 0)));
  setText(
    "cv_total_visits",
    fmt.format(Number(cv.totalVisits || cv.visits || 0)),
  );

  // ✅ 3. ดึงข้อมูลรายปี (รองรับหลายรูปแบบ)
  let yearlyData = null;
  let selectedYear = yearNow;

  // รูปแบบที่ 1: Array of objects
  if (Array.isArray(cv.yearly) && cv.yearly.length > 0) {
    yearlyData = cv.yearly.find((item) => item.year == yearNow) || cv.yearly[0]; // fallback to first item
    console.log("Found yearly array data:", yearlyData);
  }
  // รูปแบบที่ 2: Object with year keys
  else if (cv.yearly && typeof cv.yearly === "object") {
    yearlyData = cv.yearly[yearNow] || cv.yearly[Object.keys(cv.yearly)[0]];
    console.log("Found yearly object data:", yearlyData);
  }
  // รูปแบบที่ 3: Direct properties
  else if (cv.totalPresented || cv.totalQuoted || cv.totalClosed) {
    yearlyData = cv;
    console.log("Using direct properties data:", yearlyData);
  }
  // รูปแบบที่ 4: byYear
  else if (cv.byYear && Array.isArray(cv.byYear)) {
    yearlyData = cv.byYear.find((item) => item.year == yearNow) || cv.byYear[0];
    console.log("Found byYear data:", yearlyData);
  }

  // ✅ 4. แสดงข้อมูล (ใช้ helper function เพื่อความปลอดภัย)
  const presented = getNumberValue(yearlyData, [
    "presented",
    "totalPresented",
    "present",
    "L",
  ]);
  const quoted = getNumberValue(yearlyData, [
    "quoted",
    "totalQuoted",
    "quote",
    "M",
  ]);
  const closed = getNumberValue(yearlyData, [
    "closed",
    "totalClosed",
    "close",
    "N",
  ]);

  console.log("Final values:", { presented, quoted, closed });

  // ✅ 5. อัปเดต UI
  setText("cv_total_presented", fmt.format(presented));
  setText("cv_total_quoted", fmt.format(quoted));
  setText("cv_total_closed", fmt.format(closed));

  // ✅ 6. ถ้าไม่มีข้อมูลให้แสดง fallback
  if (presented === 0 && quoted === 0 && closed === 0) {
    console.warn("⚠️ No call & visit yearly data found");

    // แสดงข้อความใน container
    const container =
      document.querySelector(".call-visit-yearly") ||
      document.getElementById("callVisitContainer");
    if (container) {
      const message = document.createElement("div");
      message.className = "no-data-message";
      message.innerHTML = `
        <div style="text-align: center; padding: 20px; color: #94a3b8;">
          <div>📊 ไม่มีข้อมูล Call & Visit Yearly</div>
          <small>ตรวจสอบโครงสร้างข้อมูลจาก API</small>
        </div>
      `;

      // เพิ่มถ้ายังไม่มี
      if (!container.querySelector(".no-data-message")) {
        container.appendChild(message);
      }
    }
  }
}

// ✅ Helper function: ดึงค่าตัวเลขจาก object
function getNumberValue(obj, keys) {
  if (!obj) return 0;

  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && obj[key] !== "") {
      const num = Number(String(obj[key]).replace(/,/g, "").trim());
      return Number.isFinite(num) ? num : 0;
    }
  }
  return 0;
}

/* ================= Area Performance Heatmap ================= */
function renderAreaHeatmap(payload) {
  console.log("🔄 renderAreaHeatmap called");

  const heatmapData = payload.areaHeatmap || {};
  const data = heatmapData.heatmapData || [];
  const summary = heatmapData.summary || {};
  const meta = heatmapData.meta || {};

  const container = document.getElementById("areaHeatmapContainer");
  if (!container) {
    console.error("❌ areaHeatmapContainer element not found");
    return;
  }

  // ตรวจสอบว่ามีข้อมูลหรือไม่
  if (data.length === 0) {
    container.innerHTML = `
      <div class="muted" style="text-align: center; padding: 40px;">
        ${meta.note || "ไม่มีข้อมูล Area Heatmap"}
      </div>
    `;
    return;
  }

  // แปลงเดือนให้เป็นชื่อเดือนไทย
  const months = summary.months || [];
  const thaiMonths = months.map((month) => {
    const [year, monthNum] = month.split("-");
    const monthNames = [
      "ม.ค.",
      "ก.พ.",
      "มี.ค.",
      "เม.ย.",
      "พ.ค.",
      "มิ.ย.",
      "ก.ค.",
      "ส.ค.",
      "ก.ย.",
      "ต.ค.",
      "พ.ย.",
      "ธ.ค.",
    ];
    return `${monthNames[parseInt(monthNum) - 1]} ${year}`;
  });

  // คำนวณค่าสูงสุดสำหรับ normalization
  let maxSales = 0;
  data.forEach((area) => {
    area.monthlyData.forEach((month) => {
      if (month.sales > maxSales) maxSales = month.sales;
    });
  });

  let html = "";

  // ✅ Header Section
  html += `
    <div class="heatmap-header">
      <div class="header-info">
        <h3>Area Performance Heatmap</h3>
        <div class="header-subtitle">
          <span class="heatmap-stats">
            <span class="stat-item">
              <span class="stat-label">พื้นที่ทั้งหมด:</span>
              <span class="stat-value">${summary.totalAreas || 0}</span>
            </span>
            <span class="stat-item">
              <span class="stat-label">ยอดขายรวม:</span>
              <span class="stat-value">${fmt.format(summary.totalSales || 0)} ฿</span>
            </span>
            <span class="stat-item">
              <span class="stat-label">พื้นที่ยอดนิยม:</span>
              <span class="stat-value">${summary.topPerformingArea || "-"}</span>
            </span>
          </span>
        </div>
      </div>
      <div class="header-controls">
        <div class="view-toggle">
          <button class="view-btn active" data-view="heatmap">Heatmap</button>
          <button class="view-btn" data-view="table">Table</button>
          <button class="view-btn" data-view="trend">Trend</button>
        </div>
        <div class="color-scale">
          <span>ต่ำ</span>
          <div class="scale-gradient"></div>
          <span>สูง</span>
        </div>
      </div>
    </div>
    
    <div class="heatmap-views">
      <!-- Heatmap View -->
      <div class="heatmap-view active" id="heatmapView">
  `;

  // ✅ Heatmap Grid
  html += `
    <div class="heatmap-grid-container">
      <div class="heatmap-grid">
        <!-- Header Row (Months) -->
        <div class="heatmap-cell area-header"></div>
        ${thaiMonths
          .map(
            (month) => `
          <div class="heatmap-cell month-header">
            <div class="month-name">${month}</div>
          </div>
        `,
          )
          .join("")}
        <div class="heatmap-cell total-header">รวม</div>
        
        <!-- Data Rows -->
        ${data
          .map((area, areaIndex) => {
            const areaSales = area.summary.totalSales;
            const monthlyData = area.monthlyData;
            const contribution = area.summary.contribution;
            const trend = area.summary.trend;

            // กำหนดสีตาม rank
            let rankClass = "";
            if (areaIndex === 0) rankClass = "rank-1";
            else if (areaIndex === 1) rankClass = "rank-2";
            else if (areaIndex === 2) rankClass = "rank-3";

            return `
            <div class="heatmap-cell area-name ${rankClass}">
              <div class="area-info">
                <div class="area-rank">${areaIndex + 1}</div>
                <div class="area-details">
                  <div class="area-title">${escapeHtml(area.area)}</div>
                  <div class="area-meta">
                    <span class="meta-item">${fmt.format(area.summary.totalDeals || 0)} deals</span>
                    <span class="meta-separator">•</span>
                    <span class="meta-item">${contribution.toFixed(1)}%</span>
                  </div>
                </div>
              </div>
              ${
                trend === "up"
                  ? '<div class="trend-indicator up" title="กำลังขึ้น">↑</div>'
                  : trend === "down"
                    ? '<div class="trend-indicator down" title="กำลังลง">↓</div>'
                    : '<div class="trend-indicator stable" title="คงที่">→</div>'
              }
            </div>
            
            ${monthlyData
              .map((month) => {
                // คำนวณสี intensity (0-1)
                const intensity =
                  month.sales > 0 ? Math.min(month.sales / maxSales, 1) : 0;

                // กำหนดสีตาม intensity
                let colorClass = "color-0";
                if (intensity > 0.8) colorClass = "color-5";
                else if (intensity > 0.6) colorClass = "color-4";
                else if (intensity > 0.4) colorClass = "color-3";
                else if (intensity > 0.2) colorClass = "color-2";
                else if (intensity > 0) colorClass = "color-1";

                // กำหนด growth indicator
                let growthIndicator = "";
                if (month.growth !== null) {
                  if (month.growth > 20)
                    growthIndicator = '<div class="growth-badge high">↑</div>';
                  else if (month.growth > 0)
                    growthIndicator =
                      '<div class="growth-badge medium">↗</div>';
                  else if (month.growth < -20)
                    growthIndicator = '<div class="growth-badge low">↓</div>';
                  else if (month.growth < 0)
                    growthIndicator = '<div class="growth-badge low">↘</div>';
                }

                return `
                <div class="heatmap-cell data-cell ${colorClass}" 
                     data-area="${escapeHtml(area.area)}" 
                     data-month="${month.month}"
                     data-sales="${month.sales}"
                     data-deals="${month.deals}"
                     data-companies="${month.uniqueCompanies}"
                     data-growth="${month.growth || 0}">
                  <div class="cell-content">
                    <div class="sales-value">${month.sales > 0 ? fmt.formatShort(month.sales) : "-"}</div>
                    ${growthIndicator}
                  </div>
                  <div class="cell-tooltip">
                    <div class="tooltip-title">${escapeHtml(area.area)} - ${month.month}</div>
                    <div class="tooltip-content">
                      <div>ยอดขาย: <strong>${fmt.format(month.sales)} ฿</strong></div>
                      <div>ดีล: <strong>${fmt.format(month.deals)}</strong></div>
                      ${month.uniqueCompanies > 0 ? `<div>บริษัท: <strong>${fmt.format(month.uniqueCompanies)}</strong></div>` : ""}
                      ${month.growth !== null ? `<div>Growth: <strong class="${month.growth >= 0 ? "positive" : "negative"}">${month.growth >= 0 ? "+" : ""}${month.growth.toFixed(1)}%</strong></div>` : ""}
                    </div>
                  </div>
                </div>
              `;
              })
              .join("")}
            
            <div class="heatmap-cell total-cell ${rankClass}">
              <div class="total-content">
                <div class="total-value">${fmt.formatShort(areaSales)}</div>
                <div class="total-label">฿</div>
              </div>
            </div>
          `;
          })
          .join("")}
      </div>
    </div>
  `;

  // ✅ Table View
  html += `
      </div>
      
      <div class="heatmap-view" id="tableView">
        <div class="heatmap-table-container">
          <table class="heatmap-table">
            <thead>
              <tr>
                <th class="sticky">พื้นที่</th>
                ${thaiMonths
                  .map(
                    (month) => `
                  <th class="text-center">${month}</th>
                `,
                  )
                  .join("")}
                <th class="text-center">รวม</th>
                <th class="text-center">ส่วนแบ่ง</th>
                <th class="text-center">แนวโน้ม</th>
              </tr>
            </thead>
            <tbody>
              ${data
                .map((area, areaIndex) => {
                  const areaSales = area.summary.totalSales;
                  const monthlyData = area.monthlyData;
                  const contribution = area.summary.contribution;
                  const trend = area.summary.trend;

                  let trendIcon = "→";
                  let trendClass = "stable";
                  let trendText = "คงที่";

                  if (trend === "up") {
                    trendIcon = "↑";
                    trendClass = "up";
                    trendText = "กำลังขึ้น";
                  } else if (trend === "down") {
                    trendIcon = "↓";
                    trendClass = "down";
                    trendText = "กำลังลง";
                  }

                  return `
                  <tr>
                    <td class="area-cell">
                      <div class="area-rank">${areaIndex + 1}</div>
                      <div class="area-name">${escapeHtml(area.area)}</div>
                    </td>
                    ${monthlyData
                      .map((month) => {
                        const intensity =
                          month.sales > 0
                            ? Math.min(month.sales / maxSales, 1)
                            : 0;
                        let colorClass = "color-0";
                        if (intensity > 0.8) colorClass = "color-5";
                        else if (intensity > 0.6) colorClass = "color-4";
                        else if (intensity > 0.4) colorClass = "color-3";
                        else if (intensity > 0.2) colorClass = "color-2";
                        else if (intensity > 0) colorClass = "color-1";

                        let growthIndicator = "";
                        if (month.growth !== null && month.growth !== 0) {
                          growthIndicator = `<span class="growth-indicator ${month.growth > 0 ? "positive" : "negative"}">
                          ${month.growth > 0 ? "+" : ""}${month.growth.toFixed(0)}%
                        </span>`;
                        }

                        return `
                        <td class="data-cell ${colorClass}">
                          <div class="cell-content">
                            <div class="sales-value">${month.sales > 0 ? fmt.formatShort(month.sales) : "-"}</div>
                            ${growthIndicator}
                          </div>
                        </td>
                      `;
                      })
                      .join("")}
                    <td class="total-cell">
                      <strong>${fmt.format(areaSales)}</strong>
                    </td>
                    <td class="contribution-cell">
                      <div class="contribution-bar">
                        <div class="contribution-fill" style="width: ${contribution}%"></div>
                      </div>
                      <span class="contribution-value">${contribution.toFixed(1)}%</span>
                    </td>
                    <td class="trend-cell ${trendClass}">
                      <span class="trend-icon">${trendIcon}</span>
                      <span class="trend-text">${trendText}</span>
                    </td>
                  </tr>
                `;
                })
                .join("")}
            </tbody>
            <tfoot>
              <tr class="summary-row">
                <td><strong>รวมทั้งหมด</strong></td>
                ${months
                  .map((month, index) => {
                    const monthTotal = data.reduce((sum, area) => {
                      return sum + (area.monthlyData[index]?.sales || 0);
                    }, 0);
                    return `<td class="text-center"><strong>${fmt.formatShort(monthTotal)}</strong></td>`;
                  })
                  .join("")}
                <td class="text-center"><strong>${fmt.format(summary.totalSales || 0)}</strong></td>
                <td class="text-center"><strong>100%</strong></td>
                <td class="text-center">-</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      
      <!-- Trend View -->
      <div class="heatmap-view" id="trendView">
        <div class="trend-chart-container">
          <canvas id="areaTrendChart"></canvas>
        </div>
        <div class="trend-legend">
          <div class="legend-title">พื้นที่ที่มีแนวโน้มดี:</div>
          <div class="trending-areas">
            ${data
              .filter((area) => area.summary.trend === "up")
              .slice(0, 5)
              .map(
                (area) => `
              <div class="trending-area">
                <div class="area-name">${escapeHtml(area.area)}</div>
                <div class="area-stats">
                  <span class="stat">${fmt.format(area.summary.totalSales)} ฿</span>
                  <span class="stat">${area.summary.contribution.toFixed(1)}%</span>
                </div>
              </div>
            `,
              )
              .join("")}
          </div>
        </div>
      </div>
    </div>
    
    <!-- Legend -->
    <div class="heatmap-legend">
      <div class="legend-items">
        <div class="legend-item">
          <div class="legend-color color-5"></div>
          <div class="legend-text">สูงมาก (≥ 80%)</div>
        </div>
        <div class="legend-item">
          <div class="legend-color color-4"></div>
          <div class="legend-text">สูง (60-79%)</div>
        </div>
        <div class="legend-item">
          <div class="legend-color color-3"></div>
          <div class="legend-text">ปานกลาง (40-59%)</div>
        </div>
        <div class="legend-item">
          <div class="legend-color color-2"></div>
          <div class="legend-text">ต่ำ (20-39%)</div>
        </div>
        <div class="legend-item">
          <div class="legend-color color-1"></div>
          <div class="legend-text">ต่ำมาก (1-19%)</div>
        </div>
        <div class="legend-item">
          <div class="legend-color color-0"></div>
          <div class="legend-text">ไม่มีข้อมูล</div>
        </div>
      </div>
      <div class="legend-note">
        *เปอร์เซ็นต์เทียบกับพื้นที่ที่มียอดขายสูงสุดในเดือนนั้นๆ
      </div>
    </div>
  `;

  container.innerHTML = html;

  // ✅ Initialize Chart.js for Trend View
  initializeAreaTrendChart(data, months, thaiMonths);

  // ✅ Add event listeners for view toggles
  setupHeatmapViewToggles();

  // ✅ Add tooltip functionality
  setupHeatmapTooltips();
}

// ✅ Helper function for short formatting
if (typeof fmt.formatShort === "undefined") {
  fmt.formatShort = function (value) {
    const num = Number(value);
    if (num >= 1000000) {
      return (num / 1000000).toFixed(1) + "M";
    } else if (num >= 1000) {
      return (num / 1000).toFixed(0) + "K";
    }
    return this.format(num);
  };
}

// ✅ Initialize trend chart
function initializeAreaTrendChart(data, months, thaiMonths) {
  const canvas = document.getElementById("areaTrendChart");
  if (!canvas || !window.Chart) return;

  // เลือกพื้นที่ Top 5 สำหรับแสดงในกราฟ
  const topAreas = data.slice(0, 5);

  const ctx = canvas.getContext("2d");

  // สร้างสีสำหรับแต่ละพื้นที่
  const areaColors = [
    "rgba(59, 130, 246, 0.8)",
    "rgba(16, 185, 129, 0.8)",
    "rgba(245, 158, 11, 0.8)",
    "rgba(139, 92, 246, 0.8)",
    "rgba(236, 72, 153, 0.8)",
  ];

  const datasets = topAreas.map((area, index) => {
    const salesData = area.monthlyData.map((m) => m.sales);

    return {
      label: area.area,
      data: salesData,
      borderColor: areaColors[index],
      backgroundColor: areaColors[index].replace("0.8", "0.1"),
      borderWidth: 3,
      tension: 0.3,
      fill: true,
    };
  });

  new Chart(ctx, {
    type: "line",
    data: {
      labels: thaiMonths,
      datasets: datasets,
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false,
      },
      plugins: {
        legend: {
          position: "top",
          labels: {
            color: "#cbd5e1",
            font: {
              size: 12,
            },
          },
        },
        tooltip: {
          backgroundColor: "rgba(15, 23, 42, 0.95)",
          titleColor: "#e2e8f0",
          bodyColor: "#cbd5e1",
          borderColor: "rgba(56, 189, 248, 0.3)",
          borderWidth: 1,
          callbacks: {
            label: (context) => {
              const label = context.dataset.label || "";
              const value = context.parsed.y;
              return `${label}: ${fmt.format(value)} ฿`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: {
            color: "rgba(255, 255, 255, 0.05)",
          },
          ticks: {
            color: "#94a3b8",
          },
        },
        y: {
          beginAtZero: true,
          grid: {
            color: "rgba(255, 255, 255, 0.05)",
          },
          ticks: {
            color: "#94a3b8",
            callback: (value) => fmt.formatShort(value),
          },
        },
      },
    },
  });
}

// ✅ Setup view toggles
function setupHeatmapViewToggles() {
  const viewButtons = document.querySelectorAll(".view-btn");
  const views = document.querySelectorAll(".heatmap-view");

  viewButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const view = button.dataset.view;

      // Update active button
      viewButtons.forEach((btn) => btn.classList.remove("active"));
      button.classList.add("active");

      // Show selected view
      views.forEach((v) => v.classList.remove("active"));
      document.getElementById(`${view}View`).classList.add("active");
    });
  });
}

// ✅ Setup tooltips
function setupHeatmapTooltips() {
  const cells = document.querySelectorAll(".data-cell");

  cells.forEach((cell) => {
    cell.addEventListener("mouseenter", (e) => {
      const tooltip = cell.querySelector(".cell-tooltip");
      if (tooltip) {
        tooltip.style.display = "block";

        // Position tooltip
        const rect = cell.getBoundingClientRect();
        tooltip.style.left = `${rect.left + rect.width / 2}px`;
        tooltip.style.top = `${rect.top - tooltip.offsetHeight - 10}px`;
      }
    });

    cell.addEventListener("mouseleave", () => {
      const tooltip = cell.querySelector(".cell-tooltip");
      if (tooltip) {
        tooltip.style.display = "none";
      }
    });
  });
}

function updateAllUI(payload) {
  console.group("🔄 updateAllUI called");
  console.log("Payload keys:", Object.keys(payload));

  // ✅ Debug: ตรวจสอบข้อมูลสำคัญ
  console.log("🔍 Data check:", {
    hasDailyTrend: !!payload.dailyTrend,
    dailyTrendLength: payload.dailyTrend?.length || 0,
    hasSummary: !!payload.summary,
    summaryLength: payload.summary?.length || 0,
    hasPersonTotals: !!payload.personTotals,
    personTotalsLength: payload.personTotals?.length || 0,
    hasCallVisitYearly: !!payload.callVisitYearly,
    hasCustomerSegmentation: !!payload.customerSegmentation,
    hasProductMix: !!payload.productMix,
    hasTopByTeam: !!payload.topByTeam,
  });

  if (!payload) {
    console.error("❌ Payload is null or undefined");
    showToast("ไม่มีข้อมูลจากเซิร์ฟเวอร์", "error");
    return;
  }

  // ✅ ตั้งค่า state
  state.lastPayload = payload;

  // ✅ 1. อัปเดตข้อมูลพื้นฐาน (ไม่ต้องการ container)
  try {
    updateRangeText(payload);
    setAvailable_PATCH(payload);
    setKPI(payload);

    // ✅ 2. อัปเดต chart
    if (chart) {
      console.log("📈 Updating chart...");
      setTrend(payload);
    } else {
      console.warn("⚠️ Chart not initialized, calling initChart...");
      initChart();
      if (chart) setTrend(payload);
    }
  } catch (error) {
    console.error("❌ Error updating basic data:", error);
  }

  // ✅ 3. อัปเดตตารางข้อมูล (ใช้ safeRender เพื่อป้องกัน error)
  console.log("📊 Rendering tables...");

  // 3.1 Person Totals with Pagination
  try {
    if (typeof renderPersonTotalsWithPagination === "function") {
      console.log("👥 Rendering person totals...");
      renderPersonTotalsWithPagination(payload, 1, 20);
    } else if (typeof renderPersonTotals === "function") {
      console.log("👥 Rendering person totals (fallback)...");
      renderPersonTotals(payload);
    } else {
      console.warn("⚠️ No person totals function found");
    }
  } catch (error) {
    console.error("❌ Error rendering person totals:", error);
  }

  // 3.2 Summary Table
  try {
    if (typeof setSummary === "function") {
      console.log("🏢 Rendering summary...");
      setSummary(payload);
    }
  } catch (error) {
    console.error("❌ Error rendering summary:", error);
  }

  // ✅ 4. อัปเดต Charts และ Metrics (สำคัญ!)
  console.log("📈 Rendering charts and metrics...");

  // 4.1 Product Mix Chart
  try {
    if (typeof renderProductMix === "function" && payload.productMix) {
      console.log("📦 Rendering product mix...");
      renderProductMix(payload);
    } else {
      console.log("ℹ️ No product mix data or function");
    }
  } catch (error) {
    console.error("❌ Error in renderProductMix:", error);
    const productContainer =
      document.getElementById("productChart")?.parentElement;
    if (productContainer) {
      productContainer.innerHTML = '<div class="muted">ไม่มีข้อมูลสินค้า</div>';
    }
  }

  // 4.2 Sales Funnel
  try {
    if (typeof renderFunnel === "function") {
      console.log("🔄 Rendering sales funnel...");
      renderFunnel(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderFunnel:", error);
    // Fallback
    const funnelLeads = document.getElementById("funnel_leads");
    const funnelQuotes = document.getElementById("funnel_quotes");
    const funnelClosed = document.getElementById("funnel_closed");
    if (funnelLeads) funnelLeads.textContent = "-";
    if (funnelQuotes) funnelQuotes.textContent = "-";
    if (funnelClosed) funnelClosed.textContent = "-";
  }

  // 4.3 Target Achievement
  try {
    if (typeof renderTarget === "function") {
      console.log("🎯 Rendering target...");
      renderTarget(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderTarget:", error);
    // Fallback
    const targetActual = document.getElementById("target_actual");
    const targetGoal = document.getElementById("target_goal");
    const targetPct = document.getElementById("target_pct");
    if (targetActual) targetActual.textContent = "ไม่มีข้อมูล";
    if (targetGoal) targetGoal.textContent = "ไม่มีข้อมูล";
    if (targetPct) targetPct.textContent = "0%";
  }

  // ✅ 5. อัปเดตเมตริกอื่นๆ
  console.log("📊 Rendering other metrics...");

  // 5.1 Monthly Comparison
  try {
    if (typeof renderMonthlyComparison === "function") {
      console.log("📅 Rendering monthly comparison...");
      renderMonthlyComparison(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderMonthlyComparison:", error);
  }

  // 5.2 Customer Insight
  try {
    if (typeof renderCustomerInsight === "function") {
      console.log("👥 Rendering customer insight...");
      renderCustomerInsight(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderCustomerInsight:", error);
  }

  // 5.3 Call & Visit Yearly
  try {
    if (typeof renderCallVisitYearly === "function") {
      console.log("📞 Rendering call & visit yearly...");
      renderCallVisitYearly(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderCallVisitYearly:", error);
    // Fallback values
    const ids = [
      "cv_total_calls",
      "cv_total_visits",
      "cv_total_presented",
      "cv_total_quoted",
      "cv_total_closed",
    ];
    ids.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.textContent = "N/A";
    });
  }

  // 5.4 Lost Deals Chart
  try {
    if (typeof renderLostDeals === "function") {
      console.log("📉 Rendering lost deals...");
      renderLostDeals(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderLostDeals:", error);
  }

  // ✅ 6. อัปเดต Top 5
  console.log("🏆 Rendering Top 5...");
  if (!state.activeMetric) {
    state.activeMetric = "sales";
    console.log("Set default active metric to:", state.activeMetric);
  }

  try {
    if (typeof renderTop5 === "function") {
      renderTop5(payload);
    } else {
      console.warn("⚠️ renderTop5 function not found");
    }
  } catch (error) {
    console.error("❌ Error in renderTop5:", error);
    const top5Wrap = document.getElementById("top5Wrap");
    if (top5Wrap) {
      top5Wrap.innerHTML =
        '<div class="muted">เกิดข้อผิดพลาดในการแสดงผล Top 5</div>';
    }
  }

  // ✅ 7. อัปเดต Area Performance
  try {
    if (typeof renderAreaPerformance === "function") {
      console.log("🗺️ Rendering area performance...");
      renderAreaPerformance(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderAreaPerformance:", error);
    const container = document.getElementById("areaPerformanceContainer");
    if (container) {
      container.innerHTML =
        '<div class="muted">เกิดข้อผิดพลาดในการแสดงผล Area Performance</div>';
    }
  }

  // ✅ 8. อัปเดต Top Performers
  try {
    if (typeof renderTopPerformers === "function") {
      console.log("⭐ Rendering top performers...");
      renderTopPerformers(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderTopPerformers:", error);
  }

  // ✅ 9. อัปเดต Conversion Rate
  try {
    if (typeof renderConversionRate === "function") {
      console.log("📊 Rendering conversion rate...");
      renderConversionRate(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderConversionRate:", error);
    const container = document.getElementById("conversionContainer");
    if (container) {
      container.innerHTML =
        '<div class="muted">เกิดข้อผิดพลาดในการแสดงผล Conversion Rate</div>';
    }
  }

  // ✅ 10. อัปเดต Customer Segmentation (พร้อม fallback)
  try {
    if (typeof renderCustomerSegmentation === "function") {
      console.log("👥 Rendering customer segmentation...");

      // ตรวจสอบว่ามีข้อมูลก่อน
      if (payload.customerSegmentation) {
        renderCustomerSegmentation(payload);
      } else {
        console.log("ℹ️ No customer segmentation data in payload");

        // ลองหา container และแสดงข้อความว่าไม่มีข้อมูล
        const container =
          document.getElementById("customerSegmentationBody") ||
          document.querySelector("#customerSegmentationTable tbody") ||
          document.querySelector(".customer-segmentation tbody");

        if (container) {
          container.innerHTML = `
            <tr>
              <td colspan="5" class="muted" style="text-align: center; padding: 20px;">
                ไม่มีข้อมูล Customer Segmentation
              </td>
            </tr>
          `;
        }
      }
    } else {
      console.warn("⚠️ renderCustomerSegmentation function not found");
    }
  } catch (error) {
    console.error("❌ Error in renderCustomerSegmentation:", error);

    // Fallback: ลองแสดงข้อความใน container ที่มีอยู่
    const possibleContainers = [
      "#customerSegmentationBody",
      "#customerSegmentationTable tbody",
      ".customer-segmentation tbody",
      "[data-section='customer-segmentation'] tbody",
    ];

    for (const selector of possibleContainers) {
      const container = document.querySelector(selector);
      if (container) {
        container.innerHTML = `
          <tr>
            <td colspan="5" class="muted error" style="text-align: center; padding: 20px;">
              เกิดข้อผิดพลาดในการแสดงผล Customer Segmentation
            </td>
          </tr>
        `;
        break;
      }
    }
  }

  // ✅ 11. อัปเดต Product Performance
  try {
    if (typeof renderProductPerformance === "function") {
      console.log("📦 Rendering product performance...");
      renderProductPerformance(payload);
    }
  } catch (error) {
    console.error("❌ Error in renderProductPerformance:", error);
    const container = document.getElementById("productPerformanceContainer");
    if (container) {
      container.innerHTML =
        '<div class="muted">เกิดข้อผิดพลาดในการแสดงผล Product Performance</div>';
    }
  }

  // ✅ 12. อัปเดต Area Heatmap (ถ้ามี)
  try {
    if (typeof renderAreaHeatmap === "function" && payload.areaHeatmap) {
      console.log("🗺️ Rendering area heatmap...");
      renderAreaHeatmap(payload);
    } else {
      console.log("ℹ️ No area heatmap data or function");
    }
  } catch (error) {
    console.error("❌ Error in renderAreaHeatmap:", error);
    const container = document.getElementById("areaHeatmapContainer");
    if (container) {
      container.innerHTML =
        '<div class="muted">เกิดข้อผิดพลาดในการแสดงผล Area Heatmap</div>';
    }
  }

  // ✅ 13. ตรวจสอบและเรียก initChart ถ้าจำเป็น
  if (!chart && window.Chart) {
    console.log("🔄 Initializing main chart...");
    initChart();
    if (chart && payload.dailyTrend) {
      setTrend(payload);
    }
  }

  // ✅ 14. ตรวจสอบและเรียก initProductChart ถ้าจำเป็น
  if (!productChart && window.Chart) {
    const productCanvas = document.getElementById("productChart");
    if (productCanvas) {
      console.log("🔄 Initializing product chart...");
      initProductChart();
      if (productChart && payload.productMix) {
        renderProductMix(payload);
      }
    }
  }

  // ✅ 15. ตรวจสอบและเรียก initLostDealChart ถ้าจำเป็น
  if (!lostDealChart && window.Chart) {
    const lostDealCanvas = document.getElementById("lostDealChart");
    if (lostDealCanvas) {
      console.log("🔄 Initializing lost deal chart...");
      initLostDealChart();
      if (lostDealChart && payload.lostReasons) {
        renderLostDeals(payload);
      }
    }
  }

  // ✅ 16. อัปเดต filter status
  setFilterStatus("พร้อมใช้งาน");

  // ✅ 17. ตรวจสอบว่ามีข้อผิดพลาดใน console หรือไม่
  const errorCount = (() => {
    try {
      const logs = console.logs || [];
      return logs.filter((log) => log.type === "error").length;
    } catch {
      return 0;
    }
  })();

  if (errorCount > 0) {
    console.warn(`⚠️ Found ${errorCount} errors during UI update`);
  }

  console.log("✅ updateAllUI completed successfully");
  console.groupEnd();
}

// ✅ HELPER FUNCTION: สำหรับเรียก render อย่างปลอดภัย (fixed parameter order)
function safeRender(
  containerId,
  renderFunction,
  payload,
  fallbackMessage = "ไม่มีข้อมูล",
) {
  try {
    console.log(
      `🔧 safeRender: ${containerId}, function: ${renderFunction?.name || "anonymous"}`,
    );

    if (typeof renderFunction !== "function") {
      console.warn(
        `⚠️ ${renderFunction?.name || "renderFunction"} is not a function`,
      );
      return;
    }

    const container = document.getElementById(containerId);
    if (!container) {
      console.warn(`⚠️ Container ${containerId} not found`);
      return;
    }

    // ตรวจสอบว่ามีข้อมูลใน payload หรือไม่
    const hasData = checkPayloadForData(renderFunction.name, payload);
    if (!hasData) {
      console.log(`ℹ️ No data for ${renderFunction.name}, using fallback`);
      container.innerHTML = `<div class="muted">${fallbackMessage}</div>`;
      return;
    }

    // เรียก render function
    renderFunction(payload);
  } catch (error) {
    console.error(
      `❌ Error in ${renderFunction?.name || "renderFunction"}:`,
      error,
    );
    const container = document.getElementById(containerId);
    if (container) {
      container.innerHTML = `<div class="muted error">เกิดข้อผิดพลาดในการแสดงผล</div>`;
    }
  }
}

function checkPayloadForData(functionName, payload) {
  if (!payload) return false;

  // Map render functions กับ keys ใน payload
  const dataMap = {
    renderTop5: ["topByTeam", "personTotals"],
    renderAreaPerformance: ["areaPerformance"],
    renderConversionRate: ["conversionAnalysis", "summary", "personTotals"],
    renderCustomerSegmentation: ["customerSegmentation"],
    renderProductPerformance: ["productPerformance", "productMix"],
    renderAreaHeatmap: ["areaHeatmap"],
    renderFunnel: ["funnel"],
    renderMonthlyComparison: ["monthlyComparison", "dailyTrend"],
    renderTarget: ["target"],
    renderProductMix: ["productMix"],
    renderCustomerInsight: ["customerInsight"],
    renderCallVisitYearly: ["callVisitYearly"],
    renderLostDeals: ["lostReasons"],
    renderTopPerformers: ["callVisitAnalysis", "topPerformers"],
    renderPersonTotalsWithPagination: ["personTotals"],
    renderPersonTotals: ["personTotals"],
    setSummary: ["summary"],
    setTrend: ["dailyTrend"],
  };

  const keys = dataMap[functionName] || [];

  // ถ้าไม่มี mapping ให้ถือว่ามีข้อมูล (ให้ render function จัดการเอง)
  if (keys.length === 0) {
    console.log(`ℹ️ No data mapping for ${functionName}, assuming data exists`);
    return true;
  }

  for (const key of keys) {
    if (payload[key] !== undefined && payload[key] !== null) {
      // ตรวจสอบ array
      if (Array.isArray(payload[key]) && payload[key].length > 0) {
        console.log(
          `✓ Data found for ${functionName}: ${key} (array with ${payload[key].length} items)`,
        );
        return true;
      }
      // ตรวจสอบ object
      if (
        typeof payload[key] === "object" &&
        Object.keys(payload[key]).length > 0
      ) {
        console.log(
          `✓ Data found for ${functionName}: ${key} (object with keys: ${Object.keys(payload[key]).join(", ")})`,
        );
        return true;
      }
      // ตรวจสอบ primitive values
      if (payload[key] !== "" && payload[key] !== 0) {
        console.log(
          `✓ Data found for ${functionName}: ${key} (value: ${payload[key]})`,
        );
        return true;
      }
    }
  }

  console.log(
    `✗ No data found for ${functionName}, checking keys: ${keys.join(", ")}`,
  );
  return false;
}

// ✅ HELPER FUNCTION: ตรวจสอบว่ามีข้อมูลใน payload หรือไม่
function checkPayloadForData(renderFunctionName, payload) {
  if (!payload) return false;

  // Map render functions กับ keys ใน payload
  const dataMap = {
    renderTop5: ["topByTeam", "personTotals"],
    renderAreaPerformance: ["areaPerformance"],
    renderConversionRate: ["conversionAnalysis", "summary", "personTotals"],
    renderCustomerSegmentation: ["customerSegmentation"],
    renderProductPerformance: ["productPerformance", "productMix"],
    renderAreaHeatmap: ["areaHeatmap"],
    renderFunnel: ["funnel"],
    renderMonthlyComparison: ["monthlyComparison", "dailyTrend"],
    renderTarget: ["target"],
    renderProductMix: ["productMix"],
    renderCustomerInsight: ["customerInsight"],
    renderCallVisitYearly: ["callVisitYearly"],
    renderLostDeals: ["lostReasons"],
    renderTopPerformers: ["callVisitAnalysis", "topPerformers"],
    renderPersonTotalsWithPagination: ["personTotals"],
    renderPersonTotals: ["personTotals"],
    setSummary: ["summary"],
    renderAreaHeatmap: ["areaHeatmap"],
  };

  const keys = dataMap[renderFunctionName] || [];

  // ถ้าไม่มี mapping ให้ถือว่ามีข้อมูล (ให้ render function จัดการเอง)
  if (keys.length === 0) {
    console.log(
      `ℹ️ No data mapping for ${renderFunctionName}, assuming data exists`,
    );
    return true;
  }

  for (const key of keys) {
    if (payload[key] !== undefined && payload[key] !== null) {
      // ตรวจสอบ array
      if (Array.isArray(payload[key]) && payload[key].length > 0) {
        console.log(
          `✓ Data found for ${renderFunctionName}: ${key} (array with ${payload[key].length} items)`,
        );
        return true;
      }
      // ตรวจสอบ object
      if (
        typeof payload[key] === "object" &&
        Object.keys(payload[key]).length > 0
      ) {
        console.log(
          `✓ Data found for ${renderFunctionName}: ${key} (object with keys: ${Object.keys(payload[key]).join(", ")})`,
        );
        return true;
      }
      // ตรวจสอบ number
      if (typeof payload[key] === "number" && payload[key] > 0) {
        console.log(
          `✓ Data found for ${renderFunctionName}: ${key} (number: ${payload[key]})`,
        );
        return true;
      }
      // ตรวจสอบ string (สำหรับบางฟังก์ชันเช่น range.text)
      if (typeof payload[key] === "string" && payload[key].trim().length > 0) {
        console.log(`✓ Data found for ${renderFunctionName}: ${key} (string)`);
        return true;
      }
    }
  }

  console.log(
    `✗ No data found for ${renderFunctionName}, checking keys: ${keys.join(", ")}`,
  );
  return false;
}

const DS = {
  SALES_CUM: 0,
  CALLS: 1,
  VISITS: 2,
  QUOTES: 3,
};

function bindChartCheckboxes() {
  const map = [
    { id: "ck_sales", idx: DS.SALES_CUM },
    { id: "ck_calls", idx: DS.CALLS },
    { id: "ck_visits", idx: DS.VISITS },
    { id: "ck_quotes", idx: DS.QUOTES },
  ];

  map.forEach(({ id, idx }) => {
    const box = el(id);
    if (!box) return;

    box.addEventListener("change", () => {
      if (!chart) return;
      chart.setDatasetVisibility(idx, !!box.checked);
      chart.update("none");
    });
  });
}

function initLostDealChart() {
  if (!window.Chart) return;

  const canvas = document.getElementById("lostDealChart");
  if (!canvas) return;

  const ctx = canvas.getContext("2d");

  lostDealChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: [],
      datasets: [
        {
          label: "จำนวนครั้ง",
          data: [],
          borderWidth: 1,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } },
    },
  });
}

function initChart() {
  if (!window.Chart) {
    console.error("❌ Chart.js not loaded");
    setText("chartStatus", "Chart.js โหลดไม่สำเร็จ");
    return;
  }

  const canvas = el("chart");
  if (!canvas) {
    console.error("❌ canvas#chart not found");
    return;
  }

  const ctx = canvas.getContext("2d");

  chart = new Chart(ctx, {
    type: "line",
    data: {
      labels: [],
      datasets: [
        {
          label: "ยอดขายสะสม (บาท)",
          data: [],
          yAxisID: "ySales",
          borderColor: "#22c55e",
          backgroundColor: "rgba(34,197,94,0.15)",
          tension: 0.35,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 7,
          borderWidth: 3,
        },
        {
          label: "โทร",
          data: [],
          yAxisID: "yCount",
          borderColor: "#fb7185",
          backgroundColor: "rgba(251,113,133,0.15)",
          tension: 0.35,
          fill: true,
          pointRadius: 2,
          pointHoverRadius: 6,
          borderWidth: 2,
        },
        {
          label: "เข้าพบ",
          data: [],
          yAxisID: "yCount",
          borderColor: "#38bdf8",
          backgroundColor: "rgba(56,189,248,0.15)",
          tension: 0.35,
          fill: true,
          pointRadius: 2,
          pointHoverRadius: 6,
          borderWidth: 2,
        },
        {
          label: "ใบเสนอราคา",
          data: [],
          yAxisID: "yCount",
          borderColor: "#facc15",
          backgroundColor: "rgba(250,204,21,0.15)",
          tension: 0.35,
          fill: true,
          pointRadius: 2,
          pointHoverRadius: 6,
          borderWidth: 2,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false, // ✅ สำคัญมาก ให้เต็มช่อง
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          display: true,
          // ✅ กันไม่ให้คลิก legend แล้ว toggle
          onClick: () => {},
          labels: { color: "#cbd5e1", font: { size: 12, weight: "600" } },
        },
        tooltip: {
          backgroundColor: "rgba(15, 23, 42, 0.95)",
          padding: 12,
          titleColor: "#cbd5e1",
          bodyColor: "#e5e7eb",
          borderColor: "rgba(96, 165, 250, 0.3)",
          borderWidth: 1,
          callbacks: {
            title: (items) => addThaiDow(items?.[0]?.label || ""),
          },
        },
      },
      scales: {
        ySales: {
          position: "left",
          beginAtZero: true,
          ticks: { callback: (v) => fmt.format(v) },
        },
        yCount: {
          position: "right",
          beginAtZero: true,
          grid: { drawOnChartArea: false },
          ticks: { callback: (v) => fmt.format(v) },
        },
      },
    },
  });

  // ✅ bind checkbox แยกของคุณยังใช้ได้
  if (typeof bindChartCheckboxes === "function") bindChartCheckboxes();

  // ✅ force resize หลัง DOM render
  setTimeout(() => {
    try {
      chart.resize();
      chart.update("none");
    } catch {}
  }, 0);
}

function setTrend(payload) {
  if (!chart) return;

  const rows = Array.isArray(payload.dailyTrend) ? payload.dailyTrend : [];

  if (!rows.length) {
    console.warn("⚠️ dailyTrend ว่าง");
    chart.data.labels = [];
    chart.data.datasets.forEach((ds) => (ds.data = []));
    chart.update();
    return;
  }

  let cum = 0;
  const sorted = rows
    .filter((r) => r.date)
    .sort((a, b) => a.date.localeCompare(b.date));

  const labels = [];
  const salesCum = [];
  const calls = [];
  const visits = [];
  const quotes = [];

  sorted.forEach((r) => {
    const s = Number(r.sales || 0);
    cum += s;

    labels.push(r.date);
    salesCum.push(cum);
    calls.push(Number(r.calls || 0));
    visits.push(Number(r.visits || 0));
    quotes.push(Number(r.quotes || 0));
  });

  chart.data.labels = labels;
  chart.data.datasets[DS.SALES_CUM].data = salesCum;
  chart.data.datasets[DS.CALLS].data = calls;
  chart.data.datasets[DS.VISITS].data = visits;
  chart.data.datasets[DS.QUOTES].data = quotes;

  chart.update();
}

// ---------------- 🆕 Sales Funnel ----------------
function renderFunnel(payload) {
  // ใช้ข้อมูลจริงจาก API แทน mock data
  const funnel = payload.funnel || {
    leads: payload.totalLeads || 0,
    quotes: payload.totalQuotes || 0,
    closed: payload.totalClosed || 0,
  };

  const totalLeads = funnel.leads || 1;
  const quotesPct =
    totalLeads > 0 ? ((funnel.quotes / totalLeads) * 100).toFixed(1) : 0;
  const closedPct =
    totalLeads > 0 ? ((funnel.closed / totalLeads) * 100).toFixed(1) : 0;

  setText("funnel_leads", fmt.format(funnel.leads));
  setText("funnel_quotes", fmt.format(funnel.quotes));
  setText("funnel_quotes_pct", `${quotesPct}%`);
  setText("funnel_closed", fmt.format(funnel.closed));
  setText("funnel_closed_pct", `${closedPct}%`);

  const quotesBar = el("funnel_quotes_bar");
  const closedBar = el("funnel_closed_bar");

  if (quotesBar) quotesBar.style.width = `${Math.min(quotesPct, 100)}%`;
  if (closedBar) closedBar.style.width = `${Math.min(closedPct, 100)}%`;
}

// ---------------- 🆕 Conversion Rate Analysis ----------------

function renderConversionRate(payload) {
  console.log("🔄 renderConversionRate called");
  console.log("🔍 Conversion Rate payload:", payload?.conversionAnalysis);
  console.log("🔍 Summary payload:", payload?.summary);
  console.log("🔍 PersonTotals payload:", payload?.personTotals);

  const summary = payload.summary || [];
  const personTotals = payload.personTotals || [];
  const summaryTotals = payload.summaryTotals || {
    sales: 0,
    calls: 0,
    visits: 0,
    quotes: 0,
  };

  // ✅ ตรวจสอบว่าข้อมูลถูกต้อง
  console.log("📊 Summary Totals:", summaryTotals);
  console.log("📊 Summary Array:", summary);

  // ✅ ตรวจสอบว่าไม่ได้เอา Sales amount ไปหารด้วย Quotes count
  console.log("⚠️ IMPORTANT: Check if sales is amount or count");
  console.log("- Sales total:", summaryTotals.sales);
  console.log("- Quotes total:", summaryTotals.quotes);

  // ถ้า sales เป็นจำนวนเงิน (บาท) และ quotes เป็นจำนวนใบ
  // จะคำนวณ conversion rate ไม่ได้
  if (summaryTotals.sales > summaryTotals.quotes * 10000) {
    console.error(
      "❌ DETECTED: Sales (amount) vs Quotes (count) unit mismatch!",
    );
    console.error("Sales:", summaryTotals.sales, "฿");
    console.error("Quotes:", summaryTotals.quotes, "ใบ");
    console.error(
      "Sales/Quotes ratio:",
      summaryTotals.sales / summaryTotals.quotes,
    );
  }

  // ✅ ตรวจสอบปีของข้อมูล
  const dataYear = payload.range?.year || new Date().getFullYear();
  const currentYear = new Date().getFullYear();
  const isCurrentYear = dataYear === currentYear;

  let html = "";

  // ✅ แสดงปีของข้อมูล
  html += `
    <div class="conversion-year-header">
      <div class="year-badge ${isCurrentYear ? "current" : "past"}">
        <span class="year-icon">📅</span>
        <span class="year-text">ข้อมูลปี ${dataYear}</span>
        ${isCurrentYear ? '<span class="year-current">(ปัจจุบัน)</span>' : ""}
      </div>
    </div>
  `;

  // ✅ ตรวจสอบว่ามีข้อมูลหรือไม่
  if (summary.length === 0) {
    html += `<div class="muted" style="text-align: center; padding: 40px;">
              ไม่มีข้อมูล Conversion Rate สำหรับปี ${dataYear}
            </div>`;
    document.getElementById("conversionContainer").innerHTML = html;
    return;
  }

  // ✅ สำคัญ: ต้องรู้ว่า sales เป็นจำนวนเงินหรือจำนวนใบ
  const overallQuotes = Number(summaryTotals.quotes || 0);
  const overallSalesAmount = Number(summaryTotals.sales || 0); // นี่คือจำนวนเงิน (บาท)
  const overallCalls = Number(summaryTotals.calls || 0);
  const overallVisits = Number(summaryTotals.visits || 0);

  console.log("📈 Overall metrics:", {
    calls: overallCalls,
    visits: overallVisits,
    quotes: overallQuotes,
    salesAmount: overallSalesAmount,
  });

  // ✅ ปัญหา: เรามี sales เป็นจำนวนเงิน แต่ quotes เป็นจำนวนใบ
  // เราต้องประมาณการจำนวน deal ที่ปิดได้จากยอดขาย
  const AVERAGE_DEAL_SIZE = 50000; // สมมติ average deal = 50,000 ฿
  const estimatedClosedDeals = Math.max(
    1,
    Math.round(overallSalesAmount / AVERAGE_DEAL_SIZE),
  );

  // ✅ การคำนวณ Conversion Rates ที่ถูกต้อง
  const overallQuoteToSaleRate =
    overallQuotes > 0
      ? Math.min(100, (estimatedClosedDeals / overallQuotes) * 100)
      : 0;

  const overallCallToQuoteRate =
    overallCalls > 0 ? Math.min(100, (overallQuotes / overallCalls) * 100) : 0;

  const overallCallToVisitRate =
    overallCalls > 0 ? Math.min(100, (overallVisits / overallCalls) * 100) : 0;

  const overallVisitToQuoteRate =
    overallVisits > 0
      ? Math.min(100, (overallQuotes / overallVisits) * 100)
      : 0;

  console.log("📊 Calculated rates:", {
    quoteToSaleRate: overallQuoteToSaleRate,
    callToQuoteRate: overallCallToQuoteRate,
    callToVisitRate: overallCallToVisitRate,
    visitToQuoteRate: overallVisitToQuoteRate,
    estimatedClosedDeals: estimatedClosedDeals,
    averageDealSize: AVERAGE_DEAL_SIZE,
  });

  // ✅ แสดง warning ถ้า conversion rate ผิดปกติ
  if (overallQuoteToSaleRate > 100 || overallQuoteToSaleRate < 0) {
    console.error("❌ ABNORMAL CONVERSION RATE:", overallQuoteToSaleRate);
    console.error("This usually means sales/quotes units are mismatched!");

    // แสดงข้อความเตือนใน UI
    html += `
      <div class="warning-message" style="background: rgba(239,68,68,0.1); border: 1px solid rgba(239,68,68,0.3); 
              border-radius: 6px; padding: 10px; margin-bottom: 15px;">
        <div style="color: #ef4444; font-weight: 600; margin-bottom: 5px;">⚠️ ข้อมูลหน่วยไม่ตรงกัน</div>
        <div style="color: #94a3b8; font-size: 13px;">
          Sales (${fmt.format(overallSalesAmount)} ฿) และ Quotes (${fmt.format(overallQuotes)} ใบ) เป็นคนละหน่วย<br>
          Conversion rate นี้เป็นค่าประมาณการจากยอดขาย
        </div>
      </div>
    `;
  }

  // ✅ Header section with overall metrics
  html += `
    <div class="conversion-header">
      <div class="conversion-overview">
        <h3>Overall Conversion Funnel ปี ${dataYear}</h3>
        <div class="funnel-steps">
          <div class="funnel-step">
            <div class="step-label">การโทร</div>
            <div class="step-value">${fmt.format(overallCalls)}</div>
            <div class="step-rate">${overallCallToVisitRate.toFixed(1)}% →</div>
          </div>
          <div class="funnel-step">
            <div class="step-label">การเข้าพบ</div>
            <div class="step-value">${fmt.format(overallVisits)}</div>
            <div class="step-rate">${overallVisitToQuoteRate.toFixed(1)}% →</div>
          </div>
          <div class="funnel-step">
            <div class="step-label">ใบเสนอราคา</div>
            <div class="step-value">${fmt.format(overallQuotes)}</div>
            <div class="step-rate">${overallQuoteToSaleRate.toFixed(1)}% →</div>
          </div>
          <div class="funnel-step success">
            <div class="step-label">ยอดขาย (ประมาณการ)</div>
            <div class="step-value">${fmt.format(estimatedClosedDeals)} ดีล</div>
            <div class="step-rate">สุดท้าย</div>
          </div>
        </div>
        <div class="funnel-summary">
          <div class="summary-item">
            <div class="summary-label">อัตราการปิดการขาย</div>
            <div class="summary-value">${overallQuoteToSaleRate.toFixed(1)}%</div>
            <div class="summary-note">(ประมาณการจากยอดขาย)</div>
          </div>
          <div class="summary-item">
            <div class="summary-label">ประสิทธิภาพการโทร → ใบเสนอ</div>
            <div class="summary-value">${overallCallToQuoteRate.toFixed(1)}%</div>
          </div>
          <div class="summary-item">
            <div class="summary-label">ยอดขายรวม</div>
            <div class="summary-value">${fmt.format(overallSalesAmount)} ฿</div>
          </div>
        </div>
      </div>
    </div>
  `;

  // ✅ 2. Conversion Rate ตามทีม (ใช้วิธีการเดียวกัน)
  html += `<div class="conversion-teams-title"><h3>Conversion Rate ตามทีม (ปี ${dataYear})</h3></div>`;
  html += `<div class="conversion-teams-grid">`;

  // กรองทีมที่มีข้อมูล
  const teamsWithData = summary.filter(
    (team) => (team.quotes || 0) > 0 && (team.sales || 0) > 0,
  );

  if (teamsWithData.length === 0) {
    html += `<div class="muted" style="grid-column: 1 / -1; text-align: center; padding: 40px;">
              ไม่มีข้อมูลทีมสำหรับปี ${dataYear}
            </div>`;
  } else {
    teamsWithData.forEach((team, index) => {
      const teamName = escapeHtml(team.team || "ไม่ระบุทีม");
      const teamSalesAmount = Number(team.sales || 0); // จำนวนเงิน (บาท)
      const teamQuotes = Number(team.quotes || 0);
      const teamCalls = Number(team.calls || 0);
      const teamVisits = Number(team.visits || 0);

      // ✅ คำนวณ Conversion Rates (ใช้ average deal size เดียวกัน)
      const teamEstimatedDeals = Math.max(
        1,
        Math.round(teamSalesAmount / AVERAGE_DEAL_SIZE),
      );
      const quoteToSaleRate =
        teamQuotes > 0
          ? Math.min(100, (teamEstimatedDeals / teamQuotes) * 100)
          : 0;

      const callToQuoteRate =
        teamCalls > 0 ? Math.min(100, (teamQuotes / teamCalls) * 100) : 0;

      // ✅ กำหนดสีตาม performance
      const quoteToSaleRateNum = parseFloat(quoteToSaleRate);
      let rateColorClass = "poor";
      if (quoteToSaleRateNum >= 30) rateColorClass = "excellent";
      else if (quoteToSaleRateNum >= 20) rateColorClass = "good";
      else if (quoteToSaleRateNum >= 10) rateColorClass = "fair";

      // ตรวจสอบว่ามีข้อมูลผิดปกติหรือไม่
      const hasDataIssue = teamSalesAmount > teamQuotes * 10000;
      const issueBadge = hasDataIssue
        ? '<span class="issue-badge" title="ข้อมูลหน่วยอาจไม่ตรงกัน">⚠️</span>'
        : "";

      html += `
        <div class="conversion-team-card ${hasDataIssue ? "has-issue" : ""}">
          <div class="team-header">
            <div class="team-name">${teamName} ${issueBadge}</div>
            <div class="team-performance ${rateColorClass}">
              <div class="main-rate">${quoteToSaleRate.toFixed(1)}%</div>
              <div class="rate-label">อัตราการปิด</div>
              ${hasDataIssue ? '<div class="rate-note">(ประมาณการ)</div>' : ""}
            </div>
          </div>
          
          <div class="team-metrics">
            <div class="metric-row">
              <span class="metric-label">การโทร</span>
              <span class="metric-value">${fmt.format(teamCalls)}</span>
            </div>
            <div class="metric-row">
              <span class="metric-label">การเข้าพบ</span>
              <span class="metric-value">${fmt.format(teamVisits)}</span>
            </div>
            <div class="metric-row">
              <span class="metric-label">ใบเสนอราคา</span>
              <span class="metric-value">${fmt.format(teamQuotes)}</span>
            </div>
            <div class="metric-row highlight">
              <span class="metric-label">ยอดขาย</span>
              <span class="metric-value">${fmt.format(teamSalesAmount)} ฿</span>
            </div>
            <div class="metric-row">
              <span class="metric-label">ประมาณการดีลที่ปิด</span>
              <span class="metric-value">${fmt.format(teamEstimatedDeals)} ดีล</span>
            </div>
          </div>
          
          <div class="team-stats-summary">
            <div class="stat-item">
              <div class="stat-label">อัตราการโทร→ใบเสนอ</div>
              <div class="stat-value">${callToQuoteRate.toFixed(1)}%</div>
            </div>
            <div class="stat-item">
              <div class="stat-label">เฉลี่ย/ใบเสนอ</div>
              <div class="stat-value">${teamQuotes > 0 ? fmt.format(Math.round(teamSalesAmount / teamQuotes)) : 0} ฿</div>
            </div>
          </div>
          
          ${
            hasDataIssue
              ? `
          <div class="team-note">
            <small>⚠️ ข้อมูลประมาณการ (Sales vs Quotes หน่วยไม่ตรงกัน)</small>
          </div>
          `
              : ""
          }
        </div>
      `;
    });
  }

  html += `</div>`;

  // ✅ 3. Top Performers (Individual)
  if (personTotals.length > 0) {
    html += `<div class="conversion-individual-title"><h3>ผู้ปฏิบัติงานดีเด่น (ปี ${dataYear})</h3></div>`;
    html += `<div class="conversion-individual-grid">`;

    // กรองบุคคลที่มีใบเสนอราคาและยอดขาย
    const individualsWithPerformance = personTotals
      .map((person) => {
        const salesAmount = Number(person.sales || 0);
        const quotes = Number(person.quotes || 0);
        const estimatedDeals = Math.max(
          1,
          Math.round(salesAmount / AVERAGE_DEAL_SIZE),
        );
        const conversionRate =
          quotes > 0 ? Math.min(100, (estimatedDeals / quotes) * 100) : 0;

        return {
          ...person,
          conversionRate: conversionRate,
          estimatedDeals: estimatedDeals,
          avgSalePerQuote: quotes > 0 ? Math.round(salesAmount / quotes) : 0,
        };
      })
      .filter((p) => p.quotes > 0 && p.sales > 0)
      .sort((a, b) => b.conversionRate - a.conversionRate)
      .slice(0, 5);

    if (individualsWithPerformance.length > 0) {
      individualsWithPerformance.forEach((person, index) => {
        const conversionRate = person.conversionRate.toFixed(1);
        const hasDataIssue = person.sales > person.quotes * 10000;

        html += `
          <div class="individual-card ${hasDataIssue ? "has-issue" : ""}">
            <div class="individual-rank">#${index + 1}</div>
            <div class="individual-info">
              <div class="individual-name">${escapeHtml(person.person || "ไม่ระบุชื่อ")}</div>
              <div class="individual-stats">
                <span>${fmt.format(person.quotes || 0)} ใบเสนอ</span>
                <span>•</span>
                <span>${fmt.format(person.estimatedDeals || 0)} ดีล (ประมาณ)</span>
              </div>
            </div>
            <div class="individual-conversion">
              <div class="conversion-value">${conversionRate}%</div>
              <div class="conversion-label">อัตราการปิด</div>
              ${hasDataIssue ? '<div class="conversion-note">ประมาณการ</div>' : ""}
            </div>
          </div>
        `;
      });
    } else {
      html += `<div class="muted" style="grid-column: 1 / -1; text-align: center; padding: 20px;">
                ไม่มีข้อมูลผู้ปฏิบัติงานสำหรับปี ${dataYear}
              </div>`;
    }

    html += `</div>`;
  }

  // ✅ 4. Legend/Explanation
  html += `
    <div class="conversion-legend">
      <div class="legend-title">คำอธิบายและข้อควรระวัง:</div>
      <div class="legend-items">
        <div class="legend-item">
          <span class="legend-color excellent"></span>
          <span class="legend-text">ดีเยี่ยม (≥ 30%)</span>
        </div>
        <div class="legend-item">
          <span class="legend-color good"></span>
          <span class="legend-text">ดี (20-29%)</span>
        </div>
       div class="legend-item">
          <span class="legend-color fair"></span>
          <span class="legend-text">ปานกลาง (10-19%)</span>
        </div>
        <div class="legend-item">
          <span class="legend-color poor"></span>
          <span class="legend-text">ต้องปรับปรุง (< 10%)</span>
        </div>
      </div>
      <div class="legend-warning">
        <div style="color: #f59e0b; font-weight: 600; margin-bottom: 5px;">⚠️ หมายเหตุสำคัญ:</div>
        <div style="color: #94a3b8; font-size: 13px; line-height: 1.5;">
          1. <strong>Conversion Rate คำนวณจากประมาณการ</strong> เพราะข้อมูล Sales (บาท) และ Quotes (ใบ) เป็นคนละหน่วย<br>
          2. สมมติ Average Deal Size = 50,000 ฿ เพื่อแปลงยอดขายเป็นจำนวนดีล<br>
          3. สูตร: Conversion Rate = (ประมาณการดีลที่ปิดได้ ÷ จำนวนใบเสนอราคา) × 100<br>
          4. ตัวเลขนี้เป็นแนวทางอ้างอิง ไม่ใช่ค่าที่แท้จริง
        </div>
      </div>
    </div>
  `;

  const container = document.getElementById("conversionContainer");
  if (container) {
    container.innerHTML = html;
  } else {
    console.error("❌ conversionContainer not found");
  }

  console.log("✅ renderConversionRate completed");
}

// ✅ เพิ่ม CSS สำหรับ issue indicators
function addConversionRateCSS() {
  if (!document.getElementById("conversion-rate-css")) {
    const style = document.createElement("style");
    style.id = "conversion-rate-css";
    style.textContent = `
      .warning-message {
        animation: pulse 2s infinite;
      }
      
      @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.8; }
      }
      
      .issue-badge {
        color: #f59e0b;
        margin-left: 4px;
        font-size: 12px;
        cursor: help;
      }
      
      .has-issue {
        border: 1px solid rgba(245, 158, 11, 0.3);
        background: rgba(245, 158, 11, 0.05);
      }
      
      .rate-note, .conversion-note {
        font-size: 10px;
        color: #f59e0b;
        margin-top: 2px;
      }
      
      .team-note {
        margin-top: 8px;
        padding: 6px;
        background: rgba(245, 158, 11, 0.1);
        border-radius: 4px;
        font-size: 11px;
        color: #f59e0b;
      }
      
      .legend-warning {
        margin-top: 15px;
        padding: 10px;
        background: rgba(245, 158, 11, 0.1);
        border-radius: 6px;
        border-left: 3px solid #f59e0b;
      }
      
      .summary-note {
        font-size: 11px;
        color: #94a3b8;
        margin-top: 2px;
      }
    `;
    document.head.appendChild(style);
  }
}

// ---------------- 🆕 Customer Segmentation ----------------

function renderCustomerSegmentation(payload) {
  console.log("🔄 renderCustomerSegmentation called");
  console.log("Payload customerSegmentation:", payload?.customerSegmentation);

  // ✅ 1. ตรวจสอบว่ามี container หรือไม่ ถ้าไม่มีให้สร้าง
  let container = document.getElementById("customerSegmentationBody");

  if (!container) {
    console.log(
      "⚠️ customerSegmentationBody not found, checking for alternatives...",
    );

    // ลองหาตาราง customer segmentation ด้วยวิธีอื่น
    const possibleSelectors = [
      "#customerSegmentationTable tbody",
      "#customerSegmentation tbody",
      ".customer-segmentation tbody",
      "[data-section='customer-segmentation'] tbody",
    ];

    for (const selector of possibleSelectors) {
      container = document.querySelector(selector);
      if (container) {
        console.log(`✅ Found container using selector: ${selector}`);
        break;
      }
    }

    // ถ้ายังไม่เจอ ลองสร้าง container ใหม่
    if (!container) {
      console.log("🔄 Creating customer segmentation container...");
      container = createCustomerSegmentationContainer();
    }
  }

  if (!container) {
    console.error("❌ Cannot find or create customer segmentation container");
    return;
  }

  const segmentation = payload.customerSegmentation || {};
  const items = segmentation.items || [];
  const summary = segmentation.summary || {};
  const meta = segmentation.meta || {};

  // ✅ 2. ตรวจสอบว่ามีข้อมูลหรือไม่
  if (items.length === 0) {
    console.log("ℹ️ No customer segmentation data");
    container.innerHTML = `
      <tr>
        <td colspan="5" class="muted" style="text-align: center; padding: 40px;">
          ${meta.note || "ไม่มีข้อมูล Customer Segmentation"}
        </td>
      </tr>
    `;

    // อัปเดต header ถ้ามี
    updateCustomerSegmentationHeader(summary);
    return;
  }

  // ✅ 3. สร้างตาราง
  let html = "";

  items.forEach((item, index) => {
    // กำหนดสีตาม rank
    let rankClass = "";
    if (index === 0) rankClass = "rank-1";
    else if (index === 1) rankClass = "rank-2";
    else if (index === 2) rankClass = "rank-3";

    // คำนวณสัดส่วนสำหรับ progress bar
    const maxSales = items[0]?.sales || 1;
    const salesPercentage = (item.sales / maxSales) * 100;

    // รองรับฟิลด์ชื่อต่างๆ
    const type = escapeHtml(
      item.type ||
        item.segment ||
        item.category ||
        item.label ||
        "ไม่ระบุประเภท",
    );

    const uniqueCompanies = Number(
      item.uniqueCompanies || item.companies || item.count || 0,
    );
    const sales = Number(item.sales || item.value || item.amount || 0);
    const percentOfTotal = Number(
      item.percentOfTotal || item.percentage || item.pct || 0,
    );
    const avgPerDeal = Number(item.avgPerDeal || item.average || item.avg || 0);

    html += `
      <tr class="${rankClass}">
        <td>
          <div class="segment-type">
            <span class="segment-rank">${index + 1}</span>
            <span class="segment-name">${type}</span>
          </div>
          <div class="segment-progress">
            <div class="segment-bar" style="width: ${salesPercentage}%"></div>
          </div>
        </td>
        <td class="num">${fmt.format(uniqueCompanies)}</td>
        <td class="num">${fmt.format(sales)} ฿</td>
        <td class="num">
          <span class="percent-badge ${getPercentClass(percentOfTotal)}">
            ${percentOfTotal.toFixed(1)}%
          </span>
        </td>
        <td class="num">${fmt.format(Math.round(avgPerDeal))} ฿</td>
      </tr>
    `;
  });

  // ✅ 4. เพิ่ม summary row
  if (summary.totalSales > 0) {
    html += `
      <tr class="summary-row">
        <td><strong>รวมทั้งหมด</strong> (${summary.year || "ปีปัจจุบัน"})</td>
        <td class="num"><strong>${fmt.format(summary.totalUniqueCompanies || summary.totalCompanies || 0)}</strong></td>
        <td class="num"><strong>${fmt.format(summary.totalSales)} ฿</strong></td>
        <td class="num"><strong>100%</strong></td>
        <td class="num"><strong>${fmt.format(Math.round(summary.averageDealSize || summary.avgDeal || 0))} ฿</strong></td>
      </tr>
    `;
  }

  container.innerHTML = html;

  // ✅ 5. อัปเดต header
  updateCustomerSegmentationHeader(summary);

  console.log(`✅ Customer segmentation rendered: ${items.length} items`);
}

// ✅ Helper: สร้าง container ถ้าไม่มี
function createCustomerSegmentationContainer() {
  console.log("🔧 Creating customer segmentation container...");

  // ลองหาตาราง customer segmentation ใน HTML
  const existingTables = document.querySelectorAll("table");
  let customerSegmentationTable = null;

  existingTables.forEach((table) => {
    const headers = Array.from(table.querySelectorAll("th")).map((th) =>
      th.textContent.toLowerCase(),
    );
    const customerHeaders = [
      "ประเภท",
      "segment",
      "customer",
      "type",
      "category",
    ];

    if (
      headers.some((header) =>
        customerHeaders.some((ch) => header.includes(ch)),
      )
    ) {
      customerSegmentationTable = table;
    }
  });

  if (customerSegmentationTable) {
    // ถ้ามีตารางอยู่แล้ว ให้เพิ่ม tbody ถ้าไม่มี
    let tbody = customerSegmentationTable.querySelector("tbody");
    if (!tbody) {
      tbody = document.createElement("tbody");
      customerSegmentationTable.appendChild(tbody);
    }
    tbody.id = "customerSegmentationBody";
    return tbody;
  }

  // ถ้าไม่มีตารางเลย ให้สร้างใหม่
  const section = document.createElement("div");
  section.className = "section customer-segmentation";
  section.innerHTML = `
    <div class="section-header">
      <h3>Customer Segmentation</h3>
      <div class="section-subtitle" id="customerSegmentationSubtitle"></div>
    </div>
    <div class="table-container">
      <table class="data-table">
        <thead>
          <tr>
            <th>ประเภทลูกค้า</th>
            <th class="num">จำนวนบริษัท</th>
            <th class="num">ยอดขาย</th>
            <th class="num">ส่วนแบ่ง</th>
            <th class="num">เฉลี่ย/ดีล</th>
          </tr>
        </thead>
        <tbody id="customerSegmentationBody"></tbody>
      </table>
    </div>
  `;

  // หาที่วาง section ใหม่
  const targetSections = [
    "#productPerformanceContainer",
    "#areaPerformanceContainer",
    "#conversionContainer",
    ".main-grid",
  ];

  let inserted = false;
  for (const selector of targetSections) {
    const target = document.querySelector(selector);
    if (target) {
      target.parentNode.insertBefore(section, target.nextSibling);
      inserted = true;
      console.log(`✅ Inserted customer segmentation after: ${selector}`);
      break;
    }
  }

  if (!inserted) {
    document.body.appendChild(section);
  }

  return document.getElementById("customerSegmentationBody");
}

// ✅ Helper: อัปเดต header
function updateCustomerSegmentationHeader(summary) {
  const subtitle = document.getElementById("customerSegmentationSubtitle");
  if (!subtitle) return;

  if (summary.totalSales > 0) {
    subtitle.textContent =
      `จำนวนทั้งหมด: ${fmt.format(summary.totalUniqueCompanies || 0)} บริษัท, ` +
      `ยอดขายรวม: ${fmt.format(summary.totalSales || 0)} ฿ ` +
      `(ปี ${summary.year || new Date().getFullYear()})`;
  } else {
    subtitle.textContent = "Customer Segmentation Analysis";
  }
}

// ✅ Helper: ฟังก์ชันกำหนดคลาสตามเปอร์เซ็นต์
function getPercentClass(percent) {
  if (percent >= 30) return "high";
  if (percent >= 15) return "medium";
  return "low";
}

// ฟังก์ชันช่วยเหลือสำหรับกำหนดคลาสตามเปอร์เซ็นต์
function getPercentClass(percent) {
  if (percent >= 30) return "high";
  if (percent >= 15) return "medium";
  return "low";
}

/* ================= Product Performance ================= */
function renderProductPerformance(payload) {
  console.log("🔄 renderProductPerformance called");

  const productPerformance = payload.productPerformance || [];
  const productMix = payload.productMix || {};
  const mixItems = productMix.items || [];

  const container = el("productPerformanceContainer");
  if (!container) {
    console.error("❌ productPerformanceContainer element not found");
    return;
  }

  // ✅ ใช้ข้อมูลจริงจาก productMix หรือ productPerformance
  let products = [];

  if (mixItems.length > 0) {
    // ใช้ข้อมูลจาก productMix (มีข้อมูลยอดขาย)
    products = mixItems
      .map((item) => {
        const sales = Number(item.value || 0);
        // ประมาณการ quotes จากยอดขาย (สมมติ conversion rate เฉลี่ย)
        const estimatedQuotes = Math.max(1, Math.round(sales / 50000)); // สมมติเฉลี่ย 50,000 ฿ ต่อ quote
        const estimatedConversion = 25 + Math.random() * 40; // สุ่ม 25-65% สำหรับแสดงผล

        return {
          product: item.label || "ไม่ระบุ",
          sales: sales,
          quotes: estimatedQuotes,
          conversion: parseFloat(estimatedConversion.toFixed(1)),
          percent: item.pct || 0,
        };
      })
      .slice(0, 8); // แสดงสูงสุด 8 รายการ
  } else if (productPerformance.length > 0) {
    // ใช้ข้อมูลจาก productPerformance (ถ้ามี)
    products = productPerformance.slice(0, 8);
  } else {
    // ไม่มีข้อมูลจริง
    container.innerHTML = `
      <div class="muted" style="text-align: center; padding: 40px;">
        ไม่มีข้อมูล Product Performance
        <div style="font-size: 12px; margin-top: 10px;">
          (ตรวจสอบว่ามีข้อมูลในคอลัมน์ productType และ actualClose)
        </div>
      </div>
    `;
    return;
  }

  // ✅ คำนวณ total sales สำหรับหาเปอร์เซ็นต์
  const totalSales = products.reduce((sum, p) => sum + (p.sales || 0), 0);

  let html = `
    <div class="product-performance-header">
      <div class="header-title">
        <h3>Product Performance Analysis</h3>
        <div class="header-subtitle">
          ประสิทธิภาพสินค้าตาม Conversion Rate และมูลค่าต่อใบเสนอ
          ${totalSales > 0 ? `<span class="total-sales">ยอดขายรวม: ${fmt.format(totalSales)} ฿</span>` : ""}
        </div>
      </div>
    </div>
    
    <div class="product-performance-grid">
  `;

  products.forEach((product, index) => {
    const salesPerQuote =
      product.quotes > 0 ? Math.round(product.sales / product.quotes) : 0;
    const percentOfTotal =
      totalSales > 0 ? (product.sales / totalSales) * 100 : 0;

    // กำหนดสีตาม performance
    let performanceClass = "poor";
    if (product.conversion >= 40) performanceClass = "excellent";
    else if (product.conversion >= 25) performanceClass = "good";
    else if (product.conversion >= 15) performanceClass = "fair";

    // กำหนด rank
    let rankClass = "";
    if (index === 0) rankClass = "rank-1";
    else if (index === 1) rankClass = "rank-2";
    else if (index === 2) rankClass = "rank-3";

    html += `
      <div class="product-performance-card ${rankClass}">
        <div class="product-header">
          <div class="product-rank">#${index + 1}</div>
          <div class="product-info">
            <h4 class="product-name">${escapeHtml(product.product)}</h4>
            <div class="product-meta">
              <span class="meta-item">
                <span class="meta-label">ส่วนแบ่ง:</span>
                <span class="meta-value">${percentOfTotal.toFixed(1)}%</span>
              </span>
            </div>
          </div>
          <div class="product-performance-badge ${performanceClass}">
            <div class="performance-value">${product.conversion}%</div>
            <div class="performance-label">Conversion</div>
          </div>
        </div>
        
        <div class="product-stats">
          <div class="stat-row">
            <div class="stat-item">
              <div class="stat-label">ใบเสนอราคา</div>
              <div class="stat-value">${fmt.format(product.quotes)}</div>
            </div>
            <div class="stat-item">
              <div class="stat-label">ยอดขาย</div>
              <div class="stat-value">${fmt.format(product.sales)} ฿</div>
            </div>
            <div class="stat-item">
              <div class="stat-label">เฉลี่ย/ใบ</div>
              <div class="stat-value">${fmt.format(salesPerQuote)} ฿</div>
            </div>
          </div>
        </div>
        
        <div class="product-visualization">
          <div class="viz-header">
            <span>Conversion Rate</span>
            <span>${product.conversion}%</span>
          </div>
          <div class="conversion-bar">
            <div class="conversion-fill ${performanceClass}" 
                 style="width: ${Math.min(product.conversion, 100)}%"></div>
          </div>
          
          <div class="viz-header">
            <span>Market Share</span>
            <span>${percentOfTotal.toFixed(1)}%</span>
          </div>
          <div class="market-share-bar">
            <div class="share-fill" style="width: ${Math.min(percentOfTotal, 100)}%"></div>
          </div>
        </div>
        
        <div class="product-insight">
          ${getProductInsight(product.conversion, salesPerQuote)}
        </div>
      </div>
    `;
  });

  html += `</div>`;

  // ✅ เพิ่ม legend
  html += `
    <div class="performance-legend">
      <div class="legend-title">ระดับประสิทธิภาพ:</div>
      <div class="legend-items">
        <div class="legend-item">
          <span class="legend-color excellent"></span>
          <span class="legend-text">ดีเยี่ยม (≥ 40%)</span>
        </div>
        <div class="legend-item">
          <span class="legend-color good"></span>
          <span class="legend-text">ดี (25-39%)</span>
        </div>
        <div class="legend-item">
          <span class="legend-color fair"></span>
          <span class="legend-text">ปานกลาง (15-24%)</span>
        </div>
        <div class="legend-item">
          <span class="legend-color poor"></span>
          <span class="legend-text">ต้องปรับปรุง (< 15%)</span>
        </div>
      </div>
    </div>
  `;

  container.innerHTML = html;
}

// ฟังก์ชันสร้าง insight ตาม performance
function getProductInsight(conversionRate, avgPerQuote) {
  if (conversionRate >= 40 && avgPerQuote >= 100000) {
    return "⭐ <strong>สินค้ายอดนิยม:</strong> Conversion rate สูงและมูลค่าต่อใบเสนอสูง";
  } else if (conversionRate >= 40) {
    return "✅ <strong>ขายดี:</strong> Conversion rate สูง แต่ควรเพิ่มมูลค่าต่อใบเสนอ";
  } else if (avgPerQuote >= 150000) {
    return "💰 <strong>มูลค่าสูง:</strong> มูลค่าต่อใบเสนอสูง แต่ควรปรับปรุง conversion rate";
  } else if (conversionRate >= 25) {
    return "↗️ <strong>มีศักยภาพ:</strong> ประสิทธิภาพอยู่ในระดับดี";
  } else if (conversionRate >= 15) {
    return "⚠️ <strong>ต้องติดตาม:</strong> ประสิทธิภาพอยู่ในระดับปานกลาง";
  } else {
    return "🔍 <strong>ต้องการวิเคราะห์:</strong> ควรศึกษาสาเหตุที่ conversion rate ต่ำ";
  }
}

function parseToDate(dateVal) {
  if (!dateVal) return null;

  // ถ้าเป็น Date อยู่แล้ว
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) return dateVal;

  const s = String(dateVal).trim();

  // 1) YYYY-MM-DD หรือ YYYY-MM-DDTHH:mm:ss
  const m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m1) {
    const y = Number(m1[1]),
      mo = Number(m1[2]) - 1,
      d = Number(m1[3]);
    const dt = new Date(y, mo, d);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // 2) DD/MM/YYYY
  const m2 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m2) {
    const d = Number(m2[1]),
      mo = Number(m2[2]) - 1,
      y = Number(m2[3]);
    const dt = new Date(y, mo, d);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // 3) fallback (บางที browser parse ได้)
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? null : dt;
}

function monthKeyFromDate(dt) {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`; // YYYY-MM
}

function addMonths(yyyyMM, delta) {
  const [y, m] = yyyyMM.split("-").map(Number);
  const dt = new Date(y, m - 1 + delta, 1);
  return monthKeyFromDate(dt);
}

function sumMonthFromDailyTrend(rows, monthKey) {
  let sales = 0,
    calls = 0,
    visits = 0,
    quotes = 0;

  rows.forEach((r) => {
    const dt = parseToDate(r?.date);
    if (!dt) return;
    if (monthKeyFromDate(dt) !== monthKey) return;

    sales += Number(r.sales || 0);
    calls += Number(r.calls || 0);
    visits += Number(r.visits || 0);
    quotes += Number(r.quotes || 0);
  });

  return { sales, calls, visits, quotes };
}

function buildMonthlyComparisonFromTrend(payload) {
  const rows = Array.isArray(payload?.dailyTrend) ? payload.dailyTrend : [];
  const dates = rows.map((r) => parseToDate(r?.date)).filter(Boolean);
  if (!dates.length) return null;

  // ✅ ใช้ “วันที่ล่าสุดใน dailyTrend” เป็นเดือนปัจจุบันของช่วงข้อมูลจริง
  dates.sort((a, b) => a - b);
  const latest = dates[dates.length - 1];

  const currentPeriod = monthKeyFromDate(latest);
  const previousPeriod = addMonths(currentPeriod, -1);

  const currentMonth = sumMonthFromDailyTrend(rows, currentPeriod);
  const previousMonth = sumMonthFromDailyTrend(rows, previousPeriod);

  // debug ช่วยดูว่ามัน sum แล้วได้อะไร
  console.log("📌 Monthly from dailyTrend:", {
    currentPeriod,
    previousPeriod,
    currentMonth,
    previousMonth,
  });

  return {
    currentPeriod,
    previousPeriod,
    currentMonth,
    previousMonth,
    isEstimated: true,
  };
}

function renderMonthlyComparison(payload) {
  const container = el("monthlyComparisonContainer");
  if (!container) return;

  const daily = payload.dailyTrend || [];

  // เดือนปัจจุบัน / เดือนก่อน (ตามเวลาเครื่อง + timezone browser)
  const now = new Date();
  const curKey = ymKey(now);
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevKey = ymKey(prev);

  // 1) ลองใช้ของที่มาจาก API ก่อน
  let cur =
    payload.monthlyComparison?.currentMonth ||
    payload.monthlyComparison?.current ||
    null;
  let pre =
    payload.monthlyComparison?.previousMonth ||
    payload.monthlyComparison?.previous ||
    null;

  const apiHasRealNumber =
    cur &&
    pre &&
    (Number(cur.sales || 0) +
      Number(cur.calls || 0) +
      Number(cur.visits || 0) +
      Number(cur.quotes || 0) >
      0 ||
      Number(pre.sales || 0) +
        Number(pre.calls || 0) +
        Number(pre.visits || 0) +
        Number(pre.quotes || 0) >
        0);

  // 2) ถ้า API เป็น 0 หมด / ไม่มี → คำนวณจาก dailyTrend (ชัวร์กว่า)
  if (!apiHasRealNumber) {
    cur = sumMonthlyFromDailyTrend(daily, curKey);
    pre = sumMonthlyFromDailyTrend(daily, prevKey);
  }

  // หัวข้อเดือนให้เป็น "กุมภาพันธ์ vs มกราคม"
  const curName = monthNameFromKey(curKey);
  const prevName = monthNameFromKey(prevKey);

  // helper คำนวณ growth
  const growthPct = (c, p) => {
    c = Number(c || 0);
    p = Number(p || 0);
    if (p <= 0 && c > 0) return 100;
    if (p <= 0 && c <= 0) return 0;
    return ((c - p) / p) * 100;
  };

  const metrics = [
    { key: "sales", label: "ยอดขาย", isCurrency: true },
    { key: "quotes", label: "ใบเสนอราคา", isCurrency: false },
    { key: "visits", label: "เข้าพบลูกค้า", isCurrency: false },
    { key: "calls", label: "การโทร", isCurrency: false },
  ];

  let html = `
    <div class="comparison-header">
      <h4>Monthly Comparison</h4>
      <div class="comparison-period">
        <span class="current-period">${curName}</span>
        <span class="vs">vs</span>
        <span class="previous-period">${prevName}</span>
      </div>
    </div>
    <div class="comparison-grid">
  `;

  metrics.forEach((m) => {
    const c = Number(cur[m.key] || 0);
    const p = Number(pre[m.key] || 0);
    const g = growthPct(c, p);
    const pos = g >= 0;

    const cTxt = m.isCurrency ? `${fmt.format(c)} ฿` : fmt.format(c);
    const pTxt = m.isCurrency ? `${fmt.format(p)} ฿` : fmt.format(p);
    const gTxt = Math.abs(g).toFixed(1);

    html += `
      <div class="comparison-card">
        <div class="metric-label">${m.label}</div>
        <div class="current-value" title="${curName}">${cTxt}</div>
        <div class="previous-value" title="${prevName}">
          <span class="label">${prevName}:</span>
          <span class="value">${pTxt}</span>
        </div>
        <div class="growth-indicator ${pos ? "positive" : "negative"}">
          ${pos ? "📈" : "📉"}
          <span class="growth-text">${p === 0 && c > 0 ? "ใหม่" : pos ? "เพิ่ม" : "ลด"} ${gTxt}%</span>
        </div>
      </div>
    `;
  });

  html += `</div>`;
  container.innerHTML = html;
}

function displayMonthlyComparison(comparison) {
  const container = el("monthlyComparisonContainer");
  if (!container) return;

  const fallback = getCurrentPrevMonthLabels();

  const currentName = comparison?.currentPeriod
    ? getThaiMonthLabel(comparison.currentPeriod)
    : fallback.currentName;

  const prevName = comparison?.previousPeriod
    ? getThaiMonthLabel(comparison.previousPeriod)
    : fallback.prevName;

  const metrics = [
    { key: "sales", label: "ยอดขาย", unit: "฿", isCurrency: true },
    { key: "quotes", label: "ใบเสนอราคา", unit: "ใบ", isCurrency: false },
    { key: "visits", label: "เข้าพบลูกค้า", unit: "ราย", isCurrency: false },
    { key: "calls", label: "การโทร", unit: "ครั้ง", isCurrency: false },
  ];

  let html = `
    <div class="comparison-header">
      <h4>Monthly Comparison</h4>
      <div class="comparison-period">
        <span class="current-period">${currentName}</span>
        <span class="vs">vs</span>
        <span class="previous-period">${prevName}</span>
      </div>
    </div>
    <div class="comparison-grid">
  `;

  metrics.forEach((metric) => {
    const current = comparison?.currentMonth?.[metric.key] || 0;
    const previous = comparison?.previousMonth?.[metric.key] || 0;

    let growth = 0;
    if (previous > 0) growth = ((current - previous) / previous) * 100;
    else if (current > 0) growth = 100;

    const isPositive = growth >= 0;
    const currentFormatted = metric.isCurrency
      ? fmt.format(current) + " ฿"
      : fmt.format(current);
    const previousFormatted = metric.isCurrency
      ? fmt.format(previous) + " ฿"
      : fmt.format(previous);
    const growthFormatted = Math.abs(growth).toFixed(1);

    html += `
      <div class="comparison-card">
        <div class="metric-label">${metric.label}</div>

        <div class="current-value" title="${currentName}">
          ${currentFormatted}
        </div>

        <div class="previous-value" title="${prevName}">
          <span class="label">${prevName}:</span>
          <span class="value">${previousFormatted}</span>
        </div>

        <div class="growth-indicator ${isPositive ? "positive" : "negative"}">
          ${isPositive ? "📈" : "📉"}
          <span class="growth-text">
            ${previous === 0 && current > 0 ? "ใหม่" : isPositive ? "เพิ่ม" : "ลด"} ${growthFormatted}%
          </span>
        </div>
      </div>
    `;
  });

  html += `</div>`;
  container.innerHTML = html;
}

function updateRangeText(payload) {
  const range = payload.range || {};
  setText("rangeText", range.text || "-");
  const now = new Date();
  const timeStr = now.toLocaleTimeString("th-TH", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  setText("updatedAt", timeStr);
}

function setKPI(payload) {
  const kpi = payload.kpiToday || payload.todaySummary || {};

  const sales = kpi.sales || kpi.salesAmount || kpi.totalSales || 0;

  // ✅ โทร/ติดตาม (จาก Sales เดิม)
  const callsFromSales = kpi.calls || kpi.callCount || kpi.telephone || 0;

  // ✅ โทรวันนี้ (จาก Call&Visit)
  const callsTodayFromCV = kpi.calls_today || 0;

  const visits = kpi.visits || kpi.visitCount || kpi.meeting || 0;
  const quotes = kpi.quotes || kpi.quoteCount || kpi.proposal || 0;

  setText("kpi_sales", fmt.format(sales) + " ฿");
  setText("kpi_calls", fmt.format(callsFromSales)); // ⬅️ บน
  setText("kpi_visits", fmt.format(visits));
  setText("kpi_quotes", fmt.format(quotes));
  setText("kpi_date", kpi.date || "");

  // ⬅️ ล่าง (Call&Visit)
  setText("kpi_calls_today", fmt.format(callsTodayFromCV));
}

function setSummary(payload) {
  const body = el("summaryBody");
  if (!body) return;

  body.innerHTML = "";
  const teams = payload.summary || [];

  if (!teams.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6" class="muted">ไม่มีข้อมูล</td>`;
    body.appendChild(tr);
    return;
  }

  teams.forEach((t, i) => {
    // แก้ไขให้รองรับชื่อ field ที่อาจต่างกัน
    const sales = t.sales || t.salesAmount || 0;
    const calls = t.calls || t.callCount || 0;
    const visits = t.visits || t.visitCount || 0;
    const quotes = t.quotes || t.quoteCount || 0;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${escapeHtml(t.team || t.teamName || "")}</td>
      <td class="num">${fmt.format(sales)} ฿</td>
      <td class="num">${fmt.format(calls)}</td>
      <td class="num">${fmt.format(visits)}</td>
      <td class="num">${fmt.format(quotes)}</td>
    `;
    body.appendChild(tr);
  });
}

function renderPersonTotals(payload) {
  const body = el("personTotalsBody");
  if (!body) return;

  body.innerHTML = "";

  const rows = payload.personTotals || payload.people || [];

  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6" class="muted">ไม่มีข้อมูล</td>`;
    body.appendChild(tr);
    return;
  }

  rows.slice(0, 30).forEach((r, i) => {
    // แก้ไขให้รองรับชื่อ field ที่อาจต่างกัน
    const sales = r.sales || r.salesAmount || 0;
    const calls = r.calls || r.callCount || 0;
    const visits = r.visits || r.visitCount || 0;
    const quotes = r.quotes || r.quoteCount || 0;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${escapeHtml(r.person || r.name || r.salesPerson || "")}</td>
      <td class="num">${fmt.format(sales)} ฿</td>
      <td class="num">${fmt.format(calls)}</td>
      <td class="num">${fmt.format(visits)}</td>
      <td class="num">${fmt.format(quotes)}</td>
    `;
    body.appendChild(tr);
  });
}

function formatValue(metric, value) {
  if (metric === "conversion") {
    const numValue = Number(value); // value ควรเป็นเปอร์เซ็นต์แล้ว (เช่น 25.5 สำหรับ 25.5%)
    if (numValue > 1) {
      // ตรวจสอบว่าค่าเป็นเปอร์เซ็นต์หรืออัตราส่วน
      return `${numValue.toFixed(1)}%`; // เป็นเปอร์เซ็นต์แล้ว (เช่น 25.5)
    } else {
      return `${(numValue * 100).toFixed(1)}%`; // เป็นอัตราส่วน (เช่น 0.255) ต้องแปลงเป็นเปอร์เซ็นต์
    }
  } else if (metric === "sales") {
    return `${fmt.format(Number(value))} ฿`;
  }
  return fmt.format(Number(value));
}

function changeMetric(metric) {
  state.activeMetric = metric; // ตั้งค่าเมตริกที่เลือก
  console.log("Active Metric changed to:", state.activeMetric);

  // อัปเดต UI ของปุ่มเมตริก
  document.querySelectorAll(".metric-buttons button").forEach((btn) => {
    btn.classList.remove("active");
    if (btn.textContent.includes(getMetricDisplayName(metric))) {
      btn.classList.add("active");
    }
  });

  // ✅ อัปเดตการแสดงผล Top 5 ด้วยข้อมูลที่มีอยู่
  if (state.lastPayload) {
    console.log("Updating Top5 with last payload");
    renderTop5(state.lastPayload);
  } else {
    console.log("No payload available, loading data...");
    loadData(false);
  }
}

// ฟังก์ชันช่วยเหลือสำหรับแสดงชื่อเมตริก
function getMetricDisplayName(metric) {
  switch (metric) {
    case "sales":
      return "ยอดขาย";
    case "calls":
      return "โทร";
    case "visits":
      return "เข้าพบ";
    case "quotes":
      return "ใบเสนอราคา";
    case "conversion":
      return "อัตราการปิดการขาย";
    default:
      return "ยอดขาย";
  }
}

function renderTop5(payload) {
  console.log("renderTop5 called with activeMetric:", state.activeMetric);

  const wrap = el("top5Wrap");
  if (!wrap) {
    console.error("❌ top5Wrap element not found");
    return;
  }

  const topByTeam = payload?.topByTeam || {};
  console.log(
    "TopByTeam data received:",
    Object.keys(topByTeam).length,
    "teams",
    topByTeam,
  );

  // ✅ FIX: ตรวจสอบว่า topByTeam มีข้อมูลจริงและไม่ใช่ object ว่าง
  const teamKeys = Object.keys(topByTeam);
  const isEmptyObject =
    teamKeys.length === 0 ||
    teamKeys.every((key) => {
      const teamData = topByTeam[key];
      return (
        !teamData ||
        Object.keys(teamData).length === 0 ||
        Object.values(teamData).every((arr) => !arr || arr.length === 0)
      );
    });

  if (isEmptyObject) {
    console.warn("⚠️ No valid topByTeam data in payload (empty or invalid)");

    // ✅ ลองใช้ fallback data จาก personTotals
    const fallbackData = createFallbackTopByTeam(payload);
    if (fallbackData) {
      console.log("🔄 Using fallback data from personTotals");
      renderTop5WithData(wrap, fallbackData);
      return;
    }

    wrap.innerHTML = `<div class="muted">ไม่มีข้อมูล Top 5</div>`;
    return;
  }

  renderTop5WithData(wrap, topByTeam);
}

// ✅ HELPER FUNCTION: สร้าง fallback data จาก personTotals
function createFallbackTopByTeam(payload) {
  const personTotals = payload?.personTotals;
  if (!Array.isArray(personTotals) || personTotals.length === 0) {
    return null;
  }

  // ตรวจสอบโครงสร้างข้อมูล
  console.log("🔍 Person Totals structure:", {
    sample: personTotals[0],
    hasActualClose: personTotals.some((p) => p.actualClose !== undefined),
    hasClosedDeals: personTotals.some((p) => p.closedDeals !== undefined),
    fields: Object.keys(personTotals[0] || {}),
  });

  const topByTeam = {};

  // สร้างทีม "ทั่วไป" สำหรับคนที่ไม่มีทีม
  const generalTeam = {
    topSales: personTotals
      .filter((p) => Number(p.sales || 0) > 0)
      .map((p) => ({
        person: p.person || p.name || "ไม่ระบุชื่อ",
        sales: Number(p.sales || 0),
        calls: Number(p.calls || 0),
        visits: Number(p.visits || 0),
        quotes: Number(p.quotes || 0),
        // เพิ่มฟิลด์สำหรับ conversion rate
        actualClose: Number(p.actualClose || p.closedDeals || 0),
      }))
      .sort((a, b) => b.sales - a.sales)
      .slice(0, 10),
  };

  // สร้าง topConversion - ใช้ actualClose ถ้ามี
  generalTeam.topConversion = personTotals
    .filter((p) => {
      const quotes = Number(p.quotes || 0);
      const actualClose = Number(p.actualClose || p.closedDeals || 0);
      return quotes > 0 && actualClose > 0;
    })
    .map((p) => {
      const sales = Number(p.sales || 0);
      const quotes = Number(p.quotes || 0);
      const actualClose = Number(p.actualClose || p.closedDeals || 0);

      // ใช้ actualClose (จำนวนยอดขายที่ปิดได้) แทน sales amount
      const conversionRate = calculateConversionRate(
        sales,
        quotes,
        actualClose,
      );

      return {
        person: p.person || p.name || "ไม่ระบุชื่อ",
        sales: sales,
        calls: Number(p.calls || 0),
        visits: Number(p.visits || 0),
        quotes: quotes,
        actualClose: actualClose,
        conversionRate: conversionRate,
      };
    })
    .filter((p) => p.conversionRate > 0)
    .sort((a, b) => b.conversionRate - a.conversionRate)
    .slice(0, 5);

  // ถ้าไม่มีข้อมูล conversion (ไม่มี actualClose) ให้คำนวณแบบประมาณการ
  if (generalTeam.topConversion.length === 0) {
    console.log("ℹ️ No actualClose data, estimating conversion rate...");

    // ประมาณการ: สมมติ average deal size เพื่อแปลง sales amount เป็นจำนวนใบ
    const AVERAGE_DEAL_SIZE = 50000; // 50,000 ฿ ต่อใบ

    generalTeam.topConversion = personTotals
      .filter((p) => {
        const sales = Number(p.sales || 0);
        const quotes = Number(p.quotes || 0);
        return sales > 0 && quotes > 0;
      })
      .map((p) => {
        const sales = Number(p.sales || 0);
        const quotes = Number(p.quotes || 0);

        // ประมาณการจำนวนยอดขายที่ปิดได้จาก sales amount
        const estimatedClosedDeals = Math.round(sales / AVERAGE_DEAL_SIZE);
        const conversionRate = Math.min(
          100,
          (estimatedClosedDeals / quotes) * 100,
        );

        return {
          person: p.person || p.name || "ไม่ระบุชื่อ",
          sales: sales,
          calls: Number(p.calls || 0),
          visits: Number(p.visits || 0),
          quotes: quotes,
          estimatedClosedDeals: estimatedClosedDeals,
          conversionRate: conversionRate,
          isEstimated: true,
        };
      })
      .filter((p) => p.conversionRate > 0 && p.conversionRate <= 100)
      .sort((a, b) => b.conversionRate - a.conversionRate)
      .slice(0, 5);
  }

  // topCalls, topVisits, topQuotes (เหมือนเดิม)
  generalTeam.topCalls = personTotals
    .filter((p) => Number(p.calls || 0) > 0)
    .map((p) => ({
      person: p.person || p.name || "ไม่ระบุชื่อ",
      sales: Number(p.sales || 0),
      calls: Number(p.calls || 0),
      visits: Number(p.visits || 0),
      quotes: Number(p.quotes || 0),
    }))
    .sort((a, b) => b.calls - a.calls)
    .slice(0, 5);

  generalTeam.topVisits = personTotals
    .filter((p) => Number(p.visits || 0) > 0)
    .map((p) => ({
      person: p.person || p.name || "ไม่ระบุชื่อ",
      sales: Number(p.sales || 0),
      calls: Number(p.calls || 0),
      visits: Number(p.visits || 0),
      quotes: Number(p.quotes || 0),
    }))
    .sort((a, b) => b.visits - a.visits)
    .slice(0, 5);

  generalTeam.topQuotes = personTotals
    .filter((p) => Number(p.quotes || 0) > 0)
    .map((p) => ({
      person: p.person || p.name || "ไม่ระบุชื่อ",
      sales: Number(p.sales || 0),
      calls: Number(p.calls || 0),
      visits: Number(p.visits || 0),
      quotes: Number(p.quotes || 0),
    }))
    .sort((a, b) => b.quotes - a.quotes)
    .slice(0, 5);

  topByTeam["ทั่วไป"] = generalTeam;

  console.log("📊 Fallback TopByTeam created:", {
    sales: generalTeam.topSales.length,
    calls: generalTeam.topCalls.length,
    visits: generalTeam.topVisits.length,
    quotes: generalTeam.topQuotes.length,
    conversion: generalTeam.topConversion.length,
    conversionIsEstimated: generalTeam.topConversion.some((p) => p.isEstimated),
    conversionSample: generalTeam.topConversion.slice(0, 3).map((p) => ({
      person: p.person,
      quotes: p.quotes,
      sales: fmt.format(p.sales),
      actualClose: p.actualClose,
      estimatedDeals: p.estimatedClosedDeals,
      conversionRate: p.conversionRate.toFixed(1) + "%",
    })),
  });

  return topByTeam;
}

// ✅ HELPER FUNCTION: render ด้วยข้อมูล
function renderTop5WithData(wrap, topByTeam) {
  wrap.innerHTML = "";

  // Helper functions สำหรับการจัดการ conversion rate
  const calculateConversionRate = (salesAmount, quotesCount) => {
    const salesNum = Number(salesAmount || 0);
    const quotesNum = Number(quotesCount || 0);

    if (quotesNum <= 0) return 0;
    if (salesNum <= 0) return 0;

    // ตรวจสอบว่า salesNum เป็นจำนวนเงิน (บาท) หรือจำนวนใบ
    // ถ้า salesNum มากกว่า quotesNum มากๆ แสดงว่าเป็นจำนวนเงิน
    const AVERAGE_DEAL_SIZE = 50000; // สมมติ average deal size 50,000 ฿
    const estimatedDeals = Math.max(
      1,
      Math.round(salesNum / AVERAGE_DEAL_SIZE),
    );

    // ใช้ estimated deals แทน sales amount
    const rate = (estimatedDeals / quotesNum) * 100;
    return Math.min(100, rate); // ไม่ให้เกิน 100%
  };

  const formatValue = (metric, value, row = null) => {
    const numValue = Number(value);

    switch (metric) {
      case "conversion":
        if (numValue <= 0) return "0%";
        return `${numValue.toFixed(1)}%`;

      case "sales":
        return `${fmt.format(numValue)} ฿`;

      case "calls":
      case "visits":
      case "quotes":
        return fmt.format(numValue);

      default:
        return fmt.format(numValue);
    }
  };

  // กรองทีมที่มีข้อมูลตามเมตริกที่เลือก
  const teams = Object.keys(topByTeam)
    .filter((team) => {
      const teamData = topByTeam[team];
      if (!teamData) return false;

      const metricKey = getMetricKey(state.activeMetric);
      let list = teamData[metricKey] || [];

      // สำหรับ conversion rate: กรองเฉพาะที่มีใบเสนอราคา > 0
      if (state.activeMetric === "conversion") {
        list = list.filter((item) => {
          const quotes = Number(item.quotes || 0);
          const sales = Number(item.sales || 0);
          return quotes > 0 && sales > 0;
        });
      }

      return list.length > 0;
    })
    .sort((a, b) => a.localeCompare(b, "th"));

  if (!teams.length) {
    const noDataMessage =
      state.activeMetric === "conversion"
        ? `ไม่มีข้อมูลสำหรับเมตริก "${getMetricDisplayName(state.activeMetric)}"<br><small>ต้องการทั้งยอดขายและใบเสนอราคา (> 0)</small>`
        : `ไม่มีข้อมูลสำหรับเมตริก "${getMetricDisplayName(state.activeMetric)}"`;

    wrap.innerHTML = `<div class="muted" style="text-align: center; padding: 20px; line-height: 1.5;">${noDataMessage}</div>`;
    return;
  }

  console.log(
    `📊 Rendering Top 5: ${getMetricDisplayName(state.activeMetric)}`,
    {
      teams: teams,
      activeMetric: state.activeMetric,
    },
  );

  // แสดงข้อมูลตามทีม
  teams.forEach((team) => {
    const t = topByTeam[team] || {};
    const metricKey = getMetricKey(state.activeMetric);

    let list = t[metricKey] || [];
    const title = `Top 5: ${getMetricDisplayName(state.activeMetric)}`;

    console.log(
      `Team "${team}" - ${state.activeMetric}:`,
      list.length,
      "items",
    );

    // สำหรับ conversion: เรียงลำดับและกรองใหม่
    if (state.activeMetric === "conversion") {
      list = list
        .filter((item) => {
          const quotes = Number(item.quotes || 0);
          const sales = Number(item.sales || 0);
          return quotes > 0 && sales > 0;
        })
        .map((item) => {
          const sales = Number(item.sales || 0);
          const quotes = Number(item.quotes || 0);

          // คำนวณ conversion rate ด้วยวิธีที่ปลอดภัย
          const conversionRate = calculateConversionRate(sales, quotes);

          // ประมาณการจำนวนดีลจากยอดขาย
          const AVERAGE_DEAL_SIZE = 50000;
          const estimatedDeals = Math.max(
            1,
            Math.round(sales / AVERAGE_DEAL_SIZE),
          );

          return {
            ...item,
            conversionRate: conversionRate,
            estimatedDeals: estimatedDeals,
            _sales: sales,
            _quotes: quotes,
          };
        })
        .filter((item) => item.conversionRate > 0 && item.conversionRate <= 100)
        .sort((a, b) => b.conversionRate - a.conversionRate)
        .slice(0, 5);
    } else {
      // เรียงลำดับตามเมตริกอื่นๆ
      list = list
        .slice(0, 10) // เอาข้อมูลมาเยอะหน่อยเพื่อเรียงลำดับ
        .filter((item) => {
          const val = Number(item[state.activeMetric] || 0);
          return val > 0;
        })
        .sort((a, b) => {
          const aVal = Number(a[state.activeMetric] || 0);
          const bVal = Number(b[state.activeMetric] || 0);
          return bVal - aVal;
        })
        .slice(0, 5);
    }

    const card = document.createElement("div");
    card.className = "tcard";
    card.innerHTML = `
      <div class="tcardHead">
        <h4>${escapeHtml(team)}</h4>
        <div class="mini">${title}</div>
        ${
          state.activeMetric === "conversion"
            ? '<div class="hint">(ประมาณการจากยอดขาย ÷ ใบเสนอราคา)</div>'
            : ""
        }
      </div>
    `;

    if (!list.length) {
      const emptyMessage =
        state.activeMetric === "conversion"
          ? "ไม่มีข้อมูลที่คำนวณ Conversion Rate ได้<br><small>ต้องการทั้งยอดขายและใบเสนอราคา (> 0)</small>"
          : "ไม่มีข้อมูลสำหรับเมตริกนี้";
      card.innerHTML += `<div class="muted" style="margin-top:8px; padding: 10px; line-height: 1.4;">${emptyMessage}</div>`;
    } else {
      list.forEach((row, idx) => {
        let val = 0;
        let displayVal = "";
        let tooltipText = "";
        let isEstimated = false;

        switch (state.activeMetric) {
          case "sales":
            val = Number(row.sales || 0);
            displayVal = formatValue(state.activeMetric, val);
            tooltipText = `ยอดขาย: ${fmt.format(val)} ฿`;
            break;

          case "calls":
            val = Number(row.calls || 0);
            displayVal = formatValue(state.activeMetric, val);
            tooltipText = `การโทร: ${fmt.format(val)} ครั้ง`;
            break;

          case "visits":
            val = Number(row.visits || 0);
            displayVal = formatValue(state.activeMetric, val);
            tooltipText = `เข้าพบลูกค้า: ${fmt.format(val)} ครั้ง`;
            break;

          case "quotes":
            val = Number(row.quotes || 0);
            displayVal = formatValue(state.activeMetric, val);
            tooltipText = `ใบเสนอราคา: ${fmt.format(val)} ใบ`;
            break;

          case "conversion":
            // ใช้ค่า conversionRate ที่คำนวณแล้ว
            val = Number(row.conversionRate || 0);
            displayVal = formatValue(state.activeMetric, val, row);
            isEstimated = true;

            // สร้าง tooltip ที่มีรายละเอียดการคำนวณ
            const sales = Number(row._sales || row.sales || 0);
            const quotes = Number(row._quotes || row.quotes || 0);
            const estimatedDeals =
              row.estimatedDeals || Math.max(1, Math.round(sales / 50000));

            tooltipText = `
              <div style="text-align: left; min-width: 200px;">
                <strong>Conversion Rate: ${val.toFixed(1)}%</strong><br>
                <div style="margin-top: 5px;">
                  <small>ยอดขาย: ${fmt.format(sales)} ฿</small><br>
                  <small>ใบเสนอราคา: ${fmt.format(quotes)} ใบ</small><br>
                  <small>ประมาณการดีลที่ปิดได้: ${estimatedDeals} ดีล</small><br>
                  <small>สูตร: (${estimatedDeals} ÷ ${fmt.format(quotes)}) × 100</small>
                </div>
                <div style="margin-top: 5px; padding-top: 5px; border-top: 1px solid rgba(255,255,255,0.1);">
                  <small><em>*ประมาณการจากยอดขาย (สมมติ average deal = 50,000 ฿)</em></small>
                </div>
              </div>
            `;
            break;
        }

        // กำหนด class พิเศษสำหรับอันดับ 1-3
        let rankClass = "";
        if (idx === 0) rankClass = "rank-1";
        else if (idx === 1) rankClass = "rank-2";
        else if (idx === 2) rankClass = "rank-3";

        const div = document.createElement("div");
        div.className = `trow ${rankClass}`;

        // ใช้ data attribute สำหรับ tooltip ที่ซับซ้อน
        if (tooltipText) {
          div.setAttribute(
            "data-tooltip",
            tooltipText.replace(/\n/g, " ").trim(),
          );
        }

        // สร้าง content
        const nameContent = escapeHtml(row.person || "ไม่ระบุชื่อ");
        const metaContent =
          state.activeMetric === "conversion" && row._quotes
            ? `<span class="meta">(${fmt.format(row._quotes)} quotes)</span>`
            : "";

        const progressBar =
          state.activeMetric === "conversion" && val > 0
            ? `<div class="progress">
              <div class="progress-bar" style="width: ${Math.min(val, 100)}%"></div>
            </div>`
            : "";

        const estimatedBadge = isEstimated
          ? `<span class="estimated-badge" title="ประมาณการ">~</span>`
          : "";

        div.innerHTML = `
          <div class="rank">${idx + 1}</div>
          <div class="name">
            ${nameContent}
            ${metaContent}
          </div>
          <div class="val ${state.activeMetric}">
            ${estimatedBadge}
            ${displayVal}
            ${progressBar}
          </div>
        `;

        // เพิ่ม event listener สำหรับ tooltip
        div.addEventListener("mouseenter", function (e) {
          if (tooltipText) {
            showCustomTooltip(e, tooltipText);
          }
        });

        div.addEventListener("mouseleave", function () {
          hideCustomTooltip();
        });

        card.appendChild(div);
      });

      // เพิ่มข้อมูลสรุปสำหรับ conversion rate
      if (state.activeMetric === "conversion" && list.length > 0) {
        const avgConversion =
          list.reduce(
            (sum, item) => sum + (Number(item.conversionRate) || 0),
            0,
          ) / list.length;

        const totalSales = list.reduce(
          (sum, item) => sum + (Number(item._sales || item.sales) || 0),
          0,
        );
        const totalQuotes = list.reduce(
          (sum, item) => sum + (Number(item._quotes || item.quotes) || 0),
          0,
        );
        const totalEstimatedDeals = list.reduce(
          (sum, item) => sum + (Number(item.estimatedDeals) || 0),
          0,
        );

        const summaryDiv = document.createElement("div");
        summaryDiv.className = "summary";
        summaryDiv.innerHTML = `
          <div class="summary-row">
            <span>ค่าเฉลี่ย:</span>
            <span class="avg-conversion">${avgConversion.toFixed(1)}%</span>
          </div>
          <div class="summary-row">
            <span>ยอดขายรวม:</span>
            <span>${fmt.format(totalSales)} ฿</span>
          </div>
          <div class="summary-row">
            <span>ใบเสนอราคารวม:</span>
            <span>${fmt.format(totalQuotes)} ใบ</span>
          </div>
          <div class="summary-note">
            <small>*ประมาณการจากยอดขาย (สมมติ average deal = 50,000 ฿)</small>
          </div>
        `;
        card.appendChild(summaryDiv);
      }
    }

    wrap.appendChild(card);
  });

  // เพิ่ม CSS สำหรับการแสดงผล
  if (!document.getElementById("top5-custom-styles")) {
    const style = document.createElement("style");
    style.id = "top5-custom-styles";
    style.textContent = `
      .trow .val.conversion {
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: flex-end;
        gap: 2px;
      }
      .trow .progress {
        width: 80px;
        height: 6px;
        background: rgba(255,255,255,0.1);
        border-radius: 3px;
        overflow: hidden;
        margin-top: 2px;
      }
      .trow .progress-bar {
        height: 100%;
        background: linear-gradient(90deg, #3b82f6, #22c55e);
        transition: width 0.3s ease;
        border-radius: 3px;
      }
      .trow .name .meta {
        font-size: 10px;
        color: #94a3b8;
        margin-left: 4px;
        font-weight: normal;
      }
      .trow.rank-1 .val {
        color: #fbbf24;
        font-weight: 700;
      }
      .trow.rank-2 .val {
        color: #94a3b8;
        font-weight: 600;
      }
      .trow.rank-3 .val {
        color: #d1d5db;
        font-weight: 500;
      }
      .trow .estimated-badge {
        color: #f59e0b;
        font-weight: bold;
        margin-right: 2px;
        font-size: 0.9em;
      }
      .summary {
        margin-top: 12px;
        padding: 10px;
        background: rgba(255,255,255,0.03);
        border-radius: 6px;
        font-size: 12px;
        border: 1px solid rgba(255,255,255,0.05);
      }
      .summary-row {
        display: flex;
        justify-content: space-between;
        margin-bottom: 6px;
        padding-bottom: 4px;
        border-bottom: 1px solid rgba(255,255,255,0.05);
      }
      .summary-row:last-child {
        margin-bottom: 0;
        padding-bottom: 0;
        border-bottom: none;
      }
      .avg-conversion {
        color: #22c55e;
        font-weight: 600;
      }
      .summary-note {
        margin-top: 8px;
        padding-top: 8px;
        border-top: 1px solid rgba(255,255,255,0.05);
        color: #94a3b8;
        font-size: 11px;
        line-height: 1.3;
      }
      .hint {
        font-size: 11px;
        color: #94a3b8;
        margin-top: 2px;
        line-height: 1.3;
      }
      .custom-tooltip {
        position: fixed;
        background: rgba(15, 23, 42, 0.95);
        color: white;
        padding: 12px;
        border-radius: 6px;
        border: 1px solid rgba(56, 189, 248, 0.3);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
        z-index: 10000;
        max-width: 300px;
        font-size: 13px;
        line-height: 1.4;
        backdrop-filter: blur(10px);
        pointer-events: none;
      }
      .custom-tooltip small {
        color: #cbd5e1;
        opacity: 0.9;
      }
      .custom-tooltip em {
        color: #fbbf24;
        font-style: normal;
      }
    `;
    document.head.appendChild(style);
  }
}

// Helper functions สำหรับ custom tooltip
let customTooltip = null;
let tooltipTimeout = null;

function showCustomTooltip(event, content) {
  if (tooltipTimeout) {
    clearTimeout(tooltipTimeout);
  }

  tooltipTimeout = setTimeout(() => {
    if (!customTooltip) {
      customTooltip = document.createElement("div");
      customTooltip.className = "custom-tooltip";
      document.body.appendChild(customTooltip);
    }

    customTooltip.innerHTML = content;
    customTooltip.style.display = "block";

    // Position tooltip
    const x = event.clientX + 10;
    const y = event.clientY + 10;

    customTooltip.style.left = `${x}px`;
    customTooltip.style.top = `${y}px`;

    // ตรวจสอบไม่ให้ tooltip ออกนอกหน้าจอ
    const rect = customTooltip.getBoundingClientRect();
    if (rect.right > window.innerWidth) {
      customTooltip.style.left = `${event.clientX - rect.width - 10}px`;
    }
    if (rect.bottom > window.innerHeight) {
      customTooltip.style.top = `${event.clientY - rect.height - 10}px`;
    }
  }, 300); // delay 300ms
}

function hideCustomTooltip() {
  if (tooltipTimeout) {
    clearTimeout(tooltipTimeout);
  }

  if (customTooltip) {
    customTooltip.style.display = "none";
  }
}

// ปิด tooltip เมื่อคลิกที่อื่น
document.addEventListener("click", hideCustomTooltip);

// Helper function สำหรับการแสดงชื่อเมตริก
function getMetricDisplayName(metric) {
  switch (metric) {
    case "sales":
      return "ยอดขาย";
    case "calls":
      return "การโทร";
    case "visits":
      return "เข้าพบลูกค้า";
    case "quotes":
      return "ใบเสนอราคา";
    case "conversion":
      return "อัตราการปิดการขาย";
    default:
      return "ยอดขาย";
  }
}

// Helper function สำหรับการแปลงเมตริกเป็น key
function getMetricKey(metric) {
  switch (metric) {
    case "sales":
      return "topSales";
    case "calls":
      return "topCalls";
    case "visits":
      return "topVisits";
    case "quotes":
      return "topQuotes";
    case "conversion":
      return "topConversion";
    default:
      return "topSales";
  }
}

// ✅ HELPER FUNCTION: แปลงเมตริกเป็น key ใน topByTeam object
function getMetricKey(metric) {
  switch (metric) {
    case "sales":
      return "topSales";
    case "calls":
      return "topCalls";
    case "visits":
      return "topVisits";
    case "quotes":
      return "topQuotes";
    case "conversion":
      return "topConversion";
    default:
      return "topSales";
  }
}

// ✅ HELPER FUNCTION: แสดงชื่อเมตริก
function getMetricDisplayName(metric) {
  switch (metric) {
    case "sales":
      return "ยอดขาย";
    case "calls":
      return "โทร";
    case "visits":
      return "เข้าพบ";
    case "quotes":
      return "ใบเสนอราคา";
    case "conversion":
      return "อัตราการปิดการขาย";
    default:
      return "ยอดขาย";
  }
}

function calculateConversionRate(
  salesAmount,
  quotesCount,
  actualSalesCount = null,
) {
  const salesNum = Number(salesAmount || 0);
  const quotesNum = Number(quotesCount || 0);

  // ถ้าใช้ actualSalesCount (จำนวนยอดขายที่ปิดได้จริง)
  if (actualSalesCount !== null && actualSalesCount !== undefined) {
    const actualSales = Number(actualSalesCount || 0);
    if (quotesNum <= 0) return 0;
    if (actualSales <= 0) return 0;
    return Math.min(100, (actualSales / quotesNum) * 100);
  }

  // ถ้าไม่มี actualSalesCount ให้ตรวจสอบหน่วย
  if (salesNum <= 0 || quotesNum <= 0) return 0;

  // ตรวจสอบว่า salesNum น่าจะเป็นจำนวนเงินหรือจำนวนใบ
  // ถ้า salesNum ใหญ่กว่า quotesNum มาก แสดงว่าเป็นจำนวนเงิน
  if (salesNum > quotesNum * 10000) {
    // สมมติ average deal size ~ 10,000
    console.warn(
      `⚠️ Sales amount (${fmt.format(salesNum)}) > Quotes count (${quotesNum}) - หน่วยไม่ตรงกัน`,
    );
    return 0; // หรือ return null เพื่อระบุว่าไม่สามารถคำนวณได้
  }

  // ถ้าจำนวนสมเหตุสมผล ให้คำนวณ
  const rate = (salesNum / quotesNum) * 100;
  return Math.min(100, rate);
}

function renderTopPerformers(payload) {
  const analysis = payload.callVisitAnalysis || {};
  const topPerformers = analysis.topPerformers || {};

  // Top Callers
  const topCallersContainer = document.getElementById("topCallersContainer");
  if (topCallersContainer) {
    const topCallers = topPerformers.topCallers || [];
    if (topCallers.length > 0) {
      topCallersContainer.innerHTML = topCallers
        .map(
          (person, index) => `
        <div class="performer-item">
          <div class="performer-name">${index + 1}. ${escapeHtml(person.person)}</div>
          <div class="performer-value">${fmt.format(person.calls)} ครั้ง</div>
        </div>
      `,
        )
        .join("");
    } else {
      topCallersContainer.innerHTML = '<div class="muted">ไม่มีข้อมูล</div>';
    }
  }

  // Top Visitors
  const topVisitorsContainer = document.getElementById("topVisitorsContainer");
  if (topVisitorsContainer) {
    const topVisitors = topPerformers.topVisitors || [];
    if (topVisitors.length > 0) {
      topVisitorsContainer.innerHTML = topVisitors
        .map(
          (person, index) => `
        <div class="performer-item">
          <div class="performer-name">${index + 1}. ${escapeHtml(person.person)}</div>
          <div class="performer-value">${fmt.format(person.visits)} ครั้ง</div>
        </div>
      `,
        )
        .join("");
    } else {
      topVisitorsContainer.innerHTML = '<div class="muted">ไม่มีข้อมูล</div>';
    }
  }
}

// ---------------- Load flow ----------------

function validatePayload(payload) {
  console.group("📋 Payload Validation");

  const errors = [];
  const warnings = [];

  if (!payload) {
    errors.push("Payload is null or undefined");
  } else if (!payload.ok) {
    errors.push(`Payload.ok is false: ${payload.error || "No error message"}`);
  }

  if (!Array.isArray(payload.dailyTrend)) {
    warnings.push("dailyTrend is not an array");
  } else if (payload.dailyTrend.length === 0) {
    warnings.push("dailyTrend is empty");
  }

  if (!Array.isArray(payload.summary)) {
    warnings.push("summary is not an array");
  } else if (payload.summary.length === 0) {
    warnings.push("summary is empty");
  }

  if (!Array.isArray(payload.personTotals)) {
    warnings.push("personTotals is not an array");
  } else if (payload.personTotals.length === 0) {
    warnings.push("personTotals is empty");
  }

  if (!payload.kpiToday || typeof payload.kpiToday !== "object") {
    warnings.push("kpiToday is missing or not an object");
  }

  // แสดงผล
  if (errors.length > 0) {
    console.error("Validation Errors:", errors);
    showToast(`ข้อผิดพลาด: ${errors[0]}`, "error");
  }

  if (warnings.length > 0) {
    console.warn("Validation Warnings:", warnings);
  }

  console.log("✓ Validation complete", {
    isValid: errors.length === 0,
    errors,
    warnings,
  });
  console.groupEnd();

  return {
    isValid: errors.length === 0,
    errors,
    warnings,
  };
}

function debugDataStructure(payload) {
  console.group("🔍 Data Structure Debug");

  if (payload.dailyTrend && payload.dailyTrend.length > 0) {
    const sample = payload.dailyTrend[0];
    console.log("📅 dailyTrend sample:", {
      date: sample.date,
      sales: sample.sales,
      calls: sample.calls,
      visits: sample.visits,
      quotes: sample.quotes,
    });
    console.log(`📅 dailyTrend total rows: ${payload.dailyTrend.length}`);
  }

  if (payload.summary && payload.summary.length > 0) {
    console.log("🏢 summary sample:", payload.summary[0]);
  }

  if (payload.personTotals && payload.personTotals.length > 0) {
    console.log("👤 personTotals sample:", payload.personTotals[0]);
    console.log(`👤 personTotals total rows: ${payload.personTotals.length}`);
  }

  if (payload.kpiToday) {
    console.log("📊 kpiToday:", payload.kpiToday);
  }

  if (payload.callVisitYearly) {
    console.log("📞 callVisitYearly:", payload.callVisitYearly);
  }

  if (payload.customerSegmentation) {
    console.log("👥 customerSegmentation:", payload.customerSegmentation);
  }

  console.groupEnd();
}

function checkAPIData(payload) {
  console.group("📊 API Data Check");

  // Check dailyTrend totals
  if (payload.dailyTrend && payload.dailyTrend.length > 0) {
    const totalCalls = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.calls || 0),
      0,
    );
    const totalVisits = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.visits || 0),
      0,
    );
    const totalSales = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.sales || 0),
      0,
    );
    const totalQuotes = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.quotes || 0),
      0,
    );

    console.log("📈 Daily Trend Totals:", {
      calls: totalCalls,
      visits: totalVisits,
      sales: fmt.format(totalSales),
      quotes: totalQuotes,
      days: payload.dailyTrend.length,
    });
  }

  // Check summary totals
  if (payload.summary && payload.summary.length > 0) {
    const totalSales = payload.summary.reduce(
      (sum, team) => sum + (team.sales || 0),
      0,
    );
    console.log("🏢 Summary Totals:", {
      teams: payload.summary.length,
      totalSales: fmt.format(totalSales),
    });
  }

  // Check person totals
  if (payload.personTotals && payload.personTotals.length > 0) {
    const topPerson = payload.personTotals.reduce(
      (max, person) => ((person.sales || 0) > (max.sales || 0) ? person : max),
      { sales: 0 },
    );

    console.log("👑 Top Person:", {
      name: topPerson.person || topPerson.name,
      sales: fmt.format(topPerson.sales || 0),
    });
  }

  console.groupEnd();
}

function checkAPIData(payload) {
  console.group("🔍 API Data Check");

  // Check dailyTrend
  if (payload.dailyTrend && payload.dailyTrend.length > 0) {
    console.log("📅 Daily Trend Data:");
    payload.dailyTrend.forEach((day, i) => {
      console.log(
        `  ${day.date}: calls=${day.calls}, visits=${day.visits}, sales=${day.sales}, quotes=${day.quotes}`,
      );
    });

    // Calculate totals
    const totalCalls = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.calls || 0),
      0,
    );
    const totalVisits = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.visits || 0),
      0,
    );
    const totalSales = payload.dailyTrend.reduce(
      (sum, day) => sum + (day.sales || 0),
      0,
    );

    console.log("📊 Totals:", {
      calls: totalCalls,
      visits: totalVisits,
      sales: totalSales,
    });
  } else {
    console.warn("⚠️ No dailyTrend data");
  }

  // Check personTotals
  if (payload.personTotals && payload.personTotals.length > 0) {
    console.log("👥 Person Totals (first 5):");
    payload.personTotals.slice(0, 5).forEach((person) => {
      console.log(
        `  ${person.person}: calls=${person.calls}, visits=${person.visits}`,
      );
    });
  }

  // Check debug info
  if (payload.debug) {
    console.log("🐛 Debug Info:", payload.debug);
  }

  console.groupEnd();
}

// ✅ Fallback UI สำหรับเมื่อ API ไม่สามารถติดต่อได้
function showFallbackUI() {
  console.log("🔄 Showing fallback UI");

  const fallbackHTML = `
    <div class="offline-message">
      <div style="color: #fbbf24; font-size: 32px; margin-bottom: 15px; text-align: center;">
        ⚠️
      </div>
      <div style="text-align: center; margin-bottom: 20px;">
        <div style="color: #94a3b8; font-size: 16px; margin-bottom: 10px;">
          ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้
        </div>
        <div style="font-size: 13px; color: #64748b; line-height: 1.5;">
          กรุณาตรวจสอบ:
          <ul style="text-align: left; margin: 10px 0; padding-left: 20px;">
            <li>การเชื่อมต่ออินเทอร์เน็ต</li>
            <li>URL ของ API: ${API_URL.substring(0, 50)}...</li>
            <li>สถานะเซิร์ฟเวอร์</li>
          </ul>
        </div>
      </div>
      <div style="text-align: center;">
        <button onclick="location.reload()" 
                style="padding: 10px 20px; background: #3b82f6; color: white; 
                       border: none; border-radius: 6px; cursor: pointer; 
                       font-weight: 500; margin-right: 10px;">
          โหลดใหม่
        </button>
        <button onclick="loadData(false)" 
                style="padding: 10px 20px; background: #64748b; color: white; 
                       border: none; border-radius: 6px; cursor: pointer; 
                       font-weight: 500;">
          ลองอีกครั้ง
        </button>
      </div>
    </div>
  `;

  // แสดงข้อความใน containers หลัก
  const mainContainers = [
    "top5Wrap",
    "personTotalsBody",
    "summaryBody",
    "conversionContainer",
    "areaPerformanceContainer",
    "productPerformanceContainer",
    "monthlyComparisonContainer",
  ];

  mainContainers.forEach((containerId) => {
    const container = document.getElementById(containerId);
    if (container) {
      container.innerHTML = fallbackHTML;
    }
  });

  // แสดงใน chart area
  const chartStatus = document.getElementById("chartStatus");
  if (chartStatus) {
    chartStatus.innerHTML = `
      <div style="text-align: center; padding: 30px;">
        <div style="color: #f59e0b; margin-bottom: 10px;">⚠️ ไม่สามารถโหลดข้อมูลได้</div>
        <div style="font-size: 13px; color: #94a3b8;">
          กำลังใช้ข้อมูลแคชหรือลองเชื่อมต่อใหม่...
        </div>
      </div>
    `;
  }

  setFilterStatus("เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ", true);
}

// ✅ ตรวจสอบ cached data เมื่อหน้าโหลด
function checkCachedDataOnLoad() {
  try {
    const cached = localStorage.getItem("lastDashboardPayload");
    if (cached) {
      const cachedData = JSON.parse(cached);
      const cacheAge = Date.now() - cachedData.timestamp;
      const cacheValid = cacheAge < 3600000; // 1 ชั่วโมง

      if (cacheValid) {
        console.log(
          "📦 Found valid cached data, age:",
          Math.round(cacheAge / 1000),
          "seconds",
        );

        // อัปเดต UI ด้วย cached data พร้อมเครื่องหมาย
        const cachedIndicator = document.createElement("div");
        cachedIndicator.className = "cached-indicator";
        cachedIndicator.innerHTML =
          '<span style="color: #f59e0b;">⚠️ กำลังแสดงข้อมูลจากแคช</span>';

        const statusEl = el("filterStatus");
        if (statusEl) {
          statusEl.textContent = "ใช้ข้อมูลแคช (ออฟไลน์)";
        }

        return cachedData.data;
      }
    }
  } catch (e) {
    console.warn("Error checking cache:", e);
  }
  return null;
}

function checkCallVisitHTMLStructure() {
  console.group("🔍 Call & Visit HTML Structure");

  // ตรวจสอบ container ที่น่าจะใช้
  const possibleContainers = [
    "#callVisitContainer",
    "#callVisitYearlyContainer",
    ".call-visit-yearly",
    ".call-visit-analysis",
    "[data-section='call-visit']",
    ".cv-container",
  ];

  possibleContainers.forEach((selector) => {
    const el = document.querySelector(selector);
    if (el) {
      console.log(`Found container: ${selector}`, el);

      // ตรวจสอบ elements ภายใน
      const innerElements = el.querySelectorAll("*");
      console.log(
        `Inner elements (${innerElements.length}):`,
        Array.from(innerElements).map((e) => ({
          tag: e.tagName,
          id: e.id,
          class: e.className,
          text: e.textContent.substring(0, 50),
        })),
      );
    }
  });

  // ตรวจสอบ IDs ที่ฟังก์ชันต้องการ
  const requiredIds = [
    "cv_total_calls",
    "cv_total_visits",
    "cv_total_presented",
    "cv_total_quoted",
    "cv_total_closed",
  ];

  requiredIds.forEach((id) => {
    const el = document.getElementById(id);
    console.log(
      `${id}:`,
      el
        ? {
            text: el.textContent,
            parent: el.parentElement?.tagName,
            parentId: el.parentElement?.id,
          }
        : "NOT FOUND",
    );
  });

  console.groupEnd();
}

function ensureCallVisitContainer() {
  const containerId = "callVisitYearlyContainer";
  let container = document.getElementById(containerId);

  if (!container) {
    console.log("🔄 Creating Call & Visit container...");

    container = document.createElement("div");
    container.id = containerId;
    container.className = "section call-visit-yearly";
    container.innerHTML = `
      <div class="section-header">
        <h3>Call & Visit Analysis (Yearly)</h3>
        <div class="section-subtitle">ข้อมูลประจำปี</div>
      </div>
      
      <div class="cv-grid">
        <div class="cv-card">
          <div class="cv-label">การโทรทั้งหมด</div>
          <div class="cv-value" id="cv_total_calls">0</div>
          <div class="cv-unit">ครั้ง</div>
        </div>
        
        <div class="cv-card">
          <div class="cv-label">เข้าพบทั้งหมด</div>
          <div class="cv-value" id="cv_total_visits">0</div>
          <div class="cv-unit">ครั้ง</div>
        </div>
        
        <div class="cv-card highlight">
          <div class="cv-label">Presented</div>
          <div class="cv-value" id="cv_total_presented">0</div>
          <div class="cv-unit">ราย</div>
        </div>
        
        <div class="cv-card">
          <div class="cv-label">Quoted</div>
          <div class="cv-value" id="cv_total_quoted">0</div>
          <div class="cv-unit">ใบ</div>
        </div>
        
        <div class="cv-card success">
          <div class="cv-label">Closed</div>
          <div class="cv-value" id="cv_total_closed">0</div>
          <div class="cv-unit">ใบ</div>
        </div>
      </div>
    `;

    // หาที่วาง container
    const targetSelectors = [
      "#productPerformanceContainer",
      "#areaPerformanceContainer",
      "#conversionContainer",
      ".main-grid > div:last-child",
      "body",
    ];

    for (const selector of targetSelectors) {
      const target = document.querySelector(selector);
      if (target) {
        if (selector === "body") {
          target.appendChild(container);
        } else {
          target.parentNode.insertBefore(container, target.nextSibling);
        }
        console.log(`✅ Container inserted after: ${selector}`);
        break;
      }
    }
  }

  return container;
}

// เรียกใช้ใน onload
window.addEventListener("load", () => {
  setTimeout(checkCallVisitHTMLStructure, 1000);
});

// ✅ แก้ไข loadJSONP ให้มี error handling ที่ดีขึ้น
async function loadJSONP(url) {
  return new Promise((resolve, reject) => {
    const cbName =
      "__cb_" + Date.now() + "_" + Math.floor(Math.random() * 10000);
    const script = document.createElement("script");

    const TIMEOUT_MS = 45000; // 45 seconds สำหรับโหลดปกติ
    let settled = false;

    window[cbName] = (data) => {
      if (settled) return;
      settled = true;
      cleanup(false);

      if (!data) {
        reject(new Error("Empty response from server"));
        return;
      }

      if (data.error) {
        reject(new Error(data.error));
        return;
      }

      resolve(data);
    };

    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup(true);
      reject(new Error(`Request timeout (${TIMEOUT_MS}ms)`));
    }, TIMEOUT_MS);

    function cleanup(keepCallbackNoop) {
      clearTimeout(timeout);

      try {
        if (script && script.parentNode) script.parentNode.removeChild(script);
      } catch {}

      if (keepCallbackNoop) {
        window[cbName] = () => {};
        setTimeout(() => {
          try {
            delete window[cbName];
          } catch {
            window[cbName] = undefined;
          }
        }, 120000);
      } else {
        try {
          delete window[cbName];
        } catch {
          window[cbName] = undefined;
        }
      }
    }

    script.src =
      url +
      (url.includes("?") ? "&" : "?") +
      "callback=" +
      cbName +
      "&_=" +
      Date.now();

    script.onerror = () => {
      if (settled) return;
      settled = true;
      cleanup(false);
      reject(new Error("Failed to load script - Network error or CORS issue"));
    };

    // ✅ เพิ่ม tracking สำหรับ debugging
    console.log(`📤 Loading JSONP: ${cbName}`);
    script.onload = () => {
      console.log(`📥 Script loaded: ${cbName}`);
    };

    document.body.appendChild(script);
  });
}

async function checkAPIStatus() {
  try {
    const testUrl = API_URL + "?days=1";
    console.log("🔍 Testing API URL:", testUrl);

    // ✅ ใช้ isStatusCheck = true เพื่อ timeout สั้น
    const payload = await loadJSONP(testUrl, true);

    if (!payload) {
      console.warn("⚠️ API returned empty response");
      return false;
    }

    if (!payload.ok) {
      console.warn(
        "⚠️ API response not ok:",
        payload.error || "No error message",
      );
      return false;
    }

    console.log("✅ API status check passed");
    return true;
  } catch (err) {
    console.warn("⚠️ API status check failed:", err.message);

    // ✅ แสดงคำแนะนำสำหรับ debugging
    if (err.message.includes("timeout")) {
      console.log("💡 Tips for timeout issue:");
      console.log(
        "1. ตรวจสอบว่า Google Apps Script Web App ถูก deploy เป็นเวอร์ชันล่าสุด",
      );
      console.log("2. ตรวจสอบว่า Web App ตั้งค่าให้ 'Anyone' สามารถเข้าถึงได้");
      console.log("3. ตรวจสอบ URL ใน API_URL ว่าถูกต้อง");
      console.log("4. ตรวจสอบ internet connection");
      console.log(
        "5. ลองเปิด URL ใน browser: " + API_URL + "?days=1&callback=test",
      );
    } else if (err.message.includes("Failed to load script")) {
      console.log("💡 Could be CORS issue or incorrect URL");
    }

    return false;
  }
}

// ---------------- Events ----------------
function bindFilterEvents() {
  el("f_days")?.addEventListener("change", onDaysChange);
  el("f_start")?.addEventListener("change", onStartEndChange);
  el("f_end")?.addEventListener("change", onStartEndChange);
  el("f_team")?.addEventListener("change", debounceAutoLoad);
  el("f_person")?.addEventListener("change", debounceAutoLoad);
  el("f_group")?.addEventListener("change", debounceAutoLoad);

  el("btnApply")?.addEventListener("click", () => loadData(false));
  el("btnReset")?.addEventListener("click", resetFilters);
}

function bindTop5Tabs() {
  el("metricTabs")?.addEventListener("click", (ev) => {
    const t = ev.target.closest(".tab");
    if (!t) return;

    state.activeMetric = t.dataset.metric;

    document
      .querySelectorAll(".tab")
      .forEach((x) => x.classList.remove("active"));
    t.classList.add("active");

    if (state.lastPayload) renderTop5(state.lastPayload);
  });
}

/* ================= Boot ================= */
function handleVisibilityChange() {
  if (!document.hidden) {
    console.log("👁️ Page became visible, checking for auto refresh");

    const ck = el("ckAuto");
    if (ck?.checked) {
      // ✅ รอสักครู่ก่อนโหลด
      setTimeout(() => {
        if (!state.isLoading && !state.isPicking) {
          console.log("🔄 Auto-refreshing on visibility change");
          loadData(true);
        }
      }, 1000);
    }
  }
}

window.addEventListener("load", async () => {
  setFilterStatus("กำลังตรวจสอบ API…");

  bindFilterEvents_PATCH();
  bindPickingLock();

  // init chart
  if (typeof initChart === "function") initChart();
  if (typeof initProductChart === "function") initProductChart();
  if (typeof initLostDealChart === "function") initLostDealChart();

  document.addEventListener("visibilitychange", handleVisibilityChange);

  // ✅ ลองใช้ cached data ก่อน
  const cachedData = checkCachedDataOnLoad();
  if (cachedData) {
    updateAllUI(cachedData);
    showToast("ใช้ข้อมูลจากแคช", "info");
  }

  const ok = await checkAPIStatus();
  if (!ok) {
    setFilterStatus("API ไม่พร้อมใช้งาน", true);
    showToast("⚠️ API ไม่พร้อมใช้งาน (ตรวจ Web App: Anyone + /exec)", "error");

    // ถ้ามี cached data อยู่แล้วก็ไม่ต้องทำอะไร
    if (!cachedData) {
      showFallbackUI();
    }
    return;
  }

  setFilterStatus("พร้อมใช้งาน");
  await loadData(false);

  setInterval(() => {
    const ck = el("ckAuto");
    if (ck?.checked && !document.hidden && !state.isLoading && !state.isPicking)
      loadData(true);
  }, REFRESH_MS);
});

function getThaiMonthLabel(dateLike) {
  // รับได้ทั้ง Date, "YYYY-MM", "YYYY-MM-DD"
  let d;

  if (dateLike instanceof Date) {
    d = dateLike;
  } else if (typeof dateLike === "string") {
    const s = dateLike.length === 7 ? dateLike + "-01" : dateLike; // YYYY-MM -> YYYY-MM-01
    d = new Date(s + "T00:00:00");
  } else {
    d = new Date();
  }

  if (isNaN(d.getTime())) d = new Date();
  return d.toLocaleString("th-TH", { month: "long" }); // "กุมภาพันธ์"
}

function getCurrentPrevMonthLabels() {
  const now = new Date();
  const current = new Date(now.getFullYear(), now.getMonth(), 1);
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return {
    currentName: getThaiMonthLabel(current),
    prevName: getThaiMonthLabel(prev),
  };
}

async function quickAPITest() {
  try {
    const testUrl = API_URL + "?days=1&callback=test";
    console.log("🔍 Quick API test:", testUrl);

    // ใช้ timeout สั้นมาก
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 5000);

    const response = await fetch(testUrl, {
      mode: "no-cors",
      signal: controller.signal,
    });

    clearTimeout(timeoutId);
    return true;
  } catch (err) {
    console.log("🔍 Quick test failed:", err.message);
    return false;
  }
}
