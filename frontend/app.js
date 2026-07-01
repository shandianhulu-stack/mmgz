const $ = (id) => document.getElementById(id);
const logBox = $("logBox");

function log(msg) {
  const time = new Date().toLocaleTimeString();
  logBox.textContent = `[${time}] ${msg}\n` + logBox.textContent;
}

function money(v) {
  if (v === null || v === undefined || v === "" || isNaN(Number(v))) return "-";
  return Number(v).toFixed(2);
}

function plain(v) {
  if (v === null || v === undefined || v === "") return "-";
  return escapeHtml(v);
}

function escapeHtml(v) {
  return String(v ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function checkHealth() {
  try {
    const res = await fetch("/api/health");
    const data = await res.json();
    $("healthStatus").textContent = data.ok ? `本地服务正常｜${data.version}` : "服务异常";
    $("healthStatus").classList.toggle("ok", !!data.ok);
  } catch (e) {
    $("healthStatus").textContent = "服务未连接";
  }
}

function getScoreFile() {
  const file = $("scoreFile").files[0];
  if (!file) {
    alert("请先选择米墨评分 Excel");
    return null;
  }
  return file;
}

function getSettlementFile() {
  return $("settlementFile").files[0] || null;
}

function calcBudgetDiff(row) {
  if (row.budget === null || row.budget === undefined || row.budget === "" || isNaN(Number(row.budget))) return "-";
  return money(Number(row.budget) - Number(row.amount));
}

function renderSheet1(rows) {
  const body = $("previewBody");
  body.innerHTML = "";
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.name)}</td>
      <td>${escapeHtml(row.phone)}</td>
      <td>${money(row.amount)}</td>
      <td>${plain(row.budget)}</td>
      <td>${calcBudgetDiff(row)}</td>
      <td>${plain(row.completed_count)}</td>
      <td>${plain(row.settle_count)}</td>
      <td>${money(row.max_total)}</td>
      <td>${plain(row.can_calc)}</td>
      <td>${money(row.payout)}</td>
      <td class="${Math.abs(Number(row.diff)) <= 0.005 ? "zero" : (Number(row.diff) < 0 ? "negative" : "positive")}">${money(row.diff)}</td>
    `;
    body.appendChild(tr);
  });
}

function renderFinalAmount(rows) {
  const body = $("finalBody");
  body.innerHTML = "";
  rows.forEach((row) => {
    const details = row.details || [];
    details.forEach((item, index) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${index === 0 ? escapeHtml(row.name) : ""}</td>
        <td>${index === 0 ? money(row.amount) : ""}</td>
        <td>${money(item.task_amount)}</td>
        <td>${escapeHtml(item.score_text)}</td>
        <td>${money(item.pay_amount)}</td>
        <td></td>
      `;
      body.appendChild(tr);
    });
    const total = document.createElement("tr");
    total.className = "total-row";
    total.innerHTML = `
      <td>合计</td>
      <td>${money(row.amount)}</td>
      <td>${plain(row.settle_count)} 场</td>
      <td>最低 ${escapeHtml(row.min_score_text || "-")} / 平均 ${escapeHtml(row.avg_score_text || "-")} / 最高 ${escapeHtml(row.max_score_text || "-")}</td>
      <td>${money(row.payout)}</td>
      <td class="${Math.abs(Number(row.diff)) <= 0.005 ? "zero" : (Number(row.diff) < 0 ? "negative" : "positive")}">${money(row.diff)}</td>
    `;
    body.appendChild(total);
  });
}

async function runPreview({ silent = false } = {}) {
  const file = getScoreFile();
  if (!file) return null;
  const settlement = getSettlementFile();
  const form = new FormData();
  form.append("score_file", file);
  form.append("strategy", $("scoreStrategy").value);
  if (settlement) form.append("settlement_file", settlement);
  $("previewBtn").disabled = true;
  if (!silent) {
    log("开始预览：" + file.name + (settlement ? `；结算过程表：${settlement.name}` : "；未上传结算过程表") + `；策略：${$("scoreStrategy").selectedOptions[0].text}`);
  }
  try {
    const res = await fetch("/api/wage-score/preview", { method: "POST", body: form, cache: "no-store" });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || "预览失败");
    $("kpiCount").textContent = data.count;
    $("kpiTarget").textContent = money(data.total_target);
    $("kpiPayout").textContent = money(data.total_payout);
    $("kpiDiff").textContent = money(data.total_diff);
    $("kpiMaxAbsDiff").textContent = money(data.max_abs_diff);
    $("kpiOverpayCount").textContent = data.overpay_count ?? 0;

    renderSheet1(data.rows);
    renderFinalAmount(data.rows);

    const warn = data.missing_settlement > 0 ? `；${data.missing_settlement} 人没有匹配到结算过程字段` : "";
    log(`预览已刷新：${data.count} 人${warn}。目标金额 ${money(data.total_target)}，实付金额 ${money(data.total_payout)}，全批次总差值 ${money(data.total_diff)}，单人最大差值 ${money(data.max_abs_diff)}，微超人数 ${data.overpay_count ?? 0}`);
    return data;
  } catch (e) {
    log("预览失败：" + e.message);
    alert("预览失败：" + e.message);
    return null;
  } finally {
    $("previewBtn").disabled = false;
  }
}

async function preview() {
  await runPreview({ silent: false });
}

async function generate() {
  const score = getScoreFile();
  if (!score) return;
  const settlement = getSettlementFile();
  const name = $("outputName").value || "网约工工资明细_生成版.xlsx";
  $("generateBtn").disabled = true;
  try {
    log("生成前自动刷新导出表预览和最终金额明细...");
    const previewData = await runPreview({ silent: true });
    if (!previewData) return;

    const form = new FormData();
    form.append("score_file", score);
    form.append("strategy", $("scoreStrategy").value);
    if (settlement) form.append("settlement_file", settlement);
    form.append("output_name", name);
    log("开始生成 Excel..." + (settlement ? "已带结算过程表。" : "未带结算过程表，过程字段不会完整。") + ` 策略：${$("scoreStrategy").selectedOptions[0].text}`);
    const res = await fetch("/api/wage-score/generate", { method: "POST", body: form, cache: "no-store" });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(text);
    }
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name.endsWith(".xlsx") ? name : `${name}.xlsx`;
    document.body.appendChild(a);
    a.click();
    const downloadName = a.download;
    a.remove();
    window.URL.revokeObjectURL(url);
    log("生成完成，浏览器已下载：" + downloadName + "。页面上的 Sheet1 和最终金额明细就是本次导出预览。");
  } catch (e) {
    log("生成失败：" + e.message);
    alert("生成失败：" + e.message);
  } finally {
    $("generateBtn").disabled = false;
  }
}

$("previewBtn").addEventListener("click", preview);
$("generateBtn").addEventListener("click", generate);
$("scoreFile").addEventListener("change", () => {
  const f = $("scoreFile").files[0];
  if (f && f.name.includes("米墨评分")) {
    $("outputName").value = f.name.replace("米墨评分", "网约工工资明细").replace(/\.xlsx$/i, "_生成版.xlsx");
  }
});
$("scoreStrategy").addEventListener("change", () => {
  if ($("scoreFile").files[0]) runPreview({ silent: true });
});
$("settlementFile").addEventListener("change", () => {
  if ($("scoreFile").files[0]) runPreview({ silent: true });
});
checkHealth();
