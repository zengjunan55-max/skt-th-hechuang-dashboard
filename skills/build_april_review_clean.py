from __future__ import annotations

import importlib.util
import json
import re
from collections import Counter
from html import escape
from pathlib import Path
from types import ModuleType
from typing import Any

import pandas as pd
import yaml


_THIS_FILE = Path(__file__).resolve()
_BASE_PATH = _THIS_FILE.with_name("build_april_review_legacy.py")


def _load_base_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location("build_april_review_base", _BASE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load base generator from {_BASE_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_BASE = _load_base_module()
_BASE_BUILD_PAYLOAD = _BASE.build_payload
_BASE_RENDER_HTML = _BASE.render_html
_BASE_INFER_PERIOD = _BASE.infer_period

_COMPETITOR_SKILL_NAME = "$skintific-competitor-product-analysis"
_COMPETITOR_SKILL_DIR = _THIS_FILE.with_name("skintific-competitor-product-analysis")
_COMPETITOR_RULES_PATH = _COMPETITOR_SKILL_DIR / "references" / "competitor-rules.md"

_THAI_SKINCARE_CATEGORIES = {
    "精华",
    "面霜",
    "防晒",
    "爽肤水",
    "洁面",
    "套组",
    "面膜",
    "卸妆",
    "祛痘",
    "其他",
    "唇部护理",
    "定妆喷雾",
    "眼霜",
    "赠品",
    "泥膜",
    "小样",
    "已下架",
}
_THAI_MAKEUP_CATEGORIES = {"底妆"}


_SHEET_NAME_FALLBACKS = {
    "outer_ads": ["站外广告", "绔欏骞垮憡"],
    "shop_gmv": ["店铺GMV", "搴楅摵GMV"],
    "inner_ads": ["站内广告", "绔欏唴骞垮憡"],
    "tt_sales": ["TTSP销量数据源", "TTSP閿€閲忔暟鎹簮"],
    "april_temp": ["GPT_4月分析临时", "GPT_4鏈堝垎鏋愪复鏃?"],
}


def _safe_text_local(value: Any) -> str:
    return _BASE.safe_text(value) if hasattr(_BASE, "safe_text") else ("" if value is None else str(value).strip())


def _thai_group_from_category(value: Any) -> str:
    category = _safe_text_local(value)
    if category in _THAI_MAKEUP_CATEGORIES:
        return "底妆"
    if category in _THAI_SKINCARE_CATEGORIES:
        return "护肤"
    return "未归类"


def _is_thai_workbook(excel_file: pd.ExcelFile) -> bool:
    names = set(excel_file.sheet_names)
    return {"站外广告", "全站产品数据", "站内广告", "TT出货", "SP出货"}.issubset(names)


def _build_synthetic_temp_from_frames(
    outer: pd.DataFrame,
    shop: pd.DataFrame,
    inner: pd.DataFrame,
    tt: pd.DataFrame,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    if not shop.empty:
        shop_daily = (
            shop.groupby("日期", dropna=False)[["GMV-人民币", "链接访客", "订单数"]]
            .sum()
            .reset_index()
            .rename(columns={"日期": "date", "GMV-人民币": "GMV_RMB", "链接访客": "Visitors", "订单数": "Orders"})
        )
        shop_daily["date"] = pd.to_datetime(shop_daily["date"], errors="coerce")
        frames.append(shop_daily)

    if not outer.empty:
        outer_daily = (
            outer.groupby("Date_start", dropna=False)[["Spend", "Impressions", "Link_Click", "Purchase Value", "Conversions"]]
            .sum()
            .reset_index()
            .rename(
                columns={
                    "Date_start": "date",
                    "Spend": "OuterSpend",
                    "Impressions": "Impressions",
                    "Link_Click": "LPV",
                    "Purchase Value": "AdGMV",
                    "Conversions": "AdOrders",
                }
            )
        )
        outer_daily["date"] = pd.to_datetime(outer_daily["date"], errors="coerce")
        frames.append(outer_daily)

    if not inner.empty:
        inner_daily = (
            inner.groupby("date", dropna=False)[["Spend-人民币"]]
            .sum()
            .reset_index()
            .rename(columns={"date": "date", "Spend-人民币": "InnerSpend"})
        )
        inner_daily["date"] = pd.to_datetime(inner_daily["date"], errors="coerce")
        frames.append(inner_daily)

    if not tt.empty:
        tt_daily = (
            tt.groupby(["时间", "渠道"], dropna=False)[["总销量"]]
            .sum()
            .reset_index()
            .pivot_table(index="时间", columns="渠道", values="总销量", fill_value=0.0, aggfunc="sum")
            .reset_index()
            .rename(columns={"时间": "date", "TT": "TT", "Shopee": "Shopee"})
        )
        tt_daily["date"] = pd.to_datetime(tt_daily["date"], errors="coerce")
        frames.append(tt_daily)

    if not frames:
        return pd.DataFrame(columns=["date", "GMV_RMB", "Visitors", "Orders", "OuterSpend", "Impressions", "LPV", "AdGMV", "AdOrders", "InnerSpend", "Shopee", "TT"])

    merged = None
    for frame in frames:
        merged = frame if merged is None else merged.merge(frame, on="date", how="outer")

    merged = merged.fillna(0.0)
    merged["date"] = pd.to_datetime(merged["date"], errors="coerce")
    merged = merged.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
    for column in ["GMV_RMB", "Visitors", "Orders", "OuterSpend", "Impressions", "LPV", "AdGMV", "AdOrders", "InnerSpend", "Shopee", "TT"]:
        if column not in merged.columns:
            merged[column] = 0.0
    return merged


def _load_thai_workbook_frames(workbook_path: Path) -> dict[str, pd.DataFrame]:
    outer_raw = pd.read_excel(workbook_path, sheet_name="站外广告")
    shop_raw = pd.read_excel(workbook_path, sheet_name="全站产品数据")
    inner_raw = pd.read_excel(workbook_path, sheet_name="站内广告")
    tt_raw = pd.read_excel(workbook_path, sheet_name="TT出货")
    sp_raw = pd.read_excel(workbook_path, sheet_name="SP出货")

    outer = outer_raw.rename(
        columns={
            "日期": "Date_start",
            "花费金额（USD）": "Spend",
            "展示次数": "Impressions",
            "点击量": "Link_Click",
            "购物转化价值": "Purchase Value",
            "订单数": "Conversions",
            "产品": "产品",
            "产品": "站外产品命名",
            "产品": "站内名称",
            "店铺": "店铺区分",
            "店铺": "站内店铺命名",
            "品类": "品类",
            "广告形式": "模块",
            "广告形式2": "板块",
            "广告类型": "广告类型",
            "合创广告编号": "合创广告编号",
            "广告人群": "人群",
        }
    ).copy()
    ad_type = outer.get("广告类型", pd.Series("", index=outer.index)).fillna("").astype(str).str.strip()
    board = outer.get("板块", pd.Series("", index=outer.index)).fillna("").astype(str).str.strip()
    outer["模块"] = "其他"
    outer.loc[ad_type.eq("图片广告"), "模块"] = "图文"
    outer.loc[ad_type.eq("视频广告"), "模块"] = "视频"
    outer.loc[ad_type.eq("混合目录"), "模块"] = "混合目录"
    outer.loc[ad_type.eq("合创广告") & board.eq("收割广告"), "模块"] = "合创转化"
    outer.loc[ad_type.eq("合创广告") & board.eq("种草广告"), "模块"] = "合创曝光"
    outer["汇率"] = 7.0
    outer["分组情况"] = outer["品类"].map(_thai_group_from_category)
    outer["素材类型"] = outer.get("广告类型", "")
    outer["产品"] = outer_raw["产品"]
    outer["站外产品命名"] = outer_raw["产品"]
    outer["站内名称"] = outer_raw["产品"]
    outer["店铺区分"] = outer_raw["店铺"]
    outer["站内店铺命名"] = outer_raw["店铺"]
    outer = outer[
        [
            "Date_start",
            "Spend",
            "Impressions",
            "Link_Click",
            "Purchase Value",
            "Conversions",
            "产品",
            "站外产品命名",
            "站内名称",
            "店铺区分",
            "站内店铺命名",
            "品类",
            "分组情况",
            "模块",
            "汇率",
            "板块",
            "广告类型",
            "人群",
            "素材类型",
        ]
    ]

    shop = pd.DataFrame(
        {
            "日期": shop_raw["date"],
            "产品ID": shop_raw["商品名称"],
            "品类": shop_raw["品类"],
            "站内产品命名": shop_raw["商品名称"],
            "站外产品命名": shop_raw["链接名称"],
            "分组情况": shop_raw["品类"].map(_thai_group_from_category),
            "店铺": shop_raw["店铺2"],
            "GMV-人民币": shop_raw["GMV-RMB"],
            "链接访客": shop_raw["Product Visitors (Visit)"],
            "订单数": shop_raw["Units (Confirmed Order)"],
        }
    )

    inner = inner_raw.rename(columns={"Ads Spend(THB)": "Ads Spend(THB)"}).copy()
    inner["产品ID"] = inner["站内产品命名"]
    inner = inner[
        [
            "date",
            "产品ID",
            "类目",
            "站内产品命名",
            "站外产品命名",
            "Spend-人民币",
        ]
    ]

    tt = pd.concat(
        [
            tt_raw.rename(columns={"日期": "时间", "TT销量": "总销量", "产品名称": "产品名"}),
            sp_raw.rename(columns={"日期": "时间", "虾皮销量": "总销量", "产品名称": "产品名", "渠道": "渠道"}),
        ],
        ignore_index=True,
    )
    tt["分组"] = tt["品类"].map(_thai_group_from_category)
    tt["渠道"] = tt["渠道"].replace({"虾皮": "Shopee"})
    tt = tt[["时间", "sku", "中文产品名", "总销量", "渠道", "品类", "产品名", "分组"]]

    temp = _build_synthetic_temp_from_frames(outer=outer, shop=shop, inner=inner, tt=tt)
    return {"outer": outer, "shop": shop, "inner": inner, "tt": tt, "temp": temp}


_MODULE_PROXY = """function renderModuleSection(){
        if(typeof window.renderModuleSectionPatched === "function"){
          return window.renderModuleSectionPatched();
        }
        return null;
      }

      """


_MODULE_RENDER_PATCH = """function renderModuleSection(){
        ensureModuleFilterUi();
        const rows = buildModuleRowsScoped();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"Ad GMV",
          priorBarName:"Prior Ad GMV",
          lineName:"CVR",
          priorLineName:"Prior CVR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        drawBarLineCombo("module-gmv-ctr-chart", rows, {
          barField:"spend",
          priorBarField:"priorSpend",
          lineField:"ctr",
          priorLineField:"priorCtr",
          barName:"Outer Spend",
          priorBarName:"Prior Outer Spend",
          lineName:"CTR",
          priorLineName:"Prior CTR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        const daily = buildModuleDailySeries(state.moduleViewFilter || "");
        drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "Current CPM", "Prior CPM", "#2E5AA7", "#88C5FF", "CPM", false);
        drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "Current Impr.", "Prior Impr.", "#FFA62B", "#F8E6A0", "Impressions", false);

        const scope = moduleScopeText();
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.clicks) > 0).sort((a,b)=>safeNum(b.cvr) - safeNum(a.cvr))[0];
        const topSpend = [...rows].sort((a,b)=>safeNum(b.spend) - safeNum(a.spend))[0];
        const bestCtr = [...rows].filter((row)=>safeNum(row.impressions) > 0).sort((a,b)=>safeNum(b.ctr) - safeNum(a.ctr))[0];
        document.documentElement.setAttribute("data-module-top-gmv", String(safeNum((topGmv || {}).gmv)));
        document.documentElement.setAttribute("data-module-top-spend", String(safeNum((topSpend || {}).spend)));
        document.documentElement.setAttribute("data-module-top-cvr", String(safeNum((bestCvr || {}).cvr)));

        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          const cvrText = bestCvr ? (bestCvr.label + " CVR " + fmtPct(bestCvr.cvr)) : "No CVR sample";
          cvrNarrative.innerHTML = topGmv
            ? ("<p>" + scope + ": " + topGmv.label + " has the highest ad GMV " + fmtMoney(topGmv.gmv) + "; " + cvrText + ". Next: keep budget on modules with both GMV and CVR.</p>")
            : "<p>No module GMV data under current filters.</p>";
        }

        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          const ctrText = bestCtr ? (bestCtr.label + " CTR " + fmtPct(bestCtr.ctr)) : "No CTR sample";
          ctrNarrative.innerHTML = topSpend
            ? ("<p>" + scope + ": " + topSpend.label + " has the highest outer spend " + fmtMoney(topSpend.spend) + "; " + ctrText + ". Next: refresh creative first for high-spend low-CTR modules.</p>")
            : "<p>No module spend data under current filters.</p>";
        }

        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCpm);
            const priorAvg = avgValue(daily.priorCpm);
            const peakLabel = peakDayLabel(daily.currentCpm, daily.axisLabels) || "current period";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "no prior base";
            cpmNarrative.innerHTML = "<p>" + scope + " avg CPM " + fmtNum(currentAvg,2) + ", vs prior " + deltaText + "; peak cost day: " + peakLabel + ". Next: check audience overlap, placement and creative fatigue.</p>";
          }else{
            cpmNarrative.innerHTML = "<p>No module CPM daily data under current filters.</p>";
          }
        }

        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentImpressions, daily.axisLabels) || "current period";
            const deltaText = priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : "no prior base";
            impressionsNarrative.innerHTML = "<p>" + scope + " total impressions " + fmtNum(currentTotal) + ", vs prior " + deltaText + "; peak exposure day: " + peakLabel + ". Next: link that day back to GMV and click quality.</p>";
          }else{
            impressionsNarrative.innerHTML = "<p>No module impression daily data under current filters.</p>";
          }
        }
      }

      window.renderModuleSectionPatched = renderModuleSection;"""


_MODULE_FINAL_OVERRIDE = """
  <script>
    (function(){
      if(typeof state === "object" && typeof state.moduleViewFilter === "undefined"){
        state.moduleViewFilter = "";
      }

      window.moduleValidLabel = moduleValidLabel = function(value){
        const text = String(value || "").trim();
        if(!text) return false;
        if(typeof validLabel === "function"){
          try{
            return validLabel(text);
          }catch(err){}
        }
        return !["其他","未归类"].includes(text);
      };

      function moduleEscapeHtml(value){
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;");
      }

      window.moduleScopeText = moduleScopeText = function(){
        return state.moduleViewFilter ? `${state.moduleViewFilter}` : "\u5168\u90e8\u6a21\u5757";
      };

      function moduleOptionsSorted(){
        const currentRows = aggregateBy(
          filterRows(DATA.outer_module_daily, "current").filter((row)=>moduleValidLabel(row.module)),
          ["module"],
          ["ad_gmv", "outer_spend", "impressions"]
        );
        return currentRows
          .sort((a,b)=>
            safeNum(b.ad_gmv) - safeNum(a.ad_gmv) ||
            safeNum(b.outer_spend) - safeNum(a.outer_spend) ||
            safeNum(b.impressions) - safeNum(a.impressions) ||
            String(a.board || "").localeCompare(String(b.board || ""), "zh-CN")
          )
          .map((row)=>row.module)
          .filter((value, index, list)=>moduleValidLabel(value) && list.indexOf(value) === index);
      }

      window.ensureModuleFilterUi = ensureModuleFilterUi = function(){
        const card = document.getElementById("module-gmv-cvr-chart")?.closest(".card");
        if(!card) return null;
        const head = card.querySelector(".chart-head");
        if(!head) return null;
        head.classList.add("module-head-with-filter");
        let tools = head.querySelector(".module-filter-tools");
        if(!tools){
          tools = document.createElement("div");
          tools.className = "module-filter-tools";
          tools.innerHTML = '<span class="module-filter-label">\u6a21\u5757\u7b5b\u9009</span><select id="module-view-filter" class="module-local-filter"></select>';
          head.appendChild(tools);
        }
        const select = tools.querySelector("#module-view-filter");
        if(!select) return null;
        const options = moduleOptionsSorted();
        if(state.moduleViewFilter && !options.includes(state.moduleViewFilter)){
          state.moduleViewFilter = "";
        }
        select.innerHTML = `<option value="">\u5168\u90e8\u6a21\u5757</option>${options.map((value)=>`<option value="${moduleEscapeHtml(value)}">${moduleEscapeHtml(value)}</option>`).join("")}`;
        select.value = state.moduleViewFilter || "";
        if(!select.dataset.bound){
          select.dataset.bound = "1";
          select.addEventListener("change", (event)=>{
            state.moduleViewFilter = event.target.value || "";
            if(typeof window.renderAll === "function") window.renderAll(false);
          });
        }
        return select;
      };

      window.buildModuleRowsScoped = buildModuleRowsScoped = function(){
        const current = aggregateBy(
          filterRows(DATA.outer_module_daily, "current").filter((row)=>moduleValidLabel(row.module)),
          ["module"],
          ["outer_spend","impressions","clicks","ad_gmv","ad_orders"]
        );
        const prior = aggregateBy(
          filterRows(DATA.outer_module_daily, "prior").filter((row)=>moduleValidLabel(row.module)),
          ["module"],
          ["outer_spend","impressions","clicks","ad_gmv","ad_orders"]
        );
        const currentMap = indexBy(current, "module");
        const priorMap = indexBy(prior, "module");
        const keys = new Set([...currentMap.keys(), ...priorMap.keys()]);
        return Array.from(keys).map((key)=>{
          const row = currentMap.get(key) || {};
          const priorRow = priorMap.get(key) || {};
          const spend = safeNum(row.outer_spend);
          const impressions = safeNum(row.impressions);
          const clicks = safeNum(row.clicks);
          const orders = safeNum(row.ad_orders);
          const priorSpend = safeNum(priorRow.outer_spend);
          const priorImpressions = safeNum(priorRow.impressions);
          const priorClicks = safeNum(priorRow.clicks);
          const priorOrders = safeNum(priorRow.ad_orders);
          return {
            label: key,
            gmv: safeNum(row.ad_gmv),
            priorGmv: safeNum(priorRow.ad_gmv),
            cvr: ratio(orders, clicks),
            priorCvr: ratio(priorOrders, priorClicks),
            ctr: ratio(clicks, impressions),
            priorCtr: ratio(priorClicks, priorImpressions),
            roi: ratio(row.ad_gmv, spend),
            priorRoi: ratio(priorRow.ad_gmv, priorSpend),
            cpm: impressions ? spend / impressions * 1000 : 0,
            priorCpm: priorImpressions ? priorSpend / priorImpressions * 1000 : 0,
            spend,
            priorSpend,
            impressions,
            priorImpressions,
            clicks,
            priorClicks,
            orders,
            priorOrders
          };
        }).filter((row)=>moduleValidLabel(row.label) && (!state.moduleViewFilter || row.label === state.moduleViewFilter))
          .sort((a,b)=>safeNum(b.spend) - safeNum(a.spend));
      };

      window.buildModuleDailySeries = buildModuleDailySeries = function(moduleName){
        const days = Array.from({length:Number(DATA?.meta?.day_count || 0)}, (_, index)=>index + 1);
        const pickRows = (periodType)=>filterRows(DATA.outer_module_daily, periodType).filter((row)=>
          moduleValidLabel(row.module) && (!moduleName || row.module === moduleName)
        );
        const currentMap = indexBy(aggregateBy(pickRows("current"), ["day"], ["outer_spend", "impressions", "clicks", "ad_orders"]), "day");
        const priorMap = indexBy(aggregateBy(pickRows("prior"), ["day"], ["outer_spend", "impressions", "clicks", "ad_orders"]), "day");
        return {
          axisLabels: chartDayLabels(DATA.meta.current_start, days.length),
          priorAxisLabels: chartDayLabels(DATA.meta.prior_start, days.length),
          currentCpm: days.map((day)=>{
            const row = currentMap.get(day) || {};
            const spend = safeNum(row.outer_spend);
            const impressions = safeNum(row.impressions);
            return impressions ? spend / impressions * 1000 : 0;
          }),
          priorCpm: days.map((day)=>{
            const row = priorMap.get(day) || {};
            const spend = safeNum(row.outer_spend);
            const impressions = safeNum(row.impressions);
            return impressions ? spend / impressions * 1000 : 0;
          }),
          currentCtr: days.map((day)=>{
            const row = currentMap.get(day) || {};
            return ratio(row.clicks, row.impressions);
          }),
          priorCtr: days.map((day)=>{
            const row = priorMap.get(day) || {};
            return ratio(row.clicks, row.impressions);
          }),
          currentCvr: days.map((day)=>{
            const row = currentMap.get(day) || {};
            return ratio(row.ad_orders, row.clicks);
          }),
          priorCvr: days.map((day)=>{
            const row = priorMap.get(day) || {};
            return ratio(row.ad_orders, row.clicks);
          }),
          currentSpend: days.map((day)=>safeNum((currentMap.get(day) || {}).outer_spend)),
          priorSpend: days.map((day)=>safeNum((priorMap.get(day) || {}).outer_spend)),
          currentImpressions: days.map((day)=>safeNum((currentMap.get(day) || {}).impressions)),
          priorImpressions: days.map((day)=>safeNum((priorMap.get(day) || {}).impressions)),
          hasData: days.some((day)=>safeNum((currentMap.get(day) || {}).outer_spend) > 0 || safeNum((currentMap.get(day) || {}).impressions) > 0 || safeNum((priorMap.get(day) || {}).outer_spend) > 0 || safeNum((priorMap.get(day) || {}).impressions) > 0)
        };
      };

      function bindAnalysisDownloadsLocal(scope){
        if(!scope) return;
        scope.querySelectorAll("[data-plot-download]").forEach((btn)=>{
          btn.onclick = ()=>downloadCanvas(btn.dataset.plotDownload);
        });
      }

      function ensureSupplementalStyles(){
        if(document.getElementById("codex-extra-analysis-style")) return;
        const style = document.createElement("style");
        style.id = "codex-extra-analysis-style";
        style.textContent = `
          .analysis-extra-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;margin-top:18px}
          .analysis-block.analysis-span-two{grid-column:1 / -1}
          @media (max-width:1080px){
            .analysis-extra-grid{grid-template-columns:1fr}
            .analysis-block.analysis-span-two{grid-column:auto}
          }
        `;
        document.head.appendChild(style);
      }

      function ensureModuleCardLayout(){
        const moduleCard = document.getElementById("module-chart")?.closest(".card") || document.getElementById("module-gmv-cvr-chart")?.closest(".card");
        if(!moduleCard) return null;
        if(moduleCard.dataset.customModuleLayout !== "v3"){
          moduleCard.dataset.customModuleLayout = "v3";
          moduleCard.classList.add("wide-analysis-card");
          moduleCard.innerHTML = `
            <div class="analysis-hidden-legacy"><div id="module-chart"></div><div id="module-narrative"></div></div>
            <div class="chart-head"><div><div class="section-label">Module</div><h2>\u6a21\u5757\u5206\u6790</h2><div class="muted">\u5168\u90e8\u4f7f\u7528\u7ad9\u5916\u5e7f\u544a\u53e3\u5f84\uff0c\u805a\u7126\u5e7f\u544aGMV\u3001CTR\u3001CVR\u3001CPM\u3001\u66dd\u5149\u91cf\u4e0e\u82b1\u8d39\u65e5\u8d8b</div></div></div>
            <div class="analysis-dual-grid">
              <div class="analysis-stack">
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">\u5e7f\u544aGMV\u67f1\u72b6 + \u8f6c\u5316\u7387\u66f2\u7ebf</div><div class="analysis-block-note">\u6a21\u5757\u7ef4\u5ea6\u770b\u5f53\u524d\u6210\u4ea4\u89c4\u6a21\u4e0e\u70b9\u51fb\u540e\u8f6c\u5316\u6548\u7387</div></div><button class="analysis-download" data-plot-download="module-gmv-cvr-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                  <div id="module-gmv-cvr-chart" class="plot"></div>
                  <div id="module-gmv-cvr-narrative" class="analysis-mini-narrative"></div>
                </div>
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">CPM\u66f2\u7ebf</div><div class="analysis-block-note">\u770b\u5404\u6a21\u5757\u62ff\u91cf\u6210\u672c\u9ad8\u4f4e\uff0c\u8bc6\u522b\u9ad8\u6210\u672c\u62d6\u7d2f\u70b9</div></div><button class="analysis-download" data-plot-download="module-cpm-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                  <div id="module-cpm-chart" class="plot"></div>
                  <div id="module-cpm-narrative" class="analysis-mini-narrative"></div>
                </div>
              </div>
              <div class="analysis-stack">
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">\u7ad9\u5916\u82b1\u8d39\u67f1\u72b6 + CTR\u66f2\u7ebf</div><div class="analysis-block-note">\u6a21\u5757\u7ef4\u5ea6\u770b\u7ad9\u5916\u82b1\u8d39\u5f3a\u5ea6\u4e0e\u70b9\u51fb\u5438\u5f15\u529b</div></div><button class="analysis-download" data-plot-download="module-gmv-ctr-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                  <div id="module-gmv-ctr-chart" class="plot"></div>
                  <div id="module-gmv-ctr-narrative" class="analysis-mini-narrative"></div>
                </div>
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">\u66dd\u5149\u91cf\u66f2\u7ebf</div><div class="analysis-block-note">\u770b\u6a21\u5757\u62ff\u91cf\u89c4\u6a21\uff0c\u5224\u65ad\u9884\u7b97\u662f\u5426\u8fc7\u5ea6\u96c6\u4e2d</div></div><button class="analysis-download" data-plot-download="module-impressions-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                  <div id="module-impressions-chart" class="plot"></div>
                  <div id="module-impressions-narrative" class="analysis-mini-narrative"></div>
                </div>
              </div>
            </div>
            <div class="analysis-extra-grid">
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">CTR\u65e5\u8d8b</div><div class="analysis-block-note">\u770b\u6a21\u5757\u70b9\u51fb\u5438\u5f15\u529b\u5728\u5468\u671f\u5185\u662f\u5426\u6301\u7eed\u6ce2\u52a8</div></div><button class="analysis-download" data-plot-download="module-ctr-daily-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                <div id="module-ctr-daily-chart" class="plot"></div>
                <div id="module-ctr-daily-narrative" class="analysis-mini-narrative"></div>
              </div>
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">CVR\u65e5\u8d8b</div><div class="analysis-block-note">\u770b\u6a21\u5757\u70b9\u51fb\u540e\u6210\u4ea4\u6548\u7387\u662f\u5426\u7a33\u5b9a\u627f\u63a5</div></div><button class="analysis-download" data-plot-download="module-cvr-daily-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                <div id="module-cvr-daily-chart" class="plot"></div>
                <div id="module-cvr-daily-narrative" class="analysis-mini-narrative"></div>
              </div>
              <div class="analysis-block analysis-span-two">
                <div class="analysis-block-head"><div><div class="analysis-block-title">\u82b1\u8d39\u65e5\u8d8b</div><div class="analysis-block-note">\u770b\u7ad9\u5916\u82b1\u8d39\u8282\u594f\u4e0e\u4e0a\u6708\u540c\u671f\u5dee\u5f02\uff0c\u8bc6\u522b\u653e\u91cf\u6216\u6536\u7f29\u8282\u70b9</div></div><button class="analysis-download" data-plot-download="module-spend-daily-chart">\u4e0b\u8f7d\u56fe\u8868</button></div>
                <div id="module-spend-daily-chart" class="plot"></div>
                <div id="module-spend-daily-narrative" class="analysis-mini-narrative"></div>
              </div>
            </div>`;
          bindAnalysisDownloadsLocal(moduleCard);
        }
        return moduleCard;
      }

      function ensureShopInnerTrendCard(){
        const categoryCard = document.getElementById("category-chart")?.closest(".card") || document.getElementById("category-gmv-cvr-chart")?.closest(".card");
        if(!categoryCard || !categoryCard.parentNode) return null;
        let card = document.getElementById("shop-inner-trend-card");
        if(!card){
          card = document.createElement("div");
          card.className = "card wide-analysis-card";
          card.id = "shop-inner-trend-card";
          card.innerHTML = `
            <div class="chart-head"><div><div class="section-label">Shop Daily</div><h2>站内日趋</h2><div class="muted">基于店铺GMV口径，观察GMV、访客、CVR与客单价的日趋势及上月同期对比</div></div></div>
            <div class="analysis-dual-grid">
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">GMV日趋</div><div class="analysis-block-note">看站内成交规模日波动与上月同期差异</div></div><button class="analysis-download" data-plot-download="shop-inner-gmv-chart">下载图表</button></div>
                <div id="shop-inner-gmv-chart" class="plot"></div>
                <div id="shop-inner-gmv-narrative" class="analysis-mini-narrative"></div>
              </div>
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">访客日趋</div><div class="analysis-block-note">看站内承接流量是否同步放大</div></div><button class="analysis-download" data-plot-download="shop-inner-visitors-chart">下载图表</button></div>
                <div id="shop-inner-visitors-chart" class="plot"></div>
                <div id="shop-inner-visitors-narrative" class="analysis-mini-narrative"></div>
              </div>
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">CVR日趋</div><div class="analysis-block-note">看站内承接效率是否跑赢上月同期</div></div><button class="analysis-download" data-plot-download="shop-inner-cvr-chart">下载图表</button></div>
                <div id="shop-inner-cvr-chart" class="plot"></div>
                <div id="shop-inner-cvr-narrative" class="analysis-mini-narrative"></div>
              </div>
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">客单价日趋</div><div class="analysis-block-note">看成交结构变化是否来自高客单承接</div></div><button class="analysis-download" data-plot-download="shop-inner-aov-chart">下载图表</button></div>
                <div id="shop-inner-aov-chart" class="plot"></div>
                <div id="shop-inner-aov-narrative" class="analysis-mini-narrative"></div>
              </div>
            </div>`;
          categoryCard.parentNode.insertBefore(card, categoryCard);
          bindAnalysisDownloadsLocal(card);
        }
        return card;
      }

      function buildShopInnerDailySeries(){
        const days = Array.from({length:Number(DATA?.meta?.day_count || 0)}, (_, index)=>index + 1);
        const currentMap = indexBy(aggregateBy(filterRows(DATA.shop_daily, "current"), ["day"], ["gmv", "visitors", "orders"]), "day");
        const priorMap = indexBy(aggregateBy(filterRows(DATA.shop_daily, "prior"), ["day"], ["gmv", "visitors", "orders"]), "day");
        return {
          axisLabels: chartDayLabels(DATA.meta.current_start, days.length),
          priorAxisLabels: chartDayLabels(DATA.meta.prior_start, days.length),
          currentGmv: days.map((day)=>safeNum((currentMap.get(day) || {}).gmv)),
          priorGmv: days.map((day)=>safeNum((priorMap.get(day) || {}).gmv)),
          currentVisitors: days.map((day)=>safeNum((currentMap.get(day) || {}).visitors)),
          priorVisitors: days.map((day)=>safeNum((priorMap.get(day) || {}).visitors)),
          currentCvr: days.map((day)=>{
            const row = currentMap.get(day) || {};
            return ratio(row.orders, row.visitors);
          }),
          priorCvr: days.map((day)=>{
            const row = priorMap.get(day) || {};
            return ratio(row.orders, row.visitors);
          }),
          currentAov: days.map((day)=>{
            const row = currentMap.get(day) || {};
            return ratio(row.gmv, row.orders);
          }),
          priorAov: days.map((day)=>{
            const row = priorMap.get(day) || {};
            return ratio(row.gmv, row.orders);
          }),
          hasData: days.some((day)=>safeNum((currentMap.get(day) || {}).gmv) > 0 || safeNum((currentMap.get(day) || {}).visitors) > 0 || safeNum((priorMap.get(day) || {}).gmv) > 0 || safeNum((priorMap.get(day) || {}).visitors) > 0)
        };
      }

      window.drawModuleDailyLine = drawModuleDailyLine = function(id, axisLabels, priorAxisLabels, currentValues, priorValues, currentName, priorName, currentColor, priorColor, hoverLabel, isPercent, options){
        const el = document.getElementById(id);
        if(!el) return;
        const hasData = [...currentValues, ...priorValues].some((value)=>safeNum(value) > 0);
        if(!hasData){
          el.innerHTML = '<div class="empty-state">????????????</div>';
          return;
        }
        const hoverValue = options && options.hoverFormat ? options.hoverFormat : (isPercent ? "%{y:.1%}" : "%{y:,.0f}");
        const yTick = options && options.tickformat ? options.tickformat : (isPercent ? ".0%" : ",.0f");
        renderPlot(id,[
          {
            type:"scatter",
            mode:"lines+markers",
            name:currentName,
            x:axisLabels,
            y:currentValues,
            customdata:axisLabels,
            line:{color:currentColor,width:3},
            marker:{size:7,color:currentColor},
            hovertemplate:`<b>${currentName}</b><br>???%{customdata}<br>${hoverLabel}?${hoverValue}<extra></extra>`
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:priorName,
            x:axisLabels,
            y:priorValues,
            customdata:priorAxisLabels,
            line:{color:priorColor,width:3},
            marker:{size:7,color:priorColor},
            hovertemplate:`<b>${priorName}</b><br>???%{customdata}<br>${hoverLabel}?${hoverValue}<extra></extra>`
          }
        ],{
          height:360,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:76,r:28,t:76,b:100},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{title:"??",type:"category",categoryorder:"array",categoryarray:axisLabels,tickangle:-38,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
          yaxis:{title:hoverLabel,automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:yTick,tickfont:{color:"#64748b"}}
        },basePlotConfig());
      };

      window.avgValue = avgValue = function(list){
        return list.length ? list.reduce((sum, value)=>sum + safeNum(value), 0) / list.length : 0;
      };

      window.peakDayLabel = peakDayLabel = function(values, labels){
        if(!values.length) return "";
        let peakIndex = 0;
        values.forEach((value, index)=>{
          if(safeNum(value) > safeNum(values[peakIndex])) peakIndex = index;
        });
        return labels[peakIndex] || "";
      };

      window.renderModuleSection = renderModuleSection = function(){
        ensureModuleCardLayout();
        ensureModuleFilterUi();
        const rows = buildModuleRowsScoped();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"\u5e7f\u544aGMV",
          priorBarName:"\u4e0a\u6708\u540c\u671f\u5e7f\u544aGMV",
          lineName:"CVR",
          priorLineName:"????CVR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        drawBarLineCombo("module-gmv-ctr-chart", rows, {
          barField:"spend",
          priorBarField:"priorSpend",
          lineField:"ctr",
          priorLineField:"priorCtr",
          barName:"\u7ad9\u5916\u82b1\u8d39",
          priorBarName:"\u4e0a\u6708\u540c\u671f\u7ad9\u5916\u82b1\u8d39",
          lineName:"CTR",
          priorLineName:"????CTR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        const daily = buildModuleDailySeries(state.moduleViewFilter || "");
        drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "Current CPM", "Prior CPM", "#2E5AA7", "#88C5FF", "CPM", false);
        drawModuleDailyLine("module-ctr-daily-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCtr, daily.priorCtr, "\u5f53\u524dCTR", "\u4e0a\u6708\u540c\u671fCTR", "#22c55e", "#86efac", "CTR", true);
        drawModuleDailyLine("module-cvr-daily-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCvr, daily.priorCvr, "\u5f53\u524dCVR", "\u4e0a\u6708\u540c\u671fCVR", "#f59e0b", "#fcd34d", "CVR", true);
        drawModuleDailyLine("module-spend-daily-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentSpend, daily.priorSpend, "\u5f53\u524d\u82b1\u8d39", "\u4e0a\u6708\u540c\u671f\u82b1\u8d39", "#2E5AA7", "#B8DCFF", "\u7ad9\u5916\u82b1\u8d39", false);
        drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "Current Impr.", "Prior Impr.", "#FFA62B", "#F8E6A0", "Impressions", false);

        const scope = moduleScopeText();
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.clicks) > 0).sort((a,b)=>safeNum(b.cvr) - safeNum(a.cvr))[0];
        const topSpend = [...rows].sort((a,b)=>safeNum(b.spend) - safeNum(a.spend))[0];
        const bestCtr = [...rows].filter((row)=>safeNum(row.impressions) > 0).sort((a,b)=>safeNum(b.ctr) - safeNum(a.ctr))[0];

        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          const cvrText = bestCvr ? `${bestCvr.label} \u7684 CVR \u6700\u9ad8\uff0c\u4e3a ${fmtPct(bestCvr.cvr)}\u3002` : "\u5f53\u524d\u6682\u65e0\u53ef\u6bd4\u8f83\u7684 CVR \u6837\u672c\u3002";
          cvrNarrative.innerHTML = topGmv
            ? `<p>${scope}\u4e0b\uff0c${topGmv.label} \u7684\u5e7f\u544a GMV \u6700\u9ad8\uff0c\u4e3a ${fmtMoney(topGmv.gmv)}\uff1b${cvrText} \u4e0b\u4e00\u6b65\uff1a\u4f18\u5148\u4fdd\u7559\u9ad8 GMV \u4e14 CVR \u7a33\u5b9a\u7684\u6a21\u5757\u9884\u7b97\u3002</p>`
            : '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757\u5e7f\u544a GMV \u6570\u636e\u3002</p>';
        }

        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          const ctrText = bestCtr ? `${bestCtr.label} \u7684 CTR \u6700\u9ad8\uff0c\u4e3a ${fmtPct(bestCtr.ctr)}\u3002` : "\u5f53\u524d\u6682\u65e0\u53ef\u6bd4\u8f83\u7684 CTR \u6837\u672c\u3002";
          ctrNarrative.innerHTML = topSpend
            ? `<p>${scope}\u4e0b\uff0c${topSpend.label} \u7684\u7ad9\u5916\u82b1\u8d39\u6700\u9ad8\uff0c\u4e3a ${fmtMoney(topSpend.spend)}\uff1b${ctrText} \u4e0b\u4e00\u6b65\uff1a\u9ad8\u82b1\u8d39\u4f46 CTR \u504f\u5f31\u7684\u6a21\u5757\u4f18\u5148\u66f4\u6362\u7d20\u6750\u9996\u5c4f\u548c\u5356\u70b9\u94a9\u5b50\u3002</p>`
            : '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757\u82b1\u8d39\u6570\u636e\u3002</p>';
        }

        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCpm);
            const priorAvg = avgValue(daily.priorCpm);
            const peakLabel = peakDayLabel(daily.currentCpm, daily.axisLabels) || "\u5f53\u524d\u5468\u671f";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "\u6682\u65e0\u53ef\u6bd4\u57fa\u6570";
            cpmNarrative.innerHTML = `<p>${scope} \u7684 CPM \u65e5\u5747\u4e3a ${fmtNum(currentAvg,2)}\uff0c\u8f83\u4e0a\u6708\u540c\u671f${deltaText}\uff0c${peakLabel} \u6210\u672c\u538b\u529b\u6700\u9ad8\u3002\u4e0b\u4e00\u6b65\uff1a\u91cd\u70b9\u56de\u67e5\u9ad8 CPM \u65e5\u671f\u5bf9\u5e94\u7684\u4eba\u7fa4\u91cd\u53e0\u3001\u7248\u4f4d\u548c\u7d20\u6750\u75b2\u52b3\u3002</p>`;
          }else{
            cpmNarrative.innerHTML = '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757 CPM \u65e5\u8d8b\u52bf\u6570\u636e\u3002</p>';
          }
        }

        const ctrDailyNarrative = document.getElementById("module-ctr-daily-narrative");
        if(ctrDailyNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCtr);
            const priorAvg = avgValue(daily.priorCtr);
            const peakLabel = peakDayLabel(daily.currentCtr, daily.axisLabels) || "\u5f53\u524d\u5468\u671f";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "\u6682\u65e0\u53ef\u6bd4\u57fa\u6570";
            ctrDailyNarrative.innerHTML = `<p>${scope} \u7684 CTR \u65e5\u5747\u4e3a ${fmtPct(currentAvg)}\uff0c\u8f83\u4e0a\u6708\u540c\u671f${deltaText}\uff0c${peakLabel} \u70b9\u51fb\u5438\u5f15\u529b\u6700\u9ad8\u3002\u4e0b\u4e00\u6b65\uff1a\u56de\u770b\u9ad8 CTR \u65e5\u671f\u7d20\u6750\uff0c\u6c89\u6dc0\u53ef\u590d\u7528\u7684\u9996\u5c4f\u8868\u8fbe\u4e0e\u5356\u70b9\u7ec4\u5408\u3002</p>`;
          }else{
            ctrDailyNarrative.innerHTML = '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757 CTR \u65e5\u8d8b\u52bf\u6570\u636e\u3002</p>';
          }
        }

        const cvrDailyNarrative = document.getElementById("module-cvr-daily-narrative");
        if(cvrDailyNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCvr);
            const priorAvg = avgValue(daily.priorCvr);
            const peakLabel = peakDayLabel(daily.currentCvr, daily.axisLabels) || "\u5f53\u524d\u5468\u671f";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "\u6682\u65e0\u53ef\u6bd4\u57fa\u6570";
            cvrDailyNarrative.innerHTML = `<p>${scope} \u7684 CVR \u65e5\u5747\u4e3a ${fmtPct(currentAvg)}\uff0c\u8f83\u4e0a\u6708\u540c\u671f${deltaText}\uff0c${peakLabel} \u6210\u4ea4\u627f\u63a5\u6700\u4f73\u3002\u4e0b\u4e00\u6b65\uff1a\u7ed3\u5408\u9ad8 CVR \u65e5\u671f\u7684\u4eba\u7fa4\u4e0e\u7d20\u6750\uff0c\u653e\u5927\u66f4\u5f3a\u7684\u627f\u63a5\u7ec4\u5408\u3002</p>`;
          }else{
            cvrDailyNarrative.innerHTML = '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757 CVR \u65e5\u8d8b\u52bf\u6570\u636e\u3002</p>';
          }
        }

        const spendDailyNarrative = document.getElementById("module-spend-daily-narrative");
        if(spendDailyNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentSpend.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorSpend.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentSpend, daily.axisLabels) || "\u5f53\u524d\u5468\u671f";
            const deltaText = priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : "\u6682\u65e0\u53ef\u6bd4\u57fa\u6570";
            spendDailyNarrative.innerHTML = `<p>${scope} \u7684\u7ad9\u5916\u82b1\u8d39\u5408\u8ba1\u4e3a ${fmtMoney(currentTotal)}\uff0c\u8f83\u4e0a\u6708\u540c\u671f${deltaText}\uff0c${peakLabel} \u82b1\u8d39\u6700\u9ad8\u3002\u4e0b\u4e00\u6b65\uff1a\u628a\u9ad8\u82b1\u8d39\u65e5\u671f\u4e0e CTR\u3001CVR\u3001\u5e7f\u544a GMV \u8054\u52a8\u770b\uff0c\u907f\u514d\u65e0\u6548\u5806\u91cf\u3002</p>`;
          }else{
            spendDailyNarrative.innerHTML = '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757\u82b1\u8d39\u65e5\u8d8b\u52bf\u6570\u636e\u3002</p>';
          }
        }

        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentImpressions, daily.axisLabels) || "\u5f53\u524d\u5468\u671f";
            const deltaText = priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : "\u6682\u65e0\u53ef\u6bd4\u57fa\u6570";
            impressionsNarrative.innerHTML = `<p>${scope} \u7684\u66dd\u5149\u91cf\u5408\u8ba1\u4e3a ${fmtNum(currentTotal)}\uff0c\u8f83\u4e0a\u6708\u540c\u671f${deltaText}\uff0c${peakLabel} \u66dd\u5149\u6700\u9ad8\u3002\u4e0b\u4e00\u6b65\uff1a\u628a\u9ad8\u66dd\u5149\u65e5\u671f\u56de\u67e5\u5230\u70b9\u51fb\u8d28\u91cf\u4e0e\u5e7f\u544a GMV\uff0c\u8bc6\u522b\u662f\u5426\u6709\u65e0\u6548\u653e\u91cf\u3002</p>`;
          }else{
            impressionsNarrative.innerHTML = '<p>\u5f53\u524d\u7b5b\u9009\u4e0b\u6682\u65e0\u6a21\u5757\u66dd\u5149\u65e5\u8d8b\u52bf\u6570\u636e\u3002</p>';
          }
        }
      };

      function renderShopInnerTrendSection(){
        ensureShopInnerTrendCard();
        const daily = buildShopInnerDailySeries();
        drawModuleDailyLine("shop-inner-gmv-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentGmv, daily.priorGmv, "当前GMV", "上月同期GMV", "#2E5AA7", "#B8DCFF", "GMV", false);
        drawModuleDailyLine("shop-inner-visitors-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentVisitors, daily.priorVisitors, "当前访客", "上月同期访客", "#22c55e", "#86efac", "访客", false);
        drawModuleDailyLine("shop-inner-cvr-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCvr, daily.priorCvr, "当前CVR", "上月同期CVR", "#f59e0b", "#fcd34d", "CVR", true);
        drawModuleDailyLine("shop-inner-aov-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentAov, daily.priorAov, "当前客单价", "上月同期客单价", "#0f766e", "#99f6e4", "客单价", false);

        const totalCurrentGmv = daily.currentGmv.reduce((sum, value)=>sum + safeNum(value), 0);
        const totalPriorGmv = daily.priorGmv.reduce((sum, value)=>sum + safeNum(value), 0);
        const peakGmvDay = peakDayLabel(daily.currentGmv, daily.axisLabels) || "当前周期";
        const gmvNarrative = document.getElementById("shop-inner-gmv-narrative");
        if(gmvNarrative){
          gmvNarrative.innerHTML = daily.hasData
            ? `<p>当前筛选下站内GMV合计为 ${fmtMoney(totalCurrentGmv)}，较上月同期${totalPriorGmv ? fmtDelta(pctChange(totalCurrentGmv, totalPriorGmv)) : "暂无可比基数"}；${peakGmvDay} 为成交高峰。 下一步：回看高GMV日期对应的主力产品与价格机制，复制高峰承接打法。</p>`
            : '<p>当前筛选下暂无站内GMV日趋势数据。</p>';
        }

        const totalCurrentVisitors = daily.currentVisitors.reduce((sum, value)=>sum + safeNum(value), 0);
        const totalPriorVisitors = daily.priorVisitors.reduce((sum, value)=>sum + safeNum(value), 0);
        const peakVisitorsDay = peakDayLabel(daily.currentVisitors, daily.axisLabels) || "当前周期";
        const visitorsNarrative = document.getElementById("shop-inner-visitors-narrative");
        if(visitorsNarrative){
          visitorsNarrative.innerHTML = daily.hasData
            ? `<p>当前筛选下站内访客合计为 ${fmtNum(totalCurrentVisitors)}，较上月同期${totalPriorVisitors ? fmtDelta(pctChange(totalCurrentVisitors, totalPriorVisitors)) : "暂无可比基数"}；${peakVisitorsDay} 流量最高。 下一步：结合GMV与CVR，判断高流量日期是否真正转成有效成交。</p>`
            : '<p>当前筛选下暂无站内访客日趋势数据。</p>';
        }

        const avgCurrentCvr = avgValue(daily.currentCvr);
        const avgPriorCvr = avgValue(daily.priorCvr);
        const peakCvrDay = peakDayLabel(daily.currentCvr, daily.axisLabels) || "当前周期";
        const cvrNarrative = document.getElementById("shop-inner-cvr-narrative");
        if(cvrNarrative){
          cvrNarrative.innerHTML = daily.hasData
            ? `<p>当前筛选下站内CVR日均为 ${fmtPct(avgCurrentCvr)}，较上月同期${avgPriorCvr ? fmtDelta(pctChange(avgCurrentCvr, avgPriorCvr)) : "暂无可比基数"}；${peakCvrDay} 承接效率最高。 下一步：优先复盘高CVR日期的详情页、价格力和套组承接表现。</p>`
            : '<p>当前筛选下暂无站内CVR日趋势数据。</p>';
        }

        const avgCurrentAov = avgValue(daily.currentAov);
        const avgPriorAov = avgValue(daily.priorAov);
        const peakAovDay = peakDayLabel(daily.currentAov, daily.axisLabels) || "当前周期";
        const aovNarrative = document.getElementById("shop-inner-aov-narrative");
        if(aovNarrative){
          aovNarrative.innerHTML = daily.hasData
            ? `<p>当前筛选下站内客单价日均为 ${fmtMoney(avgCurrentAov)}，较上月同期${avgPriorAov ? fmtDelta(pctChange(avgCurrentAov, avgPriorAov)) : "暂无可比基数"}；${peakAovDay} 客单价最高。 下一步：重点看高客单日期是否由套组、加购结构或高价单品拉动。</p>`
            : '<p>当前筛选下暂无站内客单价日趋势数据。</p>';
        }
      }

      const originalRenderAll = window.renderAll;
      function syncCustomSections(attempt = 0){
        const moduleCard = ensureModuleCardLayout();
        const shopCard = ensureShopInnerTrendCard();
        if(moduleCard) document.documentElement.setAttribute("data-codex-module-layout", "ready");
        if(shopCard) document.documentElement.setAttribute("data-codex-shop-daily-layout", "ready");
        if(moduleCard){
          renderModuleSection();
        }
        if(shopCard){
          renderShopInnerTrendSection();
        }
        if((!moduleCard || !shopCard) && attempt < 20){
          setTimeout(()=>syncCustomSections(attempt + 1), 250);
        }
      }
      if(originalRenderAll && !window.__codexWrappedRenderAll){
        window.__codexWrappedRenderAll = true;
        window.renderAll = function(reset){
          ensureSupplementalStyles();
          const result = originalRenderAll.call(this, reset);
          setTimeout(()=>syncCustomSections(0), 60);
          return result;
        };
      }

      const rerender = ()=>{
        try{
          ensureSupplementalStyles();
          if(typeof window.renderAll === "function") window.renderAll(false);
          else setTimeout(()=>syncCustomSections(0), 60);
        }catch(err){
          console.error(err);
          document.documentElement.setAttribute("data-module-render-error", String(err && err.message ? err.message : err));
        }
      };
      if(document.readyState === "loading") document.addEventListener("DOMContentLoaded", ()=>setTimeout(rerender, 200));
      else setTimeout(rerender, 200);
    })();
  </script>
"""


_THAILAND_RUNTIME_OVERRIDE = """
  <script>
    (function(){
      function removeCompetitorSection(){
        const section = document.getElementById("competitor-product-analysis");
        if(section) section.remove();
      }

      function thaiTopPayload(){
        return DATA && DATA.thai_shop_top ? DATA.thai_shop_top : {shops:[], rows_by_shop:{}, rows_by_shop_and_day:{}};
      }

      function thaiTopShopNames(){
        const payload = thaiTopPayload();
        return Array.isArray(payload.shops) ? payload.shops.filter(Boolean).slice(0,2) : [];
      }

      function thaiRowsForShop(shopName){
        if(!shopName) return [];
        const payload = thaiTopPayload();
        const dayValue = String(state.shopTopDay || "");
        const rowsByDay = payload.rows_by_shop_and_day || {};
        const rowsByShop = payload.rows_by_shop || {};
        if(dayValue && rowsByDay[shopName] && Array.isArray(rowsByDay[shopName][dayValue])){
          return rowsByDay[shopName][dayValue];
        }
        return Array.isArray(rowsByShop[shopName]) ? rowsByShop[shopName] : [];
      }

      function thaiRenderShopTopFinal(){
        const select = document.getElementById("shop-top-day-filter");
        const muted = document.getElementById("shop-top-muted");
        const names = thaiTopShopNames();
        if(!names.length) return;
        const shopA = names[0];
        const shopB = names[1] || "";
        if(select){
          if(state.shopTopDay === undefined) state.shopTopDay = "";
          const monthNum = Number(String(DATA?.meta?.current_start || "").slice(5,7)) || 4;
          select.innerHTML = `<option value="">全部日期</option>${Array.from({length:Number(DATA?.meta?.day_count || 0)},(_,idx)=>`<option value="${idx+1}" ${String(state.shopTopDay)===String(idx+1)?"selected":""}>${monthNum}/${idx+1}</option>`).join("")}`;
          select.onchange = function(event){
            state.shopTopDay = event.target.value || "";
            if(typeof renderShopCardsFinal === "function") renderShopCardsFinal();
          };
        }
        if(muted){
          muted.textContent = shopB ? `查看 ${shopA} 和 ${shopB} 的 GMV Top 产品` : `查看 ${shopA} 的 GMV Top 产品`;
        }

        const blocks = [
          ["shop-top-a-title", "shop-top-a", shopA],
          ["shop-top-b-title", "shop-top-b", shopB],
        ];
        blocks.forEach(([titleId, holderId, shopName])=>{
          const title = document.getElementById(titleId);
          const holder = document.getElementById(holderId);
          if(!title || !holder) return;
          if(!shopName){
            title.textContent = "暂无第二店铺";
            holder.innerHTML = '<div class="shop-mini-empty">当前仅识别到一个主力店铺。</div>';
            return;
          }
          title.textContent = `${shopName} Top10产品`;
          const rows = thaiRowsForShop(shopName);
          if(!rows.length){
            holder.innerHTML = '<div class="shop-mini-empty">当前筛选下该店铺暂无有效产品数据。</div>';
            return;
          }
          holder.innerHTML = `<div class="shop-rank-list">${rows.map((row,idx)=>`<div class="shop-rank-item"><span class="shop-rank-no">${idx+1}</span><div class="shop-rank-name">${row.label}</div><div class="shop-rank-metric"><strong>${fmtMoney(row.gmv)}</strong><span>订单 ${fmtNum(row.orders)} | CVR ${fmtPct(row.cvr)}</span></div></div>`).join("")}</div>`;
        });
      }

      function thaiRenderShopTopNarrativeFinal(){
        const holder =
          document.getElementById("shop-top-narrative") ||
          (typeof ensureShopTopNarrativeMount === "function" ? ensureShopTopNarrativeMount() : null);
        if(!holder) return;
        holder.style.display = "";
        const names = thaiTopShopNames();
        const blocks = names.map((shopName)=>{
          const rows = thaiRowsForShop(shopName);
          if(!rows.length) return null;
          const top = rows[0];
          return {
            shopName,
            name: top.label || "",
            gmv: fmtMoney(top.gmv),
            metric: `订单 ${fmtNum(top.orders)} | CVR ${fmtPct(top.cvr)}`,
            count: rows.length,
          };
        }).filter(Boolean);
        if(!blocks.length){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无店铺 Top 品分析。</div>';
          return;
        }
        const paragraphs = blocks.map((item)=>`${item.name} 当前排在 ${item.shopName} 第一位，GMV 为 ${item.gmv}，${item.metric}。`);
        const todo = [
          "继续结合单日筛选追踪各店 Top1 和 Top3 的位次变化，识别是投放切换还是自然动销在拉动结构变化。",
        ];
        if(blocks[0]){
          todo.push(`优先盯 ${blocks[0].shopName} 头部品后的第二梯队承接品，避免流量继续过度堆在单一主力品。`);
        }
        if(blocks[1]){
          todo.push(`围绕 ${blocks[1].shopName} 的高转化头部品补素材和联投组合，把榜单优势放大到更高 GMV 体量。`);
        }
        holder.innerHTML = `<p>${paragraphs.join(" ")}</p><p><strong>下一步：</strong></p>${narrativeList(todo)}`;
      }

      function thaiFocusProducts(){
        return Array.isArray(DATA?.focus_new_products) ? DATA.focus_new_products : [];
      }

      function thaiBundleRows(){
        const rows = typeof buildBundleRows === "function" ? buildBundleRows() : [];
        const labelMap = new Map(
          thaiFocusProducts()
            .map((item)=>[String(item?.source_name || item?.name || "").trim(), String(item?.name || "").trim()])
            .filter(([sourceName, name])=>sourceName && name)
        );
        return rows.map((row)=>({
          ...row,
          label: labelMap.get(String(row?.label || "").trim()) || row?.label || "",
        }));
      }

      function thaiApplyFocusLabels(){
        const focusLabel = DATA?.meta?.focus_filter_label || "只看重点新品";
        const sectionTitle = DATA?.meta?.focus_section_title || "重点新品专项";
        const sectionNote = DATA?.meta?.focus_section_note || "聚焦 2-4 月上线新品，观察放量节奏、流量承接与转化表现";
        const itemLabel = DATA?.meta?.focus_item_label || "新品";
        const tableTitle = DATA?.meta?.focus_table_title || "重点新品诊断表";

        const focusButton = document.getElementById("focus-bundle");
        if(focusButton) focusButton.textContent = focusLabel;

        const bundleCard = document.getElementById("bundle-chart")?.closest(".card");
        if(bundleCard){
          const heading = bundleCard.querySelector("h2");
          if(heading) heading.textContent = sectionTitle;
          const muted = bundleCard.querySelector(".chart-head .muted");
          if(muted) muted.textContent = sectionNote;
          const label = bundleCard.querySelector(".section-label");
          if(label) label.textContent = "New Product Focus";
        }

        const tableHeadTitle = document.querySelector(".bundle-table-title");
        if(tableHeadTitle) tableHeadTitle.textContent = tableTitle;

        const tableHeadNote = document.querySelector(".bundle-table-note");
        if(tableHeadNote) tableHeadNote.textContent = "字段较多时支持左右滑动";

        const emptyState = document.querySelector(".bundle-empty");
        if(emptyState) emptyState.textContent = "当前筛选下暂无重点新品诊断数据。";

        const firstTh = document.querySelector(".bundle-table th");
        if(firstTh) firstTh.textContent = itemLabel;
      }

      function thaiRenderNewProductNarrative(){
        const holder = document.getElementById("bundle-narrative") || document.getElementById("bundle-box");
        if(!holder) return;
        const focusProducts = thaiFocusProducts();
        if(!focusProducts.length) return;

        const rowMap = new Map(thaiBundleRows().map((row)=>[String(row.label || ""), row]));
        const paragraphs = focusProducts.map((item)=>{
          const name = String(item?.name || "").trim();
          if(!name) return "";
          const launchText = String(item?.launch_text || "").trim();
          const note = String(item?.note || "").trim();
          const row = rowMap.get(name) || {};
          const gmv = safeNum(row.gmv);
          const visitors = safeNum(row.visitors);
          const orders = safeNum(row.orders);
          const outerSpend = safeNum(row.outerSpend);
          const impressions = safeNum(row.impressions);
          const cvr = safeNum(row.cvr);
          const launchPrefix = launchText ? `${name}（${launchText}）` : name;

          if(gmv <= 0 && orders <= 0 && outerSpend <= 0){
            return `${launchPrefix} 当前在 4 月 1 日至 4 月 20 日统计区间内暂无有效数据。${note || "先保留在重点新品池，后续补数后继续跟踪。"}`;
          }

          const summaryBits = [
            `店铺 GMV ${fmtMoney(gmv)}`,
            `访客 ${fmtNum(visitors)}`,
            `CVR ${fmtPct(cvr)}`,
          ];
          if(outerSpend > 0) summaryBits.push(`站外花费 ${fmtMoney(outerSpend)}`);
          if(impressions > 0) summaryBits.push(`站外曝光 ${fmtNum(impressions)}`);
          const direction = gmv > 0 && cvr >= 0.02
            ? "已经形成基础成交承接，可继续观察放量质量。"
            : "当前仍处在验证期，先结合素材、价格力和站内承接继续看样本积累。";
          return `${launchPrefix} 当前${summaryBits.join("，")}。${direction}`;
        }).filter(Boolean);

        const activeRows = focusProducts
          .map((item)=>rowMap.get(String(item?.name || "").trim()) || null)
          .filter((row)=>row && (safeNum(row.gmv) > 0 || safeNum(row.outerSpend) > 0 || safeNum(row.visitors) > 0));

        const topRow = [...activeRows].sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv))[0];
        const weakRow = [...activeRows].sort((a,b)=>(safeNum(a.outerSpend) + safeNum(a.gmv))-(safeNum(b.outerSpend) + safeNum(b.gmv)))[0];
        const todo = ["继续按新品维度追踪首月流量、转化与投放样本，避免把放量判断建立在单日波动上。"];
        if(topRow){
          todo.push(`优先复盘 ${topRow.label} 当前的高 GMV 承接链路，确认是素材命题、价格机制还是自然动销在拉动放量。`);
        }
        if(weakRow && String(weakRow.label || "") !== String(topRow?.label || "")){
          todo.push(`针对 ${weakRow.label} 先补齐素材与基础流量，再判断是否进入下一轮重点放量。`);
        }

        holder.innerHTML = `<p>${paragraphs.join(" ")}</p><p><strong>下一步：</strong></p>${narrativeList(todo)}`;
      }

      function thaiActionValidLabel(label){
        const text = String(label || "").trim();
        if(!text) return false;
        return !["未归类","其他","混合目录","小样","赠品"].includes(text);
      }

      function thaiRenderActionList(){
        const holder = document.getElementById("action-box");
        if(!holder) return;

        const keyActions = [
          "蓝色气垫必须继续保量，但当前总花费高于 3 月、店铺 GMV 仅基本持平，下一步不建议继续直接加预算，先稳住预算总盘，再优先调整收割侧效率和预算结构。",
          "5X面霜当前是在总花费下降的情况下基本守住销售，说明它仍是稳定保盘品；下一步优先维持当前预算结构，不急着追加外投。",
          "水油喷雾和红色磨皮精华当前都处在掉量状态，下一步应先做预算回补测试，再观察店铺销售是否被拉动；若销售反应不明显，再判断为素材问题并调整收割/种草占比。",
          "美白面膜虽然已经高于 3 月，但增长更多来自更高花费，说明它是可放大的增长点但不是无限加投款；新增预算应优先放在更能直接承接成交的收割侧。",
          "PDRN面霜继续按测试观察处理，当前更重要的是建立最小有效样本，而不是承担泰国 4 月保盘任务。"
        ];

        if(!keyActions.length){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无可输出的泰国 To Do 结论。</div>';
          return;
        }

        holder.innerHTML = narrativeList(keyActions.slice(0, 5));
      }

      function thaiRenderHighlightWeakness(){
        const highlightHolder = document.getElementById("highlight-box");
        const weaknessHolder = document.getElementById("weakness-box");
        if(!highlightHolder || !weaknessHolder) return;

        const categoryRows = typeof buildDimensionRows === "function"
          ? buildDimensionRows("category").filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const productRows = typeof buildDimensionRows === "function"
          ? buildDimensionRows("product").filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const diagnosisRows = typeof buildDiagnosisRows === "function"
          ? buildDiagnosisRows("product").filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const ttRows = typeof buildTTGapRows === "function"
          ? buildTTGapRows().filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const focusRows = thaiBundleRows().filter((row)=>thaiActionValidLabel(row?.label));

        const highlights = [];
        const weaknesses = [];

        const topCategory = categoryRows[0];
        if(topCategory){
          highlights.push(`${topCategory.label} 仍是泰国当前最稳定的规模核心，说明市场主成交需求依旧集中在这条主赛道；这不只是体量大，更意味着它承担了店铺整体 GMV 的确定性底盘。`);
        }

        const topProduct = productRows[0];
        if(topProduct){
          highlights.push(`${topProduct.label} 继续站稳头部，说明它已经不是单次爆发型单品，而是可以持续承接投放与自然流量的“稳定入口款”；泰国市场当前的放量更适合围绕这类成熟链接做扩张，而不是频繁切换承接品。`);
        }

        const topFocus = [...focusRows].sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv))[0];
        if(topFocus){
          highlights.push(`${topFocus.label} 在重点新品里率先跑出成交体量，说明它已经跨过“只有曝光没有成交”的冷启动阶段；对泰国市场来说，这类新品更适合进入结构化放量观察，而不是继续只按测试品管理。`);
        }

        const lowEffProduct = diagnosisRows
          .filter((row)=>safeNum(row.gmv) > 0 && (safeNum(row.outerSpend) + safeNum(row.innerSpend)) >= 1000)
          .sort((a,b)=>(safeNum(a.gmv) / Math.max(safeNum(a.outerSpend) + safeNum(a.innerSpend), 1)) - (safeNum(b.gmv) / Math.max(safeNum(b.outerSpend) + safeNum(b.innerSpend), 1)))[0];
        if(lowEffProduct){
          weaknesses.push(`${lowEffProduct.label} 当前的问题不只是“效率低”，更像是流量质量、卖点表达和站内承接没有形成闭环；如果继续按放量款去投，只会放大无效花费，而不会真正拉动 GMV。`);
        }

        const ttLagger = [...ttRows].sort((a,b)=>safeNum(a.gap)-safeNum(b.gap))[0];
        if(ttLagger && safeNum(ttLagger.gap) < 0){
          weaknesses.push(`${ttLagger.label} 在 TT 侧卖得动、但在 Shopee 侧承接不足，说明问题更可能出在货架转化而不是需求本身；这类 gap 本质上是渠道承接效率差，而不是产品没有市场。`);
        }

        const noDataFocus = (thaiFocusProducts() || []).find((item)=>{
          const row = focusRows.find((focusRow)=>String(focusRow?.label || "").trim() === String(item?.name || "").trim()) || {};
          return safeNum(row.gmv) <= 0 && safeNum(row.orders) <= 0 && safeNum(row.outerSpend) <= 0;
        });
        if(noDataFocus){
          weaknesses.push(`${noDataFocus.name} 目前还不能下“市场反馈弱”的结论，因为它连有效成交/投放样本都没建立起来；现阶段缺的不是复盘结论，而是最基本的验证样本。`);
        }

        highlightHolder.innerHTML = highlights.length
          ? narrativeList(highlights.slice(0, 3))
          : '<div class="empty-state">当前筛选下暂无可输出的亮点结论。</div>';
        weaknessHolder.innerHTML = weaknesses.length
          ? narrativeList(weaknesses.slice(0, 3))
          : '<div class="empty-state">当前筛选下暂无明显不足项。</div>';
      }

      function thaiStrategyProducts(){
        return {
          mustKeep: ["蓝色气垫"],
          core: ["蓝色气垫", "5X面霜", "水油喷雾", "红色磨皮精华", "美白面膜"],
          test: ["PDRN面霜"],
          harvest: ["图文", "视频", "合创转化"],
          awareness: ["合创曝光", "官号"],
          mainShop: "新旗舰店",
        };
      }

      function thaiSum(rows, field){
        return (rows || []).reduce((sum, row)=>sum + safeNum(row?.[field]), 0);
      }

      function thaiRowsInPeriod(rows, period){
        const start = DATA?.meta?.[`${period}_start`];
        const end = DATA?.meta?.[`${period}_end`];
        return (rows || []).filter((row)=>{
          const date = String(row?.date || "");
          return start <= date && date <= end;
        });
      }

      function thaiPctText(value){
        return value === null || value === undefined || Number.isNaN(Number(value))
          ? "暂无可比基数"
          : fmtDelta(value);
      }

      function thaiProductPeriodRows(productName, period, shopName = ""){
        return thaiRowsInPeriod(DATA?.shop_daily || [], period).filter((row)=>{
          const sameProduct = String(row?.product || "").trim() === String(productName || "").trim();
          const sameShop = !shopName || String(row?.shop || "").trim() === String(shopName || "").trim();
          return sameProduct && sameShop;
        });
      }

      function thaiProductSnapshot(productName, shopName = ""){
        const currentRows = thaiProductPeriodRows(productName, "current", shopName);
        const priorRows = thaiProductPeriodRows(productName, "prior", shopName);
        const currentGmv = thaiSum(currentRows, "gmv");
        const priorGmv = thaiSum(priorRows, "gmv");
        const currentVisitors = thaiSum(currentRows, "visitors");
        const priorVisitors = thaiSum(priorRows, "visitors");
        const currentOrders = thaiSum(currentRows, "orders");
        const priorOrders = thaiSum(priorRows, "orders");
        return {
          label: productName,
          currentGmv,
          priorGmv,
          gmvMom: pctChange(currentGmv, priorGmv),
          currentVisitors,
          priorVisitors,
          visitorsMom: pctChange(currentVisitors, priorVisitors),
          currentOrders,
          priorOrders,
          ordersMom: pctChange(currentOrders, priorOrders),
        };
      }

      function thaiModuleSnapshot(moduleName){
        const currentRows = thaiRowsInPeriod(DATA?.outer_module_daily || [], "current").filter((row)=>String(row?.module || "").trim() === String(moduleName || "").trim());
        const priorRows = thaiRowsInPeriod(DATA?.outer_module_daily || [], "prior").filter((row)=>String(row?.module || "").trim() === String(moduleName || "").trim());
        const currentSpend = thaiSum(currentRows, "outer_spend");
        const priorSpend = thaiSum(priorRows, "outer_spend");
        const currentGmv = thaiSum(currentRows, "ad_gmv");
        const currentImpressions = thaiSum(currentRows, "impressions");
        const currentClicks = thaiSum(currentRows, "clicks");
        return {
          label: moduleName,
          currentSpend,
          priorSpend,
          spendMom: pctChange(currentSpend, priorSpend),
          roi: ratio(currentGmv, currentSpend),
          ctr: ratio(currentClicks, currentImpressions),
          impressions: currentImpressions,
        };
      }

      function thaiProductRoiDiagnosis(row){
        const label = String(row?.label || "").trim();
        const totalSpend = safeNum(row?.outerSpend) + safeNum(row?.innerSpend);
        if(!label) return "";
        if(safeNum(row?.gmv) <= 0 && totalSpend <= 0){
          return "当前样本不足，暂不进入经营判断。";
        }

        const diagnosisMap = {
          "蓝色气垫": "核心保盘款，当前更需要稳住成交盘，而不是继续用加预算换规模。",
          "5X面霜": "成熟承接款，当前表现说明这条链接仍能承担主店稳盘角色。",
          "水油喷雾": "回补验证型品项，现阶段更需要确认预算恢复后能否把店铺成交重新拉起来。",
          "红色磨皮精华": "承压回落型品项，先验证销售弹性，再判断是否进入素材优化阶段。",
          "美白面膜": "增长放大型品项，说明具备继续放大的空间，但增长质量仍要受预算纪律约束。",
          "PDRN面霜": "测试观察型新品，目前不应用成熟放量品的标准直接判断优劣。",
        };
        if(diagnosisMap[label]) return diagnosisMap[label];

        if(safeNum(row?.fullRoi) >= 4 && safeNum(row?.adRoi) >= 1.5){
          return "当前属于可继续经营的效率款，但是否继续放量，仍要看店铺销售是否被真实带动。";
        }
        if(safeNum(row?.gmvMom) <= -0.1){
          return "当前更像承压回落型品项，优先验证预算调整后能否恢复成交承接。";
        }
        if(safeNum(row?.gmvMom) >= 0.05){
          return "当前具备放大潜力，但经营重点应放在增长质量，而不是单纯追求表面规模。";
        }
        if(safeNum(row?.harvestSpend) >= safeNum(row?.seedSpend) * 1.2){
          return "当前偏收割经营，后续关键在于验证成交承接是否稳定，而不是继续堆量。";
        }
        return "当前处在结构调整窗口，判断重点应回到预算是否真正带动主店销售。";
      }

      function thaiProductRoiTodo(row){
        const label = String(row?.label || "").trim();
        const totalSpend = safeNum(row?.outerSpend) + safeNum(row?.innerSpend);
        if(!label) return "";
        if(safeNum(row?.gmv) <= 0 && totalSpend <= 0){
          return "先补齐有效投放与成交样本，再进入经营动作判断。";
        }

        const todoMap = {
          "蓝色气垫": "继续保量，先稳住预算总盘，优先优化收割侧效率和结构，不建议再直接加预算。",
          "5X面霜": "维持当前预算结构，让它继续承担稳定保盘角色，暂不急于追加外投。",
          "水油喷雾": "先做预算回补测试，观察主店销售是否回升；若销售不跟，再转入素材优化。",
          "红色磨皮精华": "先补一轮收割预算验证销售弹性；若店铺销售反应仍弱，再调整收割与种草占比。",
          "美白面膜": "可以继续放量，但新增预算优先给收割侧，并控制节奏，避免无上限加投。",
          "PDRN面霜": "按测试观察处理，保留最小有效样本，不与成熟重点品争抢预算。",
        };
        if(todoMap[label]) return todoMap[label];

        if(safeNum(row?.fullRoi) >= 4 && safeNum(row?.adRoi) >= 1.5){
          return "维持当前有效结构，小步试探预算，并同步观察主店销售是否继续被带动。";
        }
        if(safeNum(row?.gmvMom) <= -0.1){
          return "先做预算回补测试，再看店铺销售反应，避免直接按需求走弱处理。";
        }
        if(safeNum(row?.gmvMom) >= 0.05){
          return "新增预算优先向更强承接的收割侧倾斜，同时控制放量节奏。";
        }
        return "先调整预算结构，再看主店销售是否出现同步变化，再决定是否继续放量。";
      }

      function thaiBoardSpendMap(dim, periodType){
        const rows = typeof aggregateBy === "function" && typeof filterRows === "function"
          ? aggregateBy(filterRows(DATA?.outer_module_daily || [], periodType), [dim, "board"], ["outer_spend", "ad_gmv"])
          : [];
        const map = new Map();
        rows.forEach((row)=>{
          const key = String(row?.[dim] || "").trim();
          if(!thaiActionValidLabel(key)) return;
          const board = String(row?.board || "").trim();
          const spend = safeNum(row?.outer_spend);
          const adGmv = safeNum(row?.ad_gmv);
          if(!map.has(key)) map.set(key, {seedSpend:0, harvestSpend:0, seedAdGmv:0, harvestAdGmv:0});
          const target = map.get(key);
          if(board.includes("种草")){
            target.seedSpend += spend;
            target.seedAdGmv += adGmv;
          }else if(board.includes("收割")){
            target.harvestSpend += spend;
            target.harvestAdGmv += adGmv;
          }
        });
        return map;
      }

      function thaiRoiDiagnosisText(row, levelName){
        const totalSpend = safeNum(row?.innerSpend) + safeNum(row?.outerSpend);
        if(safeNum(row?.gmv) <= 0 && totalSpend <= 0){
          return ["当前样本不足，暂不进入经营判断。", "先补齐有效投放与成交样本，再进入经营动作判断。"];
        }
        if(safeNum(row?.fullRoi) >= 4 && safeNum(row?.adRoi) >= 1.5){
          return [`${levelName}当前属于高效可经营区间，说明店铺成交与站外投放还在同向配合。`, "维持当前有效结构，小步试探预算，并继续观察主店销售是否同步被带动。"];
        }
        if(safeNum(row?.fullRoi) >= 4 && safeNum(row?.adRoi) < 1){
          return [`${levelName}全站ROI高，但站外广告ROI偏弱，当前更依赖自然承接。`, "先优化站外点击到成交链路，再做小步扩量，避免粗暴加预算。"];
        }
        if(safeNum(row?.gmvMom) <= -0.1){
          return [`${levelName}当前更像承压回落型对象，成交规模较上月同期明显走弱。`, "先看预算回补后销售是否恢复，再决定是继续投放还是转向素材与承接优化。"];
        }
        if(safeNum(row?.gmvMom) >= 0.05 && safeNum(row?.fullRoi) >= 2.5){
          return [`${levelName}当前具备增长潜力，但效率仍需要被预算纪律约束。`, "新增预算优先向更强承接的收割侧倾斜，同时控制放量节奏。"];
        }
        if(safeNum(row?.harvestSpend) >= safeNum(row?.seedSpend) * 1.2 && safeNum(row?.fullRoi) < 2.5){
          return [`${levelName}当前偏收割经营，但收割效率还不够强。`, "收缩低效收割花费，优先替换素材和价格钩子，避免继续硬吃存量人群。"];
        }
        return [`${levelName}当前处在结构调整窗口，暂未形成清晰放量结论。`, "先调整预算结构，再看主店销售是否出现同步变化，再决定是否继续放量。"];
      }

      const previousBuildRoiDiagnosisRowsThailand =
        window.buildRoiDiagnosisRows || (typeof buildRoiDiagnosisRows === "function" ? buildRoiDiagnosisRows : null);
      window.buildRoiDiagnosisRows = function(kind){
        const previousRows = typeof previousBuildRoiDiagnosisRowsThailand === "function"
          ? previousBuildRoiDiagnosisRowsThailand(kind)
          : [];
        const dim = kind === "category_roi" ? "category" : "product";
        const currentRows = typeof buildDimensionRows === "function"
          ? buildDimensionRows(dim).filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const priorOuterMap = typeof indexBy === "function" && typeof aggregateBy === "function" && typeof filterRows === "function"
          ? indexBy(aggregateBy(filterRows(DATA?.outer_daily || [], "prior"), [dim], ["outer_spend", "ad_gmv"]), dim)
          : new Map();
        const currentOuterMap = typeof indexBy === "function" && typeof aggregateBy === "function" && typeof filterRows === "function"
          ? indexBy(aggregateBy(filterRows(DATA?.outer_daily || [], "current"), [dim], ["outer_spend", "ad_gmv"]), dim)
          : new Map();
        const priorInnerMap = typeof indexBy === "function" && typeof aggregateBy === "function" && typeof filterRows === "function"
          ? indexBy(aggregateBy(filterRows(DATA?.inner_daily || [], "prior"), [dim], ["inner_spend"]), dim)
          : new Map();
        const currentBoardMap = thaiBoardSpendMap(dim, "current");
        const priorBoardMap = thaiBoardSpendMap(dim, "prior");
        const rebuiltRows = currentRows.map((row)=>{
          const key = String(row?.label || "").trim();
          const currentOuter = currentOuterMap.get(key) || {};
          const priorOuter = priorOuterMap.get(key) || {};
          const priorInner = priorInnerMap.get(key) || {};
          const currentBoard = currentBoardMap.get(key) || {seedSpend:0, harvestSpend:0};
          const priorBoard = priorBoardMap.get(key) || {seedSpend:0, harvestSpend:0};
          const fullRoi = ratio(row.gmv, safeNum(row.innerSpend) + safeNum(row.outerSpend));
          const priorFullRoi = ratio(safeNum(row.gmv) / (1 + safeNum(row.gmvMom)), safeNum(priorInner.inner_spend) + safeNum(priorOuter.outer_spend));
          const adRoi = ratio(safeNum(currentOuter.ad_gmv), safeNum(currentOuter.outer_spend));
          const priorAdRoi = ratio(safeNum(priorOuter.ad_gmv), safeNum(priorOuter.outer_spend));
          const roiNarrative = kind === "product_roi"
            ? [thaiProductRoiDiagnosis({
                ...row,
                fullRoi,
                adRoi,
                seedSpend: currentBoard.seedSpend,
                harvestSpend: currentBoard.harvestSpend,
              }), thaiProductRoiTodo({
                ...row,
                fullRoi,
                adRoi,
                seedSpend: currentBoard.seedSpend,
                harvestSpend: currentBoard.harvestSpend,
              })]
            : thaiRoiDiagnosisText({
                ...row,
                fullRoi,
                adRoi,
                seedSpend: currentBoard.seedSpend,
                harvestSpend: currentBoard.harvestSpend,
              }, `${key}品类`);
          return {
            label: key,
            gmv: safeNum(row.gmv),
            gmvMom: safeNum(row.gmvMom),
            fullRoi,
            fullRoiMom: pctChange(fullRoi, priorFullRoi),
            adRoi,
            adRoiMom: pctChange(adRoi, priorAdRoi),
            seedSpend: safeNum(currentBoard.seedSpend),
            seedSpendMom: pctChange(safeNum(currentBoard.seedSpend), safeNum(priorBoard.seedSpend)),
            harvestSpend: safeNum(currentBoard.harvestSpend),
            harvestSpendMom: pctChange(safeNum(currentBoard.harvestSpend), safeNum(priorBoard.harvestSpend)),
            innerSpend: safeNum(row.innerSpend),
            outerSpend: safeNum(row.outerSpend),
            diagnosis: roiNarrative[0],
            todo: roiNarrative[1],
          };
        }).filter((row)=>thaiActionValidLabel(row?.label));
        if(rebuiltRows.length) return rebuiltRows.sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
        return previousRows;
      };
      try{
        buildRoiDiagnosisRows = window.buildRoiDiagnosisRows;
      }catch(error){
        window.buildRoiDiagnosisRowsPatched = window.buildRoiDiagnosisRows;
      }

      function thaiRenderRoiTables(){
        if(typeof renderTable !== "function" || typeof window.buildRoiDiagnosisRows !== "function") return;
        renderTable("category_roi", window.buildRoiDiagnosisRows("category_roi"));
        renderTable("product_roi", window.buildRoiDiagnosisRows("product_roi"));
      }

      function thaiRenderManagementNarratives(){
        const overallHolder = document.getElementById("overall-narrative");
        const spendHolder = document.getElementById("spend-narrative");
        const moduleHolder = document.getElementById("module-narrative");
        const categoryHolder = document.getElementById("category-narrative");
        const productHolder = document.getElementById("product-narrative");
        const ttHolder = document.getElementById("tt-gap-narrative");
        if(!overallHolder || !spendHolder || !moduleHolder || !categoryHolder || !productHolder || !ttHolder) return;

        const strategy = thaiStrategyProducts();
        const metrics = typeof overallMetrics === "function" ? overallMetrics() : null;
        const categoryRows = typeof buildDimensionRows === "function"
          ? buildDimensionRows("category").filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const diagnosisRows = typeof buildDiagnosisRows === "function"
          ? buildDiagnosisRows("product").filter((row)=>thaiActionValidLabel(row?.label))
          : [];
        const ttRows = typeof buildTTGapRows === "function"
          ? buildTTGapRows().filter((row)=>thaiActionValidLabel(row?.label))
          : [];

        const mainCurrentRows = thaiRowsInPeriod(DATA?.shop_daily || [], "current").filter((row)=>String(row?.shop || "").trim() === strategy.mainShop);
        const mainPriorRows = thaiRowsInPeriod(DATA?.shop_daily || [], "prior").filter((row)=>String(row?.shop || "").trim() === strategy.mainShop);
        const mainCurrentGmv = thaiSum(mainCurrentRows, "gmv");
        const mainPriorGmv = thaiSum(mainPriorRows, "gmv");
        const mainGmvMom = pctChange(mainCurrentGmv, mainPriorGmv);

        const coreSnapshots = strategy.core.map((name)=>thaiProductSnapshot(name));
        const coreDown = coreSnapshots.filter((row)=>safeNum(row.gmvMom) < -0.03).sort((a,b)=>safeNum(a.gmvMom)-safeNum(b.gmvMom));
        const coreStable = coreSnapshots.filter((row)=>safeNum(row.gmvMom) >= -0.03);
        const mustKeep = thaiProductSnapshot(strategy.mustKeep[0]);
        const testProduct = thaiProductSnapshot(strategy.test[0]);
        const keySpendMap = new Map(
          [...strategy.core, ...strategy.test].map((name)=>{
            const currentOuterRows = thaiRowsInPeriod(DATA?.outer_daily || [], "current").filter((row)=>String(row?.product || "").trim() === name);
            const priorOuterRows = thaiRowsInPeriod(DATA?.outer_daily || [], "prior").filter((row)=>String(row?.product || "").trim() === name);
            const currentInnerRows = thaiRowsInPeriod(DATA?.inner_daily || [], "current").filter((row)=>String(row?.product || "").trim() === name);
            const priorInnerRows = thaiRowsInPeriod(DATA?.inner_daily || [], "prior").filter((row)=>String(row?.product || "").trim() === name);
            const currentOuter = thaiSum(currentOuterRows, "outer_spend");
            const priorOuter = thaiSum(priorOuterRows, "outer_spend");
            const currentInner = thaiSum(currentInnerRows, "inner_spend");
            const priorInner = thaiSum(priorInnerRows, "inner_spend");
            return [name, {
              currentOuter,
              priorOuter,
              currentInner,
              priorInner,
              currentTotal: currentOuter + currentInner,
              priorTotal: priorOuter + priorInner,
              outerShare: ratio(currentOuter, currentOuter + currentInner),
              outerRoi: ratio(thaiSum(currentOuterRows, "ad_gmv"), currentOuter),
            }];
          })
        );

        const harvestSnapshots = strategy.harvest.map((name)=>thaiModuleSnapshot(name));
        const awarenessSnapshots = strategy.awareness.map((name)=>thaiModuleSnapshot(name));
        const harvestSpend = harvestSnapshots.reduce((sum, row)=>sum + safeNum(row.currentSpend), 0);
        const awarenessSpend = awarenessSnapshots.reduce((sum, row)=>sum + safeNum(row.currentSpend), 0);
        const topHarvest = [...harvestSnapshots].sort((a,b)=>safeNum(b.currentSpend)-safeNum(a.currentSpend))[0];
        const weakestEff = diagnosisRows
          .filter((row)=>safeNum(row.gmv) > 0 && (safeNum(row.outerSpend) + safeNum(row.innerSpend)) >= 1000)
          .sort((a,b)=>(safeNum(a.gmv) / Math.max(safeNum(a.outerSpend) + safeNum(a.innerSpend), 1)) - (safeNum(b.gmv) / Math.max(safeNum(b.outerSpend) + safeNum(b.innerSpend), 1)))[0];
        const topCategory = categoryRows[0];
        const laggerTt = [...ttRows].sort((a,b)=>safeNum(a.gap)-safeNum(b.gap))[0];

        overallHolder.innerHTML = `<p>4 月泰国盘面的核心问题不是单一模块波动，而是整体还没有拉齐 3 月水平：当前店铺 GMV ${fmtMoney(metrics?.current?.gmv)}，较 3 月同期 ${thaiPctText(pctChange(metrics?.current?.gmv, metrics?.prior?.gmv))}；新旗舰店 GMV ${fmtMoney(mainCurrentGmv)}，较 3 月同期 ${thaiPctText(mainGmvMom)}，说明主店这条主盘仍然没有完全守住。</p><p>${coreDown.length ? `${coreDown.map((row)=>row.label).join("、")} 是本期重点品里更明显的回落来源，当前掉量并不是全盘均匀下滑，而是核心保盘品没有全部守住。` : "本期重点品整体相对稳住，掉量更多来自长尾和非核心链接收缩。"} ${coreStable.length ? `${coreStable.map((row)=>row.label).join("、")} 相对更接近守盘目标，说明泰国市场并不是需求整体塌陷，而是重点品恢复节奏不一致。` : ""}</p>`;

        spendHolder.innerHTML = `<p>从经营上看，本期更需要回答“花费有没有真实带动店铺销售”，而不是只看广告面板效率。当前总花费变化与店铺 GMV 变化没有形成同步改善，说明预算调整还没有完全转成店铺端的稳定增长。</p><p>${weakestEff ? `${weakestEff.label} 当前更像预算承压点：如果后续先调预算后，店铺销售仍没有明显波动，就应优先判断为素材问题，再回头调整收割与种草的预算占比。` : "当前没有特别突出的低效放量款，但预算判断仍应回到店铺销售是否被真实带动。"} </p>`;

        moduleHolder.innerHTML = `<p>泰国这期模块判断不能只看 ROI 高低，而要看预算结构是否服务于“先守住 3 月，再拉回增长”。当前收割模块花费 ${fmtMoney(harvestSpend)}，种草模块花费 ${fmtMoney(awarenessSpend)}，说明预算主体仍然放在收割侧，经营目标是对的，但还没有完全转成店铺端守盘结果。</p><p>${topHarvest ? `${topHarvest.label} 仍在承担主要收割职责；如果后续继续调预算后店铺销售没有明显回升，就不能简单理解成需求不行，更应优先回到素材效率和收割/种草占比上找原因。` : "当前模块投放样本有限，建议继续用预算调整来验证店铺端是否真正被带动。"} </p>`;

        categoryHolder.innerHTML = `<p>${topCategory ? `${topCategory.label} 仍是泰国最核心的成交底盘，这意味着市场主需求并没有转移，当前经营重点不是换赛道，而是先把这条主赛道的头部承接重新稳住。` : "当前品类样本不足，暂不建议做赛道层面的结构判断。"}</p><p>${topCategory && safeNum(topCategory.gmvMom) < 0 ? `${topCategory.label} 自身也低于 3 月同期，说明现在的问题不是“非核心品拖累”，而是核心盘本身还没完全恢复。` : "当前头部品类相对稳住，说明掉量更多可能来自次头部和重点品节奏分化。"} </p>`;

        const blueSpend = keySpendMap.get("蓝色气垫") || {};
        const creamSpend = keySpendMap.get("5X面霜") || {};
        const spraySpend = keySpendMap.get("水油喷雾") || {};
        const serumSpend = keySpendMap.get("红色磨皮精华") || {};
        const maskSpend = keySpendMap.get("美白面膜") || {};
        const pdrnSpend = keySpendMap.get("PDRN面霜") || {};
        const productTodo = [];

        productTodo.push(`蓝色气垫必须继续保量，但当前总花费较 3 月同期仍在增加、店铺 GMV 只基本持平，说明它暂时不适合继续单纯加预算。下一步更建议先稳住预算总盘，再优先调整收割侧的素材效率和预算结构，而不是继续放大种草投入。`);
        productTodo.push(`5X面霜目前是在总花费下降的情况下基本守住 GMV，说明这条链接的承接相对稳。下一步不建议大幅加外投，优先维持当前预算结构，让它继续承担稳定保盘角色。`);
        productTodo.push(`水油喷雾当前是“销售和花费一起回落”的状态，这类品更适合先补回一轮收割预算，看店铺销售能否被重新拉动；如果预算回补后销售仍无明显反应，再转向素材问题判断。`);
        productTodo.push(`红色磨皮精华本期掉量更深，现阶段不建议直接重压放量。更合理的是先做预算回补测试，观察主店销售是否跟随恢复；如果销售不跟，再调整收割与种草占比，而不是继续平均加投。`);
        productTodo.push(`美白面膜虽然已经高于 3 月，但当前增长更多是靠更高花费换来的，说明它是可放大的增长点，但不是无上限加投款。下一步建议控制新增预算节奏，把预算优先放在更能直接承接成交的收割侧。`);
        productTodo.push(`PDRN面霜仍按测试观察处理。它现在更需要最小有效样本，而不是承担保盘任务；预算上以可验证为主，不建议和成熟重点品抢资源。`);

        productHolder.innerHTML = `<p>从重点品视角看，这期最关键的不是谁短期冲高，而是谁在保盘。${strategy.mustKeep[0]} 较 3 月同期 ${thaiPctText(mustKeep.gmvMom)}，它必须继续作为优先保量款来稳住泰国整体盘面。</p><p>${coreDown.length ? `${coreDown[0].label}${coreDown[1] ? `、${coreDown[1].label}` : ""} 当前更像本期重点品里的主要拖累项，后面应先做预算调整，再观察店铺销售响应；如果销售不跟，就更偏素材问题，而不是直接判断需求走弱。` : "重点品暂未出现特别突出的拖累项，后续更要关注预算调整后店铺端是否能继续守平。"} ${safeNum(testProduct.currentGmv) > 0 ? `${strategy.test[0]} 属于测试观察，不应用成熟放量品的标准直接要求它承担保盘责任。` : `${strategy.test[0]} 当前仍未形成稳定承接样本，更适合继续作为测试观察品。`}</p><p><strong>下一步：</strong></p>${narrativeList(productTodo.slice(0, 6))}`;

        ttHolder.innerHTML = `<p>TT 与 Shopee 的 gap 在泰国不能直接理解成产品问题，很多时候更接近平台补贴、价格心智和渠道承接差异。当前更重要的是识别 gap 在扩大还是收窄，再决定 Shopee 侧要不要加量和调链接。</p><p>${laggerTt && safeNum(laggerTt.gap) < 0 ? `${laggerTt.label} 当前是最明显的 gap 产品，但这里更像“渠道承接差异信号”，而不是可以直接下结论说需求不行；更合理的动作是参考 TT 侧表现，在 Shopee 侧做预算和链接调整后再观察销售反应。` : "当前 TT 与 Shopee 的主要 gap 相对可控，后续仍应把它当成渠道差异信号来跟踪，而不是单独拿来否定产品表现。"} </p>`;
      }

      function hideGroupAnalysis(){
        if(!DATA?.meta?.disable_group_analysis) return;
        const groupFilter = document.getElementById("group-filter");
        const groupBox = groupFilter && groupFilter.closest(".filter-box");
        if(groupBox) groupBox.style.display = "none";
        ["group-chart", "group-compare-chart"].forEach((id)=>{
          const card = document.getElementById(id) && document.getElementById(id).closest(".card");
          if(card) card.style.display = "none";
        });
        document.querySelectorAll(".card h2").forEach((heading)=>{
          if(String(heading.textContent || "").trim() === "分组分析"){
            const card = heading.closest(".card");
            if(card) card.style.display = "none";
          }
        });
      }

      function applyThailandView(){
        removeCompetitorSection();
        hideGroupAnalysis();
        thaiRenderShopTopFinal();
        thaiRenderShopTopNarrativeFinal();
        thaiRenderRoiTables();
        thaiApplyFocusLabels();
        thaiRenderManagementNarratives();
        thaiRenderNewProductNarrative();
        thaiRenderHighlightWeakness();
        thaiRenderActionList();
      }

      const previousRenderAllThailand = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllThailand === "function") previousRenderAllThailand(reset);
        applyThailandView();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(applyThailandView, 0); });
      }else{
        setTimeout(applyThailandView, 0);
      }
    })();
  </script>
"""


def _resolve_sheet_name(excel_file: pd.ExcelFile, configured_name: str, fallback_key: str) -> str:
    candidates = [configured_name, *_SHEET_NAME_FALLBACKS.get(fallback_key, [])]
    for candidate in candidates:
        if candidate and candidate in excel_file.sheet_names:
            return candidate
    raise ValueError(f"Worksheet for {fallback_key} not found. Available sheets: {excel_file.sheet_names}")


def load_workbook_frames(config: dict[str, Any], workbook_override: str = "") -> dict[str, pd.DataFrame]:
    workbook_path = Path(workbook_override) if workbook_override else _BASE.ROOT / config["local_cache_workbook"]
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    sheets = config["worksheets"]
    excel_file = pd.ExcelFile(workbook_path)

    if _is_thai_workbook(excel_file):
        return _load_thai_workbook_frames(workbook_path)

    outer_sheet = _resolve_sheet_name(excel_file, sheets.get("outer_ads", ""), "outer_ads")
    shop_sheet = _resolve_sheet_name(excel_file, sheets.get("shop_gmv", ""), "shop_gmv")
    inner_sheet = _resolve_sheet_name(excel_file, sheets.get("inner_ads", ""), "inner_ads")
    tt_sheet = _resolve_sheet_name(excel_file, sheets.get("tt_sales", ""), "tt_sales")
    temp_sheet = _resolve_sheet_name(excel_file, sheets.get("april_temp", ""), "april_temp")

    outer = pd.read_excel(
        workbook_path,
        sheet_name=outer_sheet,
        usecols=[
            "Date_start",
            "Spend",
            "Impressions",
            "Link_Click",
            "Purchase Value",
            "Conversions",
            "产品",
            "站外产品名",
            "站内产品名",
            "店铺区分",
            "站内店铺命名",
            "品类",
            "分组情况",
            "模块",
            "汇率",
            "板块",
            "广告类型",
            "人群",
            "素材类型",
        ],
    )
    outer = outer.rename(
        columns={
            "站外产品名": "站外产品命名",
            "站内产品名": "站内名称",
        }
    )
    shop = pd.read_excel(
        workbook_path,
        sheet_name=shop_sheet,
        usecols=[
            "日期",
            "产品ID",
            "品类",
            "站内产品命名",
            "站外产品命名",
            "分组情况",
            "店铺",
            "GMV-人民币",
            "链接访客",
            "订单数",
        ],
    )
    inner = pd.read_excel(
        workbook_path,
        sheet_name=inner_sheet,
        usecols=[
            "date",
            "产品ID",
            "类目",
            "站内产品命名",
            "站外产品命名",
            "Spend-人民币",
        ],
    )
    tt = pd.read_excel(
        workbook_path,
        sheet_name=tt_sheet,
        usecols=[
            "时间",
            "sku",
            "中文产品名",
            "总销量",
            "渠道",
            "品类",
            "产品名",
            "分组",
        ],
    )
    # Keep TT columns in the exact positional order expected by prepare_tt_rows:
    # A 时间, D 总销量, E 渠道, F 品类, G 产品名, H 分组.
    tt = tt[["时间", "sku", "中文产品名", "总销量", "渠道", "品类", "产品名", "分组"]]
    temp = pd.read_excel(
        workbook_path,
        sheet_name=temp_sheet,
        usecols="A:U",
        nrows=45,
    )
    return {"outer": outer, "shop": shop, "inner": inner, "tt": tt, "temp": temp}


def infer_period(shop_df: pd.DataFrame, config: dict[str, Any] | None = None) -> dict[str, pd.Timestamp]:
    override = (config or {}).get("period_override") or {}
    if override:
        current_start = pd.Timestamp(override["current_start"]).normalize()
        current_end = pd.Timestamp(override["current_end"]).normalize()
        prior_start = pd.Timestamp(override["prior_start"]).normalize()
        prior_end = pd.Timestamp(override["prior_end"]).normalize()
        day_count = int((current_end - current_start).days) + 1
        return {
            "current_start": current_start,
            "current_end": current_end,
            "prior_start": prior_start,
            "prior_end": prior_end,
            "day_count": day_count,
        }
    return _BASE_INFER_PERIOD(shop_df)


def load_source_config(path: str | Path) -> dict[str, Any]:
    config = json.loads(json.dumps(_BASE.DEFAULT_SOURCE_CONFIG))
    config_path = Path(path)
    if not config_path.exists():
        return config

    loaded = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    if not isinstance(loaded, dict):
        return config

    for key, value in loaded.items():
        if key == "worksheets" and isinstance(value, dict):
            config.setdefault("worksheets", {}).update(value)
        elif key == "competitor_sources" and isinstance(value, dict):
            config["competitor_sources"] = value
        else:
            config[key] = value
    return config


_COMPETITOR_BOARD_KEYS = ("sales", "units", "views")
_COMPETITOR_STOPWORDS = {
    "official",
    "ori",
    "original",
    "best",
    "seller",
    "promo",
    "sale",
    "new",
    "bpom",
    "free",
    "gift",
    "bundle",
    "set",
    "pack",
    "pcs",
    "piece",
    "pieces",
    "gram",
    "ml",
    "g",
    "dan",
    "untuk",
    "dengan",
    "yang",
    "sku",
    "ready",
    "stok",
    "from",
    "the",
}
_COMPETITOR_BUCKET_KEYWORDS = {
    "brightening": ["bright", "brightening", "glow", "tone up", "toneup", "whitening", "377", "vit c", "niacinamide"],
    "acne care": ["acne", "jerawat", "pimple", "salicylic", "tea tree"],
    "soothing": ["soothing", "calming", "calm", "centella", "cica", "soothe"],
    "barrier repair": ["barrier", "repair", "ceramide", "5x", "skin barrier"],
    "moisturizing": ["moist", "moisturizing", "hydrating", "hydrate", "hyaluronic", "dew"],
    "anti-aging": ["retinol", "anti aging", "anti-aging", "aging", "firming", "wrinkle", "pdrn", "collagen"],
    "no white cast": ["no white cast", "white cast"],
    "lightweight": ["lightweight", "light weight", "ringan", "watery", "gel"],
    "oil control": ["oil control", "oil", "matte", "mattifying", "sebum"],
    "sunscreen claims": ["spf", "pa++", "pa+", "sunscreen", "sunblock", "uv"],
    "bundle value": ["bundle", "set", "duo", "2pc", "2 pcs", "3pc", "3 pcs", "value pack", "hemat"],
}
_COMPETITOR_COLUMN_ALIASES = {
    "rank": ["ranking", "rank", "排名", "名次"],
    "change_rank": ["changeofrank", "changerank", "rankchange", "change of rank", "变化", "排名变化"],
    "product_name": ["productname", "product", "title", "name", "商品名称", "产品名称"],
    "product_link": ["productlink", "link", "producturl", "url", "商品链接", "产品链接"],
    "product_id": ["productid", "itemid", "id", "产品id", "商品id"],
    "shop_name": ["shopname", "shop", "shoptitle", "brand", "brandname", "店铺", "店铺名称"],
    "shop_link": ["shoplink", "shopurl", "店铺链接"],
    "is_my_shop": ["isitmyshop", "ismyshop", "myshop", "ownbrand", "自家店", "是否自家店"],
}


def _competitor_norm_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _competitor_norm_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _competitor_slug(value: Any) -> str:
    text = _competitor_norm_text(value).lower()
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", text).strip()


def _competitor_is_true(value: Any) -> bool:
    text = _competitor_norm_text(value).lower()
    return text in {"1", "true", "yes", "y", "是", "owned", "my shop"}


def _competitor_first_non_empty(*values: Any) -> str:
    for value in values:
        text = _competitor_norm_text(value)
        if text:
            return text
    return ""


def _competitor_extract_definitions(frame: pd.DataFrame) -> dict[str, str]:
    clean = frame.dropna(how="all")
    if clean.empty or clean.shape[1] < 2:
        return {}
    result: dict[str, str] = {}
    for _, row in clean.iloc[:, :2].iterrows():
        key = _competitor_norm_text(row.iloc[0])
        value = _competitor_norm_text(row.iloc[1])
        if key and value:
            result[key] = value
    return result


def _competitor_infer_board_name(name: str) -> str:
    key = _competitor_slug(name)
    if "sales" in key:
        return "sales"
    if "units" in key or "unit" in key:
        return "units"
    if "views" in key or "view" in key:
        return "views"
    return ""


def _competitor_find_board_sheets(excel_file: pd.ExcelFile, config: dict[str, Any]) -> tuple[dict[str, str], str]:
    preferred = ((config.get("competitor_sources") or {}).get("preferred_sheet_names") or {})
    sheets = excel_file.sheet_names
    board_sheets: dict[str, str] = {}
    definitions_sheet = ""
    for board in _COMPETITOR_BOARD_KEYS:
        candidates = preferred.get(board) or [board.title()]
        for candidate in candidates:
            for sheet in sheets:
                if _competitor_norm_key(sheet) == _competitor_norm_key(candidate):
                    board_sheets[board] = sheet
                    break
            if board in board_sheets:
                break
        if board not in board_sheets:
            for sheet in sheets:
                if board in _competitor_slug(sheet):
                    board_sheets[board] = sheet
                    break
    for candidate in preferred.get("definitions") or ["Definitions"]:
        for sheet in sheets:
            if _competitor_norm_key(sheet) == _competitor_norm_key(candidate):
                definitions_sheet = sheet
                break
        if definitions_sheet:
            break
    if not definitions_sheet:
        for sheet in sheets:
            if "definition" in _competitor_slug(sheet):
                definitions_sheet = sheet
                break
    return board_sheets, definitions_sheet


def _competitor_match_columns(columns: list[str]) -> dict[str, str]:
    normalized = {_competitor_norm_key(column): column for column in columns}
    matched: dict[str, str] = {}
    for field, aliases in _COMPETITOR_COLUMN_ALIASES.items():
        for alias in aliases:
            alias_key = _competitor_norm_key(alias)
            if alias_key in normalized:
                matched[field] = normalized[alias_key]
                break
        if field in matched:
            continue
        for norm_column, raw_column in normalized.items():
            if any(_competitor_norm_key(alias) in norm_column for alias in aliases):
                matched[field] = raw_column
                break
    return matched


def _competitor_prepare_rows(frame: pd.DataFrame, board: str) -> list[dict[str, Any]]:
    clean = frame.dropna(how="all").copy()
    if clean.empty:
        return []
    clean.columns = [str(column) for column in clean.columns]
    matched = _competitor_match_columns(list(clean.columns))
    if "rank" not in matched or "product_name" not in matched:
        return []

    rows: list[dict[str, Any]] = []
    for _, row in clean.iterrows():
        product_name = _competitor_norm_text(row.get(matched["product_name"], ""))
        product_link = _competitor_norm_text(row.get(matched.get("product_link", ""), ""))
        if not product_name and not product_link:
            continue
        rank_value = pd.to_numeric(row.get(matched["rank"], None), errors="coerce")
        if pd.isna(rank_value):
            continue
        rows.append(
            {
                "module": board,
                "rank": int(rank_value),
                "change_rank": pd.to_numeric(row.get(matched.get("change_rank", ""), None), errors="coerce"),
                "product_name": product_name,
                "product_link": product_link,
                "product_id": _competitor_norm_text(row.get(matched.get("product_id", ""), "")),
                "shop_name": _competitor_norm_text(row.get(matched.get("shop_name", ""), "")),
                "shop_link": _competitor_norm_text(row.get(matched.get("shop_link", ""), "")),
                "is_my_shop": _competitor_is_true(row.get(matched.get("is_my_shop", ""), "")),
            }
        )
    return rows


def _competitor_find_sources(config: dict[str, Any], workbook_path: Path) -> list[Path]:
    cfg = config.get("competitor_sources") or {}
    if cfg.get("enabled", True) is False:
        return []
    search_dirs = cfg.get("search_dirs") or ["output", "."]
    keywords = [str(item).lower() for item in (cfg.get("filename_keywords") or ["competitor", "竞品", "rank", "ranking", "榜"])]
    ignore_dirs = {str(item).lower() for item in (cfg.get("ignore_dirs") or ["archive_indonesia"])}
    seen: set[Path] = set()
    scored: list[tuple[int, Path]] = []
    main_workbook = workbook_path.resolve()
    for raw_dir in search_dirs:
        root = (Path(raw_dir) if Path(raw_dir).is_absolute() else (_BASE.ROOT / raw_dir)).resolve()
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if not path.is_file():
                continue
            if any(part.lower() in ignore_dirs for part in path.parts):
                continue
            if path.suffix.lower() not in {".xlsx", ".xls", ".csv", ".json"}:
                continue
            resolved = path.resolve()
            if resolved == main_workbook or resolved in seen:
                continue
            seen.add(resolved)
            name_lower = path.name.lower()
            score = 0
            if any(keyword in name_lower for keyword in keywords):
                score += 4
            board_hint = _competitor_infer_board_name(path.stem)
            if board_hint:
                score += 2
            if path.suffix.lower() in {".xlsx", ".xls"}:
                try:
                    excel_file = pd.ExcelFile(path)
                    board_sheets, _ = _competitor_find_board_sheets(excel_file, config)
                    if board_sheets:
                        score += 6
                except Exception:
                    continue
            if score > 0:
                scored.append((score, path))
    scored.sort(key=lambda item: (-item[0], str(item[1]).lower()))
    return [path for _, path in scored]


def _competitor_load_sources(config: dict[str, Any], workbook_path: Path) -> dict[str, Any]:
    boards: dict[str, list[dict[str, Any]]] = {board: [] for board in _COMPETITOR_BOARD_KEYS}
    definitions: dict[str, str] = {}
    source_paths: list[Path] = []
    expected_region = _safe_text_local((config.get("competitor_sources") or {}).get("expected_region"))
    for path in _competitor_find_sources(config, workbook_path):
        try:
            if path.suffix.lower() in {".xlsx", ".xls"}:
                excel_file = pd.ExcelFile(path)
                board_sheets, definitions_sheet = _competitor_find_board_sheets(excel_file, config)
                if not board_sheets:
                    continue
                local_definitions: dict[str, str] = {}
                if definitions_sheet:
                    frame = pd.read_excel(path, sheet_name=definitions_sheet)
                    local_definitions.update(_competitor_extract_definitions(frame))
                if expected_region and _safe_text_local(local_definitions.get("Region")).upper() not in {"", expected_region.upper()}:
                    continue
                source_paths.append(path)
                for board, sheet_name in board_sheets.items():
                    frame = pd.read_excel(path, sheet_name=sheet_name)
                    boards[board].extend(_competitor_prepare_rows(frame, board))
                definitions.update(local_definitions)
            elif path.suffix.lower() == ".csv":
                board = _competitor_infer_board_name(path.stem)
                if not board:
                    continue
                frame = pd.read_csv(path)
                rows = _competitor_prepare_rows(frame, board)
                if rows:
                    source_paths.append(path)
                    boards[board].extend(rows)
            elif path.suffix.lower() == ".json":
                data = json.loads(path.read_text(encoding="utf-8", errors="ignore"))
                added = False
                if isinstance(data, dict):
                    for board in _COMPETITOR_BOARD_KEYS:
                        if isinstance(data.get(board), list):
                            frame = pd.DataFrame(data[board])
                            rows = _competitor_prepare_rows(frame, board)
                            if rows:
                                boards[board].extend(rows)
                                added = True
                    if isinstance(data.get("definitions"), dict):
                        definitions.update({str(k): _competitor_norm_text(v) for k, v in data["definitions"].items() if _competitor_norm_text(v)})
                elif isinstance(data, list):
                    board = _competitor_infer_board_name(path.stem)
                    if board:
                        frame = pd.DataFrame(data)
                        rows = _competitor_prepare_rows(frame, board)
                        if rows:
                            boards[board].extend(rows)
                            added = True
                if added:
                    source_paths.append(path)
        except Exception:
            continue
    return {"boards": boards, "definitions": definitions, "source_paths": source_paths}


def _competitor_product_key(row: dict[str, Any]) -> str:
    return _competitor_first_non_empty(row.get("product_link"), row.get("product_id"), _competitor_slug(row.get("product_name")))


def _competitor_is_skintific(product_name: str, shop_name: str, is_my_shop: bool) -> bool:
    product_text = product_name.lower()
    shop_text = shop_name.lower()
    if "skintific" in shop_text:
        return True
    if "skintific" in product_text:
        return True
    # Some exported ranking files carry polluted "Is it My Shop" flags from other brands.
    # Only trust the flag as a last fallback when the row lacks usable brand text.
    if is_my_shop and not shop_text and not product_text:
        return True
    return False


def _build_competitor_products(source_data: dict[str, Any]) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    for board, rows in source_data["boards"].items():
        for row in rows:
            key = _competitor_product_key(row)
            if not key:
                continue
            current = merged.setdefault(
                key,
                {
                    "product_name": "",
                    "product_link": "",
                    "product_id": "",
                    "shop_name": "",
                    "shop_link": "",
                    "is_skintific": False,
                    "ranks": {},
                    "board_list": [],
                },
            )
            current["product_name"] = _competitor_first_non_empty(current["product_name"], row["product_name"])
            current["product_link"] = _competitor_first_non_empty(current["product_link"], row["product_link"])
            current["product_id"] = _competitor_first_non_empty(current["product_id"], row["product_id"])
            current["shop_name"] = _competitor_first_non_empty(current["shop_name"], row["shop_name"])
            current["shop_link"] = _competitor_first_non_empty(current["shop_link"], row["shop_link"])
            current["is_skintific"] = current["is_skintific"] or _competitor_is_skintific(row["product_name"], row["shop_name"], bool(row["is_my_shop"]))
            current["ranks"][board] = min(current["ranks"].get(board, 10**9), int(row["rank"]))
            if board not in current["board_list"]:
                current["board_list"].append(board)

    products: list[dict[str, Any]] = []
    for current in merged.values():
        ranks = current["ranks"]
        rank_values = list(ranks.values())
        current["sales_rank"] = ranks.get("sales")
        current["units_rank"] = ranks.get("units")
        current["views_rank"] = ranks.get("views")
        current["board_count"] = len(current["board_list"])
        current["avg_rank"] = round(sum(rank_values) / len(rank_values), 1) if rank_values else None
        current["best_rank"] = min(rank_values) if rank_values else None
        current["rank_spread"] = (max(rank_values) - min(rank_values)) if len(rank_values) > 1 else 0
        products.append(current)
    products.sort(key=lambda item: (item["avg_rank"] if item["avg_rank"] is not None else 10**9, item["best_rank"] if item["best_rank"] is not None else 10**9, item["product_name"]))
    return products


def _competitor_tokenize_names(products: list[dict[str, Any]]) -> Counter[str]:
    counter: Counter[str] = Counter()
    for product in products:
        text = product["product_name"].lower()
        tokens = re.findall(r"[a-z0-9\+\-]{2,}", text)
        for token in tokens:
            if token in _COMPETITOR_STOPWORDS or token.isdigit():
                continue
            counter[token] += 1
    return counter


def _competitor_bucket_counts(products: list[dict[str, Any]]) -> Counter[str]:
    counter: Counter[str] = Counter()
    for product in products:
        title = product["product_name"].lower()
        for bucket, keywords in _COMPETITOR_BUCKET_KEYWORDS.items():
            if any(keyword in title for keyword in keywords):
                counter[bucket] += 1
    return counter


def _competitor_brand_summary(products: list[dict[str, Any]]) -> dict[str, Any]:
    brands = [product["shop_name"] or "Unknown Shop" for product in products]
    counter = Counter(brands)
    total = sum(counter.values())
    top_three = counter.most_common(3)
    top_three_share = (sum(count for _, count in top_three) / total) if total else 0
    if top_three_share >= 0.6:
        market_state = "集中"
    elif top_three_share <= 0.35:
        market_state = "分散"
    else:
        market_state = "中度集中"
    return {
        "counter": counter,
        "total": total,
        "top_three": top_three,
        "top_three_share": top_three_share,
        "market_state": market_state,
    }


def _competitor_format_rank(value: Any) -> str:
    return "-" if value in {None, "", 0} else str(value)


def _competitor_html_list(items: list[str]) -> str:
    if not items:
        return '<div class="empty-state">当前竞品榜单样本不足，暂不输出该小节结论。</div>'
    return "<ul>" + "".join(f"<li>{escape(item)}</li>" for item in items) + "</ul>"


def _competitor_html_table(rows: list[dict[str, str]], columns: list[tuple[str, str]]) -> str:
    if not rows:
        return '<div class="empty-state">当前没有可展示的竞品表格结果。</div>'
    header_html = "".join(f"<th>{escape(label)}</th>" for _, label in columns)
    body_html = "".join(
        "<tr>" + "".join(f"<td>{escape(str(row.get(key, '')))}</td>" for key, _ in columns) + "</tr>"
        for row in rows
    )
    return f'<div class="table-wrap"><table><thead><tr>{header_html}</tr></thead><tbody>{body_html}</tbody></table></div>'


def _render_competitor_section(analysis: dict[str, Any]) -> str:
    source_note = "；".join(escape(path) for path in analysis.get("source_names", []))
    definitions = analysis.get("definitions_text", "")
    exec_cards = "".join(
        f'<div class="competitor-mini-card"><div class="section-label">{escape(card["label"])}</div><div class="competitor-mini-value">{escape(card["value"])}</div><div class="muted">{escape(card["note"])}</div></div>'
        for card in analysis.get("summary_cards", [])
    )
    section_blocks = [
        ("Executive Summary", analysis.get("executive_summary", [])),
        ("Cross-board Winners", analysis.get("cross_board_winners", [])),
        ("Brand Concentration", analysis.get("brand_concentration", [])),
        ("Keyword Trends", analysis.get("keyword_trends", [])),
        ("Selling Point Distribution", analysis.get("selling_point_distribution", [])),
        ("Traffic-heavy Products", analysis.get("traffic_heavy_products", [])),
        ("Conversion-efficient Products", analysis.get("conversion_efficient_products", [])),
        ("Skintific Positioning", analysis.get("skintific_positioning", [])),
        ("Action Recommendations", analysis.get("recommendations", [])),
    ]
    blocks_html = "".join(
        f'<div class="competitor-analysis-block"><h3>{escape(title)}</h3>{_competitor_html_list(items)}</div>'
        for title, items in section_blocks
    )
    winner_table = _competitor_html_table(
        analysis.get("winner_rows", []),
        [("product", "Product"), ("shop", "Shop"), ("boards", "Boards"), ("avg_rank", "Avg Rank"), ("best_rank", "Best Rank")],
    )
    movement_table = _competitor_html_table(
        analysis.get("movement_rows", []),
        [("product", "Product"), ("views_rank", "Views Rank"), ("sales_rank", "Sales Rank"), ("signal", "Signal")],
    )
    warning_html = ""
    if analysis.get("warning"):
        warning_html = f'<div class="empty-state competitor-warning">{escape(analysis["warning"])}</div>'
    return f"""
<section id="competitor-product-analysis" class="table-card competitor-product-analysis">
  <div class="table-head">
    <div>
      <div class="section-label">Competitive Analysis</div>
      <h2>竞品产品排名分析</h2>
      <div class="muted">依据 {_COMPETITOR_SKILL_NAME} 规则生成；来源文件：{source_note or "未识别"}</div>
      {f'<div class="muted">{escape(definitions)}</div>' if definitions else ''}
    </div>
  </div>
  {warning_html}
  <div class="competitor-mini-grid">{exec_cards}</div>
  <div class="competitor-analysis-grid">{blocks_html}</div>
  <div class="competitor-analysis-table">
    <div class="chart-head"><div><div class="section-label">Cross-board Table</div><h3>跨榜产品表</h3></div></div>
    {winner_table}
  </div>
  <div class="competitor-analysis-table">
    <div class="chart-head"><div><div class="section-label">Structure Table</div><h3>流量/转化结构表</h3></div></div>
    {movement_table}
  </div>
</section>
"""


def build_competitor_analysis(config: dict[str, Any], workbook_override: str = "") -> dict[str, Any]:
    competitor_config = config.get("competitor_sources") or {}
    if competitor_config.get("enabled") is False:
        return {"detected": False, "html": ""}

    workbook_path = Path(workbook_override) if workbook_override else (_BASE.ROOT / config["local_cache_workbook"])
    source_data = _competitor_load_sources(config, workbook_path)
    source_paths = source_data["source_paths"]
    if not source_paths:
        return {"detected": False, "html": ""}

    products = _build_competitor_products(source_data)
    if not products:
        warning = "已检测到竞品榜单文件，但未成功解析出有效的 Sales / Units / Views 竞品排名数据。"
        return {
            "detected": True,
            "html": _render_competitor_section(
                {
                    "warning": warning,
                    "source_names": [path.name for path in source_paths],
                    "definitions_text": "",
                    "summary_cards": [
                        {"label": "Skill", "value": _COMPETITOR_SKILL_NAME, "note": "已触发"},
                        {"label": "Sources", "value": str(len(source_paths)), "note": "检测到竞品文件"},
                    ],
                    "executive_summary": [warning],
                    "cross_board_winners": [],
                    "brand_concentration": [],
                    "keyword_trends": [],
                    "selling_point_distribution": [],
                    "traffic_heavy_products": [],
                    "conversion_efficient_products": [],
                    "skintific_positioning": [],
                    "recommendations": ["先校验竞品文件字段名和榜单结构，再重新生成复盘 HTML。"],
                    "winner_rows": [],
                    "movement_rows": [],
                }
            ),
        }

    brand_summary = _competitor_brand_summary(products)
    keyword_counts = _competitor_tokenize_names(products)
    bucket_counts = _competitor_bucket_counts(products)
    cross_board_winners = [product for product in products if product["board_count"] == 3][:5]
    traffic_heavy = [
        product
        for product in products
        if product.get("views_rank") and (not product.get("sales_rank") or product["views_rank"] + 5 <= product["sales_rank"])
    ][:5]
    conversion_efficient = [
        product
        for product in products
        if product.get("sales_rank") and (not product.get("views_rank") or product["sales_rank"] + 5 <= product["views_rank"])
    ][:5]
    skintific_products = [product for product in products if product["is_skintific"]]
    skintific_cross_board = [product for product in skintific_products if product["board_count"] == 3]
    top_brand_name, top_brand_slots = brand_summary["top_three"][0] if brand_summary["top_three"] else ("-", 0)
    definitions_text = "；".join(f"{key}: {value}" for key, value in source_data["definitions"].items())

    analysis = {
        "warning": "",
        "source_names": [path.name for path in source_paths],
        "definitions_text": definitions_text,
        "summary_cards": [
            {"label": "Skill", "value": _COMPETITOR_SKILL_NAME, "note": "已调用"},
            {"label": "Sources", "value": str(len(source_paths)), "note": "竞品榜单文件"},
            {"label": "Merged Products", "value": str(len(products)), "note": "跨榜去重后商品数"},
            {"label": "Top 3 Share", "value": f'{brand_summary["top_three_share"]:.0%}', "note": f'市场{brand_summary["market_state"]}'},
        ],
        "executive_summary": [
            f"本次复盘检测到 {len(source_paths)} 份竞品榜单来源，并按 {_COMPETITOR_SKILL_NAME} 规则合并出 {len(products)} 个竞品商品。",
            f"当前榜单货架由 {top_brand_name} 领先，Top 3 品牌占比 {brand_summary['top_three_share']:.0%}，整体结构判断为{brand_summary['market_state']}。",
        ],
        "cross_board_winners": [
            f"{product['product_name']} 同时出现在 Sales / Units / Views，平均排名 {product['avg_rank']}，属于跨榜强势款。"
            for product in cross_board_winners
        ],
        "brand_concentration": [
            f"{brand} 上榜商品 {count} 个。"
            for brand, count in brand_summary["top_three"]
        ] + [f"Top 3 品牌合计占比 {brand_summary['top_three_share']:.0%}，说明市场{brand_summary['market_state']}。"],
        "keyword_trends": [
            f"{keyword} 在榜单标题中出现 {count} 次。"
            for keyword, count in keyword_counts.most_common(8)
        ],
        "selling_point_distribution": [
            f"{bucket} 相关标题命中 {count} 个。"
            for bucket, count in bucket_counts.most_common(8)
        ],
        "traffic_heavy_products": [
            f"{product['product_name']} 的 Views Rank {_competitor_format_rank(product['views_rank'])} 好于 Sales Rank {_competitor_format_rank(product['sales_rank'])}，更像强引流款。"
            for product in traffic_heavy
        ],
        "conversion_efficient_products": [
            f"{product['product_name']} 的 Sales Rank {_competitor_format_rank(product['sales_rank'])} 好于 Views Rank {_competitor_format_rank(product['views_rank'])}，更像强转化款。"
            for product in conversion_efficient
        ],
        "skintific_positioning": [
            f"Skintific 在竞品榜中识别到 {len(skintific_products)} 个商品，其中跨榜款 {len(skintific_cross_board)} 个。"
        ] + [
            f"{product['product_name']} 当前平均排名 {product['avg_rank']}，是 Skintific 的榜单主力。"
            for product in skintific_products[:3]
        ] + (
            ["当前榜单中未识别到 Skintific 商品，说明本期在竞品排名可见度上存在缺口。"] if not skintific_products else []
        ),
        "recommendations": [
            "优先围绕跨榜强势竞品梳理标题结构、利益点和套组机制，反推我方需要补强的货架表达。",
            "对流量强但销售弱的竞品，重点学习其吸点击标题与利益点；对销售强但浏览弱的竞品，重点拆其承接与转化卖点。",
            "若 Skintific 跨榜商品偏少，优先补双件套、五件套与重点功效组合的标题与套组竞争力。",
        ],
        "winner_rows": [
            {
                "product": product["product_name"],
                "shop": product["shop_name"] or "-",
                "boards": " / ".join(sorted(product["board_list"])),
                "avg_rank": str(product["avg_rank"]),
                "best_rank": str(product["best_rank"]),
            }
            for product in cross_board_winners
        ],
        "movement_rows": [
            {
                "product": product["product_name"],
                "views_rank": _competitor_format_rank(product.get("views_rank")),
                "sales_rank": _competitor_format_rank(product.get("sales_rank")),
                "signal": "Traffic-heavy",
            }
            for product in traffic_heavy
        ]
        + [
            {
                "product": product["product_name"],
                "views_rank": _competitor_format_rank(product.get("views_rank")),
                "sales_rank": _competitor_format_rank(product.get("sales_rank")),
                "signal": "Conversion-efficient",
            }
            for product in conversion_efficient
        ],
    }
    return {"detected": True, "html": _render_competitor_section(analysis), "source_paths": [str(path) for path in source_paths]}


def build_thai_shop_top_data(payload: dict[str, Any]) -> dict[str, Any]:
    meta = payload.get("meta") or {}
    current_start = _safe_text_local(meta.get("current_start"))
    current_end = _safe_text_local(meta.get("current_end"))
    if not current_start or not current_end:
        return {"shops": [], "rows_by_shop": {}, "rows_by_shop_and_day": {}}

    current_rows = [
        row for row in (payload.get("shop_daily") or [])
        if current_start <= _safe_text_local(row.get("date")) <= current_end
    ]

    shop_totals: dict[str, float] = {}
    for row in current_rows:
        shop = _safe_text_local(row.get("shop"))
        if not shop or shop == "未归类":
            continue
        shop_totals[shop] = shop_totals.get(shop, 0.0) + float(row.get("gmv") or 0.0)

    top_shops = [
        shop for shop, total in sorted(shop_totals.items(), key=lambda item: (-item[1], item[0]))
        if total > 0
    ][:2]

    def aggregate_rows(target_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        product_map: dict[str, dict[str, float]] = {}
        for row in target_rows:
            product = _safe_text_local(row.get("product"))
            if not product or product == "未归类":
                continue
            current = product_map.setdefault(product, {"gmv": 0.0, "visitors": 0.0, "orders": 0.0})
            current["gmv"] += float(row.get("gmv") or 0.0)
            current["visitors"] += float(row.get("visitors") or 0.0)
            current["orders"] += float(row.get("orders") or 0.0)

        ranked = sorted(product_map.items(), key=lambda item: (-item[1]["gmv"], item[0]))
        return [
            {
                "label": product,
                "gmv": round(values["gmv"], 2),
                "visitors": int(round(values["visitors"])),
                "orders": int(round(values["orders"])),
                "cvr": round(values["orders"] / values["visitors"], 6) if values["visitors"] else 0.0,
            }
            for product, values in ranked[:10]
            if values["gmv"] > 0
        ]

    rows_by_shop: dict[str, list[dict[str, Any]]] = {}
    rows_by_shop_and_day: dict[str, dict[str, list[dict[str, Any]]]] = {}
    for shop in top_shops:
        shop_rows = [row for row in current_rows if _safe_text_local(row.get("shop")) == shop]
        rows_by_shop[shop] = aggregate_rows(shop_rows)
        day_map: dict[str, list[dict[str, Any]]] = {}
        for day in sorted({
            str(int(row.get("day")))
            for row in shop_rows
            if str(row.get("day", "")).strip() not in {"", "0"}
        }, key=lambda value: int(value)):
            daily_rows = [row for row in shop_rows if str(int(row.get("day") or 0)) == day]
            day_map[day] = aggregate_rows(daily_rows)
        rows_by_shop_and_day[shop] = day_map

    return {
        "shops": top_shops,
        "rows_by_shop": rows_by_shop,
        "rows_by_shop_and_day": rows_by_shop_and_day,
    }


def _normalize_focus_new_products(config: dict[str, Any]) -> list[dict[str, str]]:
    def pick_text(value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return "" if text.lower() == "nan" else text

    raw_items = config.get("focus_new_products") or []
    normalized: list[dict[str, str]] = []
    for item in raw_items:
        source_name = ""
        if isinstance(item, str):
            name = pick_text(item)
            launch_date = ""
            launch_text = ""
            note = ""
        elif isinstance(item, dict):
            name = pick_text(item.get("name") or item.get("product"))
            source_name = pick_text(item.get("source_name")) or name
            launch_date = pick_text(item.get("launch_date"))
            launch_text = pick_text(item.get("launch_text"))
            note = pick_text(item.get("note"))
        else:
            continue
        if not name:
            continue
        normalized.append(
            {
                "name": name,
                "source_name": source_name or name,
                "launch_date": launch_date,
                "launch_text": launch_text,
                "note": note,
            }
        )
    return normalized


def _apply_thai_focus_products(payload: dict[str, Any], config: dict[str, Any]) -> None:
    focus_products = _normalize_focus_new_products(config)
    if not focus_products:
        return

    meta = payload.setdefault("meta", {})
    payload["focus_bundles"] = [item.get("source_name") or item["name"] for item in focus_products]
    payload["focus_new_products"] = focus_products
    meta["focus_section_title"] = "重点新品专项"
    meta["focus_filter_label"] = "只看重点新品"
    meta["focus_table_title"] = "重点新品诊断表"
    meta["focus_item_label"] = "新品"
    meta["focus_section_note"] = "聚焦 2-4 月上线新品，观察放量节奏、流量承接与转化表现"


def build_payload(config: dict[str, Any], workbook_override: str = "") -> dict[str, Any]:
    original_infer_period = _BASE.infer_period
    _BASE.infer_period = lambda shop_df: infer_period(shop_df, config)
    try:
        payload = _BASE_BUILD_PAYLOAD(config=config, workbook_override=workbook_override)
    finally:
        _BASE.infer_period = original_infer_period

    meta = payload.setdefault("meta", {})
    if config.get("report_title"):
        meta["title"] = str(config["report_title"])
    if config.get("report_subtitle"):
        meta["subtitle"] = str(config["report_subtitle"])
    if config.get("data_rule_note"):
        meta["data_rule_note"] = str(config["data_rule_note"])
    meta["disable_group_analysis"] = bool(config.get("disable_group_analysis"))
    _apply_thai_focus_products(payload, config)
    payload["thai_shop_top"] = build_thai_shop_top_data(payload)
    payload["competitor_analysis"] = build_competitor_analysis(config=config, workbook_override=workbook_override)
    return payload


def _inject_competitor_style(html: str) -> str:
    style_block = """
    .competitor-product-analysis{margin-top:18px}
    .competitor-mini-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-top:16px}
    .competitor-mini-card{padding:14px 16px;border-radius:18px;background:#f8fbff;border:1px solid rgba(47,107,255,.10)}
    .competitor-mini-value{margin:8px 0 4px;font-size:22px;font-weight:700;color:#183764}
    .competitor-analysis-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;margin-top:18px}
    .competitor-analysis-block{padding:18px;border-radius:18px;background:#fbfcff;border:1px solid rgba(47,107,255,.10)}
    .competitor-analysis-block h3{margin:0 0 10px;font-size:16px;color:#183764}
    .competitor-analysis-block ul{margin:0;padding-left:18px}
    .competitor-analysis-table{margin-top:18px}
    .competitor-warning{margin-top:16px}
    @media (max-width:1200px){.competitor-mini-grid,.competitor-analysis-grid{grid-template-columns:1fr}}
"""
    if "</head>" in html:
        return html.replace("</head>", f"<style>\n{style_block}</style>\n</head>", 1)
    if "</body>" in html:
        return html.replace("</body>", f"<style>\n{style_block}</style>\n</body>", 1)
    return html + f"\n<style>\n{style_block}</style>\n"


def _inject_competitor_section(html: str, section_html: str) -> str:
    page_script_marker = re.search(r"</div>\s*<script>\s*const DATA =", html)
    if page_script_marker:
        insert_at = page_script_marker.start()
        return html[:insert_at] + section_html + "\n" + html[insert_at:]
    body_index = html.rfind("</body>")
    if body_index != -1:
        return html[:body_index] + section_html + "\n" + html[body_index:]
    return html + "\n" + section_html


def _stabilize_module_rendering(html: str) -> str:
    module_render_start = html.find("function renderModuleSection(){")
    module_render_end = html.find("function renderTTSection(){", module_render_start)
    if module_render_start != -1 and module_render_end != -1:
        html = html[:module_render_start] + _MODULE_PROXY + html[module_render_end:]

    html = re.sub(
        r"function renderModuleSection\(\)\{\s*ensureModuleFilterUi\(\);[\s\S]*?window\.renderModuleSectionPatched = renderModuleSection;",
        _MODULE_RENDER_PATCH,
        html,
        count=1,
    )
    html = html.replace(
        "function drawBarLineCombo(id, rows, options){",
        "window.drawBarLineCombo = function drawBarLineCombo(id, rows, options){",
        1,
    )
    return html


def render_html(
    payload: dict[str, Any],
    plotly_src: str = f"./{_BASE.PLOTLY_JS_NAME}",
    plotly_inline_text: str | None = None,
) -> str:
    html = _BASE_RENDER_HTML(
        payload,
        plotly_src=plotly_src,
        plotly_inline_text=plotly_inline_text,
    )
    competitor_analysis = payload.get("competitor_analysis") or {}
    if competitor_analysis.get("detected") and competitor_analysis.get("html"):
        html = _inject_competitor_style(html)
        html = _inject_competitor_section(html, competitor_analysis["html"])
    html = _stabilize_module_rendering(html)
    if "</body>" in html:
        html = html.replace("</body>", _MODULE_FINAL_OVERRIDE + "\n" + _THAILAND_RUNTIME_OVERRIDE + "\n</body>", 1)
    return html


def write_chrome_launcher(script_path: Path, target_html: Path) -> Path:
    script_path.parent.mkdir(parents=True, exist_ok=True)
    script_path.write_text(_BASE.build_chrome_launcher(target_html), encoding="utf-8")
    return script_path


def write_desktop_chrome_launcher(target_html: Path) -> Path:
    desktop_path = Path.home() / "Desktop" / "4月Meta Shopee复盘(Chrome打开).cmd"
    return write_chrome_launcher(desktop_path, target_html)


def main() -> None:
    _BASE.load_source_config = load_source_config
    _BASE.build_payload = build_payload
    _BASE.render_html = render_html
    _BASE.load_workbook_frames = load_workbook_frames
    _BASE.write_chrome_launcher = write_chrome_launcher
    _BASE.write_desktop_chrome_launcher = write_desktop_chrome_launcher
    _BASE.main()


if __name__ == "__main__":
    main()
