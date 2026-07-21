
    (function(){
      if(typeof state === "object" && state && typeof state.moduleViewFilter === "undefined"){
        state.moduleViewFilter = "";
      }

      function moduleValidLabel(value){
        const text = String(value || "").trim();
        if(!text) return false;
        if(typeof validLabel === "function"){
          try{
            return validLabel(text);
          }catch(err){}
        }
        return !["其他","未归类"].includes(text);
      }

      function moduleEscapeHtml(value){
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;");
      }

      function moduleScopeText(){
        return state.moduleViewFilter ? `${state.moduleViewFilter} 模块` : "全部模块";
      }

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
            String(a.module || "").localeCompare(String(b.module || ""), "zh-CN")
          )
          .map((row)=>row.module)
          .filter((value, index, list)=>moduleValidLabel(value) && list.indexOf(value) === index);
      }

      function ensureModuleFilterUi(){
        const card = document.getElementById("module-gmv-cvr-chart")?.closest(".card");
        if(!card) return null;
        const head = card.querySelector(".chart-head");
        if(!head) return null;
        head.classList.add("module-head-with-filter");
        let tools = head.querySelector(".module-filter-tools");
        if(!tools){
          tools = document.createElement("div");
          tools.className = "module-filter-tools";
          tools.innerHTML = '<span class="module-filter-label">模块筛选</span><select id="module-view-filter" class="module-local-filter"></select>';
          head.appendChild(tools);
        }
        const select = tools.querySelector("#module-view-filter");
        if(!select) return null;
        const options = moduleOptionsSorted();
        if(state.moduleViewFilter && !options.includes(state.moduleViewFilter)){
          state.moduleViewFilter = "";
        }
        select.innerHTML = `<option value="">全部模块</option>${options.map((value)=>`<option value="${moduleEscapeHtml(value)}">${moduleEscapeHtml(value)}</option>`).join("")}`;
        select.value = state.moduleViewFilter || "";
        if(!select.dataset.bound){
          select.dataset.bound = "1";
          select.addEventListener("change", (event)=>{
            state.moduleViewFilter = event.target.value || "";
            window.renderAll(false);
          });
        }
        const spendTitle = document.querySelector("#module-gmv-ctr-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-title");
        const spendNote = document.querySelector("#module-gmv-ctr-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-note");
        if(spendTitle) spendTitle.textContent = "站外花费柱状 + CTR曲线";
        if(spendNote) spendNote.textContent = "按站外模块口径看点击吸引力与花费强度，支持切换单模块查看日期走势";
        const cpmNote = document.querySelector("#module-cpm-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-note");
        const impNote = document.querySelector("#module-impressions-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-note");
        if(cpmNote) cpmNote.textContent = "X轴为日期，可按模块单选查看成本日趋势";
        if(impNote) impNote.textContent = "X轴为日期，可按模块单选查看曝光日趋势";
        return select;
      }

      function buildModuleRowsScoped(){
        const baseRows = typeof buildModuleOuterRowsPure === "function"
          ? buildModuleOuterRowsPure()
          : (typeof buildModuleRows === "function" ? buildModuleRows() : []);
        return (baseRows || []).filter((row)=>moduleValidLabel(row.label) && (!state.moduleViewFilter || row.label === state.moduleViewFilter));
      }

      function buildModuleDailySeries(moduleName){
        const days = Array.from({length:Number(DATA?.meta?.day_count || 0)}, (_, index)=>index + 1);
        const pickRows = (periodType)=>filterRows(DATA.outer_module_daily, periodType).filter((row)=>
          moduleValidLabel(row.module) && (!moduleName || row.module === moduleName)
        );
        const currentMap = indexBy(aggregateBy(pickRows("current"), ["day"], ["outer_spend", "impressions", "clicks"]), "day");
        const priorMap = indexBy(aggregateBy(pickRows("prior"), ["day"], ["outer_spend", "impressions", "clicks"]), "day");
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
          currentImpressions: days.map((day)=>safeNum((currentMap.get(day) || {}).impressions)),
          priorImpressions: days.map((day)=>safeNum((priorMap.get(day) || {}).impressions)),
          hasData: days.some((day)=>safeNum((currentMap.get(day) || {}).outer_spend) > 0 || safeNum((currentMap.get(day) || {}).impressions) > 0 || safeNum((priorMap.get(day) || {}).outer_spend) > 0 || safeNum((priorMap.get(day) || {}).impressions) > 0)
        };
      }

      function drawModuleDailyLine(id, axisLabels, priorAxisLabels, currentValues, priorValues, currentName, priorName, currentColor, priorColor, hoverLabel, isPercent){
        const el = document.getElementById(id);
        if(!el) return;
        const hasData = [...currentValues, ...priorValues].some((value)=>safeNum(value) > 0);
        if(!hasData){
          el.innerHTML = '<div class="empty-state">当前筛选下暂无模块日趋势数据。</div>';
          return;
        }
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
            hovertemplate:`<b>${currentName}</b><br>日期：%{customdata}<br>${hoverLabel}：${isPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
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
            hovertemplate:`<b>${priorName}</b><br>日期：%{customdata}<br>${hoverLabel}：${isPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          }
        ],{
          height:360,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:76,r:28,t:76,b:100},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:chartLegendLayout(),
          xaxis:{
            title:chartAxisTitle("日期"),
            type:"category",
            categoryorder:"array",
            categoryarray:axisLabels,
            tickangle:-38,
            automargin:true,
            showgrid:false,
            linecolor:"#dbe3ef",
            tickfont:{color:"#64748b"}
          },
          yaxis:{
            title:chartAxisTitle(hoverLabel),
            automargin:true,
            gridcolor:"#eef3fb",
            zeroline:false,
            tickformat:isPercent ? ".0%" : ",.0f",
            tickfont:{color:"#64748b"}
          }
        },basePlotConfig());
      }

      function avgValue(list){
        return list.length ? list.reduce((sum, value)=>sum + safeNum(value), 0) / list.length : 0;
      }

      function peakDayLabel(values, labels){
        if(!values.length) return "";
        let peakIndex = 0;
        values.forEach((value, index)=>{
          if(safeNum(value) > safeNum(values[peakIndex])) peakIndex = index;
        });
        return labels[peakIndex] || "";
      }

      function renderModuleSection(){
        try{
        ensureModuleFilterUi();
        const rows = buildModuleRowsScoped();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"广告GMV",
          priorBarName:"上月同期广告GMV",
          lineName:"CVR",
          priorLineName:"上月同期CVR",
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
          barName:"站外花费",
          priorBarName:"上月同期站外花费",
          lineName:"CTR",
          priorLineName:"上月同期CTR",
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

        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          const cvrText = bestCvr ? `${bestCvr.label} 的CVR最高，为 ${fmtPct(bestCvr.cvr)}。` : "当前暂无可比较的CVR样本。";
          cvrNarrative.innerHTML = topGmv
            ? `<p>${scope}下，${topGmv.label} 的广告GMV最高，为 ${fmtMoney(topGmv.gmv)}；${cvrText} 下一步：优先保留高GMV且CVR稳定的模块预算。</p>`
            : '<p>当前筛选下暂无模块广告GMV数据。</p>';
        }

        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          const ctrText = bestCtr ? `${bestCtr.label} 的CTR最高，为 ${fmtPct(bestCtr.ctr)}。` : "当前暂无可比较的CTR样本。";
          ctrNarrative.innerHTML = topSpend
            ? `<p>${scope}下，${topSpend.label} 的站外花费最高，为 ${fmtMoney(topSpend.spend)}；${ctrText} 下一步：高花费但CTR偏弱的模块优先换素材首屏和卖点钩子。</p>`
            : '<p>当前筛选下暂无模块花费数据。</p>';
        }

        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCpm);
            const priorAvg = avgValue(daily.priorCpm);
            const peakLabel = peakDayLabel(daily.currentCpm, daily.axisLabels) || "当前周期";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "暂无可比基数";
            cpmNarrative.innerHTML = `<p>${scope}的CPM日均为 ${fmtNum(currentAvg,2)}，较上月同期${deltaText}；${peakLabel} 成本压力最高。 下一步：重点回查高CPM日期对应的人群重叠、版位和素材疲劳。</p>`;
          }else{
            cpmNarrative.innerHTML = '<p>当前筛选下暂无模块CPM日趋势数据。</p>';
          }
        }

        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentImpressions, daily.axisLabels) || "当前周期";
            const deltaText = priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : "暂无可比基数";
            impressionsNarrative.innerHTML = `<p>${scope}当前累计曝光量为 ${fmtNum(currentTotal)}，较上月同期${deltaText}；${peakLabel} 是曝光峰值日。 下一步：把曝光峰值日与GMV、点击承接联查，确认高曝光是否带来有效成交。</p>`;
          }else{
            impressionsNarrative.innerHTML = '<p>当前筛选下暂无模块曝光日趋势数据。</p>';
          }
        }
        }catch(err){
          ["module-gmv-cvr-chart","module-gmv-ctr-chart","module-cpm-chart","module-impressions-chart"].forEach((id)=>{
            const el = document.getElementById(id);
            if(el) el.innerHTML = '<div class="empty-state">Module render error: ' + String(err && err.message ? err.message : err) + '</div>';
          });
          ["module-gmv-cvr-narrative","module-gmv-ctr-narrative","module-cpm-narrative","module-impressions-narrative"].forEach((id)=>{
            const el = document.getElementById(id);
            if(el) el.innerHTML = '<p>Module render error: ' + String(err && err.message ? err.message : err) + '</p>';
          });
        }
      }

      window.renderModuleSectionPatched = renderModuleSection;
      const resetButton = document.getElementById("reset-filter");
      if(resetButton && !resetButton.dataset.moduleResetBound){
        resetButton.dataset.moduleResetBound = "1";
        resetButton.addEventListener("click", ()=>{
          state.moduleViewFilter = "";
        }, true);
      }
    })();
  