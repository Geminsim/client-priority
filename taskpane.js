const state={range:null,headers:[],mapping:null};
const $=id=>document.getElementById(id);

function setStatus(kind,title,text){$('statusCard').className='status-card '+kind;$('statusTitle').textContent=title;$('statusText').textContent=text}
function setBusy(busy,label){$('analyzeBtn').disabled=busy||!state.range;$('detectBtn').disabled=busy;$('analyzeBtn').querySelector('span').textContent=label||'分析并生成结果'}
function renderWeights(){const host=$('weights');host.innerHTML='';LeadScoring.FIELD_DEFS.forEach(f=>{const row=document.createElement('div');row.className='weight-row';row.innerHTML=`<label for="weight-${f.key}">${f.label}<small>${f.hint}</small></label><div class="weight-control"><input id="weight-${f.key}" type="number" min="0" max="100" value="${f.weight}" aria-label="${f.label}权重"><span>%</span></div>`;host.appendChild(row)})}
function weights(){return Object.fromEntries(LeadScoring.FIELD_DEFS.map(f=>[f.key,Number($(`weight-${f.key}`).value)||0]))}

async function detectRange(){
  try{setBusy(true);setStatus('','正在识别…','读取当前工作表的数据区域');await Excel.run(async context=>{const sheet=context.workbook.worksheets.getActiveWorksheet();const selected=context.workbook.getSelectedRange();const region=selected.getSurroundingRegion();region.load(['address','values','rowCount','columnCount']);await context.sync();if(region.rowCount<2)throw new Error('当前区域只有标题，没有客户数据');state.range=region.address;state.headers=region.values[0].map(String);state.mapping=LeadScoring.detect(state.headers);if(state.mapping.company<0)throw new Error('未找到“公司名称”列，请把首行设置为列名');const matched=LeadScoring.FIELD_DEFS.filter(f=>state.mapping[f.key]>=0).length;$('rangeSummary').classList.remove('hidden');$('rangeSummary').textContent=`已识别 ${region.rowCount-1} 位客户 · ${region.columnCount} 列 · ${matched}/5 个评分字段`;setStatus('ready','数据已就绪',region.address)});setBusy(false)}catch(e){state.range=null;setBusy(false);setStatus('error','无法识别数据',e.message||'请选择客户表中的任意单元格后重试')}}

async function analyze(){
  if(!state.range)return;try{setBusy(true,'正在分析…');setStatus('','正在计算优先级…','请勿关闭工作簿');await Excel.run(async context=>{
    const sheet=context.workbook.worksheets.getActiveWorksheet();const range=sheet.getRange(state.range);range.load(['values','rowCount','columnCount','rowIndex','columnIndex']);await context.sync();
    const values=range.values,headers=values[0].map(String),mapping=LeadScoring.detect(headers),w=weights();if(mapping.company<0)throw new Error('公司名称列已被移动或删除');
    const results=values.slice(1).map(row=>LeadScoring.scoreRow(row,mapping,w));const resultHeaders=['优先级得分','客户优先级','评分依据'];
    let startCol=range.columnIndex+range.columnCount;const existingStart=headers.indexOf(resultHeaders[0]);if(existingStart>=0)startCol=range.columnIndex+existingStart;
    const output=sheet.getRangeByIndexes(range.rowIndex,startCol,values.length,3);output.values=[resultHeaders,...results.map(r=>[r.score,r.priority,r.reason])];output.format.autofitColumns();output.getRow(0).format.fill.color='#0D6B57';output.getRow(0).format.font.color='#FFFFFF';output.getRow(0).format.font.bold=true;
    const scoreData=output.getColumn(0).getOffsetRange(1,0).getResizedRange(-1,0);scoreData.conditionalFormats.clearAll();const scale=scoreData.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);scale.colorScale.criteria={minimum:{color:'#FEE4E2',type:'LowestValue'},midpoint:{color:'#FEF0C7',formula:'50',type:'Percentile'},maximum:{color:'#D1FADF',type:'HighestValue'}};
    await context.sync();
    if($('sortRows').checked){const full=sheet.getRangeByIndexes(range.rowIndex,range.columnIndex,values.length,(startCol-range.columnIndex)+3);full.getSort().apply([{key:startCol-range.columnIndex,ascending:false}],false,true)}
    if($('createDashboard').checked)await buildDashboard(context,results,w);
    await context.sync();
  });setStatus('ready','分析完成','优先级、评分依据和仪表盘已更新');setBusy(false)}catch(e){setBusy(false);setStatus('error','分析未完成',e.message||'请检查数据格式后重试')}}

async function buildDashboard(context,results,w){
  const sheets=context.workbook.worksheets;let dash=sheets.getItemOrNullObject('优先级仪表盘');dash.load('isNullObject');await context.sync();if(dash.isNullObject)dash=sheets.add('优先级仪表盘');
  const used=dash.getUsedRangeOrNullObject();used.load('isNullObject');await context.sync();if(!used.isNullObject)used.clear(Excel.ClearApplyTo.all);dash.charts.deleteAll();dash.getRange('A1:F1').merge();dash.getRange('A1').values=[['潜在客户优先级仪表盘']];dash.getRange('A1').format.font={bold:true,size:20,color:'#173B35'};dash.getRange('A3:B7').values=[['优先级','客户数量'],['A - 立即跟进',results.filter(r=>r.score>=80).length],['B - 重点培育',results.filter(r=>r.score>=60&&r.score<80).length],['C - 持续观察',results.filter(r=>r.score>=40&&r.score<60).length],['D - 暂缓投入',results.filter(r=>r.score<40).length]];
  const avg=results.length?Math.round(results.reduce((s,r)=>s+r.score,0)/results.length):0;dash.getRange('D3:E6').values=[['关键指标','数值'],['客户总数',results.length],['平均得分',avg],['A级客户',results.filter(r=>r.score>=80).length]];
  dash.getRange('A3:B3').format.fill.color='#0D6B57';dash.getRange('D3:E3').format.fill.color='#0D6B57';dash.getRanges('A3:B3,D3:E3').format.font.color='#FFFFFF';dash.getRanges('A3:B3,D3:E3').format.font.bold=true;
  const chart=dash.charts.add(Excel.ChartType.doughnut,dash.getRange('A3:B7'));chart.name='优先级分布';chart.title.text='客户优先级分布';chart.legend.position=Excel.ChartLegendPosition.right;chart.setPosition('A9','F24');
  dash.getRange('H3:I8').values=[['评分维度','权重'],...LeadScoring.FIELD_DEFS.map(f=>[f.label,Number(w[f.key])||0])];const weightChart=dash.charts.add(Excel.ChartType.barClustered,dash.getRange('H3:I8'));weightChart.name='评分权重';weightChart.title.text='当前评分权重';weightChart.legend.visible=false;weightChart.setPosition('H9','N24');
  dash.getUsedRange().format.autofitColumns();dash.activate();
}

Office.onReady(info=>{renderWeights();$('detectBtn').addEventListener('click',detectRange);$('analyzeBtn').addEventListener('click',analyze);$('resetBtn').addEventListener('click',renderWeights);if(info.host===Office.HostType.Excel){setStatus('ready','已连接 Excel','请选择客户数据区域中的任意单元格');$('detectBtn').disabled=false}else setStatus('error','请在 Excel 中打开','此插件仅支持 Microsoft Excel')});
