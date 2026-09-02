const state={range:null,headers:[],mapping:null};
const $=id=>document.getElementById(id);

function setStatus(kind,title,text){$('statusCard').className='status-card '+kind;$('statusTitle').textContent=title;$('statusText').textContent=text}
function setBusy(busy,label){$('analyzeBtn').disabled=busy||!state.range;$('detectBtn').disabled=busy;$('analyzeBtn').querySelector('span').textContent=label||'Analyze and generate results'}
function renderWeights(){const host=$('weights');host.innerHTML='';LeadScoring.FIELD_DEFS.forEach(f=>{const row=document.createElement('div');row.className='weight-row';row.innerHTML=`<label for="weight-${f.key}">${f.label}<small>${f.hint}</small></label><div class="weight-control"><input id="weight-${f.key}" type="number" min="0" max="100" value="${f.weight}" aria-label="${f.label} weight"><span>%</span></div>`;host.appendChild(row)})}
function weights(){return Object.fromEntries(LeadScoring.FIELD_DEFS.map(f=>[f.key,Number($(`weight-${f.key}`).value)||0]))}

async function detectRange(){
  try{setBusy(true);setStatus('','Detecting data…','Reading the current worksheet region');await Excel.run(async context=>{const sheet=context.workbook.worksheets.getActiveWorksheet();const selected=context.workbook.getSelectedRange();const region=selected.getSurroundingRegion();region.load(['address','values','rowCount','columnCount']);await context.sync();if(region.rowCount<2)throw new Error('The selected region contains headers but no lead data');state.range=region.address;state.headers=region.values[0].map(String);state.mapping=LeadScoring.detect(state.headers);if(state.mapping.company<0)throw new Error('No company column was found. Add a header such as “Company” to the first row');const matched=LeadScoring.FIELD_DEFS.filter(f=>state.mapping[f.key]>=0).length;$('rangeSummary').classList.remove('hidden');$('rangeSummary').textContent=`Detected ${region.rowCount-1} leads · ${region.columnCount} columns · ${matched}/5 scoring fields`;setStatus('ready','Data is ready',region.address)});setBusy(false)}catch(e){state.range=null;setBusy(false);setStatus('error','Could not detect data',e.message||'Select any cell in the lead table and try again')}}

async function analyze(){
  if(!state.range)return;try{setBusy(true,'Analyzing…');setStatus('','Calculating priorities…','Keep the workbook open');await Excel.run(async context=>{
    const sheet=context.workbook.worksheets.getActiveWorksheet();const range=sheet.getRange(state.range);range.load(['values','rowCount','columnCount','rowIndex','columnIndex']);await context.sync();
    const values=range.values,headers=values[0].map(String),mapping=LeadScoring.detect(headers),w=weights();if(mapping.company<0)throw new Error('The company column was moved or deleted');
    const results=values.slice(1).map(row=>LeadScoring.scoreRow(row,mapping,w));const resultHeaders=['Priority Score','Lead Priority','Scoring Rationale'];
    let startCol=range.columnIndex+range.columnCount;const existingStart=headers.indexOf(resultHeaders[0]);if(existingStart>=0)startCol=range.columnIndex+existingStart;
    const output=sheet.getRangeByIndexes(range.rowIndex,startCol,values.length,3);output.values=[resultHeaders,...results.map(r=>[r.score,r.priority,r.reason])];output.format.autofitColumns();output.getRow(0).format.fill.color='#0D6B57';output.getRow(0).format.font.color='#FFFFFF';output.getRow(0).format.font.bold=true;
    const scoreData=output.getColumn(0).getOffsetRange(1,0).getResizedRange(-1,0);scoreData.conditionalFormats.clearAll();const scale=scoreData.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);scale.colorScale.criteria={minimum:{color:'#FEE4E2',type:'LowestValue'},midpoint:{color:'#FEF0C7',formula:'50',type:'Percentile'},maximum:{color:'#D1FADF',type:'HighestValue'}};
    await context.sync();
    if($('sortRows').checked){const full=sheet.getRangeByIndexes(range.rowIndex,range.columnIndex,values.length,(startCol-range.columnIndex)+3);full.getSort().apply([{key:startCol-range.columnIndex,ascending:false}],false,true)}
    if($('createDashboard').checked)await buildDashboard(context,results,w);
    await context.sync();
  });setStatus('ready','Analysis complete','Priorities, rationale, and dashboard have been updated');setBusy(false)}catch(e){setBusy(false);setStatus('error','Analysis could not be completed',e.message||'Check the data format and try again')}}

async function buildDashboard(context,results,w){
  const sheets=context.workbook.worksheets;let dash=sheets.getItemOrNullObject('Lead Priority Dashboard');dash.load('isNullObject');await context.sync();if(dash.isNullObject)dash=sheets.add('Lead Priority Dashboard');
  const used=dash.getUsedRangeOrNullObject();used.load('isNullObject');await context.sync();if(!used.isNullObject)used.clear(Excel.ClearApplyTo.all);dash.charts.deleteAll();dash.getRange('A1:F1').merge();dash.getRange('A1').values=[['Lead Priority Dashboard']];dash.getRange('A1').format.font={bold:true,size:20,color:'#173B35'};dash.getRange('A3:B7').values=[['Priority','Lead Count'],['A - Follow up now',results.filter(r=>r.score>=80).length],['B - Nurture actively',results.filter(r=>r.score>=60&&r.score<80).length],['C - Monitor',results.filter(r=>r.score>=40&&r.score<60).length],['D - Deprioritize',results.filter(r=>r.score<40).length]];
  const avg=results.length?Math.round(results.reduce((s,r)=>s+r.score,0)/results.length):0;dash.getRange('D3:E6').values=[['Key Metric','Value'],['Total leads',results.length],['Average score',avg],['A-priority leads',results.filter(r=>r.score>=80).length]];
  dash.getRange('A3:B3').format.fill.color='#0D6B57';dash.getRange('D3:E3').format.fill.color='#0D6B57';dash.getRanges('A3:B3,D3:E3').format.font.color='#FFFFFF';dash.getRanges('A3:B3,D3:E3').format.font.bold=true;
  const chart=dash.charts.add(Excel.ChartType.doughnut,dash.getRange('A3:B7'));chart.name='Priority Distribution';chart.title.text='Lead Priority Distribution';chart.legend.position=Excel.ChartLegendPosition.right;chart.setPosition('A9','F24');
  dash.getRange('H3:I8').values=[['Scoring Dimension','Weight'],...LeadScoring.FIELD_DEFS.map(f=>[f.label,Number(w[f.key])||0])];const weightChart=dash.charts.add(Excel.ChartType.barClustered,dash.getRange('H3:I8'));weightChart.name='Scoring Weights';weightChart.title.text='Current Scoring Weights';weightChart.legend.visible=false;weightChart.setPosition('H9','N24');
  dash.getUsedRange().format.autofitColumns();dash.activate();
}

Office.onReady(info=>{renderWeights();$('detectBtn').addEventListener('click',detectRange);$('analyzeBtn').addEventListener('click',analyze);$('resetBtn').addEventListener('click',renderWeights);if(info.host===Office.HostType.Excel){setStatus('ready','Connected to Excel','Select any cell in the lead data region');$('detectBtn').disabled=false}else setStatus('error','Open this add-in in Excel','This add-in supports Microsoft Excel only')});
