let selected=null;
let query='';
export function renderLabelTableView(container,labels) {
  const parameters=[...new Set(labels.flatMap(item=>item.parameters))];
  if(selected===null) selected=new Set(parameters.filter(p=>['Thick','Wicking','NaCl'].includes(p)));
  container.classList.add('label-browser');
  container.replaceChildren();
  const toolbar=document.createElement('div');toolbar.className='label-view-toolbar';
  const search=document.createElement('input');search.type='search';search.placeholder='Find Lot, MR or complaint';search.setAttribute('aria-label','Search saved labels');search.value=query;
  const summary=document.createElement('span');
  toolbar.append(search,summary);
  const chooser=document.createElement('details');chooser.className='label-column-picker';
  const title=document.createElement('summary');title.textContent='Visible parameters';chooser.append(title);
  const hint=document.createElement('p');hint.textContent='Choose columns to display. Saved labels and Excel exports keep all parameters.';chooser.append(hint);
  const choices=document.createElement('div');choices.className='label-column-choices';
  const tableHost=document.createElement('div');tableHost.className='label-result-scroll';
  function draw() {
    tableHost.replaceChildren();
    const visible=parameters.filter(p=>selected.has(p));
    const filtered=labels.filter(item=>[item.lot,item.batch,item.label,item.notes].join(' ').toLowerCase().includes(query.toLowerCase()));
    summary.textContent=`${filtered.length} of ${labels.length} labeled regions · ${visible.length} parameters shown`;
    const table=document.createElement('table');const head=table.createTHead().insertRow();
    for(const name of ['Lot','MR','Zone','Complaint reason / Label',...visible,'Details']) {const th=document.createElement('th');th.textContent=name;th.scope='col';head.append(th);}
    const body=table.createTBody();
    for(const item of filtered) {
      const row=body.insertRow();
      for(const value of [item.lot,item.batch,`Zone ${item.zone}`,item.label,...visible.map(p=>item.parameters.includes(p)?(item.values?.[p]??'—'):'—')]) {const cell=row.insertCell();cell.textContent=String(value??'—');}
      const cell=row.insertCell();const details=document.createElement('details');const label=document.createElement('summary');label.textContent='View';details.append(label);
      for(const [name,value] of [['Comment',item.comment],['Source / Notes',item.notes],['Updated',item.updated]]) if(value) {const p=document.createElement('p');const strong=document.createElement('strong');strong.textContent=name;const text=document.createElement('div');text.textContent=value;p.append(strong,text);details.append(p);}
      cell.append(details);
    }
    tableHost.append(table);
    if(!filtered.length){const empty=document.createElement('p');empty.textContent='No matching labels.';tableHost.append(empty);}
  }
  for(const parameter of parameters) {
    const label=document.createElement('label');const input=document.createElement('input');input.type='checkbox';input.checked=selected.has(parameter);input.setAttribute('aria-label',`Show ${parameter}`);
    input.addEventListener('change',()=>{input.checked?selected.add(parameter):selected.delete(parameter);draw();});label.append(input,document.createTextNode(parameter));choices.append(label);
  }
  for(const [name,all] of [['Show all',true],['Hide all',false]]) {const button=document.createElement('button');button.type='button';button.textContent=name;button.addEventListener('click',()=>{selected=new Set(all?parameters:[]);choices.querySelectorAll('input').forEach(input=>{input.checked=all;});draw();});chooser.append(button);}
  chooser.append(choices);search.addEventListener('input',()=>{query=search.value;draw();});
  container.append(toolbar,chooser,tableHost);draw();
}
