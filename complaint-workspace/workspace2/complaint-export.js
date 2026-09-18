const COLORS=['FFEAF3FA','FFF0F6E8','FFFFF2DF','FFF1ECF8'];
const key=row=>JSON.stringify([row[0],row[1]]);
export function buildComplaintWorkbook(ExcelJS, values, review, retain) {
  const book=new ExcelJS.Workbook();
  book.creator='IPW Workspace 2';
  const headers=values[0];
  const rows=values.slice(1).map(row=>[...row]).sort((a,b)=>String(a[0]).localeCompare(String(b[0]),'en',{numeric:true}) || String(a[1]).localeCompare(String(b[1]),'en',{numeric:true}) || Number(a[2])-Number(b[2]) || Number(a[3])-Number(b[3]));
  function add(name,table,grouped=false,colorGroups=false) {
    const sheet=book.addWorksheet(name,{views:[{state:'frozen',ySplit:1,xSplit:grouped?6:0}]});
    sheet.addRows(table);
    sheet.columns=table[0].map(header=>({width:/Complaint reason|Problem|Evidence|Result|Source file|Missing/.test(header)?42:/Complaint|Source|source|MR level/.test(header)?20:/^(MR|FR|Zone|Type)$/.test(header)?8:14}));
    sheet.getRow(1).height=32;
    sheet.getRow(1).eachCell(cell=>{cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF123F61'}};cell.alignment={vertical:'middle',wrapText:true};});
    if(!grouped) sheet.autoFilter={from:{row:1,column:1},to:{row:table.length,column:table[0].length}};
    let group=-1,previous=null,start=2;
    function merge(end) {
      if(!grouped || end<=start) return;
      // Only share identical case metadata. Location and parameter cells stay separate.
      for(const col of [1,2,6]) {
        const first=sheet.getCell(start,col).value;
        if(Array.from({length:end-start+1},(_,i)=>sheet.getCell(start+i,col).value).every(v=>v===first)) sheet.mergeCells(start,col,end,col);
      }
    }
    for(let r=2;r<=table.length;r++) {
      const current=colorGroups?key(table[r-1]):String(r%2);
      if(current!==previous){merge(r-1);start=r;group++;previous=current;}
      const row=sheet.getRow(r); row.height=34;
      for(let c=1;c<=table[0].length;c++) {
        const cell=row.getCell(c);cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:COLORS[group%COLORS.length]}};
        cell.font={name:'Calibri',size:11,color:{argb:'FF203449'}};
        cell.alignment={vertical:'middle',wrapText:true};
        if(r===start)cell.border={top:{style:'thin',color:{argb:'FFAABBCD'}}};
      }
    }
    merge(table.length);
    sheet.properties.defaultRowHeight=34;
    return sheet;
  }
  add('Complaint_Grouped',[headers,...rows],true,true);
  add('Problem_Zone_Values',[headers,...rows],false,true);
  add('Complaint_Import_Review',review);
  add('Retain_Zone_Check',retain);
  return book;
}
