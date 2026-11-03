importScripts("./vendor/xlsx.full.min.js");

self.addEventListener("message",event=>{
  try{
    const output=self.XLSX.write(event.data.workbook,event.data.options);
    const buffer=output instanceof ArrayBuffer?output:output.buffer.slice(output.byteOffset,output.byteOffset+output.byteLength);
    self.postMessage({buffer},[buffer]);
  }catch(error){
    self.postMessage({error:error?.message||String(error)});
  }
});
