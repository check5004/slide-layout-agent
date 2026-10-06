// Optional development QA only. Uses an isolated headless Chrome profile.
import { mkdirSync, writeFileSync, readFileSync } from 'node:fs';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../catalog/upstream/lib_cdp.mjs';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const inventory=JSON.parse(readFileSync(join(root,'catalog/inventory.json'),'utf8'));
const dest=join(root,'catalog/source-previews'); mkdirSync(dest,{recursive:true});
for(const file of inventory.source_html_files.filter(x=>x.classification==='layout_catalog')) {
  const page=await openPage(join(root,'catalog/upstream',file.path),{width:1600,height:1000,media:'print'});
  try {
    const boxes=await page.evaluate(()=>[...document.querySelectorAll('section')].map(s=>{
      const r=s.getBoundingClientRect();return {x:r.x+scrollX,y:r.y+scrollY,width:r.width,height:r.height,scale:1};
    }));
    const layouts=inventory.layouts.filter(x=>x.source_file===file.path);
    for(let i=0;i<boxes.length;i++){
      const shot=await page.send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true,clip:boxes[i]});
      writeFileSync(join(dest,`${layouts[i].layout_id}.png`),Buffer.from(shot.data,'base64'));
    }
  } finally { await page.close(); }
}
console.log('Rendered all 62 pinned source HTML sections');
