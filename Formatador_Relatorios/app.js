/* Formatador de Relatórios: .xlsx direto com JSZip + .xls convertido com SheetJS.
   Os dados são processados localmente; ao usar .xls, a biblioteca conversora
   será carregada da internet caso não exista uma cópia local. */
(() => {
  'use strict';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const $=id=>document.getElementById(id);
  const selected = [];
  const customTitles = new Map();
  let outputs = [];
  const norm = s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase();
  const children = (node, name) => Array.from(node?.children||[]).filter(e=>e.localName===name);
  const first = (node,name) => children(node,name)[0]||null;
  const node = (doc,name,attrs={}) => {const el=doc.createElementNS(NS,name);Object.entries(attrs).forEach(([k,v])=>el.setAttribute(k,String(v)));return el};
  const xml = str => {const d=new DOMParser().parseFromString(str,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('XML inválido no Excel');return d};
  const serialize=d=>new XMLSerializer().serializeToString(d);
  const previewSize = size => size<1024?size+' B':size<1048576?(size/1024).toFixed(1)+' KB':(size/1048576).toFixed(1)+' MB';
  const escapeHTML=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  function notify(message, error=false){$('status').textContent=message;$('status').className='status'+(error?' error':'')}
  function addFiles(files){
    const existing=new Set(selected.map(f=>f.name+'_'+f.size+'_'+f.lastModified));
    for(const file of Array.from(files||[])){
      if(!/\.(xlsx|xls)$/i.test(file.name)||file.name.startsWith('~$'))continue;
      const key=file.name+'_'+file.size+'_'+file.lastModified;
      if(!existing.has(key)){selected.push(file);existing.add(key);customTitles.set(key,'')}
    }
    outputs=[];$('results').hidden=true;renderQueue();notify('');
  }
  const fileKey=f=>f.name+'_'+f.size+'_'+f.lastModified;
  function renderQueue(){
    $('selectedCount').textContent=selected.length;
    $('queueDescription').textContent=selected.length?`${selected.length} arquivo(s) aguardando formatação`:'Nenhum arquivo selecionado';
    $('runBtn').disabled=!selected.length;
    $('fileList').innerHTML=selected.length?selected.map((f,i)=>`<div class="file-item"><div class="file-icon">▤</div><div class="file-info"><strong>${escapeHTML(f.name)}</strong><small>${previewSize(f.size)} · Título definido por você</small><label class="title-field"><span>Título do relatório</span><input class="title-input" data-title="${i}" type="text" maxlength="180" placeholder="Ex.: APLICAÇÕES EM ABERTO" value="${escapeHTML(customTitles.get(fileKey(f))||'')}"></label><div class="field-note">Deixe vazio para manter o título existente, quando houver.</div></div><button class="remove" type="button" aria-label="Remover arquivo" title="Remover" data-remove="${i}">×</button></div>`).join(''):'<div class="empty">Os arquivos selecionados aparecerão aqui.</div>';
    document.querySelectorAll('[data-title]').forEach(field=>field.addEventListener('input',e=>customTitles.set(fileKey(selected[+e.target.dataset.title]),e.target.value)));
    document.querySelectorAll('[data-remove]').forEach(btn=>btn.addEventListener('click',()=>{customTitles.delete(fileKey(selected[+btn.dataset.remove]));selected.splice(+btn.dataset.remove,1);outputs=[];$('results').hidden=true;renderQueue()}));
  }
  $('chooseBtn').addEventListener('click',e=>{e.stopPropagation();$('fileInput').click()});
  $('dropzone').addEventListener('click',()=>$('fileInput').click());
  $('dropzone').addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('fileInput').click()}});
  $('fileInput').addEventListener('change',e=>{addFiles(e.target.files);e.target.value='' });
  $('clearBtn').addEventListener('click',()=>{selected.length=0;customTitles.clear();outputs=[];renderQueue();$('results').hidden=true;notify('')});
  for(const ev of ['dragenter','dragover'])$('dropzone').addEventListener(ev,e=>{e.preventDefault();$('dropzone').classList.add('dragover')});
  for(const ev of ['dragleave','drop'])$('dropzone').addEventListener(ev,e=>{e.preventDefault();$('dropzone').classList.remove('dragover')});
  $('dropzone').addEventListener('drop',e=>addFiles(e.dataTransfer.files));

  const getText = (cell,shared) => {
    if(!cell)return '';
    const t=cell.getAttribute('t'), v=first(cell,'v');
    if(t==='s')return shared[+(v?.textContent||0)]||'';
    if(t==='inlineStr')return Array.from(cell.getElementsByTagNameNS(NS,'t')).map(x=>x.textContent).join('');
    if(t==='str'||t==='e')return v?.textContent||'';
    if(first(cell,'f'))return v?.textContent||'';
    return v?.textContent||'';
  };
  async function readShared(zip){
    const p=zip.file('xl/sharedStrings.xml');if(!p)return [];
    const d=xml(await p.async('string'));
    return children(d.documentElement,'si').map(si=>Array.from(si.getElementsByTagNameNS(NS,'t')).map(x=>x.textContent).join(''));
  }
  const colNo=ref=>{const letters=(ref.match(/^[A-Z]+/i)||['A'])[0].toUpperCase();let n=0;for(const ch of letters)n=n*26+ch.charCodeAt(0)-64;return n};
  const colLetters=n=>{let result='';while(n){n--;result=String.fromCharCode(65+(n%26))+result;n=Math.floor(n/26)}return result};
  const rowCells=r=>children(r,'c');
  function buildStyles(styles){
    const root=styles.documentElement;
    const ensure=(name,defaultChildren=[])=>{let el=first(root,name);if(!el){el=node(styles,name,{count:0});defaultChildren.forEach(c=>el.appendChild(c));root.appendChild(el)}return el};
    const fonts=ensure('fonts',[node(styles,'font')]);
    const fills=ensure('fills',[node(styles,'fill')]);
    const borders=ensure('borders',[node(styles,'border')]);
    const xfs=ensure('cellXfs',[node(styles,'xf',{numFmtId:0,fontId:0,fillId:0,borderId:0,xfId:0})]);
    const count=(el,child)=>children(el,child).length;
    const mkFont=(size,bold)=>{const f=node(styles,'font');f.appendChild(node(styles,'sz',{val:size}));f.appendChild(node(styles,'name',{val:'Arial'}));f.appendChild(node(styles,'color',{rgb:'FF111827'}));if(bold)f.appendChild(node(styles,'b'));const id=count(fonts,'font');fonts.appendChild(f);fonts.setAttribute('count',count(fonts,'font'));return id};
    const bodyFont=mkFont(10,false),headFont=mkFont(10,true),titleFont=mkFont(13,true);
    const fill=node(styles,'fill'),pattern=node(styles,'patternFill',{patternType:'solid'});pattern.appendChild(node(styles,'fgColor',{rgb:'FFF0F0F0'}));pattern.appendChild(node(styles,'bgColor',{indexed:64}));fill.appendChild(pattern);const headerFill=count(fills,'fill');fills.appendChild(fill);fills.setAttribute('count',count(fills,'fill'));
    const border=node(styles,'border');for(const side of ['left','right','top','bottom']){const edge=node(styles,side,{style:'thin'});edge.appendChild(node(styles,'color',{rgb:'FF9DA7B1'}));border.appendChild(edge)}border.appendChild(node(styles,'diagonal'));const borderId=count(borders,'border');borders.appendChild(border);borders.setAttribute('count',count(borders,'border'));
    const originalXfs=children(xfs,'xf');
    const ids=new Map();
    function style(role,originalId=0,align='left'){
      const key=[role,originalId,align].join(':');if(ids.has(key))return ids.get(key);
      const orig=originalXfs[originalId]||originalXfs[0];
      const xf=orig?orig.cloneNode(true):node(styles,'xf',{numFmtId:0,fontId:0,fillId:0,borderId:0,xfId:0});
      xf.setAttribute('fontId',role==='title'?titleFont:role==='header'?headFont:bodyFont);
      xf.setAttribute('fillId',role==='header'?headerFill:0);xf.setAttribute('borderId',borderId);
      xf.setAttribute('applyFont','1');xf.setAttribute('applyFill','1');xf.setAttribute('applyBorder','1');xf.setAttribute('applyAlignment','1');
      let al=first(xf,'alignment');if(!al){al=node(styles,'alignment');xf.appendChild(al)}
      al.setAttribute('horizontal',role==='title'||role==='header'?'center':align);
      al.setAttribute('vertical','center');al.setAttribute('wrapText','1');
      const id=count(xfs,'xf');xfs.appendChild(xf);ids.set(key,id);return id;
    }
    return {style,originalXfs,finalize(){xfs.setAttribute('count',count(xfs,'xf'))}};
  }
  // Insere o título sem apagar o cabeçalho ou qualquer célula existente.
  // Para folhas que começam diretamente pelo cabeçalho, uma linha é inserida.
  // Arquivos com fórmulas e sem linha de título são recusados para evitar alterar referências.
  function setSheetTitle(doc,shared,title){
    if(!title)return;
    const root=doc.documentElement, data=first(root,'sheetData');
    if(!data)throw new Error('A aba não contém dados editáveis.');
    const rows=children(data,'row');
    const row1=rows.find(r=>r.getAttribute('r')==='1');
    const row2=rows.find(r=>r.getAttribute('r')==='2');
    const cells1=row1?rowCells(row1).filter(c=>getText(c,shared).trim()!==''):[];
    const merges=first(root,'mergeCells');
    const mergedTitle=children(merges,'mergeCell').some(m=>/^A1:[A-Z]+1$/.test(m.getAttribute('ref')||''));
    const secondCount=row2?rowCells(row2).filter(c=>getText(c,shared).trim()!=='').length:0;
    const existingTitle=Boolean(row1&&(mergedTitle || cells1.length===1&&secondCount>=2));
    if(!existingTitle){
      if(root.getElementsByTagNameNS(NS,'f').length)throw new Error('Esta planilha contém fórmulas e não tem uma linha de título identificável. Para preservar as referências, adicione uma linha de título no Excel antes de formatar.');
      // Suspende inserção quando outros recursos do Excel têm referências complexas.
      if(['tableParts','drawing','pivotTableParts','conditionalFormatting','dataValidations'].some(tag=>root.getElementsByTagNameNS(NS,tag).length))throw new Error('Esta planilha tem elementos avançados e não possui linha de título. Adicione uma linha para título antes de formatar.');
      for(const row of [...rows].reverse()){
        const old=Number(row.getAttribute('r'));
        row.setAttribute('r',old+1);
        for(const cell of rowCells(row)){
          const ref=cell.getAttribute('r');if(ref)cell.setAttribute('r',ref.replace(/\d+$/,(n)=>String(Number(n)+1)));
        }
      }
      // Atualiza referências simples dos intervalos pertencentes à aba.
      const shiftRef=val=>String(val).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(_,c,r)=>c+(+r+1));
      for(const tag of ['dimension','mergeCell','autoFilter','hyperlink','selection']){
        for(const el of Array.from(root.getElementsByTagNameNS(NS,tag))){
          for(const attr of ['ref','sqref','activeCell'])if(el.hasAttribute(attr))el.setAttribute(attr,shiftRef(el.getAttribute(attr)));
        }
      }
    }
    let titleRow=children(data,'row').find(r=>r.getAttribute('r')==='1');
    if(!titleRow){titleRow=node(doc,'row',{r:1});data.insertBefore(titleRow,data.firstChild)}
    // Remover conteúdo apenas da primeira linha se ela já for a linha de título.
    for(const cell of rowCells(titleRow))titleRow.removeChild(cell);
    const cell=node(doc,'c',{r:'A1',t:'inlineStr'}),str=node(doc,'is'),t=node(doc,'t');
    t.textContent=title;str.appendChild(t);cell.appendChild(str);titleRow.appendChild(cell);
  }
  function customizeSheet(doc,styles,shared,opt){
    const root=doc.documentElement, sheetData=first(root,'sheetData');if(!sheetData)return;
    const rows=children(sheetData,'row');if(!rows.length)return;
    const firstRow=rows[0],secondRow=rows[1];
    const firstCells=rowCells(firstRow).filter(c=>getText(c,shared).trim()!=='');
    const secondCells=secondRow?rowCells(secondRow).filter(c=>getText(c,shared).trim()!==''):[];
    const merges=first(root,'mergeCells');
    const hasTitle=Boolean(secondRow&&(firstCells.length===1&&secondCells.length>=2 || children(merges,'mergeCell').some(c=>(c.getAttribute('ref')||'').match(/^A1:[A-Z]+1$/))));
    const headerIdx=hasTitle?1:0;const headerRow=rows[headerIdx];
    if(!headerRow)return;
    const headers=new Map(rowCells(headerRow).map(c=>[colNo(c.getAttribute('r')||'A'),norm(getText(c,shared))]));
    const maxCol=Math.max(1,...rows.slice(0,Math.min(rows.length,150)).flatMap(r=>rowCells(r).map(c=>colNo(c.getAttribute('r')||'A'))));
    // Excel column widths, strictly copied from each identified reference layout.
    if(!opt.keepWidths){
      let cols=first(root,'cols');if(!cols){cols=node(doc,'cols');root.insertBefore(cols,sheetData)}
      const current=new Map();for(const c of children(cols,'col')){for(let n=+c.getAttribute('min');n<=+c.getAttribute('max')&&n<=maxCol;n++)current.set(n,+c.getAttribute('width'))}
      while(cols.firstChild)cols.removeChild(cols.firstChild);
      for(let n=1;n<=maxCol;n++){
        const header=headers.get(n)||'';
        const guessed=/DESCRICAO|OBSERVA|PRODUTO|INSUMO|HISTORICO/.test(header)?42:/FAZENDA|TALHAO|ATIVIDADE|OPERACAO/.test(header)?27:/DATA|VALOR|QTD|QUANTIDADE|AREA|DOSAGEM/.test(header)?14:19;
        const width=Math.min(52,Math.max(10,current.get(n)||guessed));
        cols.appendChild(node(doc,'col',{min:n,max:n,width,customWidth:1}));
      }
    }
    const widths=Array.from({length:maxCol},(_,i)=>{
      const cols=first(root,'cols');const c=children(cols,'col').find(x=>+x.getAttribute('min')<=i+1&&+x.getAttribute('max')>=i+1);return +(c?.getAttribute('width')||18)
    });
    for(let index=0;index<rows.length;index++){
      const row=rows[index],rn=+(row.getAttribute('r')||index+1);
      const role=hasTitle&&index===0?'title':index===headerIdx?'header':'body';
      let height=role==='title'?31:role==='header'?31:22;
      let maxLines=1;
      for(const cell of rowCells(row)){
        const c=colNo(cell.getAttribute('r')||'A');const original=+(cell.getAttribute('s')||0);
        const text=getText(cell,shared);const head=headers.get(c)||'';
        const numId=+(styles.originalXfs[original]?.getAttribute('numFmtId')||0);
        const isDate=/DATA|DT\b|VENCIMENTO|EMISSAO/.test(head)||[14,15,16,17,22,45,46,47].includes(numId)||numId>=164&&/DATA|DT|VENCIMENTO|EMISSAO/.test(head);
        const isNumber=cell.getAttribute('t')!=='s'&&cell.getAttribute('t')!=='inlineStr'&&cell.getAttribute('t')!=='str'&&first(cell,'v')&&Number.isFinite(+first(cell,'v').textContent);
        cell.setAttribute('s',styles.style(role,original,isDate?'center':isNumber?'right':'left'));
        if(role==='body'&&opt.adjustHeight){
          const width=widths[c-1]||18;
          const capacity=Math.max(7,Math.floor(width*1.06)-2);
          const parts=String(text).split(/\r?\n/);
          const lines=parts.reduce((sum,p)=>sum+Math.max(1,Math.ceil(p.length/capacity)),0);
          maxLines=Math.max(maxLines,Math.min(lines,10));
        }
      }
      if(role==='body'&&opt.adjustHeight)height=Math.min(176,Math.max(26,Math.ceil(maxLines*15.4+9)));
      if(role!=='body'||opt.adjustHeight){row.setAttribute('ht',height);row.setAttribute('customHeight','1')}
    }
    // For a title-only row, merge only if it has one value and no other merges intersect row 1.
    if(hasTitle && !children(merges,'mergeCell').some(m=>/^A1:[A-Z]+1$/.test(m.getAttribute('ref')||''))){
      let container=merges;
      if(!container){container=node(doc,'mergeCells');const ordered=['phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'];const next=Array.from(root.children).find(x=>ordered.includes(x.localName));if(next)root.insertBefore(container,next);else root.appendChild(container)}
      container.appendChild(node(doc,'mergeCell',{ref:`A1:${colLetters(maxCol)}1`}));container.setAttribute('count',children(container,'mergeCell').length);
    }
  }
  // XLS (BIFF/OLE) não é ZIP e não pode ser lido por JSZip.
  // Carregamos SheetJS somente quando um arquivo legado .xls for selecionado.
  let sheetJsPromise = null;
  function loadScript(url){
    return new Promise((resolve,reject)=>{
      const script=document.createElement('script');script.src=url;script.async=true;
      script.onload=()=>typeof XLSX!=='undefined'?resolve():reject(new Error('Biblioteca XLS carregada sem a API esperada.'));
      script.onerror=()=>{script.remove();reject(new Error('Falha ao carregar '+url))};
      document.head.appendChild(script);
    });
  }
  async function ensureSheetJS(){
    if(typeof XLSX!=='undefined')return;
    if(!sheetJsPromise){
      sheetJsPromise=(async()=>{
        const sources=[
          'lib/xlsx.full.min.js',
          'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
          'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
        ];
        for(const src of sources){
          try{await loadScript(src);return}catch(e){console.warn(e.message)}
        }
        throw new Error('Não foi possível carregar o conversor .xls. Conecte-se à internet e tente novamente. Arquivos .xlsx continuam funcionando offline.');
      })().catch(e=>{sheetJsPromise=null;throw e});
    }
    await sheetJsPromise;
  }
  async function toXlsxBuffer(file){
    const original=await file.arrayBuffer();
    if(/\.xlsx$/i.test(file.name))return original;
    await ensureSheetJS();
    let book;
    try{
      book=XLSX.read(original,{type:'array',cellStyles:true,cellFormula:true,cellNF:true,cellDates:false});
    }catch(e){throw new Error('Não foi possível ler o arquivo .xls: '+e.message)}
    if(!book.SheetNames?.length)throw new Error('Nenhuma aba foi encontrada no arquivo .xls.');
    try{
      return XLSX.write(book,{bookType:'xlsx',type:'array',compression:true,cellStyles:true});
    }catch(e){throw new Error('Falha ao converter .xls para .xlsx: '+e.message)}
  }
  async function formatFile(file,opt,title){
    const wasXls=/\.xls$/i.test(file.name);
    const zip=await JSZip.loadAsync(await toXlsxBuffer(file));
    if(!zip.file('xl/workbook.xml')||!zip.file('xl/styles.xml'))throw new Error('O arquivo não parece ser um .xlsx válido.');
    const shared=await readShared(zip), stylesDoc=xml(await zip.file('xl/styles.xml').async('string'));
    const styles=buildStyles(stylesDoc);
    const sheets=Object.keys(zip.files).filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort((a,b)=>+a.match(/sheet(\d+)/)[1]-+b.match(/sheet(\d+)/)[1]);
    if(!sheets.length)throw new Error('Nenhuma aba do Excel foi encontrada.');
    for(const path of opt.allSheets?sheets:sheets.slice(0,1)){
      const doc=xml(await zip.file(path).async('string'));
      setSheetTitle(doc,shared,title);
      customizeSheet(doc,styles,shared,opt);
      zip.file(path,serialize(doc));
    }
    styles.finalize();zip.file('xl/styles.xml',serialize(stylesDoc));
    const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:5}});
    return {name:file.name.replace(/\.(xls|xlsx)$/i,'')+'_FORMATADO.xlsx',blob,model:wasXls?'XLS convertido + título personalizado':'Título personalizado'};
  }
  function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),2000)}
  function showResults(){
    $('results').hidden=false;
    $('resultList').innerHTML=outputs.map((r,i)=>`<div class="file-item"><div class="file-icon">✓</div><div class="file-info"><strong>${escapeHTML(r.name)}</strong><small>${previewSize(r.blob.size)} · Padrão: ${escapeHTML(r.model)}</small></div><button class="download" data-download="${i}">↓ Baixar</button></div>`).join('');
    document.querySelectorAll('[data-download]').forEach(b=>b.addEventListener('click',()=>{const r=outputs[+b.dataset.download];download(r.blob,r.name)}));
  }
  $('runBtn').addEventListener('click',async()=>{
    if(!selected.length)return;
    $('runBtn').disabled=true;outputs=[];$('results').hidden=true;
    const opt={adjustHeight:$('adjustHeight').checked,allSheets:$('allSheets').checked,keepWidths:$('keepWidths').checked};
    const failures=[];
    for(let i=0;i<selected.length;i++){
      const f=selected[i];notify(`${/\.xls$/i.test(f.name)?'Convertendo .xls e formatando':'Formatando'} ${i+1} de ${selected.length}: ${f.name}`);
      await new Promise(r=>setTimeout(r,15));
      try{outputs.push(await formatFile(f,opt,(customTitles.get(fileKey(f))||'').trim()))}catch(e){console.error(e);failures.push(`${f.name}: ${e.message}`)}
      $('processedCount').textContent=outputs.length;
    }
    $('runBtn').disabled=false;
    if(outputs.length)showResults();
    notify(`${outputs.length} arquivo(s) formatado(s) com sucesso.${failures.length?' Erro(s): '+failures.join(' | '):' Os arquivos originais foram preservados.'}`,!!failures.length);
  });
  $('downloadAll').addEventListener('click',async()=>{
    if(!outputs.length)return;
    $('downloadAll').disabled=true;$('downloadAll').textContent='Preparando ZIP...';
    try{const zip=new JSZip();for(const r of outputs)zip.file(r.name,r.blob);download(await zip.generateAsync({type:'blob'}),'Relatorios_Formatados.zip')}
    catch(e){notify('Não foi possível criar o ZIP: '+e.message,true)}
    finally{$('downloadAll').disabled=false;$('downloadAll').textContent='↓ Baixar todos (.zip)'}
  });
  renderQueue();
})();
