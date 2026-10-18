// PCP Bandas · Rubbernew
// Lê as planilhas enviadas para o repositório (CSV ou Excel), aplica as regras
// combinadas e gera os dados do app. Roda sozinho no GitHub a cada envio.
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import XLSX from 'xlsx';
const LINHA = {'BE 45':'PREMIUM','BR 80':'ALTER','BS 70':'SUPER PREMIUM','OT 75':'ECOMAX','BC 01':'ECOMAX'};
const PROD = JSON.parse(fs.readFileSync(new URL('./produtos.json', import.meta.url),'utf8'));
const DIAS_IGNORADOS = ['2026-10-06'];
const PRENSA_NUM = {3:'A',4:'B',89:'C',119:'D',127:'E'};
const semAcento = s => String(s??'').normalize('NFD').replace(/[̀-ͯ]/g,'');
const N = s => semAcento(s).toUpperCase().replace(/\s+/g,' ').trim();
const U = s => String(s??'').toUpperCase().replace(/\s+/g,' ').trim();
const isoDe = br => br.slice(6,10)+'-'+br.slice(3,5)+'-'+br.slice(0,2);
const brDe = iso => iso.slice(8,10)+'/'+iso.slice(5,7)+'/'+iso.slice(0,4);
function numBR(v){
  if(typeof v==='number') return v;
  let s=String(v??'').trim(); if(!s||s.includes('#')) return null;
  if(s.includes(',')&&s.includes('.')) s=s.replace(/\./g,'').replace(',','.');
  else if(s.includes(',')) s=s.replace(',','.');
  const n=parseFloat(s); return isNaN(n)?null:n;
}
function keycode(c){
  const m=String(c??'').trim().replace(/^'+/,'').trim().match(/^(\d{13}\.\d{4})\s*(.*)$/);
  if(!m) return null; const suf=m[2].replace(/[\s\-']+/g,'').toUpperCase();
  return m[1]+(suf?'-'+suf:'');
}
function achaProduto(raw){
  const k=keycode(raw); if(!k) return null;
  if(PROD[k] && !k.endsWith('-ANTIGA')) return [k,PROD[k]];
  const base=k.split('-')[0];
  if(PROD[base+'-NOVA']) return [base+'-NOVA',PROD[base+'-NOVA']];
  const vs=Object.keys(PROD).filter(c=>c.startsWith(base+'-')&&!c.endsWith('-ANTIGA'));
  return vs.length===1?[vs[0],PROD[vs[0]]]:null;
}
// índice por descrição (sem PREMIUM/ALTER/ECOMAX/BFE) + massa, para os pedidos
const PALAVRAS_LINHA = new Set(['PREMIUM','ALTER','ECOMAX','BFE','SUPER']);
const chaveDesc = d => N(d).split(' ').filter(w=>w&&!PALAVRAS_LINHA.has(w)).sort().join(' ');
const IDX_DESC = {};
for(const [c,p] of Object.entries(PROD)){ const k=chaveDesc(p[0])+'|'+p[1]; if(!IDX_DESC[k]) IDX_DESC[k]=c; }
function massasDaLinha(desc){
  const w=new Set(N(desc).split(' '));
  if(w.has('SUPER')&&w.has('PREMIUM')) return ['BS 70'];
  if(w.has('PREMIUM')) return ['BE 45'];
  if(w.has('ALTER')) return ['BR 80'];
  if(w.has('ECOMAX')) return ['OT 75','BC 01','BS 70'];
  return [];
}


function csvLinhas(txt){
  const rows=[]; let row=[], cel='', q=false;
  for(let i=0;i<txt.length;i++){const ch=txt[i];
    if(q){ if(ch==='"'){ if(txt[i+1]==='"'){cel+='"';i++;} else q=false; } else cel+=ch; }
    else if(ch==='"') q=true;
    else if(ch===','){row.push(cel);cel='';}
    else if(ch==='\n'){row.push(cel);rows.push(row);row=[];cel='';}
    else if(ch!=='\r') cel+=ch;
  }
  if(cel||row.length){row.push(cel);rows.push(row);}
  return rows;
}

function lerArquivo(caminho, marcador){
  const nome=caminho.toLowerCase(), buf=fs.readFileSync(caminho);
  if(nome.endsWith('.csv')||nome.endsWith('.txt')){
    let txt=new TextDecoder('utf-8').decode(buf);
    if(txt.includes(String.fromCharCode(0xFFFD))) txt=new TextDecoder('windows-1252').decode(buf);
    let rows=csvLinhas(txt);
    if(rows.length && rows.filter(r=>r.length<3).length>rows.length*.8 && txt.includes(';')) rows=csvLinhas(txt.replace(/;/g,','));
    return rows;
  }
  const wb=XLSX.read(buf,{type:'buffer',cellDates:false});
  const folhas=wb.SheetNames.map(n=>{const ws=wb.Sheets[n]; if(!ws['!ref']) return [n,[]];
    const r=XLSX.utils.decode_range(ws['!ref']); r.s.c=0; r.s.r=0;
    return [n,XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:'',range:r,dateNF:'dd/mm/yyyy'})];});
  const boa=folhas.find(([n,rows])=>rows.some(r=>marcador(r)));
  if(!boa) throw new Error('nenhuma aba do arquivo tem o formato esperado');
  return boa[1];
}
function mapaColunas(row, nomes){
  const m={}; row.forEach((c,i)=>{const k=N(c); for(const [campo,alvos] of Object.entries(nomes)) if(m[campo]==null && alvos.includes(k)) m[campo]=i;});
  return m;
}

/* ---------- regras da EXTRUSAGEM ---------- */
const COLS_EXT={codigo:['CODIGO TOTVS'],massa:['MASSA'],cor:['COR DA LISTRA'],qtde:['QTDE'],prensa:['PRENSA'],plato:['PLATO'],uso:['MASSA UTILIZADA']};
const ehCabExt = r => r.some(c=>N(c)==='CODIGO TOTVS');
function processaExtrusagem(rows){
  let col={codigo:2,massa:4,cor:5,qtde:6,prensa:16,plato:18,uso:19}, data=null;
  const dias={}, blocos={}, semCadastro={}, ignorados=new Set();
  for(const r of rows){
    const linha=r.map(c=>String(c??'')).join(' ');
    const md=linha.match(/DATA\s*:?\s*(\d{2}\/\d{2}\/\d{4})/i);
    if(md && /PROG|DATA\s*:/i.test(linha)){ data=md[1]; blocos[data]=(blocos[data]||0)+1; continue; }
    if(ehCabExt(r)){ col={...col,...mapaColunas(r,COLS_EXT)}; continue; }
    const cod=String(r[col.codigo]??'').trim();
    if(!/^'?\s*\d{13}\.\d{4}/.test(cod) || !data) continue;
    const iso=isoDe(data);
    if(DIAS_IGNORADOS.includes(iso)){ ignorados.add(data); continue; }
    const achado=achaProduto(cod), p=achado&&achado[1];
    if(!p) semCadastro[cod]=(semCadastro[cod]||0)+1;
    const qt=numBR(r[col.qtde])||0, mp=N(r[col.massa]);
    const massa = LINHA[mp]? mp : (p? p[1] : null);
    const pr=String(r[col.prensa]??'').trim();
    const prensa = PRENSA_NUM[parseInt(pr,10)] || pr;
    (dias[iso]||(dias[iso]={data,iso,itens:[]})).itens.push({
      codigo:cod, desc:p?p[0]:null, massa, linha:LINHA[massa]||null, cor:N(r[col.cor]), qtde:qt,
      prensa, plato:String(r[col.plato]??'').trim(), uso:N(r[col.uso]),
      peso_medio:p?p[2]:null, kg:p?Math.round(qt*p[2]*100)/100:null, larg:p?p[3]:null, esp:p?p[4]:null });
  }
  return {dias:Object.values(dias).sort((a,b)=>a.iso.localeCompare(b.iso)), semCadastro, ignorados:[...ignorados],
    repetidos:Object.entries(blocos).filter(([d,n])=>n>1&&!DIAS_IGNORADOS.includes(isoDe(d))).map(([d])=>d)};
}

/* ---------- regras dos PEDIDOS ---------- */
const COLS_PED={pedido:['NO PEDIDO','Nº PEDIDO','N PEDIDO'],emissao:['DATA EMISSAO'],cliente:['CLIENTE'],loja:['LOJA'],nome:['NOME'],item:['ITEM'],
  desc:['DESCRICAO'],entrega:['ENTREGA'],qtde:['QTDE PENDENTE'],posicao:['POSICAO'],producao:['DATA DE PRODUCAO'],obs:['OBSERVACOES']};
const ehCabPed = r => r.some(c=>N(c)==='QTDE PENDENTE') && r.some(c=>N(c)==='DESCRICAO');
function processaPedidos(rows){
  let col=null; const itens=[], fora=[];
  for(const r of rows){
    if(!col){ if(ehCabPed(r)){ col=mapaColunas(r.map(c=>String(c).replace('º','O')),COLS_PED); } continue; }
    const desc=U(r[col.desc]); if(!desc) continue;
    const qt=numBR(r[col.qtde]); if(qt==null) continue;
    const massas=massasDaLinha(desc); let codigo=null,massa=null;
    for(const m of massas){ const c=IDX_DESC[chaveDesc(desc)+'|'+m]; if(c){codigo=c;massa=m;break;} }
    const base={pedido:String(r[col.pedido]??'').trim(), emissao:String(r[col.emissao]??'').trim(), cliente:String(r[col.cliente]??'').trim(),
      nome:U(r[col.nome]), item:String(r[col.item]??'').trim(), desc, entrega:String(r[col.entrega]??'').trim(), qtde:qt};
    if(!codigo){ fora.push(base); continue; }
    const pos=U(r[col.posicao])||'PRODUZIR';
    const datas=(String(r[col.producao]??'').match(/\d{2}\/\d{2}\/\d{4}/g)||[]);
    const p=PROD[codigo];
    itens.push({...base, linha:LINHA[massa], massa, codigo, produto:p[0], posicao:pos,
      inicio:datas[0]||null, fim:datas.at(-1)||null, obs:String(r[col.obs]??'').trim(),
      peso_medio:p[2], kg:Math.round(qt*p[2]*100)/100});
  }
  if(!col) throw new Error('não encontrei o cabeçalho dos pedidos (colunas DESCRIÇÃO e QTDE PENDENTE)');
  return {itens, fora};
}
function dataDoNome(nome){
  const m=nome.match(/(\d{2})[-_.](\d{2})[-_.](\d{2,4})/); if(!m) return null;
  const a=m[3].length===2?'20'+m[3]:m[3]; return `${a}-${m[2]}-${m[1]}`;
}
const resumoPed = itens => ({itens:itens.length, pedidos:new Set(itens.map(i=>i.pedido)).size,
  bandas:itens.reduce((a,i)=>a+i.qtde,0), kg:Math.round(itens.reduce((a,i)=>a+i.kg,0)*10)/10,
  posicao:itens.reduce((o,i)=>(o[i.posicao]=(o[i.posicao]||0)+i.qtde,o),{})});



// ===================================================================
//  Geração dos dados do app a partir das planilhas da raiz do repositório
// ===================================================================
const RAIZ = path.dirname(new URL(import.meta.url).pathname);
const SAIDA = path.join(RAIZ, 'site');
const EXTS = ['.csv', '.xlsx', '.xls'];
const isoLocal = ms => new Date(ms - 3 * 3600e3).toISOString().slice(0, 10); // horário de Brasília
function quandoEnviado(p) {
  let q = 0;
  try { q = +execFileSync('git', ['log', '-1', '--format=%ct', '--', p], { cwd: RAIZ, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim() * 1000; } catch {}
  return q || fs.statSync(p).mtimeMs;
}
const planilhas = fs.readdirSync(RAIZ)
  .filter(n => EXTS.includes(path.extname(n).toLowerCase()))
  .map(n => ({ nome: n, caminho: path.join(RAIZ, n), quando: quandoEnviado(path.join(RAIZ, n)) }))
  .sort((a, b) => a.quando - b.quando || a.nome.localeCompare(b.nome)); // a enviada por último prevalece

const base = { gerado_em: new Date().toISOString(), extrusora: [], pedidos: [], arquivos: { extrusora: [], pedidos: [] } };
const dias = {}, fotos = {};
for (const a of planilhas) {
  const enviado_em = new Date(a.quando).toISOString();
  // o tipo da planilha é reconhecido pelo cabeçalho
  let tipo = null, rows = null, erro = null;
  try { rows = lerArquivo(a.caminho, ehCabPed); tipo = 'pedidos'; }
  catch { try { rows = lerArquivo(a.caminho, ehCabExt); tipo = 'extrusora'; } catch (e) { erro = e.message; } }
  if (tipo === 'pedidos' && !rows.some(ehCabPed)) { try { rows = lerArquivo(a.caminho, ehCabExt); tipo = 'extrusora'; } catch {} }
  if (tipo === 'extrusora' && !rows.some(ehCabExt)) { tipo = null; erro = 'planilha não reconhecida (nem extrusagem nem pedidos)'; }
  if (!tipo) { base.arquivos.extrusora.push({ nome: a.nome, enviado_em, dias: [], erro: erro || 'planilha não reconhecida' }); continue; }
  if (tipo === 'extrusora') {
    const info = { nome: a.nome, enviado_em, dias: [], semCadastro: [], erro: null };
    try {
      const r = processaExtrusagem(rows);
      for (const d of r.dias) { dias[d.iso] = { ...d, arquivo: a.nome, importado_em: enviado_em }; info.dias.push(d.data); }
      info.semCadastro = Object.keys(r.semCadastro); info.repetidos = r.repetidos;
      if (!r.dias.length) info.erro = 'nenhum dia de programação encontrado';
    } catch (e) { info.erro = e.message; }
    base.arquivos.extrusora.push(info);
  } else {
    const ref = dataDoNome(a.nome) || isoLocal(a.quando);
    const info = { nome: a.nome, enviado_em, ref, itens: 0, fora: 0, erro: null };
    try {
      const r = processaPedidos(rows);
      fotos[ref] = { ref, arquivo: a.nome, importado_em: enviado_em, resumo: resumoPed(r.itens), fora: r.fora, itens: r.itens };
      info.itens = r.itens.length; info.fora = r.fora.length;
    } catch (e) { info.erro = e.message; }
    base.arquivos.pedidos.push(info);
  }
}
base.extrusora = Object.values(dias).sort((a, b) => a.iso.localeCompare(b.iso));
base.pedidos = Object.values(fotos).sort((a, b) => a.ref.localeCompare(b.ref));
fs.mkdirSync(path.join(SAIDA, 'dados'), { recursive: true });
fs.writeFileSync(path.join(SAIDA, 'dados', 'base.json'), JSON.stringify(base));
fs.copyFileSync(path.join(RAIZ, 'index.html'), path.join(SAIDA, 'index.html'));
console.log(`Extrusagem: ${base.extrusora.length} dias · Pedidos: ${base.pedidos.length} carteiras`);
for (const t of ['extrusora', 'pedidos']) for (const a of base.arquivos[t]) if (a.erro) console.log(`AVISO ${a.nome}: ${a.erro}`);
