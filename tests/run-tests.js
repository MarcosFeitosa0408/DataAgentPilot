const {JSDOM,VirtualConsole}=require('jsdom');const fs=require('fs'),crypto=require('crypto'),XLSX=require('xlsx');
const html=fs.readFileSync(require('path').join(__dirname,'..','index.html'),'utf8');
const errs=[];
function boot(h=html,nosub){const vc=new VirtualConsole();vc.on('jsdomError',e=>{if(/Could not parse CSS stylesheet/.test(String(e.message)))return;/* jsdom não entende @page com caixas de margem (CSS de impressão); navegadores entendem */errs.push(String(e.message).slice(0,200))});
 const w=new JSDOM(h,{runScripts:'dangerously',pretendToBeVisual:true,virtualConsole:vc,beforeParse(w){Object.defineProperty(w,'crypto',{value:nosub?{}:crypto.webcrypto,configurable:true});w.TextDecoder=TextDecoder;w.TextEncoder=TextEncoder;w.XLSX=XLSX;w.confirm=()=>true;w.print=()=>{};}}).window;
 w.eval("save=(n,t,x)=>{window.__saved={n,t,x}}");return w}
const ab=b=>{b=Buffer.isBuffer(b)?b:Buffer.from(b,'utf8');return b.buffer.slice(b.byteOffset,b.byteOffset+b.length)};
const R=[];const t=(n,c,d)=>R.push([c?'PASS':'FAIL',n,c?'':d||'']);
const q=(w,s)=>w.document.querySelector(s);
async function go(csv,name='t.csv',tol){const w=boot();if(tol!==undefined)q(w,'#tol').value=tol;await w.load(ab(csv),name);return w}
const res=w=>w.eval('S.res'),rule=(w,id)=>res(w).rules.find(r=>r[0]===id),cnt=(w,id)=>res(w).D.filter(d=>d.rule===id).length;
(async()=>{
 const H='Produto,Receita,Custo,Lucro\n';
 // 1 CSV válido (BR, ;)
 let w=await go('Produto;Receita;Custo;Lucro\nA;1.234,56;234,56;1.000,00\nB;100,00;40,00;60,00\n');
 t('CSV válido (; e formato BR)',res(w).analyzed===2&&res(w).fin.checked===2&&res(w).fin.det.length===0&&w.eval('S.notes').join().includes('ponto e vírgula'),JSON.stringify(res(w).fin));
 t('Dados originais preservados (string bruta intacta)',res(w).rows[0].v[1]==='1.234,56'&&w.eval('S.raw')[1][1]==='1.234,56');
 // SHA
 const buf=Buffer.from('Produto;Receita\nA;1\n');w=await go(buf);
 t('SHA-256 sobre bytes originais',w.eval('S.trace.sha256')===crypto.createHash('sha256').update(buf).digest('hex'),w.eval('S.trace.sha256'));
 // exemplo do prompt
 w=await go(H+'Produto A,"100,00","20,00","80,00"\nProduto B,"-20,00","0,00","-20,00"\nProduto C,"0,00","-5,00","4,00"\nProduto D,Nulo,Nulo,"0,00"\n');
 let r=res(w),d=r.fin.det[0];
 t('Exemplo do prompt: 4 lidas/3 verificadas/1 não/1 divergência',r.read===4&&r.fin.checked===3&&r.fin.unver===1&&r.fin.det.length===1,JSON.stringify([r.read,r.fin]));
 t('Registro 3 Produto C: calc 5,00 / dif 1,00',d&&d.n===3&&d.prod==='Produto C'&&d.R===0&&d.C===-5&&d.L===4&&d.calc===5&&d.diff===1,JSON.stringify(d));
 t('Negativos e zeros = Informativo (não erro)',cnt(w,'R06')>0&&r.D.filter(x=>x.rule==='R06').every(x=>x.sev==='Informativo')&&rule(w,'R06')[2]==='Atenção');
 t('Receita/Custo/Lucro ausentes não viram erro financeiro',r.D.filter(x=>x.rule==='R05').length===1);
 // 0,01 e tolerância
 const c1=H+'A,"100,00","20,00","79,99"\n';
 w=await go(c1);t('Divergência de R$ 0,01 detectada (tol 0)',res(w).fin.det.length===1&&res(w).fin.det[0].diff===0.01,JSON.stringify(res(w).fin.det));
 w=await go(c1,'t.csv','0.01');t('Tolerância R$ 0,01 aceita a diferença',res(w).fin.det.length===0&&rule(w,'R05')[2]==='Aprovada');
 w=await go(H+'A,"100,00","20,00","79,00"\n');t('Divergência de R$ 1,00 detectada',res(w).fin.det[0]&&res(w).fin.det[0].diff===1&&rule(w,'R05')[2]==='Falha'&&res(w).D[0].sev==='Crítico');
 w=await go(H+'A,0.1,0.2,-0.1\nB,1.1,0.1,1.0\n');t('Sem falso erro de ponto flutuante (0,1−0,2 / 1,1−0,1)',res(w).fin.det.length===0,JSON.stringify(res(w).fin.det));
 t('Arquivo sem divergências: R05 Aprovada',rule(w,'R05')[2]==='Aprovada');
 // ausência de colunas financeiras
 w=await go('Produto,Qtd\nA,1\n');t('Sem colunas financeiras: R05/R06/R04 "Não avaliada"',['R04','R05','R06'].every(i=>rule(w,i)[2]==='Não avaliada')&&q(w,'#fi').textContent.includes('Não avaliado'));
 // parcialmente preenchido
 w=await go(H+'A,100,,80\nB,,,\nC,50,10,40\n');t('Parcialmente preenchido: 1 verificada, 2 não',res(w).fin.checked===1&&res(w).fin.unver===2);
 // ausentes + duplicados
 w=await go('Produto,Qtd,Cat\nA,1,x\nA,1,x\nB,,y\nC,3,\n');r=res(w);
 t('Valores ausentes (R01=2) e duplicados (R02=1, nada removido)',cnt(w,'R01')===2&&r.dups===1&&r.analyzed===4&&rule(w,'R02')[2]==='Atenção',JSON.stringify([cnt(w,'R01'),r.dups,r.analyzed]));
 // misto + injeção csv
 w=await go('Id,Valor\n1,10\n2,20\n3,30\n4,=1+1\n');r=res(w);
 t('Mistura texto/número (coluna Misto, 1 diag R03)',r.prof[1].mixed&&cnt(w,'R03')===1&&r.prof[1].type==='Misto',JSON.stringify(r.prof[1]));
 q(w,'#xc').click();const csv=w.__saved.x;
 t('Exportação CSV de diagnósticos (cabeçalho + linhas, BOM)',csv.charCodeAt(0)===0xFEFF&&csv.split('\r\n').length===r.D.length+1&&csv.includes('regra,nome'),csv.slice(0,120));
 t('CSV protegido contra injeção de fórmula',csv.includes('"\'=1+1"'),csv);
 // calendário
 w=await go('Ano,Mes,Dia\n2025,13,5\n2025,2,30\n2025,5,32\n2025,2,28\n2024,2,29\n');r=res(w);
 const cal=r.D.filter(x=>x.rule==='R04').map(x=>x.n+':'+x.name);
 t('Mês fora do intervalo (reg 1)',cal.includes('1:Mês fora do intervalo'),cal.join());
 t('Data inexistente 30/02 (reg 2)',cal.includes('2:Data inexistente'),cal.join());
 t('Dia fora do intervalo (reg 3)',cal.includes('3:Dia fora do intervalo'),cal.join());
 t('Datas válidas (28/02/2025, 29/02/2024) não flagradas',!cal.some(x=>/^[45]:/.test(x))&&rule(w,'R04')[2]==='Falha',cal.join());
 // contagem de linhas
 w=await go('A,B,C\n1,2,3\n\n4,5\n7,8,9\n');r=res(w);
 t('Linhas lidas/analisadas/não analisadas = 3/2/1 + diag R00 no reg 2',r.read===3&&r.analyzed===2&&w.eval('S.trace.linhas_nao_analisadas')===1&&r.D.some(x=>x.rule==='R00'&&x.n===2),JSON.stringify([r.read,r.analyzed]));
 // vazio / inválido
 w=boot();await w.load(ab(Buffer.alloc(0)),'v.csv');t('Arquivo vazio: mensagem clara',q(w,'#msg').textContent.includes('vazio'),q(w,'#msg').textContent);
 w=boot();await w.load(ab('Produto,Receita\n'),'v.csv');t('Só cabeçalho: sem registros',q(w,'#msg').textContent.includes('Não há registros'),q(w,'#msg').textContent);
 w=boot();await w.load(ab('abc'),'v.txt');t('Formato incompatível (.txt)',q(w,'#msg').textContent.includes('incompatível'),q(w,'#msg').textContent);
 w=boot();await w.load(ab(crypto.randomBytes(300)),'lixo.xlsx');const m=q(w,'#msg').textContent;t('.xlsx corrompido não gera análise enganosa',m.length>0&&!w.eval('S.res'),'msg="'+m+'" res='+JSON.stringify(w.eval('S.res&&S.res.analyzed')));
 // Excel válido, 2 abas
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Produto','Receita','Custo','Lucro'],['A',100,20,80],['B',50,10,41]]),'Vendas');XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['X'],[1]]),'Outra');
 w=boot();await w.load(ab(XLSX.write(wb,{type:'buffer',bookType:'xlsx'})),'v.xlsx');r=res(w);
 t('Excel válido: leitura, divergência e seletor de abas',r.analyzed===2&&r.fin.det.length===1&&!q(w,'#sheetBox').classList.contains('hide'),JSON.stringify([r.analyzed,r.fin.det.length]));
 const sel=q(w,'#sheet');sel.value='Outra';sel.dispatchEvent(new w.Event('change'));t('Troca de aba reanalisa',res(w).cols===1&&res(w).analyzed===1);
 // fuso
 w=boot();t('UTC→Brasília 15:00Z = 12:00',w.eval('brt')(new Date('2026-01-15T15:00:00Z')).includes('12:00:00'),w.eval('brt')(new Date('2026-01-15T15:00:00Z')));
 t('UTC→Brasília cruza o dia (02:30Z → 23:30 do dia anterior)',w.eval('brt')(new Date('2026-06-15T02:30:00Z')).startsWith('14/06/2026'),w.eval('brt')(new Date('2026-06-15T02:30:00Z')));
 // XSS + filtros + exports no mesmo arquivo
 w=await go(H+'"<img src=x onerror=window.__x=1>",100,20,79\nB,100,20,80\nC,5,1,1\nC,5,1,1\n');
 t('XSS: HTML da planilha vira texto (sem <img> no DOM)',!q(w,'#out img')&&!w.__x&&q(w,'#fi').textContent.includes('<img'));
 const fr=q(w,'#fr');fr.value='R02';fr.dispatchEvent(new w.Event('change'));
 let rows=[...w.document.querySelectorAll('#dg tbody tr')];t('Filtro por regra (R02) mostra só R02',rows.length>0&&rows.every(x=>x.cells[0].textContent==='R02'),rows.length);
 fr.value='';fr.dispatchEvent(new w.Event('change'));const fs=q(w,'#fs');fs.value='Crítico';fs.dispatchEvent(new w.Event('change'));
 rows=[...w.document.querySelectorAll('#dg tbody tr')];t('Filtro por severidade (Crítico)',rows.length>0&&rows.every(x=>x.cells[2].textContent==='Crítico'),rows.length);
 fs.value='';fs.dispatchEvent(new w.Event('change'));
 q(w,'#xm').click();t('Descrição da análise (.md) gerada',w.__saved.n.endsWith('.md')&&w.__saved.x.includes('## Contrato de qualidade')&&w.__saved.x.includes(w.eval('S.trace.id')));
 q(w,'#xj').click();let j;try{j=JSON.parse(w.__saved.x)}catch(e){}t('Relatório JSON válido com rastreabilidade e diagnósticos',j&&j.rastreabilidade&&j.diagnosticos.length===w.eval('S.res.D.length'));
 q(w,'#xp').click();t('PDF: bloco de impressão preenchido com ID e SHA',q(w,'#pr').textContent.includes(w.eval('S.trace.id'))&&q(w,'#pr').textContent.includes('SHA-256'));
 q(w,'#xh').click();const dash=w.__saved.x;errs.length=0;
 const w2=boot(dash);
 t('Dashboard HTML: sem erros de script ao abrir',errs.length===0,errs.join(' | '));
 t('Dashboard HTML: KPIs, regras e diagnósticos renderizados',w2.document.querySelectorAll('#kpis .card').length>=5&&w2.document.querySelectorAll('#rules tbody tr').length===7&&w2.document.querySelectorAll('#dg tbody tr').length>0);
 const f2=w2.document.querySelector('#fs');f2.value='Crítico';f2.dispatchEvent(new w2.Event('change'));rows=[...w2.document.querySelectorAll('#dg tbody tr')];
 t('Dashboard HTML: filtro funcional',rows.length>0&&rows.every(x=>x.cells[2].textContent==='Crítico'),rows.length);
 t('Dashboard HTML: sem injeção (img) e sem dependência externa',!w2.document.querySelector('img')&&!/src="http|href="http/.test(dash));
 // desempenho 30k
 let big=H;for(let i=0;i<30000;i++)big+='P'+(i%500)+','+(100+i%7)+',20,'+(80+i%7)+'\n';
 let t0=Date.now();w=await go(big);const ms=Date.now()-t0;
 t('30.000 registros: análise em '+ms+' ms e DOM enxuto',res(w).analyzed===30000&&ms<15000&&w.document.querySelectorAll('*').length<1500,w.document.querySelectorAll('*').length);
 
 // ===== Conclusão das limitações =====
 w=boot();const sj=w.eval('sha256js');let okAll=true,bad='';
 for(const n of [0,1,3,55,56,63,64,65,119,1000,100000]){const b=crypto.randomBytes(n);const h1=sj(ab(b)),h2=crypto.createHash('sha256').update(b).digest('hex');if(h1!==h2){okAll=false;bad+=n+' '}}
 t('SHA-256 próprio idêntico ao nativo (tamanhos 0 a 100 mil bytes)',okAll,'divergiu em: '+bad);
 const wn=boot(html,true);await wn.load(ab(Buffer.from('Produto,Qtd\nA,1\n')),'n.csv');
 t('Sem crypto.subtle: hash real calculado pelo fallback e método informado',wn.eval('S.trace.sha256')===crypto.createHash('sha256').update('Produto,Qtd\nA,1\n').digest('hex')&&wn.eval('S.trace.sha256_metodo').includes('própria')&&q(wn,'#tr').textContent.includes('método'),wn.eval('S.trace.sha256'));
 t('Com crypto.subtle: método = Web Crypto API',(await go('A,B\n1,2\n')).eval('S.trace.sha256_metodo')==='Web Crypto API');
 w=await go('Data,Receita\n2025-03-15,10\n2025-02-30,20\n15/13/2025,5\nabc,1\n31/04/2025,3\n29/02/2024,4\n');
 let c4=res(w).D.filter(x=>x.rule==='R04').map(x=>x.n+':'+x.name);
 t('Coluna de data completa validada (ISO e DD/MM/AAAA)',JSON.stringify(c4)===JSON.stringify(['2:Data inexistente','3:Data inexistente','4:Formato de data não reconhecido','5:Data inexistente']),c4.join());
 t('Coluna de data: regra R04 avaliada (Falha) e datas válidas passam',rule(w,'R04')[2]==='Falha'&&!c4.some(x=>/^[16]:/.test(x)));
 w=await go('Data,Receita,Categoria\n2025-01-10,100,A\n2025-01-20,50,B\n2025-02-01,70,A\n');
 t('Gráfico mensal gerado a partir da coluna Data',w.eval('S.ch_c2')&&JSON.stringify(w.eval('S.ch_c2[1]'))==='[150,70]'&&w.eval('S.ch_c2[0]').join()==='01/2025,02/2025',JSON.stringify(w.eval('S.ch_c2')));
 t('Gráfico por categoria (Receita+Categoria)',JSON.stringify(w.eval('S.ch_c1[1]'))==='[170,50]');
 w=await go('Produto,Valor Bruto,Gasto,Resultado\nA,100,20,80\nB,50,10,41\n');
 t('Sem nomes reconhecidos: financeiro "Não avaliada" (nada inventado)',rule(w,'R05')[2]==='Não avaliada'&&q(w,'#mp_rev').value==='-1');
 for(const [k,i] of [['rev',1],['cost',2],['prof',3]]){const e=q(w,'#mp_'+k);e.value=String(i);e.dispatchEvent(new w.Event('change'))}
 t('Mapeamento manual de colunas ativa o financeiro (1 divergência no reg 2)',res(w).fin.checked===2&&res(w).fin.det.length===1&&res(w).fin.det[0].n===2&&rule(w,'R05')[2]==='Falha',JSON.stringify(res(w).fin));
 const e2=q(w,'#mp_rev');e2.value='-1';e2.dispatchEvent(new w.Event('change'));
 t('Remover mapeamento desliga o módulo de novo',rule(w,'R05')[2]==='Não avaliada');
 w=await go('Faturamento,Cost,Profit,Year,Month,Day,Segmento\n100,20,80,2025,1,10,x\n');
 t('Sinônimos (Faturamento/Cost/Profit/Year/Month/Day/Segmento) reconhecidos',res(w).finRun&&rule(w,'R04')[2]==='Aprovada'&&res(w).M.cat===6,JSON.stringify(res(w).M));
 const vm=require('vm'),cx={};vm.runInNewContext(require('fs').readFileSync(require('path').join(__dirname,'..','libs','xlsx.full.min.js'),'utf8'),cx);
 t('libs/xlsx.full.min.js local é funcional (Excel offline)',cx.XLSX&&typeof cx.XLSX.read==='function'&&cx.XLSX.version==='0.18.5',cx.XLSX&&cx.XLSX.version);
 console.log(R.map(x=>x.join(' | ')).join('\n'));console.log('\nPASS',R.filter(x=>x[0]==='PASS').length,'FAIL',R.filter(x=>x[0]==='FAIL').length);
})();
