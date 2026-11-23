// Uso: node gerar_freq.js dados.json saida.docx
const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,ShadingType,ImageRun,BorderStyle,VerticalAlign,HeightRule,TableLayoutType}=require('docx');
const [,,inp,out]=process.argv; const D=JSON.parse(fs.readFileSync(inp,'utf8'));
const F='Arial', NAVY='282C62';
const b={style:BorderStyle.SINGLE,size:6,color:'000000'}; const borders={top:b,bottom:b,left:b,right:b};
const W=10466; // A4 com margens de 0,6 pol (larg. útil em DXA)
const t=(text,o={})=>new TextRun({text,font:F,size:o.size||20,bold:!!o.bold,color:o.color,italics:!!o.it});
const cell=(text,w,o={})=>new TableCell({borders,width:{size:w,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
  shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,columnSpan:o.span,
  margins:{top:40,bottom:40,left:80,right:80},
  children:[new Paragraph({alignment:o.align||AlignmentType.LEFT,children:[t(text,o)]})]});
// Cabeçalho (rótulo | valor | rótulo | valor)
const L=1900,V=3333;
const hrow=(a,av,b2,bv)=>new TableRow({children:[cell(a,L,{bold:true,fill:'E8ECF5'}),cell(av,V),cell(b2,L,{bold:true,fill:'E8ECF5'}),cell(bv,V)]});
const cab=new Table({width:{size:W,type:WidthType.DXA},columnWidths:[L,V,L,V],layout:TableLayoutType.FIXED,rows:[
  hrow('MUNICÍPIO','TUPANATINGA - PE','GRE',D.gre||''),
  hrow('LOCAL',D.local,'ARTICULADOR MUNICIPAL','ANTONIO CARLOS PEREIRA'),
  hrow('PÚBLICO',D.publico,'DATA',D.data),
  new TableRow({children:[cell('TEMA',L,{bold:true,fill:'E8ECF5'}),cell(D.tema,V+L+V,{span:3})]}),
  hrow('HORÁRIO',D.horario,'PREVISTOS',String(D.rows.length)),
  hrow('AUSÊNCIAS JUSTIFICADAS',String(D.rows.filter(r=>r.obs).length),'PRESENTES',''),
]});
// Lista
const cN=600,cNome=4300,cAss=3466,cObs=2100;
const head=new TableRow({tableHeader:true,children:['Nº','NOME','ASSINATURA','OBS'].map((h,i)=>cell(h,[cN,cNome,cAss,cObs][i],{bold:true,fill:NAVY,color:'FFFFFF',align:AlignmentType.CENTER}))});
const linhas=D.rows.map((r,i)=>{const aus=!!r.obs, f=aus?'D9D9D9':undefined;
  return new TableRow({height:{value:470,rule:HeightRule.ATLEAST},cantSplit:true,children:[
   cell(String(i+1),cN,{align:AlignmentType.CENTER,fill:f}),cell(r.nome,cNome,{fill:f,size:19}),
   cell(aus?'AUSENTE':'',cAss,{fill:f,bold:aus,align:AlignmentType.CENTER}),cell(r.obs||'',cObs,{fill:f,size:16})]});});
const lista=new Table({width:{size:W,type:WidthType.DXA},columnWidths:[cN,cNome,cAss,cObs],layout:TableLayoutType.FIXED,rows:[head,...linhas]});
const img=(f,w,h)=>new ImageRun({type:'png',data:fs.readFileSync(f),transformation:{width:w,height:h}});
const logos=new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[img('logo-sme.png',78,70),t('          '),img('logo-crianca-alfabetizada.png',110,43)]});
const nb={style:BorderStyle.NONE,size:0,color:'FFFFFF'}; const nbs={top:nb,bottom:nb,left:nb,right:nb};
const ac=(nome,cargo)=>new TableCell({borders:nbs,width:{size:W/2,type:WidthType.DXA},children:[
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:600},children:[t('_________________________________')]}),
  new Paragraph({alignment:AlignmentType.CENTER,children:[t(nome,{bold:true,size:18})]}),
  new Paragraph({alignment:AlignmentType.CENTER,children:[t(cargo,{size:18})]})]});
const assinaturas=new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W/2,W/2],layout:TableLayoutType.FIXED,
  rows:[new TableRow({cantSplit:true,children:[ac('ANTONIO CARLOS PEREIRA','Articulador Municipal'),ac('ALESSANDRA DE SOUZA F. RODRIGUES','Diretora de Ensino')]})]});
const doc=new Document({styles:{default:{document:{run:{font:F,size:20}}}},sections:[{
  properties:{page:{size:{width:11906,height:16838},margin:{top:720,bottom:720,left:720,right:720}}},
  children:[logos,
   new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:160},children:[t(D.titulo,{bold:true,size:26,color:NAVY})]}),
   cab,new Paragraph({spacing:{after:120},children:[]}),lista,
   assinaturas]}]});
Packer.toBuffer(doc).then(buf=>fs.writeFileSync(out,buf));
