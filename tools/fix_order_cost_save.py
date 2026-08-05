from pathlib import Path
import re

p = Path('App.html')
s = p.read_text(encoding='utf-8')


def replace_function(text, name, new):
    m = re.search(r'function ' + re.escape(name) + r'\([^)]*\)\{', text)
    if not m:
        raise SystemExit(name + ' not found')
    i = m.end()
    depth = 1
    quote = None
    escaped = False
    while i < len(text) and depth:
        c = text[i]
        if quote:
            if escaped:
                escaped = False
            elif c == '\\':
                escaped = True
            elif c == quote:
                quote = None
        else:
            if c in "'\"`":
                quote = c
            elif c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
        i += 1
    return text[:m.start()] + new + text[i:]

capture = r'''function captureOrderCostsFromScreen(){
  const f=$('#orderForm');
  const currentItem=findMainInventory(f.producto.value),qty=num(f.cantidad.value)||1;
  mainMaterial=currentItem?{id:currentItem.id,cantidad:qty}:null;
  inkType=$('#inkType')?$('#inkType').value:'Sin tinta';
  inkCost=$('#inkCost')?num($('#inkCost').value):0;
  materials=[...document.querySelectorAll('#materialRows .material-row')].map(row=>{
    const select=row.querySelector('select'),input=row.querySelector('input[type="number"]');
    return{id:select?select.value:'',cantidad:input?num(input.value):0}
  }).filter(x=>x.id&&x.cantidad>0);
  otherCosts=[...document.querySelectorAll('#otherCostRows .other-row')].map(row=>{
    const inputs=row.querySelectorAll('input');
    return{nombre:inputs[0]?String(inputs[0].value||'').trim():'',costo:inputs[1]?num(inputs[1].value):0}
  }).filter(x=>x.nombre||x.costo>0);
  recalculateOrder();
  return{mainMaterial,materials,otherCosts,inkType,inkCost}
}'''

if 'function captureOrderCostsFromScreen()' not in s:
    s = s.replace('function saveOrderClient(e)', capture + '\nfunction saveOrderClient(e)', 1)

new_save = r'''function saveOrderClient(e){
  e.preventDefault();
  const f=e.target,b=$('#saveOrderBtn'),start=Date.now(),backup=JSON.parse(JSON.stringify(data));
  captureOrderCostsFromScreen();
  const o=Object.fromEntries(new FormData(f).entries()),isEdit=!!o.id;
  o._mode=isEdit?'edit':'create';
  o.clientKey=isEdit?'':'CLIENT-'+Date.now()+'-'+Math.random().toString(36).slice(2,10);
  const all=mainMaterial?[mainMaterial,...materials]:[...materials],costs=[...otherCosts];
  if(inkType!=='Sin tinta'||num(inkCost)>0)costs.push({nombre:'Tinta '+inkType,costo:num(inkCost),tipoCosto:'tinta',tintaTipo:inkType});
  o.materialesJson=JSON.stringify(all);
  o.otrosCostosJson=JSON.stringify(costs);
  const calculatedCost=all.reduce((sum,x)=>sum+fifoItemCost(x.id,x.cantidad),0)+costs.reduce((sum,x)=>sum+num(x.costo),0);
  const localOrder={...o,costoTotal:calculatedCost,ganancia:num(o.pagado)-calculatedCost,materialesJson:o.materialesJson,otrosCostosJson:o.otrosCostosJson};
  if(isEdit){const i=data.orders.findIndex(x=>x.id===o.id);if(i>=0)data.orders[i]={...data.orders[i],...localOrder}}
  setSpin(b,true);
  setTimeout(()=>{closeModal('orderModal');renderAll()},500);
  (async()=>{
    try{
      const old=String(o.foto||'').split('|||').filter(Boolean),fresh=[];
      for(const file of pendingFiles)fresh.push(await upload(file,o.id||o.clientKey,'orden'));
      o.foto=[...old,...fresh].join('|||');
      google.script.run.withSuccessHandler(async d=>{await waitMin(start,500);data=d;pendingFiles=[];resetButton(b);renderAll()}).withFailureHandler(async x=>{await waitMin(start,500);data=backup;resetButton(b);renderAll();showModal('orderModal');showError('orderError',x)}).saveOrder(o)
    }catch(x){await waitMin(start,500);data=backup;resetButton(b);renderAll();showModal('orderModal');showError('orderError',x)}
  })()
}'''

s = replace_function(s, 'saveOrderClient', new_save)
p.write_text(s, encoding='utf-8')
