const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'google_apps_script.js'), 'utf8');
const secret = 'only-a-unit-test-password';
function harness() {
  const rows = [], notes = [], requests = [];
  const props = { ADMIN_SECRET: secret, META_CAPI_ACCESS_TOKEN: 'unit-test-token' };
  let locked = false, busy = false, response = { code: 200, body: { events_received: 1 } };
  const sheet = {
    getLastRow: () => rows.length + 1,
    appendRow: row => rows.push([...row]),
    getRange(row, col, height = 1, width = 1) {
      return {
        getValues: () => Array.from({length: height}, (_, i) => (rows[row - 2 + i] || []).slice(col - 1, col - 1 + width)),
        getNotes: () => Array.from({length: height}, (_, i) => [notes[row - 2 + i] || '']),
        setValue(value) { rows[row - 2][col - 1] = value; },
        setValues(values) { values.forEach((r, i) => r.forEach((v, j) => rows[row - 2 + i][col - 1 + j] = v)); },
        setNote(value) { notes[row - 2] = value; }
      };
    }
  };
  const context = vm.createContext({
    Date, JSON, Math,
    PropertiesService: {getScriptProperties: () => ({
      getProperty: k => props[k] || null, setProperty: (k, v) => {props[k] = v;}
    })},
    LockService: {getScriptLock: () => ({
      tryLock: () => {if (busy) return false; locked = true; return true;},
      hasLock: () => locked, releaseLock: () => {locked = false;}
    })},
    SpreadsheetApp: {flush() {}},
    Utilities: {DigestAlgorithm: {SHA_256: 'sha256'}, Charset: {UTF_8:'utf8'},
      computeDigest: (_, text) => Array.from(crypto.createHash('sha256').update(text).digest())},
    ContentService: {MimeType:{JSON:'json'}, createTextOutput: text => ({setMimeType: () => JSON.parse(text)})},
    UrlFetchApp: {fetch(url, options) {
      requests.push({url, options, payload: JSON.parse(options.payload)});
      if (response.throws) throw new Error('connection interrupted');
      return {getResponseCode: () => response.code, getContentText: () => JSON.stringify(response.body)};
    }}
  });
  vm.runInContext(source, context);
  context.setupSheet = () => {};
  context.getTargetSpreadsheet = () => ({getSheetByName: () => sheet});
  const post = (action, data = {}, authenticated = true) =>
    context.doPost({postData:{contents:JSON.stringify({action, ...data, ...(authenticated ? {admin_secret:secret} : {})})}});
  const order = {id:'VL-test-unique-0001', name:'Prueba Integracion', phone:'900000001', dni:'00000001',
    city:'Lima', address:'Agencia de prueba', units:2, price:1,
    attribution:{fbp:'fb.1.1234567890000.12345', client_user_agent:'Buyer test agent', landing_url:'https://velorabeautype.store/'}};
  return {context, post, order, rows, props, requests, notes,
    create: () => post('create_order', order, false),
    setResponse: v => {response = v;}, setBusy: () => {busy = true;}, locked: () => locked};
}
test('scripts del frontend y backend tienen sintaxis válida', () => {
  new vm.Script(source);
  new vm.Script(fs.readFileSync(path.join(root,'api-client.js'),'utf8'));
  for (const file of ['index.html','admin.html']) {
    const html = fs.readFileSync(path.join(root,file),'utf8');
    for (const match of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) new vm.Script(match[1], {filename:file});
    assert.ok(html.includes('src="api-client.js"'));
  }
  assert.ok(!source.includes('EAAd'));
  assert.ok(source.includes('GRAPH_API_VERSION = "v26.0"'));
});
test('GET no revela pedidos y POST privado exige contraseña', () => {
  const h = harness(); h.create();
  assert.equal(h.context.doGet({parameter:{action:'get_orders'}}).orders, undefined);
  for (const action of ['get_orders','mark_paid','retry_capi','set_config','get_config_status','change_password']) {
    assert.equal(h.post(action, {id:h.order.id}, false).code,'unauthorized');
  }
  assert.equal(h.requests.length,0);
});
test('pedido guardado con precio del catálogo, atribución y sin Purchase', () => {
  const h=harness(), r=h.create();
  assert.equal(r.saved,true);
  assert.equal(h.rows[0][9],169.90);
  assert.equal(h.rows[0][12],'Pendiente');
  assert.equal(JSON.parse(h.rows[0][16]).client_user_agent,'Buyer test agent');
  assert.equal(h.requests.length,0);
});
test('reintentar el mismo pedido no agrega otra fila; colisión se rechaza', () => {
  const h=harness(); h.create(); assert.equal(h.create().duplicate,true);
  assert.equal(h.rows.length,1);
  assert.equal(h.post('create_order',{...h.order,phone:'900000002'},false).status,'error');
});
test('datos inválidos se rechazan en servidor', () => {
  const h=harness();
  for(const override of [{units:4},{phone:'123'},{dni:'123'},{name:'Uno'},{id:'invalid'},{city:''}]) {
    assert.equal(h.post('create_order',{...h.order,...override},false).status,'error');
  }
  assert.equal(h.rows.length,0);
});
test('cambiar logística no dispara compras y retry exige pago real', () => {
  const h=harness(); h.create();
  assert.equal(h.post('update_status',{id:h.order.id,status:'Confirmado'}).status,'success');
  assert.equal(h.post('retry_capi',{id:h.order.id}).status,'error');
  assert.equal(h.requests.length,0);
});
test('pago dispara Purchase con PEN, valor exacto, teléfono hash y sin DNI', () => {
  const h=harness(); h.create();
  const r=h.post('mark_paid',{id:h.order.id});
  assert.equal(r.capi_status,'sent');
  const request=h.requests[0], event=request.payload.data[0];
  assert.equal(event.event_name,'Purchase');
  assert.equal(event.event_id,'purchase_'+h.order.id);
  assert.equal(event.custom_data.value,169.90);
  assert.equal(event.custom_data.currency,'PEN');
  assert.equal(event.action_source,'chat');
  assert.match(event.user_data.ph[0],/^[a-f0-9]{64}$/);
  assert.equal(event.user_data.client_user_agent,'Buyer test agent');
  assert.equal(event.user_data.dni,undefined);
  assert.equal(event.event_time,Math.floor(new Date(r.paid_at).getTime()/1000));
  assert.ok(!request.url.includes('unit-test-token'));
  assert.equal(h.locked(),false);
});
test('pago repetido y retry de evento aceptado no vuelven a enviar', () => {
  const h=harness();h.create();h.post('mark_paid',{id:h.order.id});
  assert.equal(h.post('mark_paid',{id:h.order.id}).already_sent,true);
  assert.equal(h.post('retry_capi',{id:h.order.id}).already_sent,true);
  assert.equal(h.requests.length,1);
});
test('fallo Meta conserva pago y retry conserva event_id y fecha originales', () => {
  const h=harness();h.create();h.setResponse({code:400,body:{error:{message:'Token inválido'}}});
  const failed=h.post('mark_paid',{id:h.order.id});
  assert.equal(failed.payment_status,'Pagado');assert.equal(failed.capi_status,'failed');
  assert.match(h.notes[0],/Token inválido/);
  h.setResponse({code:200,body:{events_received:1}});
  const retry=h.post('retry_capi',{id:h.order.id});
  assert.equal(retry.capi_status,'sent');
  assert.equal(retry.paid_at,failed.paid_at);
  assert.equal(retry.capi_event_id,failed.capi_event_id);
});
test('200 sin events_received y fallo de red no son éxito', () => {
  for(const response of [{code:200,body:{}},{code:200,body:{events_received:0}},{throws:true}]) {
    const h=harness();h.create();h.setResponse(response);
    assert.equal(h.post('mark_paid',{id:h.order.id}).capi_status,'failed');
  }
});
test('modo de prueba no marca sent y se puede enviar luego a producción', () => {
  const h=harness();h.create();h.props.META_TEST_EVENT_CODE='TEST_ONLY';
  assert.equal(h.post('mark_paid',{id:h.order.id}).capi_status,'test_sent');
  assert.equal(h.requests[0].payload.test_event_code,'TEST_ONLY');
  h.props.META_TEST_EVENT_CODE='';
  assert.equal(h.post('retry_capi',{id:h.order.id}).capi_status,'sent');
});
test('falta de token, Pixel distinto, fecha inválida y fecha antigua', () => {
  const h=harness();h.create();delete h.props.META_CAPI_ACCESS_TOKEN;
  assert.equal(h.post('mark_paid',{id:h.order.id}).capi_status,'failed');
  h.props.META_CAPI_ACCESS_TOKEN='unit-test-token';
  h.rows[0][13]='invalid';
  assert.equal(h.post('retry_capi',{id:h.order.id}).capi_status,'failed');
  h.rows[0][13]='2020-01-01T00:00:00Z';
  assert.equal(h.post('retry_capi',{id:h.order.id}).capi_status,'failed');
  h.props.META_PIXEL_ID='different';
  assert.equal(h.post('retry_capi',{id:h.order.id}).status,'error');
  assert.equal(h.requests.length,0);
});
test('fecha antigua de Lima se interpreta con -05:00', () => {
  const h=harness();
  assert.equal(h.context.normalizePaidAt('2026-09-05 10:30'),'2026-09-05T10:30-05:00');
});
test('lock ocupado no cambia estado ni llama a Meta', () => {
  const h=harness();h.create();h.setBusy();
  assert.equal(h.post('mark_paid',{id:h.order.id}).status,'error');
  assert.equal(h.rows[0][12],'Pendiente');assert.equal(h.requests.length,0);
});
test('contraseña se cambia en servidor y la anterior deja de funcionar', () => {
  const h=harness();
  assert.equal(h.post('login',{admin_secret:secret.toUpperCase()},false).code,'unauthorized');
  assert.equal(h.post('change_password',{new_password:'another-long-unit-test-secret'}).status,'success');
  assert.equal(h.post('login').code,'unauthorized');
});
test('cliente rechaza backend antiguo y respuesta opaca en lugar de simular éxito', async () => {
  const js=fs.readFileSync(path.join(root,'api-client.js'),'utf8');
  let body={status:'success',orders:[]};
  const context=vm.createContext({window:{},localStorage:{getItem:()=>null},crypto,
    AbortController,setTimeout,clearTimeout,
    fetch:async()=>({ok:true,text:async()=>JSON.stringify(body)})});
  vm.runInContext(js,context);
  await assert.rejects(context.window.VeloraAPI.request('mark_paid',{}),/desactualizado/);
  body={status:'success',api_version:'2026-09-05.2',id:'VL-test',saved:true};
  assert.equal((await context.window.VeloraAPI.request('create_order',{})).saved,true);
  body={status:'error',api_version:'2026-09-05.2',message:'No autorizado'};
  await assert.rejects(context.window.VeloraAPI.request('get_orders',{}),/No autorizado/);
});
