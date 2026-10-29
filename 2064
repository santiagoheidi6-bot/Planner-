// Mi Planner — backend de Google Sheets. Cambia SECRET y despliega como Aplicación web.
const SECRET = 'CAMBIA-ESTA-CLAVE';
const HEAD = {
  'Actividades': ['ID','Fecha','Hora inicio','Hora fin','Actividad','Categoría','Estado','Tipo','Creada','Finalización real','Observaciones','Reprogramaciones','Fecha original','Color','Serie'],
  'Rutinas': ['ID','Actividad','Horario habitual','Días de repetición','Fecha inicio','Fecha fin','Estado','Excepciones'],
  'Historial de cambios': ['ID actividad','Fecha del cambio','Horario anterior','Horario nuevo','Estado anterior','Estado nuevo','Motivo','ID cambio']
};
const AK = ['id','date','start','end','name','cat','status','type','created','doneAt','notes','resched','orig','color','rid'];

function sh_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.getRange(1,1,1,HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function upsert_(s, rows, keyIdx) {
  if (!rows.length) return;
  const n = s.getLastRow();
  const ids = n > 1 ? s.getRange(2, keyIdx + 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
  const pos = {}; ids.forEach((v, i) => pos[v] = i + 2);
  const add = [];
  rows.forEach(r => {
    const k = String(r[keyIdx]);
    if (pos[k]) s.getRange(pos[k], 1, 1, r.length).setNumberFormat('@').setValues([r]);
    else { add.push(r); pos[k] = -1; }
  });
  if (add.length) {
    const start = s.getLastRow() + 1;
    s.getRange(start, 1, add.length, add[0].length).setNumberFormat('@').setValues(add);
  }
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== SECRET) return out_({ ok: false, error: 'Clave incorrecta' });
    if ((d.del || []).length) {
      const s = sh_('Actividades'), n = s.getLastRow();
      if (n > 1) { const ids = s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]));
        for (let i = ids.length - 1; i >= 0; i--) if (d.del.indexOf(ids[i]) >= 0) s.deleteRow(i + 2); }
    }
    upsert_(sh_('Actividades'), (d.acts || []).map(a => AK.map(k => a[k] == null ? '' : a[k])), 0);
    upsert_(sh_('Rutinas'), (d.routines || []).map(r => [r.id, r.name, r.start + '-' + (r.end || ''), (r.days || []).join(','), r.from, r.to, r.status, (r.exc || []).join(',')]), 0);
    upsert_(sh_('Historial de cambios'), (d.hist || []).map(h => [h.id, h.at, h.oldT, h.newT, h.oldS, h.newS, h.reason || '', h.hid]), 7);
    return out_({ ok: true, acts: (d.acts || []).length });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}
function doGet(e) {
  if (!e.parameter || e.parameter.key !== SECRET) return out_({ ok: false, error: 'Clave incorrecta' });
  const rd = n => { const s = sh_(n), r = s.getLastRow(); return r > 1 ? s.getRange(2, 1, r - 1, HEAD[n].length).getDisplayValues() : []; };
  return out_({
    ok: true,
    acts: rd('Actividades').map(r => { const o = {}; AK.forEach((k, i) => o[k] = r[i]); return o; }),
    routines: rd('Rutinas').map(r => { const [a, b] = r[2].split('-'); return { id: r[0], name: r[1], start: a, end: b, days: r[3].split(',').filter(Boolean).map(Number), from: r[4], to: r[5], status: r[6], exc: (r[7] || '').split(',').filter(Boolean) }; }),
    hist: rd('Historial de cambios').map(r => ({ id: r[0], at: r[1], oldT: r[2], newT: r[3], oldS: r[4], newS: r[5], reason: r[6], hid: r[7] }))
  });
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// Ejecuta UNA vez (Ejecutar → setup) para crear las hojas y el Dashboard con fórmulas.
// El Dashboard cubre 24 meses desde el mes actual; vuelve a ejecutarlo si quieres moverlo.
function setup() {
  Object.keys(HEAD).forEach(sh_);
  const ss = SpreadsheetApp.getActive();
  const d = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  d.clear();
  const t = new Date(Date.now() - 5 * 36e5), y = t.getUTCFullYear(), m = t.getUTCMonth(), N = 731, A = 'Actividades!';
  d.getRange('A1:H1').setValues([['Fecha','Computables','Completadas','% cumplimiento','Imprevistas','Reprogramadas','Canceladas','Día productivo']]).setFontWeight('bold');
  const rows = [];
  for (let i = 0; i < N; i++) {
    const r = i + 2, f = new Date(Date.UTC(y, m, 1 + i)).toISOString().slice(0, 10);
    rows.push([f,
      `=COUNTIFS(${A}$B:$B,$A${r},${A}$G:$G,"<>cancelada")`,
      `=COUNTIFS(${A}$B:$B,$A${r},${A}$G:$G,"completada")`,
      `=IF(B${r}=0,"Sin datos",C${r}/B${r})`,
      `=COUNTIFS(${A}$B:$B,$A${r},${A}$H:$H,"imprevista")`,
      `=COUNTIFS(${A}$B:$B,$A${r},${A}$L:$L,">0")`,
      `=COUNTIFS(${A}$B:$B,$A${r},${A}$G:$G,"cancelada")`,
      `=IF(B${r}=0,"Sin datos",IF(C${r}/B${r}>=$P$1,"Productivo","Menos avance"))`]);
  }
  const L = N + 1;
  d.getRange(`A2:A${L}`).setNumberFormat('@');
  d.getRange(2, 1, N, 8).setValues(rows);
  d.getRange('O1:P1').setValues([['Umbral (0.7 = 70%)', 0.7]]);
  d.getRange(`D2:D${L}`).setNumberFormat('0%');
  d.getRange('I1:N1').setValues([['Mes','Computables','Completadas','% cumplimiento','Imprevistas','Días productivos']]).setFontWeight('bold');
  const mf = [];
  for (let i = 0; i < 24; i++) {
    const k = new Date(Date.UTC(y, m + i, 1)).toISOString().slice(0, 7), r = i + 2;
    mf.push([k, `=SUMIFS(B$2:B$${L},$A$2:$A$${L},"${k}*")`, `=SUMIFS(C$2:C$${L},$A$2:$A$${L},"${k}*")`, `=IF(J${r}=0,"Sin datos",K${r}/J${r})`, `=SUMIFS(E$2:E$${L},$A$2:$A$${L},"${k}*")`, `=COUNTIFS($A$2:$A$${L},"${k}*",$H$2:$H$${L},"Productivo")`]);
  }
  d.getRange(2, 9, 24, 6).setFormulas(mf);
  d.getRange('I26:N26').setFormulas([['TOTAL', `=SUM(B2:B${L})`, `=SUM(C2:C${L})`, '=IF(J26=0,"Sin datos",K26/J26)', `=SUM(E2:E${L})`, `=COUNTIF(H2:H${L},"Productivo")`]]).setFontWeight('bold');
  d.getRange('L2:L26').setNumberFormat('0%');
}
