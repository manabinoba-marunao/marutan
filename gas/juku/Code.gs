/**
 * マル単 塾ネット(他塾向け)— Google Apps Script
 * ------------------------------------------------------------
 * 設置手順(最初の1回だけ)
 *  1. 新しいスプレッドシートを作る(名前の例:マル単 塾ネット)
 *  2. 拡張機能 → Apps Script を開き、このファイルの中身をぜんぶ貼りつけて保存
 *  3. 上の関数の選択で「setup」を選んで ▶実行(権限の確認が出たら許可)
 *     → シート「塾」「生徒」「記録」「まちがい」「設定」ができる
 *     → 「設定」シートの B1 が管理者(直樹さん)の合言葉。好きな言葉に変えてOK
 *  4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       次のユーザーとして実行:自分 / アクセスできるユーザー:全員
 *     → 出てきたURLを index.html と juku-admin.html の JUKU_GAS_URL に貼る
 *  ※ コードを直したときは「デプロイを管理 → 編集 → 新バージョン」でURLを変えずに更新
 *
 * しくみ
 *  - 直樹さん(管理者合言葉)… すべての塾を見られる・塾を発行できる
 *  - 塾長(塾コード+塾長パスワード)… 自分の塾の生徒の登録と成績だけ
 *  - 生徒(塾コード+名前+暗証番号4けた)… マル単を使う。記録は自分の塾にだけ入る
 */

const SH = {juku: "塾", stu: "生徒", log: "記録", wrong: "まちがい", conf: "設定"};
const HEAD = {
  juku:  ["塾コード", "塾名", "塾長パスワード", "状態", "発行日", "メモ"],
  stu:   ["塾コード", "生徒ID", "名前", "学年", "暗証番号", "状態", "登録日", "マルの数", "最後に使った日"],
  log:   ["日時", "塾コード", "生徒ID", "名前", "セット", "形式", "問題数/時間", "解いた", "正解", "正答率", "秒", "まちがえた語", "記録ID"],
  wrong: ["日時", "塾コード", "生徒ID", "名前", "英語", "意味", "セット"]
};
const ACTIVE = "有効", STOPPED = "停止", LEFT = "退会";

/* ========== 最初に1回だけ実行 ========== */
function setup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEAD).forEach(k => {
    let sh = ss.getSheetByName(SH[k]);
    if(!sh) sh = ss.insertSheet(SH[k]);
    if(sh.getLastRow() === 0){ sh.appendRow(HEAD[k]); sh.setFrozenRows(1); sh.getRange(1,1,1,HEAD[k].length).setFontWeight("bold"); }
  });
  // 塾コード・生徒ID・暗証番号は「001」などの先頭0が消えないよう文字列に
  ss.getSheetByName(SH.juku).getRange("A:A").setNumberFormat("@");
  ss.getSheetByName(SH.stu).getRange("A:B").setNumberFormat("@");
  ss.getSheetByName(SH.stu).getRange("E:E").setNumberFormat("@");
  let conf = ss.getSheetByName(SH.conf);
  if(!conf){
    conf = ss.insertSheet(SH.conf);
    conf.getRange("A1:B1").setValues([["管理者の合言葉", randStr(10)]]);
    conf.getRange("A2:B2").setValues([["(説明)", "B1が直樹さん用の合言葉。juku-admin.html の「管理者」から入るときに使う"]]);
  }
  const p = PropertiesService.getScriptProperties();
  if(!p.getProperty("SECRET")) p.setProperty("SECRET", randStr(32));
}

/* ========== 入口 ========== */
function doGet(e){ return out(route((e && e.parameter) || {})); }
function doPost(e){
  let b = {};
  try{ b = JSON.parse((e && e.postData && e.postData.contents) || "{}"); }catch(err){ return out({ok:false, error:"bad-json"}); }
  return out(route(b));
}
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function route(p){
  try{
    const a = String(p.action || "");
    switch(a){
      // ---- 生徒 ----
      case "jukuInfo":  return jukuInfo(p);
      case "login":     return studentLogin(p);
      case "log":       return saveLog(p);
      case "ranking":   return ranking(p);
      // ---- マル単GP ----
      case "gpCreate":  return gpCreate(p);
      case "gpJoin":    return gpJoin(p);
      case "gpState":   return gpState(p);
      case "gpStart":   return gpStart(p);
      case "gpEnd":     return gpEnd(p);
      case "gpScore":   return gpScore(p);
      case "gpHistory": return gpHistory(p);
      // ---- 塾長・管理者 ----
      case "auth":      return authInfo(p);
      case "students":  return listStudents(p);
      case "addStudents": return addStudents(p);
      case "updateStudent": return updateStudent(p);
      case "report":    return report(p);
      // ---- 管理者だけ ----
      case "jukuList":  return jukuList(p);
      case "addJuku":   return addJuku(p);
      case "updateJuku": return updateJuku(p);
    }
    return {ok:false, error:"unknown-action"};
  }catch(err){
    return {ok:false, error:String(err && err.message || err)};
  }
}

/* ========== 共通 ========== */
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(k){ return ss_().getSheetByName(SH[k]); }
function rows(k){
  const sh = sheet(k); const n = sh.getLastRow();
  if(n < 2) return [];
  return sh.getRange(2, 1, n-1, HEAD[k].length).getValues();
}
function norm(s){ return String(s == null ? "" : s).normalize("NFKC").replace(/\s+/g, "").toLowerCase(); }
function normCode(s){ return String(s == null ? "" : s).normalize("NFKC").replace(/\s+/g, "").toUpperCase(); }
function randStr(n){
  const c = "abcdefghjkmnpqrstuvwxyz23456789"; let s = "";
  for(let i=0;i<n;i++) s += c.charAt(Math.floor(Math.random()*c.length));
  return s;
}
function randPin(){ return String(Math.floor(Math.random()*10000)).padStart(4, "0"); }
function today(){ return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"); }
function adminKey(){ return String(sheet("conf").getRange("B1").getValue() || "").trim(); }
function isAdmin(p){ const k = adminKey(); return !!k && String(p.admin || "") === k; }
function secret(){ return PropertiesService.getScriptProperties().getProperty("SECRET") || "marutan"; }
function tokenFor(code, sid, pin){
  const sig = Utilities.computeHmacSha256Signature(code + "|" + sid + "|" + pin, secret());
  return Utilities.base64EncodeWebSafe(sig).slice(0, 22);
}
function findJuku(code){
  const c = normCode(code); if(!c) return null;
  const r = rows("juku");
  for(let i=0;i<r.length;i++) if(normCode(r[i][0]) === c)
    return {row:i+2, code:String(r[i][0]), name:String(r[i][1]), pass:String(r[i][2]), status:String(r[i][3]||ACTIVE)};
  return null;
}
function studentObj(r, i){
  return {row:i+2, code:String(r[0]), sid:String(r[1]), name:String(r[2]), grade:String(r[3]||""), pin:String(r[4]).padStart(4,"0"),
          status:String(r[5]||ACTIVE), joined:fmtDay(r[6]), maru:+r[7]||0, last:fmtDay(r[8])};
}
function fmtDay(v){ if(!v) return ""; if(v instanceof Date) return Utilities.formatDate(v, "Asia/Tokyo", "yyyy-MM-dd"); return String(v); }
function studentsOf(code){
  const c = normCode(code);
  return rows("stu").map(studentObj).filter(s => normCode(s.code) === c);
}
/* 生徒の確かめ(記録・ランキング用):暗証番号から作ったトークンで照合 */
function checkStudent(p){
  const j = findJuku(p.code); if(!j || j.status !== ACTIVE) return null;
  const s = studentsOf(j.code).find(x => x.sid === String(p.sid));
  if(!s || s.status !== ACTIVE) return null;
  if(String(p.token || "") !== tokenFor(j.code, s.sid, s.pin)) return null;
  return {juku:j, stu:s};
}
/* 塾長または管理者か。塾長なら自分の塾だけ */
function teacherJuku(p){
  const j = findJuku(p.code); if(!j) return null;
  if(isAdmin(p)) return j;
  if(j.status !== ACTIVE) return null;
  if(String(p.pass || "") && String(p.pass) === j.pass) return j;
  return null;
}

/* ========== 生徒 ========== */
function jukuInfo(p){
  const j = findJuku(p.code);
  if(!j || j.status !== ACTIVE) return {ok:false, error:"code"};
  return {ok:true, code:j.code, name:j.name};
}

function studentLogin(p){
  const j = findJuku(p.code);
  if(!j || j.status !== ACTIVE) return {ok:false, error:"code"};
  const cache = CacheService.getScriptCache();
  const lockKey = "fail:" + j.code + ":" + norm(p.name);
  const fails = +(cache.get(lockKey) || 0);
  if(fails >= 8) return {ok:false, error:"locked"};
  const s = studentsOf(j.code).find(x => x.status === ACTIVE && norm(x.name) === norm(p.name));
  if(!s || s.pin !== String(p.pin || "")){
    cache.put(lockKey, String(fails+1), 600);   // 8回まちがえたら10分お休み
    return {ok:false, error: s ? "pin" : "name"};
  }
  cache.remove(lockKey);
  return {ok:true, code:j.code, jukuName:j.name, sid:s.sid, name:s.name, token:tokenFor(j.code, s.sid, s.pin)};
}

function saveLog(p){
  const who = checkStudent(p); if(!who) return {ok:false, error:"auth"};
  let e = p.entry;
  if(typeof e === "string"){ try{ e = JSON.parse(e); }catch(err){ return {ok:false, error:"bad-entry"}; } }
  if(!e || !e.id) return {ok:false, error:"bad-entry"};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const lg = sheet("log");
    // 同じ記録を二重に入れない(最近の300件を確認)
    const n = lg.getLastRow();
    if(n > 1){
      const from = Math.max(2, n-299);
      const ids = lg.getRange(from, 13, n-from+1, 1).getValues();
      if(ids.some(r => String(r[0]) === String(e.id))) return {ok:true, dup:true};
    }
    const ts = +e.ts || Date.now();
    const answered = +e.answered || 0, correct = +e.correct || 0;
    const wrong = Array.isArray(e.wrong) ? e.wrong.slice(0, 200) : [];
    const S = who.stu, code = who.juku.code;
    lg.appendRow([new Date(ts), code, S.sid, S.name, String(e.label||""), String(e.mode||""), String(e.limit||""),
      answered, correct, answered ? Math.round(correct/answered*100) : 0, +e.seconds||0,
      wrong.map(w => w.en).join(", "), String(e.id)]);
    if(wrong.length){
      const W = wrong.map(w => [new Date(ts), code, S.sid, S.name, String(w.en||""), String(w.ja||""), String(w.set||"")]);
      const wg = sheet("wrong");
      wg.getRange(wg.getLastRow()+1, 1, W.length, W[0].length).setValues(W);
    }
    const st = sheet("stu");
    st.getRange(S.row, 9).setValue(today());
    if(e.maru != null && isFinite(+e.maru)) st.getRange(S.row, 8).setValue(+e.maru);
  } finally { lock.releaseLock(); }
  return {ok:true};
}

/* 塾内ランキング:今週(月曜はじまり)と今月の正解数 */
function ranking(p){
  const who = checkStudent(p); if(!who) return {ok:false, error:"auth"};
  const code = who.juku.code;
  const now = new Date();
  const jst = new Date(now.getTime() + 9*3600e3);
  const dow = (jst.getUTCDay() + 6) % 7;
  const weekStart = Date.UTC(jst.getUTCFullYear(), jst.getUTCMonth(), jst.getUTCDate() - dow) - 9*3600e3;
  const monthStart = Date.UTC(jst.getUTCFullYear(), jst.getUTCMonth(), 1) - 9*3600e3;
  const stus = studentsOf(code).filter(s => s.status === ACTIVE);
  const names = {}; stus.forEach(s => names[s.sid] = s.name);
  const agg = {week:{}, month:{}};
  rows("log").forEach(r => {
    if(normCode(r[1]) !== normCode(code)) return;
    const t = r[0] instanceof Date ? r[0].getTime() : +new Date(r[0]);
    if(t < monthStart) return;
    const sid = String(r[2]); if(!(sid in names)) return;
    const add = (o) => { if(!o[sid]) o[sid] = {sid, name:names[sid], correct:0, answered:0, times:0}; const x = o[sid]; x.correct += +r[8]||0; x.answered += +r[7]||0; x.times++; };
    add(agg.month); if(t >= weekStart) add(agg.week);
  });
  const sortTop = o => Object.values(o).sort((a,b) => b.correct - a.correct || b.answered - a.answered);
  const maru = stus.map(s => ({sid:s.sid, name:s.name, maru:s.maru})).filter(x => x.maru > 0).sort((a,b) => b.maru - a.maru);
  return {ok:true, jukuName:who.juku.name, me:who.stu.sid, week:sortTop(agg.week), month:sortTop(agg.month), maru};
}

/* ========== 塾長・管理者 ========== */
function authInfo(p){
  if(isAdmin(p) && !p.code) return {ok:true, admin:true};
  const j = teacherJuku(p); if(!j) return {ok:false, error:"auth"};
  return {ok:true, admin:isAdmin(p), code:j.code, name:j.name, status:j.status};
}

function listStudents(p){
  const j = teacherJuku(p); if(!j) return {ok:false, error:"auth"};
  return {ok:true, code:j.code, name:j.name, students:studentsOf(j.code).map(s => { delete s.row; return s; })};
}

/* names: 改行区切りの名前(「名前,学年」でもOK) */
function addStudents(p){
  const j = teacherJuku(p); if(!j) return {ok:false, error:"auth"};
  const lines = String(p.names || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if(!lines.length) return {ok:false, error:"empty"};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const cur = studentsOf(j.code);
    const used = new Set(cur.filter(s => s.status !== LEFT).map(s => norm(s.name)));
    let maxNo = 0; cur.forEach(s => { const m = s.sid.match(/(\d+)$/); if(m) maxNo = Math.max(maxNo, +m[1]); });
    const added = [], skipped = [];
    const R = [];
    lines.forEach(l => {
      const [nm, gr] = l.split(/[,，、\t]/).map(x => (x||"").trim());
      if(!nm){ return; }
      if(nm.length > 20 || used.has(norm(nm))){ skipped.push(nm); return; }
      used.add(norm(nm)); maxNo++;
      const sid = j.code + "-" + String(maxNo).padStart(3, "0"), pin = randPin();
      R.push([j.code, sid, nm, gr || String(p.grade || ""), pin, ACTIVE, today(), 0, ""]);
      added.push({sid, name:nm, pin, grade:gr || String(p.grade || "")});
    });
    if(R.length){ const st = sheet("stu"); st.getRange(st.getLastRow()+1, 1, R.length, R[0].length).setValues(R); }
    return {ok:true, added, skipped};
  } finally { lock.releaseLock(); }
}

/* name / grade / status(有効・退会)/ resetPin:true */
function updateStudent(p){
  const j = teacherJuku(p); if(!j) return {ok:false, error:"auth"};
  const s = studentsOf(j.code).find(x => x.sid === String(p.sid));
  if(!s) return {ok:false, error:"not-found"};
  const st = sheet("stu");
  if(p.name != null && String(p.name).trim() && norm(p.name) !== norm(s.name)){
    const nm = String(p.name).trim();
    if(studentsOf(j.code).some(x => x.sid !== s.sid && x.status !== LEFT && norm(x.name) === norm(nm))) return {ok:false, error:"dup-name"};
    st.getRange(s.row, 3).setValue(nm);
  }
  if(p.grade != null) st.getRange(s.row, 4).setValue(String(p.grade));
  if(p.status === ACTIVE || p.status === LEFT) st.getRange(s.row, 6).setValue(p.status);
  let pin = s.pin;
  if(p.resetPin){ pin = randPin(); st.getRange(s.row, 5).setValue(pin); }
  return {ok:true, pin};
}

/* teacher.html と同じ形で返す
   sessions: [ts,id,name,label,mode,limit,answered,correct,rate,seconds,wrongText]
   wrongs:   [ts,id,name,en,ja,set] */
function report(p){
  const j = teacherJuku(p); if(!j) return {ok:false, error:"auth"};
  const days = +p.days || 0;
  const since = days ? Date.now() - days*864e5 : 0;
  const c = normCode(j.code);
  const tms = v => v instanceof Date ? v.getTime() : +new Date(v);
  const sessions = rows("log").filter(r => normCode(r[1]) === c && tms(r[0]) >= since)
    .map(r => [tms(r[0]), String(r[2]), String(r[3]), r[4], r[5], r[6], +r[7]||0, +r[8]||0, +r[9]||0, +r[10]||0, r[11]]);
  const wrongs = rows("wrong").filter(r => normCode(r[1]) === c && tms(r[0]) >= since)
    .map(r => [tms(r[0]), String(r[2]), String(r[3]), r[4], r[5], r[6]]);
  return {ok:true, code:j.code, name:j.name, sessions, wrongs};
}

/* ========== 管理者だけ ========== */
function jukuList(p){
  if(!isAdmin(p)) return {ok:false, error:"auth"};
  const days = +p.days || 7, since = Date.now() - days*864e5;
  const st = rows("stu"), lg = rows("log");
  const list = rows("juku").map(r => {
    const code = String(r[0]), c = normCode(code);
    const ss = st.filter(x => normCode(x[0]) === c);
    const L = lg.filter(x => normCode(x[1]) === c);
    const recent = L.filter(x => (x[0] instanceof Date ? x[0].getTime() : +new Date(x[0])) >= since);
    const ans = recent.reduce((a,x) => a + (+x[7]||0), 0), cor = recent.reduce((a,x) => a + (+x[8]||0), 0);
    const lastT = L.reduce((m,x) => Math.max(m, x[0] instanceof Date ? x[0].getTime() : +new Date(x[0])), 0);
    return {code, name:String(r[1]), pass:String(r[2]), status:String(r[3]||ACTIVE), created:fmtDay(r[4]), memo:String(r[5]||""),
            students:ss.filter(x => String(x[5]||ACTIVE) === ACTIVE).length,
            activeStudents:new Set(recent.map(x => String(x[2]))).size,
            times:recent.length, answered:ans, correct:cor, last:lastT || 0};
  });
  return {ok:true, days, juku:list};
}

function addJuku(p){
  if(!isAdmin(p)) return {ok:false, error:"auth"};
  const name = String(p.name || "").trim(); if(!name) return {ok:false, error:"empty"};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    let code = normCode(p.code || "");
    if(code){
      if(!/^[A-Z0-9]{3,10}$/.test(code)) return {ok:false, error:"code-format"};
      if(findJuku(code)) return {ok:false, error:"code-used"};
    } else {
      const C = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
      do { code = ""; for(let i=0;i<5;i++) code += C.charAt(Math.floor(Math.random()*C.length)); } while(findJuku(code));
    }
    const pass = randStr(8);
    sheet("juku").appendRow([code, name, pass, ACTIVE, today(), String(p.memo || "")]);
    return {ok:true, code, name, pass};
  } finally { lock.releaseLock(); }
}

/* name / status(有効・停止)/ resetPass:true / memo */
function updateJuku(p){
  if(!isAdmin(p)) return {ok:false, error:"auth"};
  const j = findJuku(p.code); if(!j) return {ok:false, error:"not-found"};
  const sh = sheet("juku");
  if(p.name != null && String(p.name).trim()) sh.getRange(j.row, 2).setValue(String(p.name).trim());
  if(p.status === ACTIVE || p.status === STOPPED) sh.getRange(j.row, 4).setValue(p.status);
  if(p.memo != null) sh.getRange(j.row, 6).setValue(String(p.memo));
  let pass = j.pass;
  if(p.resetPass){ pass = randStr(8); sh.getRange(j.row, 3).setValue(pass); }
  return {ok:true, pass};
}

/* ========== マル単GP(みんなで同時に競う部屋)==========
   部屋と点数はキャッシュに置く(6時間で自動で消える)。シートは使わない。
   マル直の生徒も他塾の生徒も、この塾ネットのGASで部屋を共有する。 */
const GP_TTL = 21600;
function gpC_(){ return CacheService.getScriptCache(); }
function gpGet_(code){ const s = gpC_().get("gp:" + code); return s ? JSON.parse(s) : null; }
function gpPut_(r){ gpC_().put("gp:" + r.code, JSON.stringify(r), GP_TTL); }
function gpPub_(r){ const o = JSON.parse(JSON.stringify(r)); delete o.host; return o; }
function gpLock_(fn){ const lock = LockService.getScriptLock(); lock.waitLock(15000); try{ return fn(); } finally { lock.releaseLock(); } }
function gpCreate(p){
  return gpLock_(() => {
    let code = "";
    for(let i = 0; i < 40; i++){ code = String(1000 + Math.floor(Math.random() * 9000)); if(!gpGet_(code)) break; }
    const r = {code, host: randStr(16), org: String(p.org || "").slice(0, 40), title: String(p.title || "").slice(0, 80), mode: p.mode === "free" ? "free" : "same",
      sets: (Array.isArray(p.sets) ? p.sets : []).slice(0, 300).map(String),
      dir: ["en2ja", "ja2en", "mix"].indexOf(p.dir) >= 0 ? p.dir : "en2ja",
      sec: Math.max(30, Math.min(1200, Number(p.sec) || 300)),
      seed: Math.floor(Math.random() * 2147483647), status: "lobby", players: [], created: Date.now(), startAt: 0, endAt: 0};
    gpPut_(r);
    return {ok: true, code, host: r.host, now: Date.now(), room: gpPub_(r)};
  });
}
function gpJoin(p){
  const code = String(p.code || "").trim(), name = String(p.name || "").trim().slice(0, 20);
  if(!name) return {ok: false, error: "name"};
  return gpLock_(() => {
    const r = gpGet_(code); if(!r) return {ok: false, error: "noroom"};
    let pid = String(p.pid || "");
    if(!r.players.some(x => x.pid === pid)){
      if(r.status !== "lobby") return {ok: false, error: "started"};
      if(r.players.length >= 80) return {ok: false, error: "full"};
      pid = randStr(10); r.players.push({pid, name, yid: String(p.yid || "").slice(0, 40)}); gpPut_(r);
    }
    return {ok: true, pid, now: Date.now(), room: gpPub_(r)};
  });
}
function gpScores_(r){
  const keys = r.players.map(x => "gp:" + r.code + ":s:" + x.pid);
  const sc = keys.length ? gpC_().getAll(keys) : {};
  return r.players.map(x => { const s = sc["gp:" + r.code + ":s:" + x.pid]; const o = s ? JSON.parse(s) : {};
    return {pid: x.pid, name: x.name, score: o.score || 0, ok: o.ok || 0, n: o.n || 0, done: !!o.done}; });
}
function gpState(p){
  let r = gpGet_(String(p.code || "")); if(!r) return {ok: false, error: "noroom"};
  if(r.status === "run" && Date.now() > r.endAt + 5000) r.status = "done";
  // 終わったら1回だけ、結果を記録して、トップ3のボーナス(やりきリストのボスへのダメージ)を決める
  //  全員の最後の点数が届くまで待つ(先生が途中で終了したときも、生徒の画面が止まって点数を送るまで少しかかる)
  let ps = gpScores_(r);
  if(r.status === "done" && !r.awards && (ps.every(x => x.done || x.n === 0) && Date.now() > r.endAt + 3000 || Date.now() > r.endAt + 15000)){
    try{ r = gpAward_(r.code) || r; ps = gpScores_(r); }catch(err){ console.error(err); }
  }
  const room = gpPub_(r);
  room.players = ps;
  return {ok: true, now: Date.now(), room};
}
/* ---- トップ3のボーナス:やりきリストのボスへのダメージ(マルナ)。1日の上限600とは別枠。
   送るのは生徒の端末(マル単)。ここでは「だれに・いくつ」を決めて、1人1日2回までに数える ---- */
const GP_BONUS_DMG = [300, 200, 100], GP_BONUS_PER_DAY = 2;
const GP_BONUS_SH = "GPボーナス";
function gpBonusSheet_(){
  const ss = ss_(); let sh = ss.getSheetByName(GP_BONUS_SH);
  if(!sh){ sh = ss.insertSheet(GP_BONUS_SH); sh.appendRow(["日付", "やりきリストID", "名前", "順位", "ダメージ", "部屋", "日時"]); sh.setFrozenRows(1); sh.getRange("A:B").setNumberFormat("@"); }
  return sh;
}
function gpAward_(code){
  return gpLock_(() => {
    const r = gpGet_(code); if(!r) return null;
    if(r.awards) return r;
    const all = gpScores_(r).sort((a, b) => b.score - a.score || b.ok - a.ok || a.n - b.n);
    const yidOf = pid => (r.players.find(y => y.pid === pid) || {}).yid || "";
    // ボーナスは塾生(やりきリストの名簿の人)の中での1〜3位に
    const ranked = all.filter(x => x.score > 0 && yidOf(x.pid)).slice(0, 3);
    const day = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
    const sh = gpBonusSheet_(); const n = sh.getLastRow();
    const given = n > 1 ? sh.getRange(2, 1, n - 1, 2).getValues().filter(v => String(v[0]) === day).map(v => String(v[1])) : [];
    r.awards = ranked.map((x, i) => {
      const yid = yidOf(x.pid);
      const a = {pid: x.pid, rank: i + 1, dmg: GP_BONUS_DMG[i], status: "ok"};
      if(given.filter(g => g === yid).length >= GP_BONUS_PER_DAY) a.status = "today";
      else { sh.appendRow([day, yid, x.name, i + 1, GP_BONUS_DMG[i], r.code, new Date()]); given.push(yid); }
      return a;
    });
    // あとで見返せるように、結果をシート「GP結果」に残す
    try{
      const hs = gpHistSheet_(), at = new Date(), aw = {}; r.awards.forEach(a => aw[a.pid] = a);
      const rowsOut = all.map((x, i) => [at, r.code, r.org || "", r.title, r.sec, r.mode === "same" ? "みんな同じ" : "それぞれ別",
        i + 1, x.name, x.score, x.ok, x.n, aw[x.pid] && aw[x.pid].status === "ok" ? aw[x.pid].dmg : ""]);
      if(rowsOut.length) hs.getRange(hs.getLastRow() + 1, 1, rowsOut.length, rowsOut[0].length).setValues(rowsOut);
    }catch(err){ console.error(err); }
    gpPut_(r);
    return r;
  });
}

function gpStart(p){
  return gpLock_(() => {
    const r = gpGet_(String(p.code || "")); if(!r) return {ok: false, error: "noroom"};
    if(r.host !== String(p.host || "")) return {ok: false, error: "auth"};
    if(r.status === "lobby"){ r.status = "run"; r.startAt = Date.now() + 6000; r.endAt = r.startAt + r.sec * 1000; gpPut_(r); }
    return {ok: true, now: Date.now(), room: gpPub_(r)};
  });
}
function gpEnd(p){
  return gpLock_(() => {
    const r = gpGet_(String(p.code || "")); if(!r) return {ok: false, error: "noroom"};
    if(r.host !== String(p.host || "")) return {ok: false, error: "auth"};
    if(r.status === "run") r.endAt = Math.min(r.endAt, Date.now());
    r.status = "done"; gpPut_(r);
    return {ok: true, now: Date.now(), room: gpPub_(r)};   // ボーナスは、遅れて届く点数を待ってから gpState で記録する
  });
}
function gpScore(p){
  const code = String(p.code || ""), pid = String(p.pid || "");
  const r = gpGet_(code); if(!r) return {ok: false, error: "noroom"};
  if(!r.players.some(x => x.pid === pid)) return {ok: false, error: "player"};
  if(r.status === "lobby" || Date.now() > r.endAt + 20000) return {ok: false, error: "closed"};
  const o = {score: Math.max(0, Math.round(Number(p.score) || 0)), ok: Math.max(0, Number(p.ok) || 0), n: Math.max(0, Number(p.n) || 0), done: !!p.done};
  gpC_().put("gp:" + code + ":s:" + pid, JSON.stringify(o), GP_TTL);
  return {ok: true, now: Date.now()};
}
/* ---- 結果の記録と見返し ---- */
const GP_HIST_SH = "GP結果";
const GP_HIST_HEAD = ["日時", "部屋", "塾", "タイトル", "秒", "問題", "順位", "名前", "点数", "正解", "解いた", "ボーナス"];
function gpHistSheet_(){
  const ss = ss_(); let sh = ss.getSheetByName(GP_HIST_SH);
  if(!sh){ sh = ss.insertSheet(GP_HIST_SH); sh.appendRow(GP_HIST_HEAD); sh.setFrozenRows(1); sh.getRange("B:C").setNumberFormat("@"); }
  return sh;
}
/* 自分の塾の回だけ、新しい順に(最大30回分) */
function gpHistory(p){
  const org = String(p.org || ""), sh = ss_().getSheetByName(GP_HIST_SH);
  if(!sh || sh.getLastRow() < 2) return {ok: true, games: []};
  const n = sh.getLastRow() - 1, from = Math.max(2, sh.getLastRow() - 3000 + 1);
  const vals = sh.getRange(from, 1, sh.getLastRow() - from + 1, GP_HIST_HEAD.length).getValues();
  const games = {}, order = [];
  vals.forEach(v => { if(String(v[2]) !== org) return;
    const at = v[0] instanceof Date ? v[0].getTime() : new Date(v[0]).getTime(), key = String(v[1]) + "@" + Math.floor(at / 60000);
    if(!games[key]){ games[key] = {at, code: String(v[1]), title: String(v[3]), sec: Number(v[4]) || 0, mode: String(v[5]), players: []}; order.push(key); }
    games[key].players.push({rank: Number(v[6]), name: String(v[7]), score: Number(v[8]) || 0, ok: Number(v[9]) || 0, n: Number(v[10]) || 0, bonus: Number(v[11]) || 0});
  });
  return {ok: true, games: order.reverse().slice(0, 30).map(k => games[k])};
}
