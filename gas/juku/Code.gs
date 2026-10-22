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
