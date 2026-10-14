// Google Apps Script backend для расширения «Товары из Китая» (EnvyCRM).
// Одна Google Таблица. Листы:
//   products        — каталог фабрик; photo = URL картинки (Drive/хостинг) или data URL
//   specs           — спецификация (параметры подбора) по каждому товару, ключ product_id
//   rates           — курсы валют, введённые вручную, по датам
//   selections      — сохранённые подборы клиентов (снимки цен/курсов, дата расчёта)
//   selection_items — позиции подборов (поиск «какому клиенту подбирали товар»)
// Опубликовать: Extensions → Apps Script → Deploy → Web app (Execute as: Me, Access: Anyone).
// URL и TOKEN подставить в china-app.html (константа API).

// Токен и папка Диска хранятся в Script Properties и создаются автоматически при setup().
// После деплоя один раз выполните setup() в редакторе — или я сделаю это через clasp.
function props_(){ return PropertiesService.getScriptProperties(); }
function token_(){
  var t = props_().getProperty('TOKEN');
  if(!t){ t = Utilities.getUuid(); props_().setProperty('TOKEN', t); }
  return t;
}
function photoFolder_(){
  var id = props_().getProperty('PHOTO_FOLDER_ID');
  if(id){ try{ return DriveApp.getFolderById(id); }catch(e){} }
  var f = DriveApp.createFolder('Фото товаров (расширение)');
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props_().setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
// Одноразовая настройка: создаёт таблицу (если скрипт standalone — создаст новую),
// папку фото и токен. Вернёт TOKEN — его подставить в china-app.html (API.token).
function setup(){
  var ss;
  try{ ss = SpreadsheetApp.getActiveSpreadsheet(); }catch(e){ ss = null; }
  if(!ss){
    ss = SpreadsheetApp.create('Товары из Китая — данные расширения');
  }
  Object.keys(SHEETS).forEach(sheet_); // создать листы с заголовками
  photoFolder_();
  var t = token_();
  Logger.log(JSON.stringify({ok:true, token:t, spreadsheetId:ss.getId(), folderId:props_().getProperty('PHOTO_FOLDER_ID')}));
  return t;
}
// ping без токена — проверка, что веб-приложение живо (для диагностики).
// setup=1 (без токена) запускает setup() и возвращает TOKEN один раз — после этого
// токен живёт только в Script Properties. Если setup уже был, возвращает уже_настроено.
function doGet(e){
  if(e && e.parameter && e.parameter.setup==='1'){
    var existed = props_().getProperty('TOKEN');
    var t = setup();
    return out_({ok:true, token: existed ? '(уже был настроен)' : t,
      spreadsheetId: props_().getProperty('SPREADSHEET_ID')||'привязанная',
      folderId: props_().getProperty('PHOTO_FOLDER_ID')});
  }
  return out_({ok:true, service:'china-goods', time:new Date().toISOString()});
}
const SHEETS = {
  products:  ['id','factory','article','name','price','currency','weight','tnved','duty','vat','price_date','photo','notes','updated_at'],
  specs:     ['id','product_id','params_json','updated_at'],
  rates:     ['id','currency','rate','date'],
  selections:['id','deal_id','client_name','created_at','calc_date','delivery','items_json','total_rub','source'],
  selection_items:['id','selection_id','product_id','deal_id','client_name','name','qty','price_rub','calc_date'],
};

function ss_(){
  var ss = null;
  try{ ss = SpreadsheetApp.getActiveSpreadsheet(); }catch(e){}
  if(ss) return ss;
  var id = props_().getProperty('SPREADSHEET_ID');
  if(id){ try{ return SpreadsheetApp.openById(id); }catch(e){} }
  ss = SpreadsheetApp.create('Товары из Китая — данные расширения');
  props_().setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}
function sheet_(name){
  var sh = ss_().getSheetByName(name);
  if(!sh){
    sh = ss_().insertSheet(name);
    sh.getRange(1,1,1,SHEETS[name].length).setValues([SHEETS[name]]);
  }
  return sh;
}
function rows_(name){
  var sh = sheet_(name);
  var vals = sh.getDataRange().getValues();
  var head = vals.shift() || SHEETS[name];
  return vals.filter(function(r){return r.join('')!=='';}).map(function(r){
    var o={}; head.forEach(function(h,i){o[h]=r[i];}); return o;
  });
}
function uid_(){ return Utilities.getUuid(); }

function doPost(e){
  try{
    var p = JSON.parse(e.postData.contents);
    if(p.token !== token_()) return out_({ok:false, error:'bad_token'});
    var a = p.action, sheet = p.sheet;

    if(a==='list'){
      if(!SHEETS[sheet]) return out_({ok:false, error:'bad_sheet'});
      var rows = rows_(sheet);
      if(p.where){
        rows = rows.filter(function(r){
          return Object.keys(p.where).every(function(k){return String(r[k])===String(p.where[k]);});
        });
      }
      return out_({ok:true, rows:rows});
    }
    if(!SHEETS[sheet] && a!=='add_selection') return out_({ok:false, error:'bad_sheet'});
    if(a==='add'){
      var obj = p.data||{}; if(!obj.id) obj.id = uid_();
      append_(sheet, obj); return out_({ok:true, id:obj.id});
    }
    if(a==='update'){ return out_(update_(sheet, p.data||{})); }
    if(a==='upsert'){
      var d = p.data||{}, keys = p.keys||['id'];
      var found = rows_(sheet).filter(function(r){
        return keys.every(function(k){return String(r[k]||'')===String(d[k]||'');});
      });
      if(found.length){ d.id = found[0].id; return out_(update_(sheet, d)); }
      if(!d.id) d.id = uid_(); append_(sheet, d); return out_({ok:true, id:d.id, created:true});
    }
    // add_selection: подбор + позиции одной операцией
    if(a==='add_selection'){
      var sel = p.selection||{}; if(!sel.id) sel.id = uid_();
      append_('selections', sel);
      (p.items||[]).forEach(function(it){
        it.id = it.id || uid_();
        it.selection_id = sel.id; it.deal_id = sel.deal_id; it.client_name = sel.client_name;
        append_('selection_items', it);
      });
      return out_({ok:true, id:sel.id});
    }
    // upload_photo: сохранить фото на Google Диск (папка создаётся автоматически), вернуть URL.
    // p.data = {base64: "...", mime: "image/jpeg", name: "артикул.jpg"}
    if(a==='upload_photo'){
      var d2 = p.data||{};
      var blob = Utilities.newBlob(Utilities.base64Decode(d2.base64), d2.mime||'image/jpeg', d2.name||'photo.jpg');
      var file = photoFolder_().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
      return out_({ok:true, id:file.getId(), url:url});
    }
    return out_({ok:false, error:'bad_action'});
  }catch(err){ return out_({ok:false, error:String(err)}); }
}

function append_(sheet, obj){
  var cols = SHEETS[sheet];
  sheet_(sheet).appendRow(cols.map(function(c){return obj[c]!==undefined?obj[c]:'';}));
}
function update_(sheet, obj){
  if(!obj.id) return {ok:false, error:'no_id'};
  var sh = sheet_(sheet), cols = SHEETS[sheet];
  var ids = sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).getValues();
  for(var i=0;i<ids.length;i++){
    if(String(ids[i][0])===String(obj.id)){
      cols.forEach(function(c,j){
        if(obj[c]!==undefined) sh.getRange(i+2,j+1).setValue(obj[c]);
      });
      return {ok:true, id:obj.id, updated:true};
    }
  }
  return {ok:false, error:'not_found'};
}
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
