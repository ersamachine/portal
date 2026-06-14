/*** ERSA GROUP · Muhasebe (Cari Hesap) — Apps Script Web App  (v2 / tüm modüller) ***
 * MUHASEBE Google Sheet'ine bağlı Apps Script projesine yapıştırın.
 * Sheet hesap motorudur (KPI/bakiye/çek USD/kasa-ortak yürüyen bakiye = formül).
 * Uygulama hesaplanmış değerleri okur; ham kayıtları yazar/günceller/siler.
 *********************************************************************************/
const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = "Europe/Istanbul";

function doGet(e){
  try{
    return json({ ok:true,
      panel: readPanel(), cariOzet: readCariOzet(), musteriler: readMusteriler(),
      faturalar: readFaturalar(), odemeler: readOdemeler(), cekler: readCekler(),
      taksitler: readTaksitler(), kasa: readKasa(), ortak: readOrtak(),
      talepler: readTalepler()
    });
  }catch(err){ return json({ok:false, error:String(err)}); }
}
function doPost(e){
  try{
    const p = JSON.parse(e.postData.contents);
    if(p.action==="save")   return json(saveRec(p.type, p.row||null, p.data||{}));
    if(p.action==="delete") return json(delRec(p.type, p.row));
    if(p.action==="talep")  return json(talepUpdate(p.tip, p.row, p.durum, p.sifre));
    return json({ok:false, error:"Bilinmeyen işlem"});
  }catch(err){ return json({ok:false, error:String(err)}); }
}

/* yardımcılar */
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh(n){ return SS.getSheetByName(n); }
function num(v){ return (typeof v==="number") ? v : (parseFloat(v)||0); }
function dstr(v){ if(!v) return ""; if(Object.prototype.toString.call(v)==="[object Date]") return Utilities.formatDate(v,TZ,"yyyy-MM-dd"); return String(v); }
function dval(s){ if(!s) return ""; const d=new Date(s); return isNaN(d.getTime())?s:d; }
function firstEmptyRow(s,keyCol,start){ const last=s.getLastRow(); for(let r=start;r<=last;r++){const v=s.getRange(r,keyCol).getValue(); if(v===""||v===null) return r;} return last+1; }

/* OKUMA */
function readPanel(){ const p=sh("Panel"); if(!p)return{}; const g=a=>p.getRange(a).getValue();
  return {fatura:num(g("B8")),tahsilat:num(g("D8")),oran:num(g("F8")),bekleyen:num(g("I8")),gun30:num(g("K8")),kurFarki:num(g("M8")),netAlacak:num(g("P8")),kasaTL:num(g("R8")),kasaUSD:num(g("T8"))}; }
function readCariOzet(){ const s=sh("Cari Özet"); if(!s)return[]; const last=s.getLastRow(); if(last<6)return[];
  return s.getRange(6,1,last-5,7).getValues().filter(r=>r[1]).map(r=>({sira:r[0],musteri:String(r[1]),sorumlu:r[2]||"",toplamFatura:num(r[3]),toplamTahsilat:num(r[4]),kurFarki:num(r[5]),bakiye:num(r[6])})); }
function readMusteriler(){ const s=sh("Müşteriler"); if(!s)return[]; const last=s.getLastRow(); if(last<5)return[];
  return s.getRange(5,1,last-4,10).getValues().map((r,i)=>({_row:5+i,sira:r[0],sirket:String(r[1]||""),sorumlu:r[2]||"",telefon:r[3]||"",eposta:r[4]||"",adres:r[5]||"",yetkili:r[6]||"",vd:r[7]||"",vno:r[8]||"",acilis:num(r[9])})).filter(m=>m.sirket); }
function readFaturalar(){ const s=sh("Faturalar"); if(!s)return[]; const last=s.getLastRow(); if(last<5)return[];
  return s.getRange(5,1,last-4,6).getValues().map((r,i)=>({_row:5+i,musteri:String(r[0]||""),no:r[1]||"",faturaTarihi:dstr(r[2]),vade:dstr(r[3]),tutar:num(r[4]),aciklama:r[5]||""})).filter(f=>f.musteri); }
function readOdemeler(){ const s=sh("Ödemeler"); if(!s)return[]; const last=s.getLastRow(); if(last<5)return[];
  return s.getRange(5,1,last-4,7).getValues().map((r,i)=>({_row:5+i,odemeTuru:r[0]||"",musteri:String(r[1]||""),tarih:dstr(r[2]),tl:num(r[3]),kur:num(r[4]),usd:num(r[5]),aciklama:r[6]||""})).filter(o=>o.musteri); }
function readCekler(){ const s=sh("Çek Takibi"); if(!s)return[]; const last=s.getLastRow(); if(last<5)return[];
  return s.getRange(5,1,last-4,13).getValues().map((r,i)=>({_row:5+i,musteri:String(r[0]||""),cekNo:r[1]||"",alisTarihi:dstr(r[2]),cekTL:num(r[3]),alisKuru:num(r[4]),usdAlis:num(r[5]),vade:dstr(r[6]),vadeKuru:num(r[7]),usdVade:num(r[8]),kurFarki:num(r[9]),durum:r[10]||"",yer:r[11]||"",verilis:dstr(r[12])})).filter(c=>c.musteri); }
function readTaksitler(){ const s=sh("Taksitler"); if(!s)return[]; const last=s.getLastRow(); if(last<9)return[];
  return s.getRange(9,1,last-8,8).getValues().map((r,i)=>({_row:9+i,musteri:String(r[0]||""),taksit:r[1]||"",vade:dstr(r[2]),tutar:num(r[3]),durum:r[4]||"",odemeTarihi:dstr(r[5]),kalanGun:r[6],aciklama:r[7]||""})).filter(t=>t.musteri); }
function readKasa(){ const s=sh("Kasa"); if(!s)return[]; const last=s.getLastRow(); if(last<8)return[];
  return s.getRange(8,1,last-7,10).getValues().map((r,i)=>({_row:8+i,sira:r[0],tarih:dstr(r[1]),aciklama:r[2]||"",kimden:r[3]||"",islemTuru:r[4]||"",yon:r[5]||"",tutar:num(r[6]),pb:r[7]||"",tlBakiye:num(r[8]),usdBakiye:num(r[9])})).filter(k=>k.aciklama||k.tutar); }
function readOrtak(){ const s=sh("Ortak Cari"); if(!s)return[]; const last=s.getLastRow(); if(last<7)return[];
  return s.getRange(7,1,last-6,11).getValues().map((r,i)=>({_row:7+i,sira:r[0],tarih:dstr(r[1]),aciklama:r[2]||"",tur:r[3]||"",yon:r[4]||"",tutar:num(r[5]),pb:r[6]||"",kur:num(r[7]),usd:num(r[8]),tlBakiye:num(r[9]),usdBakiye:num(r[10])})).filter(o=>o.aciklama||o.tutar); }

/* her modül: sayfa, veri başlangıç satırı, anahtar kolon, ham kolon yazıcı(d) */
const CFG = {
  fatura:{sheet:"Faturalar", start:5, key:1, cols:d=>({1:d.musteri||"",2:d.no||"",3:dval(d.faturaTarihi),4:dval(d.vade),5:num(d.tutar),6:d.aciklama||""})},
  odeme: {sheet:"Ödemeler", start:5, key:2, cols:d=>{let u=num(d.usd); if(!u&&num(d.tl)&&num(d.kur))u=num(d.tl)/num(d.kur); return {1:d.odemeTuru||"",2:d.musteri||"",3:dval(d.tarih),4:d.tl?num(d.tl):"",5:d.kur?num(d.kur):"",6:u,7:d.aciklama||""};}},
  cek:   {sheet:"Çek Takibi", start:5, key:1, cols:d=>({1:d.musteri||"",2:d.cekNo||"",3:dval(d.alisTarihi),4:num(d.cekTL),5:d.alisKuru?num(d.alisKuru):"",7:dval(d.vade),8:d.vadeKuru?num(d.vadeKuru):"",11:d.durum||"Bekliyor",12:d.yer||"",13:d.verilis?dval(d.verilis):""})},
  taksit:{sheet:"Taksitler", start:9, key:1, cols:d=>({1:d.musteri||"",2:d.taksit||"",3:dval(d.vade),4:num(d.tutar),5:d.durum||"Bekliyor",6:d.odemeTarihi?dval(d.odemeTarihi):"",8:d.aciklama||""})},
  kasa:  {sheet:"Kasa", start:8, key:7, cols:d=>({2:dval(d.tarih),3:d.aciklama||"",4:d.kimden||"",5:d.islemTuru||"",6:d.yon||"",7:num(d.tutar),8:d.pb||"TL"})},
  ortak: {sheet:"Ortak Cari", start:7, key:3, cols:d=>({2:dval(d.tarih),3:d.aciklama||"",4:d.tur||"",5:d.yon||"",6:num(d.tutar),7:d.pb||"TL",8:d.kur?num(d.kur):""})},
  musteri:{sheet:"Müşteriler", start:5, key:2, cols:d=>({2:d.sirket||"",3:d.sorumlu||"",4:d.telefon||"",5:d.eposta||"",6:d.adres||"",7:d.yetkili||"",8:d.vd||"",9:d.vno||"",10:d.acilis?num(d.acilis):""})}
};

function saveRec(type, row, d){
  const cfg = CFG[type]; if(!cfg) return {ok:false,error:"Tip yok: "+type};
  const s = sh(cfg.sheet);
  const r = row ? Number(row) : firstEmptyRow(s, cfg.key, cfg.start);
  const map = cfg.cols(d);
  Object.keys(map).forEach(c => s.getRange(r, Number(c)).setValue(map[c]));
  if(type==="musteri" && !row){ s.getRange(r,1).setValue(r - (cfg.start-1)); } // Sıra (sadece ekleme)
  return {ok:true, row:r};
}
function delRec(type, row){
  const cfg = CFG[type]; if(!cfg) return {ok:false,error:"Tip yok"};
  const s = sh(cfg.sheet); const r = Number(row);
  const cols = Object.keys(cfg.cols({})).map(Number);
  if(type==="musteri") cols.push(1);
  cols.forEach(c => s.getRange(r, c).setValue(""));
  return {ok:true, row:r};
}

/* ===== PORTAL TALEPLERİ (müşteri portalından gelen) ===== */
function nrm_(v){ return String(v==null?"":v).trim(); }
function rows_(name,start,width){ const s=sh(name); if(!s)return[]; const last=s.getLastRow(); if(last<start)return[]; return s.getRange(start,1,last-start+1,width).getValues(); }

function readTalepler(){
  const out={bildirimler:[],sifreler:[],teklifler:[],servisler:[]};
  rows_("Ödeme Bildirimleri",2,8).forEach((r,i)=>{ if(nrm_(r[1])) out.bildirimler.push({_row:2+i,zaman:dstr(r[0]),musteri:nrm_(r[1]),tutar:num(r[2]),pb:nrm_(r[3]),tarih:dstr(r[4]),yontem:nrm_(r[5]),aciklama:nrm_(r[6]),durum:nrm_(r[7])||"Yeni"}); });
  rows_("Şifre Talepleri",2,5).forEach((r,i)=>{ if(nrm_(r[1])) out.sifreler.push({_row:2+i,zaman:dstr(r[0]),musteri:nrm_(r[1]),makine:nrm_(r[2]),durum:nrm_(r[3])||"Yeni",sifre:nrm_(r[4])}); });
  rows_("Teklif Talepleri",2,6).forEach((r,i)=>{ if(nrm_(r[1])) out.teklifler.push({_row:2+i,zaman:dstr(r[0]),musteri:nrm_(r[1]),urun:nrm_(r[2]),adet:nrm_(r[3]),aciklama:nrm_(r[4]),durum:nrm_(r[5])||"Yeni"}); });
  rows_("Servis Talepleri",2,5).forEach((r,i)=>{ if(nrm_(r[1])) out.servisler.push({_row:2+i,zaman:dstr(r[0]),musteri:nrm_(r[1]),makine:nrm_(r[2]),sorun:nrm_(r[3]),durum:nrm_(r[4])||"Yeni"}); });
  out.bildirimler.reverse(); out.sifreler.reverse(); out.teklifler.reverse(); out.servisler.reverse();
  return out;
}

// tip: bildirim|sifre|teklif|servis ; durum sütununu (ve şifre için Şifre sütununu) günceller
function talepUpdate(tip,row,durum,sifre){
  const MAP={ bildirim:{sheet:"Ödeme Bildirimleri",durumCol:8},
              sifre:{sheet:"Şifre Talepleri",durumCol:4,sifreCol:5},
              teklif:{sheet:"Teklif Talepleri",durumCol:6},
              servis:{sheet:"Servis Talepleri",durumCol:5} };
  const cfg=MAP[tip]; if(!cfg) return {ok:false,error:"Tip yok"};
  const s=sh(cfg.sheet); if(!s) return {ok:false,error:"Sayfa yok"};
  const r=Number(row); if(!r) return {ok:false,error:"Satır yok"};
  if(durum!=null) s.getRange(r,cfg.durumCol).setValue(durum);
  if(cfg.sifreCol && sifre!=null) s.getRange(r,cfg.sifreCol).setValue(String(sifre));
  return {ok:true};
}
