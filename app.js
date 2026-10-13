/* Thông báo phạt nhà xe vận chuyển
 * Dữ liệu Bên A / Bên B đọc từ MASTER.xlsx (sheet "BÊN A" và "BÊN B").
 */

// Mã lò dùng trong số thông báo, VD 001/2026-TBVC-CPVN-LA
const LO_CODE = { "Sagri": "SG", "Lộc An": "LA", "Cẩm Tú": "CT" };
const DIEU_KHOAN_MAC_DINH = "Điều 10, mục 3, khoản (i)";

let BEN_A = [], BEN_B = [];

const $=id=>document.getElementById(id);
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fmtDate=v=>{if(!v)return"__/__/____";const[y,m,d]=v.split("-");return`${d}/${m}/${y}`};
const money=n=>Math.round(n).toLocaleString("en-US");
const num=n=>(+n).toLocaleString("en-US",{maximumFractionDigits:2});
const noAccent=s=>String(s??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/đ/g,"d").replace(/Đ/g,"D");

// Vietnamese number to words
function docSo(n){
  n=Math.round(n); if(n===0) return "Không đồng";
  const so=["không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"], dv=["","nghìn","triệu","tỷ"];
  const doc3=(x,full)=>{const t=Math.floor(x/100),c=Math.floor(x%100/10),d=x%10,out=[];
    if(full||t) out.push(so[t]+" trăm");
    if(c===0){ if(d&&(full||t)) out.push("linh"); }
    else if(c===1) out.push("mười"); else out.push(so[c]+" mươi");
    if(d){ if(d===1&&c>1) out.push("mốt"); else if(d===5&&c>0) out.push("lăm"); else out.push(so[d]); }
    return out.join(" ");};
  const groups=[]; let x=n; while(x>0){groups.push(x%1000);x=Math.floor(x/1000);}
  const parts=[];
  for(let i=groups.length-1;i>=0;i--){ if(!groups[i]) continue; parts.push(doc3(groups[i],i<groups.length-1)+(dv[i%4]?" "+dv[i%4]:"")); }
  const s=parts.join(" ")+" đồng"; return s[0].toUpperCase()+s.slice(1);
}

const pad2 = n => String(n).padStart(2, "0");
const isoDate = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const loCode = lo => LO_CODE[lo] || noAccent(lo).split(/\s+/).map(w => w[0] || "").join("").toUpperCase();

// ---- MASTER.xlsx ----
function cellDate(v) {
  if (v instanceof Date) return isoDate(v);
  if (typeof v === "number") { const d = XLSX.SSF.parse_date_code(v); return d ? `${d.y}-${pad2(d.m)}-${pad2(d.d)}` : ""; }
  const m = String(v ?? "").match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  return m ? `${m[3]}-${pad2(m[2])}-${pad2(m[1])}` : "";
}
const clean = v => String(v ?? "").trim();
function sheetRows(wb, name) {
  const key = wb.SheetNames.find(n => noAccent(n).toUpperCase().replace(/\s+/g, "") === noAccent(name).toUpperCase().replace(/\s+/g, ""));
  if (!key) throw new Error(`Không tìm thấy sheet "${name}" trong file MASTER.`);
  return XLSX.utils.sheet_to_json(wb.Sheets[key], { defval: "" });
}
function parseMaster(buf) {
  const wb = XLSX.read(buf, { type: "array" }); // ngày giữ dạng số serial của Excel để không lệch múi giờ
  const a = sheetRows(wb, "BÊN A").filter(r => clean(r["Tên Lò Giết Mổ"])).map(r => ({
    lo: clean(r["Tên Lò Giết Mổ"]), cty: clean(r["Tên Công Ty"]), mst: clean(r["MST"]), dc: clean(r["Địa Chỉ"]),
    dd: clean(r["Người Đại Diện"]), stk: clean(r["Số Tài Khoản Ngân Hàng"]), tentk: clean(r["Tên Tài Khoản Ngân Hàng"]),
    nh: clean(r["Tên Ngân Hàng"]), ck: clean(r["Nội Dung Chuyển Khoản"])
  }));
  const b = sheetRows(wb, "BÊN B").filter(r => clean(r["Số Xe"])).map(r => ({
    xe: clean(r["Số Xe"]), dd: clean(r["Người Đại Diện"]), ncc: clean(r["Tên Nhà Cung Cấp"]), hd: clean(r["Số Hợp Đồng"]),
    dk: clean(r["Điều Khoản"]) || DIEU_KHOAN_MAC_DINH, ky: cellDate(r["Ngày Ký Hợp Đồng"])
  }));
  if (!a.length || !b.length) throw new Error("File MASTER không có dữ liệu Bên A hoặc Bên B.");
  return { a, b };
}
function applyMaster(data, label) {
  BEN_A = data.a; BEN_B = data.b;
  const la = $("loA").value, lb = $("soXe").value;
  $("loA").innerHTML = ""; $("soXe").innerHTML = "";
  BEN_A.forEach((x, i) => $("loA").add(new Option(x.lo, i)));
  BEN_B.forEach((x, i) => $("soXe").add(new Option(`${x.xe} · ${x.ncc}`, i)));
  if (la && BEN_A[la]) $("loA").value = la;
  if (lb && BEN_B[lb]) $("soXe").value = lb;
  fillA(); fillB();
  $("masterInfo").textContent = `${label} · ${BEN_A.length} lò, ${BEN_B.length} xe`;
  render();
}
async function initMaster() {
  try {
    const res = await fetch("MASTER.xlsx", { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    applyMaster(parseMaster(await res.arrayBuffer()), "MASTER.xlsx");
  } catch (e) {
    $("masterInfo").textContent = location.protocol === "file:"
      ? "Không đọc được MASTER.xlsx khi mở file trực tiếp. Hãy chạy qua server (xem README)."
      : "Không tải được MASTER.xlsx: " + (e.message || e);
    $("masterInfo").hidden = false;
  }
}

// ---- Master-driven fields: faded + read-only, unlocked with the pencil ----
// Master-driven fields: faded + read-only, unlocked with the pencil
const PEN='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>';
document.querySelectorAll("input.m").forEach(inp=>{
  const w=document.createElement("div"); w.className="lock"; inp.replaceWith(w); w.appendChild(inp);
  const b=document.createElement("button"); b.type="button"; b.className="pen"; b.innerHTML=PEN; b.title="Sửa"; b.setAttribute("aria-label","Sửa ô này");
  b.onclick=()=>setLock(inp,!inp.readOnly); w.appendChild(b); setLock(inp,true);
});
function setLock(inp,locked){ inp.readOnly=locked; const b=inp.nextElementSibling; b.classList.toggle("on",!locked); b.title=locked?"Sửa":"Khóa lại"; if(!locked) inp.focus(); }
const relock=ids=>ids.forEach(id=>setLock($(id),true));
function fillCK() { const a = BEN_A[$("loA").value]; if (a) $("noiDungCK").value = (noAccent($("ncc").value.trim()) + " " + a.ck).trim(); }
function fillA() { const a = BEN_A[$("loA").value]; if (!a) return; $("daiDienA").value = a.dd; $("diaChiA").value = a.dc; fillCK(); relock(["daiDienA", "diaChiA", "noiDungCK"]); }
function fillB() { const b = BEN_B[$("soXe").value]; if (!b) return; $("ncc").value = b.ncc; $("soHD").value = b.hd; $("ngayHD").value = b.ky; $("dieuKhoan").value = b.dk; $("daiDienB").value = b.dd; fillCK(); relock(["ncc", "soHD", "ngayHD", "dieuKhoan", "daiDienB", "noiDungCK"]); }
function endOfMonth(v) { if (!v) return ""; const [y, m] = v.split("-").map(Number); return `${y}-${pad2(m)}-${new Date(y, m, 0).getDate()}`; }
$("loA").onchange = () => { fillA(); render(); };
$("soXe").onchange = () => { fillB(); render(); };
$("ngay").addEventListener("input", () => { $("han").value = endOfMonth($("ngay").value); });
$("ncc").addEventListener("input", () => { if ($("noiDungCK").readOnly) fillCK(); });

// Default values
const today = new Date();
$("ngay").value = isoDate(today);
$("han").value = endOfMonth($("ngay").value);
$("soTB").value = `001/${today.getFullYear()}`;

// ---- State for files ----
let photos=[];      // {id,url,w,h}
let bienBan=null;   // {name,pages:[{url,w,h}]}
let pricePdf=null;  // {name,bytes:Uint8Array}
let priceCanvases=[];

function splitCty(c){ const i=c.search(/\s+Chi nhánh/i); return i>0?[c.slice(0,i).trim(),c.slice(i).trim()]:[c.trim(),""]; }

function render(){
  const A=BEN_A[$("loA").value], B=BEN_B[$("soXe").value];
  if(!A||!B){ $("scaler").innerHTML='<div class="page pdfp empty"><b>Chưa có dữ liệu MASTER</b><span>Chọn file MASTER.xlsx ở đầu trang</span></div>'; $("pcount").textContent=""; fit(); return; }
  const xe=B.xe;
  const v=k=>$(k).value.trim();
  const ncc=v("ncc"), kg=+$("kg").value||0, gia=+$("gia").value||0, tien=Math.round(kg*gia);
  const ngay=fmtDate($("ngay").value);
  $("tongTien").textContent=money(tien)+" đ";
  const soTB=`${v("soTB")}-TBVC-CPVN-${loCode(A.lo)}`;
  $("soTBFull").textContent="→ "+soTB;
  const [cty1,cty2]=splitCty(A.cty);
  const sign=n=>esc(n.toUpperCase());

  const p1=`<div class="page gen">
    <div class="hdr"><div>${esc(cty1)}</div>${cty2?`<div>${esc(cty2)}</div>`:""}<div>MST: ${esc(A.mst)}</div></div>
    <div class="ttl"><b>THÔNG BÁO PHẠT VI PHẠM HỢP ĐỒNG</b>(Số: ${esc(soTB)})</div>
    <div class="kg">Kính gửi: ${esc(ncc.toUpperCase())}</div>
    <p>Hôm nay ngày ${ngay}, tại lò giết mổ ${esc(A.lo)}. Địa chỉ: ${esc(v("diaChiA"))}.</p>
    <h3><span>I.</span>Căn cứ:</h3>
    <ul>
      <li>Hợp đồng vận chuyển số: <b>${esc(v("soHD"))}</b> ký ngày ${fmtDate($("ngayHD").value)} (sau đây gọi tắt là Hợp đồng) ký giữa ${esc(cty1)} (Bên A-đơn vị thuê dịch vụ vận chuyển) với ${esc(ncc)} (Bên B-đơn vị cung cấp dịch vụ vận chuyển).</li>
      <li>Biên bản xác nhận (vv heo chết trên xe vận chuyển) ngày ${ngay};</li>
      <li>Và thông tin từ các bộ phận liên quan;</li>
    </ul>
    <h3><span>II.</span>Xác định nội dung vi phạm</h3>
    <p class="indent">Trong quá trình xe ${esc(xe)} thuộc NCC ${esc(ncc)}, vận chuyển heo thịt từ trại ${esc(v("trai"))} về lò giết mổ ${esc(A.lo)} tổng cộng ${esc($("tong").value)} con, trong đó có ${esc($("soChet").value)} con heo chết trên xe, trọng lượng ${num(kg)}kg.</p>
    <p class="indent">&gt;&gt; Căn cứ theo <b>${esc(v("dieuKhoan"))}</b> của <b>Hợp đồng vận chuyển số: ${esc(v("soHD"))}</b>, Nhà vận chuyển sẽ đền bù thiệt hại con heo chết với mức giá theo giá heo thị trường tại thời điểm phát sinh.</p>
    <h3><span>III.</span>Mức phạt áp dụng</h3>
    <p><b>Bên A quyết định phạt Bên B như sau:</b></p>
    <p class="pl">Nhà xe đền bù thiệt hại đối với con heo chết:</p>
    <p class="pl">Giá đền bù = giá heo thịt Swine bán cho khách hàng <b>${money(gia)} đ/kg</b></p>
    <p class="pl"><b>Tổng số tiền đền bù: ${num(kg)}kg x ${money(gia)}đ/kg = ${money(tien)} đ</b></p>
    <p class="pl"><i>(Số tiền bằng chữ: ${esc(docSo(tien))}).</i></p>
  </div>`;

  const p2=`<div class="page gen">
    <p><b>Phương thức thanh toán:</b></p>
    <p class="pl" style="padding-left:12px">Bên B đóng tiền vào Tài khoản</p>
    <ul>
      <li>Số Tài khoản: ${esc(A.stk)}</li>
      <li>Tên Tài khoản: ${esc(A.tentk)}</li>
      <li>Ngân hàng: ${esc(A.nh)}</li>
      <li>Nội dung CK: ${esc(v("noiDungCK"))}</li>
    </ul>
    <h3>IV. Đề nghị thực hiện</h3>
    <p>Đề nghị ${esc(ncc)}</p>
    <ul>
      <li>Thanh toán khoản tiền phạt và bồi thường nêu trên trước ngày ${fmtDate($("han").value)},</li>
      <li>Trường hợp quá hạn, số tiền trên sẽ được cấn trừ vào cước vận chuyển chưa thanh toán theo ${esc(v("dieuQuaHan"))} của Hợp đồng.</li>
    </ul>
    <h3>V. Hiệu lực</h3>
    <p>Thông báo này được lập thành <b>02 bản</b> có giá trị pháp lý như nhau, mỗi bên giữ 01 bản để làm cơ sở thực hiện và hạch toán.</p>
    <div class="sign"><div>ĐẠI DIỆN BÊN A</div><div>ĐẠI DIỆN BÊN B</div><div class="space"></div><div class="space"></div><div>${sign(v("daiDienA"))}</div><div>${sign(v("daiDienB"))}</div></div>
  </div>`;

  // Photo pages: 6 photos per page
  const head=`<p class="photo-h">Hình ảnh heo chết ngày ${ngay}, số lượng ${esc($("soChet").value)} con, trọng lượng ${num(kg)}kg</p>
    <p class="photo-h">Xe vận chuyển ${esc(xe)} – NCC ${esc(ncc)}${v("taiXe")?` – tài xế ${esc(v("taiXe"))}`:""}</p>`;
  let photoPages="";
  if(!photos.length){
    photoPages=`<div class="page gen phpage">${head}<div class="photos"><div class="phrow"><div class="ph empty">Chưa có hình</div><div class="ph empty">Chọn hình ở mục 4</div></div></div></div>`;
  } else {
    photoPages=`<div class="page gen phpage">${head}${photoLayout(photos)}</div>`;
  }

  // Biên bản sự việc: mỗi trang ảnh/PDF canh vừa một trang A4
  const bbHead=`<p class="photo-h">Biên bản sự việc ngày ${ngay}</p>`;
  const bbPages=bienBan
    ? bienBan.pages.map((p,i)=>`<div class="page gen bb">${i===0?bbHead:""}<div class="bbbox">${fitImg(p,634,i===0?955:993)}</div></div>`).join("")
    : `<div class="page gen bb">${bbHead}<div class="bbbox empty">Chọn biên bản ở mục 5</div></div>`;

  const priceSlots=pricePdf? priceCanvases.map((_,i)=>`<div class="page pdfp" data-pi="${i}"></div>`).join("")
    : `<div class="page pdfp empty"><b>Trang bảng giá</b><span>Chọn file PDF ở mục 6</span></div>`;

  $("scaler").innerHTML=p1+p2+photoPages+bbPages+priceSlots;
  document.querySelectorAll(".page.pdfp[data-pi]").forEach(el=>el.appendChild(priceCanvases[+el.dataset.pi]));
  const total=document.querySelectorAll("#scaler .page").length;
  $("pcount").textContent=`${total} trang`;
  fit();
}

// Scale the 794px-wide pages to fit the preview column
function fit(){
  const w=$("viewport").clientWidth, s=Math.min(1,w/794), sc=$("scaler");
  sc.style.transform=`scale(${s})`;
  $("viewport").style.height=(sc.scrollHeight*s)+"px";
}
new ResizeObserver(fit).observe($("viewport"));
document.querySelectorAll(".side input:not([type=file]),.side select").forEach(el=>el.addEventListener("input",render));

// Ảnh có kích thước tính sẵn để vừa khung (html2canvas không hỗ trợ object-fit)
function fitImg(p,bw,bh){ const k=Math.min(bw/p.w,bh/p.h); return `<img src="${p.url}" alt="" style="width:${Math.floor(p.w*k)}px;height:${Math.floor(p.h*k)}px">`; }

// Chọn cách xếp 1–4 hình sao cho hình to và đều nhau nhất trong khung trang
const PH_W=674, PH_H=870, PH_GAP=12;
const PH_LAYOUTS={1:[[1]],2:[[1,1],[2]],3:[[1,1,1],[2,1],[1,2]],4:[[2,2]]};
function photoLayout(list){
  let best=null;
  for(const rows of PH_LAYOUTS[list.length]){
    const rh=(PH_H-PH_GAP*(rows.length-1))/rows.length; let i=0, areas=[];
    const cells=rows.map(k=>{ const cw=(PH_W-PH_GAP*(k-1))/k;
      const row=list.slice(i,i+=k).map(p=>{ const s=Math.min(cw/p.w,rh/p.h); areas.push(p.w*p.h*s*s); return {p,cw,rh,h:p.h*s}; });
      const h=Math.max(...row.map(c=>c.h)); row.forEach(c=>c.rh=h); return row; }); // dồn hàng sát nhau, khoảng cách đều
    const score=areas.length*Math.min(...areas)+areas.reduce((a,b)=>a+b,0)*1e-3; // ưu tiên các hình to đều nhau
    if(!best||score>best.score) best={score,cells};
  }
  return `<div class="photos">${best.cells.map(r=>`<div class="phrow">${r.map(c=>`<div class="ph" style="width:${Math.floor(c.cw)}px;height:${Math.floor(c.rh)}px">${fitImg(c.p,c.cw,c.rh)}</div>`).join("")}</div>`).join("")}</div>`;
}

// ---- Photos ----
function shrink(file,m=1600){ return new Promise((res,rej)=>{
  const img=new Image(), u=URL.createObjectURL(file);
  img.onload=()=>{ const k=Math.min(1,m/Math.max(img.naturalWidth,img.naturalHeight));
    const c=document.createElement("canvas"); c.width=Math.round(img.naturalWidth*k); c.height=Math.round(img.naturalHeight*k);
    c.getContext("2d").drawImage(img,0,0,c.width,c.height); URL.revokeObjectURL(u); res({url:c.toDataURL("image/jpeg",0.88),w:c.width,h:c.height}); };
  img.onerror=()=>{URL.revokeObjectURL(u);rej()}; img.src=u; }); }
const MAX_PHOTOS=4;
async function addPhotos(files){
  const imgs=files.filter(f=>f.type.startsWith("image/"));
  if(photos.length+imgs.length>MAX_PHOTOS) status(`Tối đa ${MAX_PHOTOS} hình. Chỉ lấy ${Math.max(0,MAX_PHOTOS-photos.length)} hình đầu.`,true);
  for(const f of imgs.slice(0,Math.max(0,MAX_PHOTOS-photos.length))){
    try{ photos.push({id:Math.random().toString(36).slice(2),...await shrink(f)}); }catch(e){ status(`Không đọc được hình ${f.name}.`,true); } }
  drawThumbs(); render();
}
function drawThumbs(){
  $("dropImg").hidden=photos.length>=MAX_PHOTOS;
  $("thumbs").innerHTML=photos.map(p=>`<div class="thumb"><img src="${p.url}" alt=""><button data-id="${p.id}" aria-label="Xóa hình">×</button></div>`).join("");
}
$("thumbs").onclick=e=>{const id=e.target.dataset?.id; if(id){photos=photos.filter(p=>p.id!==id);drawThumbs();render();}};

// ---- Price PDF ----
if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
async function setPdf(file){
  if(!file) return;
  if(file.type!=="application/pdf"&&!/\.pdf$/i.test(file.name)){status("File bảng giá phải là PDF.",true);return;}
  status("Đang đọc bảng giá…");
  const bytes=new Uint8Array(await file.arrayBuffer());
  try{
    const doc=await pdfjsLib.getDocument({data:bytes.slice()}).promise; priceCanvases=[];
    for(let i=1;i<=doc.numPages;i++){
      const pg=await doc.getPage(i), vp0=pg.getViewport({scale:1}), vp=pg.getViewport({scale:1588/vp0.width});
      const c=document.createElement("canvas"); c.width=vp.width; c.height=vp.height;
      await pg.render({canvasContext:c.getContext("2d"),viewport:vp,intent:"print"}).promise; priceCanvases.push(c);
    }
    pricePdf={name:file.name,bytes};
    $("pdfChip").innerHTML=`<div class="filechip"><span>${esc(file.name)} · ${doc.numPages} trang</span><button id="rmPdf">Bỏ</button></div>`;
    $("rmPdf").onclick=()=>{pricePdf=null;priceCanvases=[];$("pdfChip").innerHTML="";render();};
    status(""); render();
  }catch(e){ status("Không mở được file PDF này. Hãy thử file khác.",true); }
}

// ---- Biên bản sự việc (ảnh hoặc PDF) ----
async function setBienBan(file){
  if(!file) return;
  const isPdf=file.type==="application/pdf"||/\.pdf$/i.test(file.name);
  if(!isPdf&&!file.type.startsWith("image/")){status("Biên bản phải là hình hoặc file PDF.",true);return;}
  status("Đang đọc biên bản…");
  try{
    let pages=[];
    if(isPdf){
      const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
      for(let i=1;i<=doc.numPages;i++){
        const pg=await doc.getPage(i), vp0=pg.getViewport({scale:1}), vp=pg.getViewport({scale:1800/Math.max(vp0.width,vp0.height)});
        const c=document.createElement("canvas"); c.width=vp.width; c.height=vp.height;
        const ctx=c.getContext("2d"); ctx.fillStyle="#fff"; ctx.fillRect(0,0,c.width,c.height);
        await pg.render({canvasContext:ctx,viewport:vp,intent:"print"}).promise;
        pages.push({url:c.toDataURL("image/jpeg",0.9),w:c.width,h:c.height});
      }
    } else pages=[await shrink(file,2200)];
    bienBan={name:file.name,pages};
    $("bbChip").innerHTML=`<div class="filechip"><span>${esc(file.name)}${pages.length>1?` · ${pages.length} trang`:""}</span><button id="rmBB">Bỏ</button></div>`;
    $("rmBB").onclick=()=>{bienBan=null;$("bbChip").innerHTML="";render();};
    status(""); render();
  }catch(e){ status("Không mở được file biên bản này. Hãy thử file khác.",true); }
}

// ---- Drop zones ----
function wireDrop(zone,input,handler){
  zone.onclick=()=>input.click();
  zone.onkeydown=e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();input.click();}};
  input.onchange=()=>{handler([...input.files]);input.value="";};
  zone.ondragover=e=>{e.preventDefault();zone.classList.add("over")};
  zone.ondragleave=()=>zone.classList.remove("over");
  zone.ondrop=e=>{e.preventDefault();zone.classList.remove("over");handler([...e.dataTransfer.files]);};
}
wireDrop($("dropImg"),$("fImg"),addPhotos);
wireDrop($("dropBB"),$("fBB"),f=>setBienBan(f[0]));
wireDrop($("dropPdf"),$("fPdf"),f=>setPdf(f[0]));

function status(t,err){ $("status").textContent=t; $("status").classList.toggle("err",!!err); }

// ---- Export ----

$("export").onclick=async()=>{
  const btn=$("export"); btn.disabled=true;
  try{
    status("Đang tạo PDF…");
    const {PDFDocument}=PDFLib, out=await PDFDocument.create();
    const stage=$("stage"); stage.innerHTML="";
    for(const pg of document.querySelectorAll("#scaler .page.gen")){
      const clone=pg.cloneNode(true); stage.appendChild(clone);
      const cv=await html2canvas(clone,{scale:2,backgroundColor:"#ffffff",useCORS:true,logging:false});
      stage.removeChild(clone);
      const jpg=await out.embedJpg(cv.toDataURL("image/jpeg",0.9));
      const p=out.addPage([595.28,841.89]); p.drawImage(jpg,{x:0,y:0,width:595.28,height:841.89});
    }
    if(pricePdf){
      const src=await PDFDocument.load(pricePdf.bytes,{ignoreEncryption:true});
      (await out.copyPages(src,src.getPageIndices())).forEach(p=>out.addPage(p));
    }
    const bytes=await out.save();
    const name=`Thong bao phat heo chet ngay ${fmtDate($("ngay").value).replace(/\//g,".")} - ${noAccent($("ncc").value.trim())}.pdf`;
    const url=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));
    const a=document.createElement("a"); a.href=url; a.download=name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),10000);
    status(`Đã lưu ${name}`);
  }catch(e){
    status("Không tạo được PDF: "+(e?.message||e),true);
  }finally{ btn.disabled=false; }
};

render();
initMaster();
