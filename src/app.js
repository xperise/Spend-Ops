import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { SB } from './supabase.js';
window.XLSX = XLSX; window.JSZip = JSZip;
/* ===================== Tiện ích ===================== */
const KEY = 'mlx-spend-proto-v7';
const $ = (s, r=document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => Math.round(n || 0).toLocaleString('vi-VN');
const money = n => fmt(n) + ' đ';
const short = n => n >= 1e9 ? (n/1e9).toLocaleString('vi-VN',{maximumFractionDigits:2}) + ' tỷ'
  : n >= 1e6 ? (n/1e6).toLocaleString('vi-VN',{maximumFractionDigits:1}) + ' triệu'
  : fmt(n) + ' đ';
const digits = s => Number(String(s||'').replace(/\D/g,'')) || 0;
const uid = p => p + Math.random().toString(36).slice(2,8);
const now = () => Date.now();
const H = 3600e3;
const p2 = n => String(n).padStart(2,'0');
function dt(t){ if(!t) return '—'; const d=new Date(t); return `${p2(d.getDate())}/${p2(d.getMonth()+1)} ${p2(d.getHours())}:${p2(d.getMinutes())}`; }
function dOnly(t){ if(!t) return '—'; const d=new Date(t); return `${p2(d.getDate())}/${p2(d.getMonth()+1)}/${d.getFullYear()}`; }
function ago(t){ const m=Math.round((now()-t)/60000); if(m<1) return 'vừa xong'; if(m<60) return m+' phút trước'; const h=Math.round(m/60); if(h<24) return h+' giờ trước'; return Math.round(h/24)+' ngày trước'; }
const clone = o => JSON.parse(JSON.stringify(o));
const vnJoin = (arr, last=' và ') => arr.length < 2 ? arr.join('') : arr.slice(0,-1).join(', ') + last + arr[arr.length-1];

function initials(name){ const w = name.split(' ').filter(Boolean); return (w.length>1 ? w[w.length-2][0]+w[w.length-1][0] : w[0].slice(0,2)).toUpperCase(); }
function avatar(id, size=''){ const u = user(id); if (!u) return ''; return `<span class="av ${size} ${id===S.me?'me':''}" aria-hidden="true">${esc(initials(u.name))}</span>`; }

const IC = {
  home:'<path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
  file:'<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h4"/>',
  check:'<path d="M9 11l3 3 8-8"/><path d="M20 12v6a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h9"/>',
  wallet:'<path d="M3 7a2 2 0 0 1 2-2h13v4"/><path d="M3 7v11a2 2 0 0 0 2 2h15V9H5a2 2 0 0 1-2-2z"/><circle cx="16" cy="14.5" r="1.2"/>',
  gear:'<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  plus:'<path d="M12 5v14M5 12h14"/>',
  x:'<path d="M6 6l12 12M18 6 6 18"/>',
  chev:'<path d="M8 10l4-4 4 4M8 14l4 4 4-4"/>',
  copy:'<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V6a2 2 0 0 1 2-2h9"/>',
  clip:'<path d="m21 11-8.6 8.6a5 5 0 0 1-7-7L14 4a3.3 3.3 0 0 1 4.7 4.7L10 17.3a1.7 1.7 0 0 1-2.4-2.4L15.5 7"/>',
  upload:'<path d="M12 16V4M7 9l5-5 5 5M4 20h16"/>',
  swap:'<path d="M7 7h11l-3-3M17 17H6l3 3"/>',
  chart:'<path d="M4 20V10M10 20V4M16 20v-7M21 20H3"/>',
  users:'<circle cx="9" cy="8" r="3.2"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><path d="M16 4.5a3 3 0 0 1 0 6M21 20c0-2.6-1.6-4.9-4-5.7"/>'
};
const icon = (n, s=18) => `<svg viewBox="0 0 24 24" width="${s}" height="${s}" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${IC[n]}</svg>`;

/* ===================== Dữ liệu mẫu ===================== */
const STATUS = {
  draft:['Nháp',''], pending:['Chờ duyệt','st-pending'], returned:['Trả lại bổ sung','st-returned'],
  rejected:['Từ chối','st-rejected'], approved:['Chờ thanh toán','st-approved'], in_batch:['Trong đợt chi','st-in_batch'],
  paid:['Đã chi','st-paid'], cancelled:['Đã huỷ','st-cancelled']
};
const BSTATUS = {
  review:['Chờ duyệt đợt','st-review'], approved:['Đã duyệt, chờ chi','st-signed'], paying:['Đang chi','st-paying'],
  paid:['Đã chi xong','st-paid'], returned:['Trả lại kế toán','st-returned'], cancelled:['Đã huỷ','st-cancelled']
};
const EVIDENCE = {
  vat:['Hoá đơn VAT','Khấu trừ được VAT'],
  sale:['Hoá đơn bán hàng','Không khấu trừ VAT'],
  bill:['Chỉ có bill','Không hoá đơn hợp lệ'],
  later:['Chưa có hoá đơn','Có ĐNTT, bổ sung hoá đơn sau khi chi'],
  advance:['Tạm ứng','Không cần hoá đơn, quyết toán sau']
};
const OUTLAY = {transfer:'Chuyển khoản', cash:'Tiền mặt', card:'Thẻ'};
const PAYM = {transfer:'Chuyển khoản', visa:'Visa công ty', cash:'Tiền mặt'};
const FKIND = {vat:'Hoá đơn VAT', sale:'Hoá đơn bán hàng', dntt:'Đề nghị thanh toán', contract:'Hợp đồng', quote:'Báo giá', accept:'Biên bản nghiệm thu', bill:'Bill, biên nhận', statement:'Sao kê ngân hàng', other:'Khác'};
const RATES = {'10':'10%', '8':'8%', '5':'5%', '0':'0%', 'kct':'KCT'};
const CATS0 = [
  ['Văn phòng', ['Văn phòng phẩm','Điện thoại · internet','Thuê văn phòng · điện nước']],
  ['Đi lại và công tác', ['Đi lại · taxi','Công tác phí','Thuê xe · vận chuyển']],
  ['Công nghệ', ['Phần mềm & SaaS','Thiết bị CNTT']],
  ['Kinh doanh và đối ngoại', ['Tiếp khách & đối ngoại','Marketing · quảng cáo','Sự kiện · đào tạo']],
  ['Nhân sự', ['Phúc lợi · team building','Tuyển dụng']],
  ['Khác', ['Phí ngân hàng','Khác']]
];
const REFUND_CAT = 'Hoàn tiền khách chuyển nhầm';
const cats = () => S.cats;
const catGroup = c => c === REFUND_CAT ? 'Không phải chi phí' : (S.cats.find(g => g[1].includes(c)) || ['Chưa phân nhóm'])[0];
const catOptions = sel => S.cats.filter(g => g[1].length).map(g => `<optgroup label="${esc(g[0])}">${g[1].map(c => `<option ${c===sel?'selected':''}>${esc(c)}</option>`).join('')}</optgroup>`).join('') + (sel && sel !== REFUND_CAT && !S.cats.some(g => g[1].includes(sel)) ? `<option selected>${esc(sel)}</option>` : '');
const INV_DAYS = 7;
const ADV_DAYS = 15; // hạn quyết toán tạm ứng, tính từ ngày chi
function guessKind(name){ const n = noAccent(fname(name)).toLowerCase();
  if (/\.xml$/.test(n) || /gtgt|hd_|hoa_?don|invoice|vat/.test(n)) return 'vat';
  if (/dntt|de_?nghi/.test(n)) return 'dntt'; if (/hop_?dong|contract/.test(n)) return 'contract'; if (/bao_?gia|quote/.test(n)) return 'quote';
  if (/nghiem_?thu/.test(n)) return 'accept'; if (/sao_?ke|statement/.test(n)) return 'statement'; if (/bill|receipt|bien_?nhan|\.(png|jpe?g|heic)$/.test(n)) return 'bill';
  return 'other'; }
function rateNum(k){ return k === 'kct' ? 0 : Number(k) / 100; }

function defaultBatch(){ return {bands:[
  {id:'bb1', max:100000000, levels:[{approvers:['u7','u8'], mode:'any'}]},
  {id:'bb2', max:null, levels:[{approvers:['u7'], mode:'any'},{approvers:['u8'], mode:'any'}]}
]}; }
function seedBase(){
  return {
    me:'u1', guideDone:{},
    depts:[
      {id:'OPS',name:'Vận hành'},{id:'SAL',name:'Kinh doanh'},{id:'FIN',name:'Tài chính – Kế toán'},
      {id:'TEC',name:'Công nghệ'},{id:'BOD',name:'Ban điều hành'}
    ],
    users:[
      {id:'u1',name:'Nguyễn Minh Anh',title:'Chuyên viên vận hành',dept:'OPS',acct:false,signer:false,admin:false,active:true},
      {id:'u2',name:'Trần Quốc Bảo',title:'Chuyên viên kinh doanh',dept:'SAL',acct:false,signer:false,admin:false,active:true},
      {id:'u3',name:'Lê Thu Hà',title:'Trưởng phòng Vận hành',dept:'OPS',acct:false,signer:false,admin:false,active:true},
      {id:'u4',name:'Phạm Đức Long',title:'Trưởng phòng Kinh doanh',dept:'SAL',acct:false,signer:false,admin:false,active:true},
      {id:'u5',name:'Võ Ngọc Mai',title:'Kế toán trưởng',dept:'FIN',acct:true,signer:false,admin:false,active:true},
      {id:'u6',name:'Đặng Hoàng Nam',title:'Kế toán thanh toán',dept:'FIN',acct:true,signer:false,admin:false,active:true},
      {id:'u7',name:'Hồ Thanh Tùng',title:'Tổng giám đốc',dept:'BOD',acct:false,signer:true,admin:false,active:true},
      {id:'u8',name:'Bùi Khánh Linh',title:'Phó tổng giám đốc',dept:'BOD',acct:false,signer:true,admin:false,active:true},
      {id:'u9',name:'Quản trị hệ thống',title:'Admin Spend Ops',dept:'TEC',acct:false,signer:false,admin:true,active:true}
    ],
    vendors:[
      {id:'v1',name:'Công ty TNHH Văn phòng phẩm Hồng Hà',tax:'0312345678',accName:'CONG TY TNHH VAN PHONG PHAM HONG HA',bank:'Vietcombank',acc:'0071000123456',verified:true},
      {id:'v2',name:'Công ty CP Phần mềm MISA',tax:'0101243150',accName:'CONG TY CO PHAN MISA',bank:'BIDV',acc:'31010000987654',verified:true},
      {id:'v3',name:'Vận tải Phương Nam',tax:'0309876543',accName:'CONG TY TNHH VAN TAI PHUONG NAM',bank:'MSB',acc:'03401010099887',verified:true},
      {id:'v4',name:'VNPT TP.HCM',tax:'0100684378-001',accName:'VIEN THONG TP HO CHI MINH',bank:'VietinBank',acc:'112000556677',verified:true},
      {id:'v5',name:'Nhà hàng Hương Sen',tax:'0314567890',accName:'HO KINH DOANH HUONG SEN',bank:'ACB',acc:'24681357',verified:false},
      {id:'v6',name:'Công ty CP Mắt Bão',tax:'0302712571',accName:'CONG TY CO PHAN MAT BAO',bank:'Techcombank',acc:'MATBAO8899',verified:true},
      {id:'v7',name:'Công ty TNHH Quảng cáo Sao Việt',tax:'0316789012',accName:'CONG TY TNHH QUANG CAO SAO VIET',bank:'MB Bank',acc:'0988776655',verified:true}
    ],
    cats: JSON.parse(JSON.stringify(CATS0)),
    userBanks:{u1:{bank:'Techcombank',acc:'19033445566017'},u2:{bank:'MB Bank',acc:'0901234567'}},
    config:{
      version:1, updatedAt: now()-72*H, updatedBy:'u9', batch:defaultBatch(),
      rules:[
        {id:'rOPS', depts:['OPS'], people:[], bands:[
          {id:'b1',max:5000000,levels:[{approvers:['u3'],mode:'any'}]},
          {id:'b2',max:30000000,levels:[{approvers:['u3'],mode:'any'},{approvers:['u5'],mode:'any'}]},
          {id:'b3',max:null,levels:[{approvers:['u3'],mode:'any'},{approvers:['u5'],mode:'any'},{approvers:['u7','u8'],mode:'any'}]}]},
        {id:'rSAL', depts:['SAL'], people:[], bands:[
          {id:'b4',max:5000000,levels:[{approvers:['u4'],mode:'any'}]},
          {id:'b5',max:30000000,levels:[{approvers:['u4'],mode:'any'},{approvers:['u5'],mode:'any'}]},
          {id:'b6',max:null,levels:[{approvers:['u4'],mode:'any'},{approvers:['u5'],mode:'any'},{approvers:['u7','u8'],mode:'any'}]}]},
        {id:'rDEF', def:true, depts:[], people:[], bands:[
          {id:'b7',max:5000000,levels:[{approvers:['u7'],mode:'any'}]},
          {id:'b8',max:30000000,levels:[{approvers:['u7'],mode:'any'},{approvers:['u5'],mode:'any'}]},
          {id:'b9',max:null,levels:[{approvers:['u7'],mode:'any'},{approvers:['u5'],mode:'any'},{approvers:['u8'],mode:'any'}]}]}
      ],
      blockSelf:true, fallback:['u7'], noInvoice:{on:true,approvers:['u5']}, cashLimit:5000000
    },
    reqs:[], batches:[], delegations:[], seq:{A:0,PB:0}
  };
}

function seed(){
  S = seedBase();
  const t0 = now();
  const mk = (o, by, at) => { const r = newReq(Object.assign({requester:by}, o)); r.createdAt = at; S.reqs.push(r); return r; };
  const r6 = mk({title:'Internet văn phòng tháng 10/2026',amount:2400000,dept:'OPS',cat:'Điện thoại · internet',evidence:'vat',inv:'0045871',payee:{type:'vendor',vendorId:'v4'},files:['HD_VNPT_T10.pdf']}, 'u1', t0-170*H);
  submitReq(r6,'u1',t0-170*H); act(r6,'u3','approve','',t0-168*H);
  const r5 = mk({title:'Thuê xe đưa đón sự kiện khách hàng VCB',amount:22000000,dept:'OPS',cat:'Thuê xe · vận chuyển',evidence:'vat',payee:{type:'vendor',vendorId:'v3'},inv:'0000018',rate:'8',files:['Bao_gia_Phuong_Nam.pdf','HD_GTGT_0018.xml']}, 'u1', t0-60*H);
  submitReq(r5,'u1',t0-60*H); act(r5,'u4','approve','Đúng kế hoạch sự kiện tháng 11.',t0-58*H,'u3'); act(r5,'u5','approve','',t0-50*H);
  const r9 = mk({title:'Cước điện thoại phòng kinh doanh tháng 10/2026',amount:1650000,dept:'SAL',cat:'Điện thoại · internet',evidence:'vat',payee:{type:'vendor',vendorId:'v4'},files:['HD_VNPT_SAL_T10.pdf']}, 'u2', t0-40*H);
  submitReq(r9,'u2',t0-40*H); act(r9,'u4','approve','',t0-38*H);
  const r4 = mk({title:'Taxi đi gặp đối tác tại Quận 7',amount:850000,dept:'OPS',cat:'Đi lại · taxi',evidence:'bill',payee:{type:'me',outlay:'card'},files:['grab_receipt.png']}, 'u1', t0-30*H);
  submitReq(r4,'u1',t0-30*H); act(r4,'u3','approve','',t0-28*H); act(r4,'u5','approve','Trong hạn mức đi lại.',t0-27*H);
  const r1 = mk({title:'Văn phòng phẩm quý 4/2026',amount:3200000,dept:'OPS',cat:'Văn phòng phẩm',evidence:'vat',payee:{type:'vendor',vendorId:'v1'},files:['Bao_gia_HongHa.pdf'],note:'Giấy in, mực, bìa hồ sơ cho 3 bộ phận.'}, 'u1', t0-3*H);
  submitReq(r1,'u1',t0-3*H);
  const r2 = mk({title:'Gia hạn phần mềm kế toán MISA năm 2027',amount:48000000,dept:'FIN',cat:'Phần mềm & SaaS',evidence:'vat',payee:{type:'vendor',vendorId:'v2'},files:['HopDong_MISA_2027.pdf']}, 'u6', t0-26*H);
  submitReq(r2,'u6',t0-26*H); act(r2,'u7','approve','Đồng ý gia hạn, đề nghị đàm phán giảm 5%.',t0-20*H);
  const r3 = mk({title:'Tiếp khách đối tác Grab tại nhà hàng Hương Sen',amount:12500000,dept:'SAL',cat:'Tiếp khách & đối ngoại',evidence:'bill',payee:{type:'vendor',vendorId:'v5'},files:['bill_huongsen.jpg']}, 'u2', t0-22*H);
  submitReq(r3,'u2',t0-22*H); act(r3,'u4','return','Bổ sung danh sách khách tham dự và xin hoá đơn VAT từ nhà hàng — họ có xuất được.',t0-18*H);
  const r8 = mk({title:'Mua 6 ghế công thái học phòng kinh doanh',amount:9000000,dept:'SAL',cat:'Khác',evidence:'later',payee:{type:'vendor',vendorId:'v1'},files:['Bao_gia_ghe.pdf']}, 'u2', t0-80*H);
  submitReq(r8,'u2',t0-80*H); act(r8,'u4','reject','Ngoài ngân sách Q4, đưa vào kế hoạch 2027.',t0-75*H);
  mk({title:'In tài liệu đào tạo tài xế mới',amount:1800000,dept:'OPS',cat:'Sự kiện · đào tạo',evidence:'later',payee:{type:'vendor',vendorId:'v1'},files:[]}, 'u1', t0-1*H);
  const r13 = mk({title:'Đặt cọc 50% thi công biển hiệu văn phòng SATRA',amount:15000000,invTotal:30000000,rate:'8',dept:'OPS',cat:'Marketing · quảng cáo',evidence:'later',payee:{type:'vendor',vendorId:'v7'},files:['DNTT_SaoViet_dot1.pdf','Hop_dong_bien_hieu.pdf'],note:'Thanh toán đợt 1 theo hợp đồng, NCC xuất hoá đơn khi thanh toán đủ.'}, 'u1', t0-175*H);
  submitReq(r13,'u1',t0-175*H); act(r13,'u3','approve','',t0-172*H); act(r13,'u5','approve','',t0-170*H);
  const pb0 = createBatch([r6.id, r13.id],'u6',t0-160*H); approveBatch(pb0,'u7','',t0-150*H); payLines(pb0,[r6.id],'u6',['UNC_PB-0001_01_SP-0001.pdf'],t0-140*H); payLines(pb0,[r13.id],'u6',['UNC_PB-0001_02_'+r13.code+'.pdf'],t0-140*H);
  const r14 = mk({title:'Gia hạn tên miền và hosting xperise.com 2027',amount:3300000,dept:'TEC',cat:'Phần mềm & SaaS',evidence:'vat',inv:'0211457',payMethod:'visa',payee:{type:'vendor',vendorId:'v6'},files:['HD_MatBao_0211457.pdf']}, 'u9', t0-36*H);
  submitReq(r14,'u9',t0-36*H); act(r14,'u7','approve','',t0-35*H);
  const r10 = mk({title:'Mực in và giấy A4 tháng 11/2026',amount:2850000,dept:'OPS',cat:'Văn phòng phẩm',evidence:'vat',payee:{type:'vendor',vendorId:'v1'},inv:'0001102',rate:'8',files:['HD_HongHa_1102.pdf']}, 'u1', t0-34*H);
  submitReq(r10,'u1',t0-34*H); act(r10,'u3','approve','',t0-33*H);
  const r11 = mk({title:'Cước vận chuyển hồ sơ khách hàng tháng 11',amount:4600000,dept:'SAL',cat:'Thuê xe · vận chuyển',evidence:'vat',payee:{type:'vendor',vendorId:'v3'},inv:'0000412',rate:'8',files:['HD_PhuongNam_0412.pdf']}, 'u2', t0-32*H);
  submitReq(r11,'u2',t0-32*H); act(r11,'u4','approve','',t0-31*H);
  const r12 = mk({title:'Thuê bao phần mềm hoá đơn điện tử 2027',amount:8400000,dept:'TEC',cat:'Phần mềm & SaaS',evidence:'vat',payee:{type:'vendor',vendorId:'v2'},inv:'0098231',files:['HD_MISA_eInvoice.pdf']}, 'u9', t0-29*H);
  submitReq(r12,'u9',t0-29*H); act(r12,'u7','approve','',t0-27*H); act(r12,'u5','approve','',t0-26*H);
  createBatch([r5.id, r9.id, r10.id, r11.id, r12.id, r14.id],'u6',t0-6*H, localDate(t0 + 24*H));
  const dd = n => localDate(t0 + n*24*H);
  S.delegations.push({id:uid('g'), from:'u3', to:'u4', start:dd(-4), end:dd(-2), note:'Nghỉ phép', createdAt:t0-100*H, revokedAt:null});
  S.delegations.push({id:uid('g'), from:'u5', to:'u3', start:dd(-1), end:dd(4), note:'Công tác Hà Nội', createdAt:t0-24*H, revokedAt:null});
  const tmA = (days) => t0 - days*24*H;
  const rA = mk({title:'Tạm ứng công tác Đà Nẵng tháng 10',amount:5000000,dept:'SAL',cat:'Đi lại · taxi',evidence:'advance',payee:{type:'me',outlay:'transfer'},files:[],lines:[]}, 'u2', tmA(21));
  submitReq(rA,'u2',tmA(21)); act(rA,'u4','approve','',tmA(20.9));
  const pbA = createBatch([rA.id],'u6',tmA(20.5)); approveBatch(pbA,'u7','',tmA(20.4)); payLines(pbA,[rA.id],'u6',['UNC_TAMUNG_0001.pdf'],tmA(20));
  const rB = mk({title:'Tạm ứng mua vật tư sự kiện tháng 11',amount:3000000,dept:'OPS',cat:'Văn phòng phẩm',evidence:'advance',payee:{type:'me',outlay:'cash'},files:[],lines:[]}, 'u1', tmA(3));
  submitReq(rB,'u1',tmA(3)); act(rB,'u3','approve','',tmA(2.9));
  const pbB = createBatch([rB.id],'u6',tmA(2)); approveBatch(pbB,'u7','',tmA(1.9)); payLines(pbB,[rB.id],'u6',['UNC_TAMUNG_0002.pdf'],tmA(1));
  S.me = 'u1';
  save();
}

/* ===================== Trạng thái ===================== */
let S;
let UI = {view:'home', id:null, listTab:'mine', cfgTab:'request', draft:null, sim:null, form:null, errors:{}, inboxSel:{}, batchSel:{}, payDraft:null, flt:{q:'',from:'',to:'',pm:'',cat:'',who:''}, bflt:{q:'',from:'',to:'',pm:''}, qflt:{pm:''}, rflt:{q:'',from:'',to:'',pm:'',cat:'',who:''}, cfgRule:null, vq:'', rtab:'history', paySel:{}, bpm:''};
/* save(): ghi lên Supabase sau 0,4 giây không thao tác (xem khối Lưu trữ ở cuối file). */
function save(){ if (!S || !SB) return; dirty = true; clearTimeout(saveTimer); saveTimer = setTimeout(pushState, 400); }

const user = id => S.users.find(u => u.id === id);
const dept = id => S.depts.find(d => d.id === id);
const vendor = id => S.vendors.find(v => v.id === id);
const me = () => user(S.me);
const req = id => S.reqs.find(r => r.id === id);
const batch = id => S.batches.find(b => b.id === id);
function log(r, by, text, at){ r.log.push({at:at||now(), by, text}); }


/* ===================== Bộ máy chuỗi duyệt ===================== */
function sortedBands(cfg){ return [...cfg.bands].sort((a,b)=> (a.max==null) - (b.max==null) || (a.max||0)-(b.max||0)); }
function bandFor(amount, cfg){ const bs = sortedBands(cfg); return bs.find(b => b.max == null || amount <= b.max) || bs[bs.length-1]; }
function bandRange(b, cfg){ const bs = sortedBands(cfg); const i = bs.indexOf(b); const lo = i>0 ? bs[i-1].max : 0;
  if (b.max == null) return lo ? `trên ${short(lo)}` : 'mọi mức tiền';
  return lo ? `${short(lo)} – ${short(b.max)}` : `đến ${short(b.max)}`; }

/**
 * Dựng chuỗi duyệt cho một đề xuất theo cấu hình.
 * Thứ tự xử lý ở mỗi cấp: danh sách người → lọc cùng phòng → bỏ người đề xuất → bỏ người đã duyệt bước trước.
 * Cấp nào rỗng sau khi lọc thì nhờ người duyệt dự phòng.
 */
/* Chọn quy định: người đề xuất cụ thể > phòng chịu chi phí > mặc định. */
function ruleFor(input, cfg){
  return cfg.rules.find(r => !r.def && r.people.includes(input.requester)) || cfg.rules.find(r => !r.def && r.depts.includes(input.dept)) || cfg.rules.find(r => r.def) || cfg.rules[0];
}
function ruleLabel(r){ if (r.def) return 'Mặc định';
  const parts = r.depts.map(id => (dept(id) || {}).name || id).concat(r.people.map(id => (user(id) || {}).name || '?')); return parts.length ? parts.join(', ') : 'Chưa chọn đối tượng'; }
function buildChain(input, cfg, users){
  users = users || S.users;
  const U = id => users.find(u => u.id === id);
  const ok = id => { const u = U(id); return u && u.active && !u.deleted; };
  const rule = ruleFor(input, cfg); const band = bandFor(input.amount, rule);
  const used = new Set(); const steps = []; const problems = [];
  const fallback = () => cfg.fallback.filter(id => ok(id) && id !== input.requester && !used.has(id));
  band.levels.forEach((lv, i) => {
    let cand = lv.approvers.filter(ok); const notes = [];
    if (lv.sameDept) {
      const same = cand.filter(id => U(id).dept === input.dept);
      if (same.length) cand = same;
      else { cand = fallback(); notes.push('Phòng chịu chi phí chưa có người duyệt ở bước này — người duyệt dự phòng ký thay.'); }
    }
    if (cfg.blockSelf && cand.includes(input.requester)) {
      cand = cand.filter(id => id !== input.requester);
      if (!cand.length) { cand = fallback(); notes.push('Người đề xuất là người duyệt của bước này — chuyển người duyệt dự phòng.'); }
      else notes.push('Đã bỏ người đề xuất khỏi danh sách duyệt bước này.');
    }
    const fresh = cand.filter(id => !used.has(id));
    if (fresh.length && fresh.length < cand.length) { cand = fresh; notes.push('Đã bỏ người đã duyệt ở bước trước.'); }
    else if (!fresh.length && cand.length) {
      const fb = fallback();
      if (fb.length) { cand = fb; notes.push('Người duyệt bước này đã ký ở bước trước — chuyển người duyệt dự phòng.'); }
      else notes.push('Trùng người duyệt bước trước, không còn người thay.');
    }
    if (!cand.length) problems.push(`Cấp ${i+1}: chưa có người duyệt hợp lệ.`);
    cand.forEach(id => used.add(id));
    steps.push({key:'L'+(i+1), label:`Cấp ${i+1}`, mode: cand.length > 1 ? lv.mode : 'any', approvers:cand, notes});
  });
  if (cfg.noInvoice.on && input.evidence === 'bill') {
    const all = cfg.noInvoice.approvers.filter(id => ok(id) && id !== input.requester);
    const fresh = all.filter(id => !used.has(id));
    if (!all.length) problems.push('Kiểm soát chi không hoá đơn: chưa có người duyệt.');
    else if (!fresh.length) steps.forEach(s => { if (s.approvers.some(id => all.includes(id))) s.notes.push('Khoản không hoá đơn — người kiểm soát đã có ở bước này.'); });
    else steps.push({key:'NI', label:'Kiểm soát chi không hoá đơn', mode:'any', approvers:fresh, notes:['Thêm bước vì khoản chi không có hoá đơn hợp lệ.']});
  }
  return {rule, band, steps, problems};
}

function currentStep(r){ const rd = r.rounds[r.rounds.length-1]; if (!rd) return null; return rd.chain.find(s => s.status === 'current') || null; }
/* Ủy quyền: người được ủy quyền ký thay vào ô của người ủy quyền; một người chỉ ký một lần mỗi bước và không ký thay nếu đã ký ở bước trước. */
function localDate(t){ const d = new Date(t || now()); return `${d.getFullYear()}-${p2(d.getMonth()+1)}-${p2(d.getDate())}`; }
function delegActive(g, t){ if (g.revokedAt) return false; const d = localDate(t); if (g.start && d < g.start) return false; if (g.end && d > g.end) return false; return true; }
function delegState(g){ if (g.revokedAt) return ['Đã thu hồi','st-cancelled']; const d = localDate(); if (g.start && d < g.start) return ['Chưa tới ngày','']; if (g.end && d > g.end) return ['Hết hạn','']; return ['Đang hiệu lực','st-approved']; }
const principalsOf = id => (S.delegations||[]).filter(g => g.to === id && delegActive(g)).map(g => g.from);
const delegatesOf = id => (S.delegations||[]).filter(g => g.from === id && delegActive(g));
function actorsFor(r, id){
  if (r.status !== 'pending') return [];
  if (r.requester === id && S.config.blockSelf) return [];
  const s = currentStep(r); if (!s) return [];
  if (Object.values(s.decisions).some(d => (d.by || '') === id)) return [];
  const rd = r.rounds[r.rounds.length-1];
  const signedEarlier = rd.chain.some(x => x !== s && Object.values(x.decisions).some(d => d.by === id));
  const out = [];
  if (s.approvers.includes(id) && !s.decisions[id]) out.push(id);
  if (!signedEarlier && r.requester !== id) principalsOf(id).forEach(p => { if (s.approvers.includes(p) && !s.decisions[p] && p !== r.requester && !out.includes(p)) out.push(p); });
  return out;
}
const canAct = (r, id) => actorsFor(r, id).length > 0;
function waitingOn(r){ const s = currentStep(r); if (!s) return []; return s.approvers.filter(id => !s.decisions[id]); }

/* ===================== Hành động trên đề xuất ===================== */
function newReq(o){
  const r = Object.assign({id:uid('r'), code:null, type:'opex', title:'', amount:0, dept:'OPS', cat:'Khác', evidence:'vat', payMethod:'transfer', payee:{type:'vendor',vendorId:'v1',outlay:'transfer'}, needBy:'', note:'', files:[], lines:null, requester:S.me, status:'draft', createdAt:now(), rounds:[], comments:[], log:[], batchId:null, unc:[]}, o);
  r.files = r.files.map(f => typeof f === 'string' ? {name:f, kind:guessKind(f)} : f);
  if (!r.lines) { // dựng một dòng tiền thuế từ số tổng (dữ liệu mẫu, khoản chi trực tiếp)
    const tot = r.invTotal || r.amount; const rate = r.rate || (r.evidence === 'vat' || r.evidence === 'later' ? '10' : 'kct');
    const pre = Math.round(tot / (1 + rateNum(rate)));
    r.lines = tot ? [{inv:r.inv || '', desc:r.title, pre, rate, tax:tot - pre, total:tot}] : [];
  }
  delete r.inv; delete r.rate; delete r.invTotal;
  return r;
}
const linesSum = (ls, k) => (ls || []).reduce((a,l) => a + (l[k] || 0), 0);
const invNos = r => [...new Set((r.lines || []).map(l => (l.inv || '').trim()).filter(Boolean))];
function nextCode(p){ S.seq[p] = (S.seq[p]||0) + 1; return (p==='A' ? 'SP-' : 'PB-') + String(S.seq[p]).padStart(4,'0'); }
function submitReq(r, by, at){
  at = at || now();
  if (!r.code) r.code = nextCode('A');
  const res = buildChain({amount:r.amount, dept:r.dept, requester:r.requester, evidence:r.evidence}, S.config);
  const chain = res.steps.map((s,i) => Object.assign({}, s, {decisions:{}, status: i===0 ? 'current' : 'waiting'}));
  const round = {no:r.rounds.length+1, at, cfgVersion:S.config.version, band:bandRange(res.band, res.rule), rule:ruleLabel(res.rule), chain};
  r.rounds.push(round);
  r.status = 'pending'; r.submittedAt = at;
  log(r, by, round.no > 1 ? `Gửi lại lần ${round.no} sau khi bổ sung` : 'Gửi đề xuất', at);
}
function act(r, by, kind, note, at, principal){
  at = at || now(); principal = principal || by;
  const s = currentStep(r); if (!s) return;
  s.decisions[principal] = {act:kind, note, at, by};
  const rd = r.rounds[r.rounds.length-1];
  const who = principal !== by ? ` thay ${user(principal).name}` : '';
  if (kind === 'approve') {
    const done = s.mode === 'all' ? s.approvers.every(id => s.decisions[id] && s.decisions[id].act === 'approve') : true;
    log(r, by, `Duyệt ${s.label.toLowerCase()}${who}${note ? ': ' + note : ''}`, at);
    if (done) {
      s.status = 'approved'; s.doneAt = at;
      const nx = rd.chain.find(x => x.status === 'waiting');
      if (nx) nx.status = 'current';
      else { r.status = 'approved'; r.approvedAt = at; log(r, null, 'Đủ chữ ký, chuyển kế toán lập đợt chi', at); }
    }
  } else if (kind === 'reject') {
    s.status = 'rejected'; r.status = 'rejected'; log(r, by, `Từ chối${who}: ${note}`, at);
  } else if (kind === 'return') {
    s.status = 'returned'; r.status = 'returned'; log(r, by, `Trả lại bổ sung${who}: ${note}`, at);
  }
}
function cancelReq(r, by){ r.status = 'cancelled'; log(r, by, 'Huỷ đề xuất'); }


/* ===================== Đợt chi ===================== */
function blog(b, by, text, at){ (b.log = b.log || []).push({at:at||now(), by, text}); }
/* Chuỗi duyệt đợt chi: tách riêng khỏi đề xuất, chọn theo tổng tiền đợt (cấu hình "Đợt chi"). */
function buildBatchChain(total, requester, cfg){
  const B = (cfg || S.config).batch; const band = bandFor(total, B);
  const ok = id => { const u = user(id); return !!(u && u.active && !u.deleted && u.signer && id !== requester); };
  const used = new Set(); const steps = []; const problems = [];
  band.levels.forEach((lv, i) => {
    const cand = lv.approvers.filter(id => ok(id) && !used.has(id));
    if (!cand.length) problems.push(`Đợt chi, cấp ${i+1}: chưa có người duyệt hợp lệ.`);
    cand.forEach(id => used.add(id));
    steps.push({key:'L'+(i+1), label:`Cấp ${i+1}`, mode: cand.length > 1 ? lv.mode : 'any', approvers:cand, notes:[], decisions:{}, status:'waiting'});
  });
  if (steps.length) steps[0].status = 'current';
  return {band, steps, problems};
}
function ensureBatchChain(b){
  if (!b.chain) b.chain = b.status === 'review' ? buildBatchChain(batchTotal(b), b.createdBy).steps : [];
  return b.chain;
}
function currentBatchStep(b){ return ensureBatchChain(b).find(s => s.status === 'current') || null; }
function batchChainText(b){
  const ch = ensureBatchChain(b);
  return ch.length ? ch.map((s,i) => `Cấp ${i+1}: ${vnJoin(s.approvers.map(id => (user(id) || {}).name || '?'), s.mode === 'all' ? ' và ' : ' hoặc ')}`).join(' → ') : 'Chưa có chuỗi duyệt';
}
function createBatch(ids, by, at, payDate){
  at = at || now();
  const total = ids.reduce((a,id) => a + ((req(id) || {}).amount || 0), 0);
  const plan = buildBatchChain(total, by);
  const b = {id:uid('b'), code:nextCode('PB'), reqIds:[...ids], status:'review', createdBy:by, createdAt:at, payDate: payDate || localDate(at), approvedBy:null, approvedAt:null, approveNote:'', lines:{}, removed:[], log:[], chain:plan.steps, cfgVersion:S.config.version};
  ids.forEach(id => { b.lines[id] = {paidAt:null, paidBy:null, unc:[]}; const r = req(id); r.status = 'in_batch'; r.batchId = b.id; log(r, by, `Đưa vào đợt chi ${b.code}, hẹn thanh toán ${dOnly(b.payDate + 'T00:00')}`, at); });
  blog(b, by, `Lập đợt ${ids.length} khoản, hẹn thanh toán ${dOnly(b.payDate + 'T00:00')}`, at);
  S.batches.push(b); return b;
}
function approveBatch(b, by, note, at, principal){
  at = at || now(); principal = principal || by;
  const s = currentBatchStep(b); if (!s) return;
  s.decisions[principal] = {act:'approve', note: note || '', at, by};
  const who = principal !== by ? ' thay ' + user(principal).name : '';
  blog(b, by, `Duyệt ${s.label.toLowerCase()} của đợt${who}${note ? ': ' + note : ''}`, at);
  const done = s.mode === 'all' ? s.approvers.every(id => s.decisions[id] && s.decisions[id].act === 'approve') : true;
  if (!done) return;
  s.status = 'approved'; s.doneAt = at;
  const nx = b.chain.find(x => x.status === 'waiting');
  if (nx) { nx.status = 'current'; return; }
  b.status = 'approved'; b.approvedBy = by; b.approvedFor = principal !== by ? principal : null; b.approvedAt = at; b.approveNote = note || '';
  b.reqIds.forEach(id => log(req(id), by, `Duyệt đợt chi ${b.code}${who}`, at));
  blog(b, by, `Đủ chữ ký, chuyển kế toán chi tiền${who}`, at);
}
function returnBatch(b, by, note, principal){
  principal = principal || by;
  const at = now(); b.status = 'returned'; b.returnedBy = by; b.returnedFor = principal !== by ? principal : null; b.returnedAt = at; b.returnNote = note;
  const who = b.returnedFor ? ' thay ' + user(principal).name : '';
  b.reqIds.forEach(id => { const r = req(id); if (b.lines[id].paidAt) return; r.status = 'approved'; r.batchId = null; log(r, by, `Trả lại đợt ${b.code}${who}: ${note}`, at); });
  blog(b, by, `Trả lại kế toán${who}: ${note}`, at);
}
function payLines(b, ids, by, filesFor, at){
  at = at || now();
  ids.forEach(id => { const files = typeof filesFor === 'function' ? filesFor(id) : filesFor; b.lines[id] = {paidAt:at, paidBy:by, unc:[...files]}; const r = req(id); r.status = 'paid'; r.paidAt = at; r.unc = [...files];
    if (r.evidence === 'later' && !r.invDone) r.invDue = at + INV_DAYS*24*H;
    if (r.evidence === 'advance' && !r.settleDue) r.settleDue = at + ADV_DAYS*24*H;
    log(r, by, `${r.payMethod === 'visa' ? 'Đã chi bằng thẻ Visa' : r.payMethod === 'cash' ? 'Đã chi tiền mặt' : 'Đã chuyển khoản'}, chứng từ ${files.join(', ')}`, at); });
  blog(b, by, `Xác nhận đã chi ${ids.length} khoản: ${ids.map(id => req(id).code).join(', ')}`, at);
  const all = b.reqIds.every(id => b.lines[id].paidAt);
  b.status = all ? 'paid' : 'paying'; if (all) b.paidAt = at;
}
/* Gỡ một khoản khỏi đợt: kind 'removed' trả về hàng chờ thanh toán, 'cancelled' huỷ hẳn lệnh. */
function dropLine(b, id, by, reason, kind){
  const r = req(id); if (!b || b.lines[id]?.paidAt) return;
  b.reqIds = b.reqIds.filter(x => x !== id); delete b.lines[id]; b.removed.push({id, by, at:now(), reason, kind});
  blog(b, by, `${kind === 'cancelled' ? 'Huỷ' : 'Gỡ'} ${r.code} khỏi đợt: ${reason}`);
  if (!b.reqIds.length) { b.status = 'cancelled'; blog(b, by, 'Đợt không còn khoản nào, tự huỷ'); }
  else if (b.reqIds.every(x => b.lines[x].paidAt)) { b.status = 'paid'; b.paidAt = now(); }
}
function cancelOrder(r, by, reason){
  const b = batch(r.batchId);
  if (b && r.status === 'in_batch') dropLine(b, r.id, by, reason, 'cancelled');
  r.status = 'cancelled'; r.batchId = null; r.cancelReason = reason; r.cancelledAt = now(); r.cancelledBy = by; log(r, by, `Huỷ lệnh: ${reason}`);
}
function cancelBatch(b, by, reason){
  b.reqIds.forEach(id => { const r = req(id); if (b.lines[id].paidAt) return; r.status = 'approved'; r.batchId = null; log(r, by, `Đợt ${b.code} bị huỷ: ${reason}. Khoản về lại chờ thanh toán`); });
  b.status = 'cancelled'; b.cancelReason = reason; blog(b, by, `Huỷ đợt: ${reason}`);
}
const canCancelOrder = (r, u) => ['pending','approved','in_batch','returned'].includes(r.status) && (r.requester === u.id || u.acct || u.admin);
const batchTotal = b => b.reqIds.reduce((a,id) => a + (req(id)?.amount || 0), 0);
const paidCount = b => b.reqIds.filter(id => b.lines[id]?.paidAt).length;
const noAccent = s => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D');
const bankText = s => noAccent(s).toUpperCase().replace(/[^A-Z0-9 .,\-\/]/g,' ').replace(/\s+/g,' ').trim();
function payeeOf(r){
  if (r.payee.type === 'me') { const u = user(r.requester); const ub = S.userBanks[r.requester] || {bank:'Vietcombank', acc:'0011002233445'};
    return {name:u.name, accName:bankText(u.name), note:'Hoàn ứng', bank:ub.bank, acc:ub.acc, ok:true}; }
  if (r.payee.type === 'other') return {name:r.payee.name, accName:bankText(r.payee.accName || r.payee.name), note:'Hoàn tiền', bank:r.payee.bank, acc:r.payee.acc, ok:true};
  const v = vendor(r.payee.vendorId) || {};
  return {name:v.name || '—', accName:v.accName || bankText(v.name || ''), note:'', bank:v.bank || '—', acc:v.acc || '—', tax:v.tax, ok:!!v.verified, vendorId:v.id};
}
/* Nội dung chuyển khoản chuẩn ngân hàng: IN HOA không dấu, "[Mã] MLX TT [nội dung] theo HD [số hoá đơn]". */
function memoOf(r){
  const inv = invNos(r);
  const core = r.type === 'refund' ? `${r.code} MLX HOAN TIEN ${r.title.replace(/^hoàn tiền\s*/i,'')}` : `${r.code} MLX TT ${r.title}${inv.length ? ' THEO HD ' + inv.slice(0,3).join(' ') + (inv.length > 3 ? ' VA ' + (inv.length-3) + ' HD KHAC' : '') : ''}`;
  return bankText(core).slice(0,140);
}
const uncName = (b, id) => `UNC_${b.code}_${p2(b.reqIds.indexOf(id) + 1)}_${req(id).code}.pdf`;
const activeUsers = () => S.users.filter(u => u.active && !u.deleted);
const isSigner = id => { const u = user(id); return !!(u && u.signer && u.active && !u.deleted); };
const signers = () => activeUsers().filter(u => u.signer).map(u => u.id);
function batchActors(b, id){
  if (b.status !== 'review' || id === b.createdBy) return [];
  const s = currentBatchStep(b); if (!s) return [];
  if (Object.values(s.decisions).some(d => (d.by || '') === id)) return [];
  const signedEarlier = ensureBatchChain(b).some(x => x !== s && Object.values(x.decisions).some(d => d.by === id));
  const out = [];
  if (s.approvers.includes(id) && !s.decisions[id]) out.push(id);
  if (!signedEarlier) principalsOf(id).forEach(p => { if (s.approvers.includes(p) && !s.decisions[p] && p !== b.createdBy && !out.includes(p)) out.push(p); });
  return out;
}
const canApproveBatch = (b, id) => batchActors(b, id).length > 0;
const batchesToApprove = id => S.batches.filter(b => canApproveBatch(b, id));
const batchesToPay = () => S.batches.filter(b => b.status === 'approved' || b.status === 'paying');
const dayDiff = d => Math.round((new Date(d + 'T00:00').getTime() - new Date(localDate() + 'T00:00').getTime()) / (24*H));
function dueText(d){ const n = dayDiff(d); return n === 0 ? 'hôm nay' : n > 0 ? `còn ${n} ngày` : `quá hạn ${-n} ngày`; }
function batchNext(b){
  if (b.status === 'review') {
    const s = currentBatchStep(b); const ch = ensureBatchChain(b); const idx = ch.indexOf(s) + 1;
    return s ? `Chờ ${vnJoin(s.approvers.filter(id => !s.decisions[id]).map(id => user(id).name), s.mode === 'all' ? ' và ' : ' hoặc ')} duyệt${ch.length > 1 ? ` (cấp ${idx}/${ch.length})` : ''}` : 'Chờ duyệt';
  }
  if (b.status === 'approved' || b.status === 'paying') return `Kế toán chi tiền ${paidCount(b)}/${b.reqIds.length}, hẹn ${dOnly(b.payDate + 'T00:00')} (${dueText(b.payDate)})`;
  if (b.status === 'paid') return `Xong ${dt(b.paidAt)}`;
  if (b.status === 'cancelled') return 'Đã huỷ';
  return 'Các khoản đã về lại chờ thanh toán';
}
/* Hoá đơn bổ sung sau khi chi */
const invPending = r => r.evidence === 'later' && r.status === 'paid' && !r.invDone;
const invWaitList = () => S.reqs.filter(invPending).sort((a,b) => a.invDue - b.invDue);

/* ===================== Quyền xem ===================== */
function canSeeAll(u){ return u.admin || u.acct || u.signer; }
function visibleReqs(){
  const u = me();
  return S.reqs.filter(r => {
    if (r.requester === u.id) return true;
    if (r.status === 'draft') return false;
    if (canSeeAll(u)) return true;
    const ids = [u.id, ...principalsOf(u.id)];
    return r.rounds.some(rd => rd.chain.some(s => s.approvers.some(a => ids.includes(a))));
  });
}
const inboxFor = id => S.reqs.filter(r => canAct(r, id));
const seesBatches = u => u.acct || isSigner(u.id) || u.admin;

/* ===================== Điều hướng ===================== */
function go(view, id, extra){
  UI.view = view; UI.id = id || null; UI.errors = {};
  if (view === 'new') UI.form = formFrom(newReq({dept: me().dept}));
  if (view === 'edit') UI.form = formFrom(req(id));
  if (view === 'config') { if (!me().admin && !['cats','vendors'].includes(UI.cfgTab)) UI.cfgTab = 'cats'; UI.draft = {config: clone(S.config)}; UI.sim = UI.sim || {requester:'u1', dept:'OPS', amount:12000000, evidence:'vat'}; }
  if (view === 'batch') { UI.paySel = {}; UI.bpm = ''; }
  if (view === 'batches') UI.batchSel = UI.batchSel || {};
  if (view === 'delegate') UI.dg = {from:S.me, to:'', mode:'none', start:localDate(), end:'', note:'', err:''};
  Object.assign(UI, extra || {});
  render(); window.scrollTo(0,0);
  $('#main').focus({preventScroll:true});
}
function setPersona(id){ S.me = id; save(); toast(`Đang dùng vai ${user(id).name}`); go(UI.view === 'detail' || UI.view === 'batch' ? UI.view : 'home', UI.view === 'detail' || UI.view === 'batch' ? UI.id : null); }

let toastTimer;
function toast(msg){ const l = $('#toast'); l.innerHTML = `<div class="toast" role="status">${esc(msg)}</div>`; clearTimeout(toastTimer); toastTimer = setTimeout(() => { l.innerHTML = ''; }, 3000); }
function closeModal(){ $('#layer').innerHTML = ''; }

/* ===================== Khung trang ===================== */
function render(){ renderSide(); renderMain(); }
function roleLine(u){
  const parts = [];
  if (inConfig(u.id)) parts.push('Người duyệt');
  if (u.acct) parts.push('kế toán');
  if (isSigner(u.id)) parts.push('duyệt đợt chi');
  if (u.admin) parts.push('quản trị');
  const pr = principalsOf(u.id);
  let line = parts.length ? parts.join(', ') : 'Người đề xuất';
  line = line[0].toUpperCase() + line.slice(1);
  if (pr.length) line += ` · đang duyệt thay ${vnJoin(pr.map(id => user(id).name))}`;
  return line;
}
function renderSide(){
  document.body.classList.remove('auth');
  const u = me();
  const inbox = inboxFor(u.id).length;
  const bcount = (u.acct ? S.reqs.filter(r => r.status === 'approved').length + batchesToPay().length : 0) + batchesToApprove(u.id).length;
  const nav = [
    ['home','home','Tổng quan',0],
    ['list','file','Đề xuất chi',0],
    ['inbox','check','Chờ tôi duyệt',inbox],
    ...(seesBatches(u) ? [['batches','wallet','Đợt chi',bcount]] : []),
    ['delegate','swap','Ủy quyền duyệt',0],
    ['reports','chart','Lịch sử, báo cáo',0],
    ...(canDash(u) ? [['dash','chart','Dashboard quản trị',0]] : []),
    ...(u.admin ? [['users','users','Người dùng',0]] : []), ...((u.admin || u.acct) ? [['config','gear','Cấu hình',0]] : [])
  ];
  const active = ({new:'list', edit:'list', detail:'list', batch:'batches'})[UI.view] || UI.view;
  $('#side').innerHTML = `
    <div class="brand"><span class="logo" role="img" aria-label="Xperise"></span><small>Spend Ops · Đề xuất chi</small></div>
    <label class="persona">${avatar(u.id)}<span class="who"><b>${esc(u.name)}</b><span>${esc(u.title)}</span></span><span class="chev">${icon('chev',16)}</span>
      </label>
    <p class="persona-cap">${esc(roleLine(u))}</p>
    <nav class="nav" aria-label="Điều hướng">${nav.map(n => `<button class="${active===n[0]?'on':''}" data-go="${n[0]}">${icon(n[1])}<span>${n[2]}</span>${n[3] ? `<span class="cnt">${n[3]}</span>` : ''}</button>`).join('')}</nav>
    <div class="side-foot"><span>Đăng nhập: ${esc(SESSION_EMAIL || '')}</span><button class="linkbtn" data-action="logout">Đăng xuất</button></div>`;
}
function renderMain(){
  const v = UI.view; const m = $('#main');
  const views = {home:viewHome, list:viewList, inbox:viewInbox, new:viewForm, edit:viewForm, detail:viewDetail, batches:viewBatches, batch:viewBatch, delegate:viewDelegate, reports:viewReports, dash:() => canDash(me()) ? viewDash() : `<div class="card empty"><b>Chỉ người được phân quyền mới xem được dashboard</b>Nhờ quản trị bật quyền “Dashboard quản trị” cho bạn.</div>`, users:() => me().admin ? viewUsers() : `<div class="card empty"><b>Chỉ Quản trị quản lý người dùng</b></div>`, config:() => (me().admin || me().acct) ? viewConfig() : `<div class="card empty"><b>Chỉ Quản trị và Kế toán mở được cấu hình</b>Đổi vai sang Quản trị hệ thống ở góc trái.</div>`};
  m.innerHTML = (views[v] || viewHome)();
  if (v === 'new' || v === 'edit') updatePreview();
  if (v === 'config') updateSim();
}
const pill = st => `<span class="pill ${STATUS[st][1]}">${STATUS[st][0]}</span>`;
const bpill = st => `<span class="pill ${BSTATUS[st][1]}">${BSTATUS[st][0]}</span>`;
function waitLabel(r){
  if (r.status === 'pending') { const s = currentStep(r); return `${esc(vnJoin(waitingOn(r).map(id => user(id).name), s.mode==='all'?' và ':' hoặc '))}`; }
  if (r.status === 'approved') return 'Kế toán lập đợt chi';
  if (invPending(r)) return `Bổ sung hoá đơn, hạn ${dOnly(r.invDue).slice(0,5)}`;
  if (r.status === 'in_batch') { const b = batch(r.batchId); return b ? `${b.status === 'review' ? 'Duyệt đợt' : 'Chi tiền'} ${b.code}, hẹn ${dOnly(b.payDate+'T00:00').slice(0,5)}` : '—'; }
  if (r.status === 'returned' || r.status === 'draft') return `${esc(user(r.requester).name)} sửa và gửi`;
  return '—';
}
const head = (eyebrow, title, sub, right='') => `<div class="head"><div>${eyebrow?`<p class="eyebrow">${eyebrow}</p>`:''}<h1>${title}</h1>${sub?`<p>${sub}</p>`:''}</div>${right}</div>`;

/* ===================== Tổng quan ===================== */
function viewHome(){
  const u = me();
  const sum = a => a.reduce((x,r) => x + r.amount, 0);
  const mine = S.reqs.filter(r => r.requester === u.id);
  const inbox = inboxFor(u.id);
  const minePending = mine.filter(r => ['pending','approved','in_batch'].includes(r.status));
  const mineFix = mine.filter(r => r.status === 'returned' || r.status === 'draft');
  const approved = S.reqs.filter(r => r.status === 'approved');
  const toApprove = batchesToApprove(u.id), toPay = batchesToPay();
  const paidMonth = S.reqs.filter(r => r.status === 'paid' && new Date(r.paidAt).getMonth() === new Date().getMonth());
  const k = (go, l, v, s, act) => `<button class="kpi ${act?'act':''}" data-go="${go}"><span class="l">${l}</span><span class="v">${v}</span><span class="s">${s}</span></button>`;
  const kpis = [
    k('inbox','Chờ tôi duyệt', inbox.length, inbox.length ? short(sum(inbox)) : 'Không có', inbox.length),
    k('list','Phiếu của tôi đang xử lý', minePending.length, minePending.length ? short(sum(minePending)) : 'Không có', false),
    k('list','Cần tôi sửa hoặc gửi', mineFix.length, mineFix.length ? 'Nháp, phiếu trả lại' : 'Không có', mineFix.length)
  ];
  if (u.acct) kpis.push(k('batches','Chờ thanh toán / đợt cần chi', `${approved.length} / ${toPay.length}`, `${short(sum(approved))} chờ thanh toán`, approved.length || toPay.length));
  else if (isSigner(u.id) || toApprove.length) kpis.push(k('batches','Đợt chi chờ tôi duyệt', toApprove.length, toApprove.length ? short(toApprove.reduce((a,b)=>a+batchTotal(b),0)) : 'Không có', toApprove.length));
  else kpis.push(`<div class="kpi static"><span class="l">Công ty đã chi tháng này</span><span class="v">${paidMonth.length}</span><span class="s">${short(sum(paidMonth))}</span></div>`);

  const todo = [];
  inbox.forEach(r => { const p = actorsFor(r, u.id)[0]; todo.push({open:r.id, t:r.title, m:`${r.code} · ${p !== u.id ? 'duyệt thay ' + user(p).name : 'chờ bạn duyệt'} ${currentStep(r).label.toLowerCase()}`, a:r.amount}); });
  invWaitList().filter(r => r.requester === u.id).forEach(r => todo.push({open:r.id, t:r.title, m:`${r.code} · đã chi, cần bạn bổ sung hoá đơn trước ${dOnly(r.invDue)} (${dueText(localDate(r.invDue))})`, a:r.amount, hot:true}));
  const advMine = S.reqs.filter(r => r.requester === u.id && ['waiting','overdue'].includes(advState(r)));
  advMine.forEach(r => todo.push({open:r.id, t:r.title, m:`${r.code} · ${advState(r)==='overdue' ? 'quá hạn' : 'cần'} quyết toán tạm ứng, hạn ${dOnly(r.settleDue)}`, a:r.amount, hot:advState(r)==='overdue'}));
  if (u.acct) {
    S.reqs.filter(r => advState(r) === 'refund_waiting').forEach(r => todo.push({open:r.id, t:r.title, m:`${r.code} · chờ nhận hoàn ứng ${money(r.settle.refund.amount)} từ ${user(r.requester).name}`, a:r.settle.refund.amount}));
    S.reqs.filter(r => advState(r) === 'overdue' && r.requester !== u.id).forEach(r => todo.push({open:r.id, t:r.title, m:`${r.code} · ${user(r.requester).name} quá hạn quyết toán tạm ứng`, a:r.amount, hot:true}));
  }
  mineFix.forEach(r => todo.push({open:r.id, t:r.title, m: r.status === 'returned' ? `${r.code} · bị trả lại, cần sửa và gửi lại` : 'Nháp chưa gửi', a:r.amount}));
  toApprove.forEach(b => { const p = batchActors(b, u.id)[0]; todo.push({batch:b.id, t:`Đợt chi ${b.code}`, m:`${b.reqIds.length} khoản · ${p !== u.id ? 'duyệt thay ' + user(p).name : 'chờ bạn duyệt đợt'}`, a:batchTotal(b)}); });
  if (u.acct) {
    toPay.sort((a,b) => a.payDate < b.payDate ? -1 : 1).forEach(b => todo.push({batch:b.id, t:`Đợt chi ${b.code}`, m:`Hẹn thanh toán ${dOnly(b.payDate+'T00:00')}, ${dueText(b.payDate)} · đã chi ${paidCount(b)}/${b.reqIds.length}`, a:batchTotal(b), hot:dayDiff(b.payDate) <= 0}));
    if (approved.length) todo.push({go:'batches', t:`${approved.length} khoản chờ thanh toán`, m:'Gộp thành đợt chi và hẹn ngày thanh toán', a:sum(approved)});
    invWaitList().filter(r => r.requester !== u.id).forEach(r => todo.push({open:r.id, t:r.title, m:`${r.code} · chờ ${user(r.requester).name} bổ sung hoá đơn, hạn ${dOnly(r.invDue)} (${dueText(localDate(r.invDue))})`, a:r.amount, hot:r.invDue < now()}));
    S.reqs.filter(r => r.refundBack && r.refundBack.status === 'waiting').forEach(r => todo.push({open:r.id, t:r.title, m:`${r.code} · chờ NCC hoàn tiền chuyển nhầm`, a:r.refundBack.amount}));
  }
  const G = S.guideDone || {};
  const steps = [
    {k:'create', who:'u1', t:'Tạo đề xuất chi', d:'Nhập số tiền để thấy chuỗi duyệt đổi theo mức tiền.', go:'new'},
    {k:'approve', who:'u3', t:'Duyệt cấp 1', d:'Duyệt hoặc trả lại phiếu văn phòng phẩm.', go:'inbox'},
    {k:'delegated', who:'u3', t:'Duyệt thay người khác', d:'Võ Ngọc Mai đã ủy quyền, duyệt thay phiếu phần mềm 48 triệu.', go:'inbox'},
    {k:'resubmit', who:'u2', t:'Sửa phiếu bị trả lại và gửi lại', d:'Phiếu tiếp khách đang chờ bổ sung.', go:'list'},
    {k:'mkdeleg', who:'u7', t:'Tạo ủy quyền', d:'Ủy quyền duyệt cho người khác, có hoặc không có thời hạn.', go:'delegate'},
    {k:'batch', who:'u6', t:'Kế toán lập đợt chi', d:'Gộp các khoản đã duyệt để trình duyệt.', go:'batches'},
    {k:'sign', who:'u8', t:'Duyệt đợt chi', d:'Xem tổng đợt và từng khoản rồi duyệt.', go:'batches'},
    {k:'pay', who:'u6', t:'Chuyển tiền và đính kèm UNC', d:'Xem thông tin chuyển khoản, đánh dấu đã chuyển.', go:'batches'},
    {k:'users', who:'u9', t:'Quản lý người dùng', d:'Thêm, sửa, xoá, phân quyền.', go:'users'},
    {k:'config', who:'u9', t:'Đặt ma trận phê duyệt', d:'Quy định theo phòng hoặc người, mốc tiền, cấp duyệt.', go:'config'}
  ];
  return `${head('Spend Ops', `Chào ${esc(u.name.split(' ').slice(-1)[0])}`, `${esc(u.title)}, ${esc(dept(u.dept).name)}`, `<button class="btn pri" data-go="new">${icon('plus',16)} Tạo đề xuất chi</button>`)}
  <div class="stack">
    <div class="g4">${kpis.join('')}</div>
    <div class="g75">
      <div class="card"><div class="ch"><h2 class="card-title">Việc cần làm</h2><span class="card-note">${todo.length ? todo.length + ' việc' : ''}</span></div>
        ${todo.length ? todo.map(x => `<div class="li click" ${x.open?`data-open="${x.open}"`:x.batch?`data-batch="${x.batch}"`:`data-go="${x.go}"`} role="button" tabindex="0">
          <div class="grow"><div class="t">${esc(x.t)}</div><div class="m" ${x.hot ? 'style="color:var(--watch)"' : ''}>${esc(x.m)}</div></div><span class="amt">${short(x.a)}</span></div>`).join('')
        : `<div class="empty"><b>Không có việc nào chờ bạn</b>Đổi vai ở góc trái để xem luồng từ phía người khác.</div>`}
      </div>
          </div>
  </div>`;
}


/* ===================== Bộ lọc, xuất file ===================== */
const BLOBS = {};   // nội dung file tải lên trong phiên này (prototype không lưu file)
let DL = null;
(async () => { try { DL = window.claude && window.claude.use ? await window.claude.use('downloads') : null; } catch (e) { DL = null; } })();
async function saveFile(filename, data){
  if (!DL) { try { const blob = data instanceof Blob ? data : new Blob([data]); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000); toast(`Đã tải ${filename}`); return true; } catch (err) { toast('Không tải được file trên trình duyệt này.'); return false; } }
  try { await DL.save({filename, data}); toast(`Đã lưu ${filename}`); return true; }
  catch (e) { if (!e || e.code !== 'declined') toast(`Không lưu được file (${(e && e.code) || 'lỗi'})`); return false; }
}
function exportXlsx(filename, sheets){
  if (typeof XLSX === 'undefined') { toast('Chưa tải được thư viện Excel, thử lại sau vài giây.'); return; }
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => { const ws = XLSX.utils.aoa_to_sheet(s.rows);
    ws['!cols'] = (s.rows[0] || []).map((h,i) => ({wch: Math.min(60, Math.max(String(h).length + 2, ...s.rows.slice(1,200).map(r => String(r[i] ?? '').length + 2)))}));
    XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0,31)); });
  saveFile(filename, XLSX.write(wb, {type:'array', bookType:'xlsx'}));
}
function pdfStub(lines){
  const txt = lines.map((l,i) => `BT /F1 ${i ? 11 : 15} Tf 60 ${770 - i*24} Td (${noAccent(l).replace(/[()\\]/g,'')}) Tj ET`).join('\n');
  const objs = ['<< /Type /Catalog /Pages 2 0 R >>','<< /Type /Pages /Kids [3 0 R] /Count 1 >>','<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',`<< /Length ${txt.length} >>\nstream\n${txt}\nendstream`,'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
  let out = '%PDF-1.4\n'; const offs = [];
  objs.forEach((o,i) => { offs.push(out.length); out += `${i+1} 0 obj\n${o}\nendobj\n`; });
  const x = out.length;
  return out + `xref\n0 ${objs.length+1}\n0000000000 65535 f \n` + offs.map(o => String(o).padStart(10,'0') + ' 00000 n \n').join('') + `trailer\n<< /Size ${objs.length+1} /Root 1 0 R >>\nstartxref\n${x}\n%%EOF`;
}
const ATTACH_BUCKET = 'attachments'; const MAX_FILE_BYTES = 10 * 1024 * 1024;
const UPLOADS = {};   // key -> {ok, err, p}: trạng thái tải lên Supabase Storage
const fname = key => (S && S.fileNames && S.fileNames[key]) || String(key).replace(/^[0-9a-f]{8}_/, '');   // tên gốc người dùng chọn (lưu trong dữ liệu)
function newKey(name){
  const id = (globalThis.crypto && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(16).slice(2) + Date.now().toString(16)).replace(/-/g,'').slice(0,8);
  return `${id}_${String(name).replace(/[^\w.\-]+/g,'_').slice(0,120)}`;
}
/* Chọn file là tải lên ngay (nền). Trả về khoá lưu trong dữ liệu, hoặc null nếu quá lớn. */
function addBlob(file){
  if (file.size > MAX_FILE_BYTES) { toast(`${file.name} lớn hơn 10 MB, chọn file nhỏ hơn.`); return null; }
  const key = newKey(file.name); BLOBS[key] = file;
  S.fileNames = S.fileNames || {}; S.fileNames[key] = file.name;
  const entry = {ok:false, err:null, p:null}; UPLOADS[key] = entry;
  entry.p = Promise.resolve(SB.storage.from(ATTACH_BUCKET).upload(`files/${key}`, file, {contentType: file.type || 'application/octet-stream', upsert:false}))
    .then(res => { if (res.error) throw res.error; entry.ok = true; })
    .catch(err => { entry.err = (err && err.message) || 'lỗi tải lên'; });
  return key;
}
/* Chờ các file đã chọn tải xong. File lỗi thì bỏ khỏi form và báo lại. */
async function flushUploads(keys){
  const list = (keys || []).filter(k => UPLOADS[k]);
  await Promise.all(list.map(k => UPLOADS[k].p));
  const bad = list.filter(k => !UPLOADS[k].ok);
  if (!bad.length) return true;
  bad.forEach(k => { delete UPLOADS[k]; delete BLOBS[k]; });
  if (UI.form && UI.form.files) UI.form.files = UI.form.files.filter(x => !bad.includes(x.name));
  if (UI.payDraft && UI.payDraft.files) {
    const P = UI.payDraft; const old = P.files; P.files = old.filter(x => !bad.includes(x));
    for (const id of Object.keys(P.map)) { const ni = P.files.indexOf(old[P.map[id]]); if (ni < 0) delete P.map[id]; else P.map[id] = ni; }
  }
  UI.settleFiles = (UI.settleFiles || []).filter(x => !bad.includes(x));
  toast(`Chưa tải lên được ${bad.length} file: ${bad.map(fname).join(', ')}. Đã bỏ khỏi danh sách, hãy chọn lại.`);
  renderMain(); return false;
}
/* Lấy nội dung file: từ bộ nhớ nếu vừa chọn, không thì tải từ Supabase Storage. */
async function fileBlob(key){
  if (BLOBS[key]) return BLOBS[key];
  const { data, error } = await SB.storage.from(ATTACH_BUCKET).download(`files/${key}`);
  if (error || !data) throw new Error(`Không tải được ${fname(key)} từ máy chủ`);
  return data;
}
async function downloadKey(key){
  try { const b = await fileBlob(key); await saveFile(fname(key), b); }
  catch (e) { toast(e.message || 'Không tải được file'); }
}
async function zipKeys(filename, items){ // items: [{name, key}]
  try { const got = []; for (const it of items) got.push({name: it.name, data: await fileBlob(it.key)}); await zipUnc(filename, got); }
  catch (e) { toast(e.message || 'Không tải được file để nén'); }
}
function safeName(n){ return /\.(pdf|png|jpe?g|gif|webp|zip|xlsx|csv|txt)$/i.test(n) ? n : n + '.pdf'; }
async function zipUnc(filename, items){ // items: [{name, data}]
  if (typeof JSZip === 'undefined') { toast('Chưa tải được thư viện nén file, thử lại sau vài giây.'); return; }
  if (!items.length) { toast('Chưa có UNC nào để tải.'); return; }
  const z = new JSZip(); items.forEach(it => z.file(it.name, it.data));
  saveFile(filename, await z.generateAsync({type:'blob'}));
}
const inRange = (t, f) => { if (!t) return !(f.from || f.to); const d = localDate(t); return (!f.from || d >= f.from) && (!f.to || d <= f.to); };
function filterBar(key, o){
  const f = UI[key];
  return `<div class="fbar">
    ${o.q ? `<input class="in" type="search" placeholder="${o.q}" value="${esc(f.q||'')}" data-flt="${key}.q" aria-label="Tìm kiếm" style="flex:1;min-width:180px">` : ''}
    <label class="fdate">Từ ngày<input class="in" type="date" value="${esc(f.from||'')}" data-flt="${key}.from"></label>
    <label class="fdate">Đến ngày<input class="in" type="date" value="${esc(f.to||'')}" data-flt="${key}.to"></label>
    ${o.pm ? `<select class="in" data-flt="${key}.pm" aria-label="Hình thức thanh toán" style="width:170px"><option value="">Mọi hình thức</option>${Object.entries(PAYM).map(([k,v]) => `<option value="${k}" ${f.pm===k?'selected':''}>${v}</option>`).join('')}</select>` : ''}
    ${o.cat ? `<select class="in" data-flt="${key}.cat" aria-label="Hạng mục chi phí" style="width:190px"><option value="">Mọi hạng mục</option>${S.cats.map(g => `<optgroup label="${esc(g[0])}"><option value="g:${esc(g[0])}" ${f.cat==='g:'+g[0]?'selected':''}>Cả nhóm ${esc(g[0])}</option>${g[1].map(c => `<option value="${esc(c)}" ${f.cat===c?'selected':''}>${esc(c)}</option>`).join('')}</optgroup>`).join('')}<option value="${REFUND_CAT}" ${f.cat===REFUND_CAT?'selected':''}>${REFUND_CAT}</option></select>` : ''}
    ${o.who ? `<select class="in" data-flt="${key}.who" aria-label="Người đề xuất" style="width:180px"><option value="">Mọi người đề xuất</option>${S.users.filter(x => !x.deleted && S.reqs.some(r => r.requester === x.id)).map(x => `<option value="${x.id}" ${f.who===x.id?'selected':''}>${esc(x.name)}</option>`).join('')}</select>` : ''}
    ${(f.q || f.from || f.to || f.pm || f.cat || f.who) ? `<button class="btn ghost sm" data-fltclear="${key}">Xoá lọc</button>` : ''}
    ${o.extra || ''}</div>`;
}
const catMatch = (r, c) => !c || (c.startsWith('g:') ? catGroup(r.cat) === c.slice(2) : r.cat === c);
const pmOf = r => r.payee.type === 'me' ? 'transfer' : (r.payMethod || 'transfer');
const invTag = r => invPending(r) ? ` <span class="pill st-returned" title="Hạn bổ sung ${dOnly(r.invDue)}">Chờ hoá đơn</span>` : '';

/* ===================== Danh sách ===================== */
function listRows(){
  const u = me(); const f = UI.flt; const q = noAccent(f.q || '').toLowerCase();
  return visibleReqs().filter(r => inRange(r.createdAt, f) && (!f.pm || pmOf(r) === f.pm) && catMatch(r, f.cat) && (!f.who || r.requester === f.who) &&
    (!q || noAccent([r.code, r.title, user(r.requester).name, payeeOf(r).name, r.cat, invNos(r).join(' ')].join(' ')).toLowerCase().includes(q)));
}
function viewList(){
  const u = me(); const all = listRows();
  const tabs = [
    ['mine','Của tôi', r => r.requester === u.id],
    ['open','Đang xử lý', r => ['pending','returned','approved','in_batch'].includes(r.status)],
    ['done','Đã chi', r => r.status === 'paid'],
    ['closed','Từ chối, huỷ', r => r.status === 'rejected' || r.status === 'cancelled'],
    ['all','Tất cả', () => true]
  ];
  const t = tabs.find(x => x[0] === UI.listTab) || tabs[0];
  const rows = all.filter(t[2]).sort((a,b) => b.createdAt - a.createdAt);
  UI.listShown = rows.map(r => r.id);
  return `${head('Spend Ops','Đề xuất chi', canSeeAll(u) ? 'Toàn bộ đề xuất của công ty.' : 'Phiếu của bạn và phiếu bạn được giao duyệt hoặc duyệt thay.', `<button class="btn pri" data-go="new">${icon('plus',16)} Tạo đề xuất chi</button>`)}
  <div class="tabs" role="tablist">${tabs.map(x => `<button role="tab" aria-selected="${x[0]===t[0]}" class="${x[0]===t[0]?'on':''}" data-listtab="${x[0]}">${x[1]}<span class="c">${all.filter(x[2]).length}</span></button>`).join('')}</div>
  ${filterBar('flt', {q:'Tìm mã, nội dung, NCC, số hoá đơn', pm:true, cat:true, who:canSeeAll(u), extra:`<span class="sp"></span><button class="btn sec sm" data-action="xlist">Xuất Excel</button><button class="btn sec sm" data-action="zuncs">Tải UNC (.zip)</button>`})}
  ${rows.length ? `<div class="tw"><table>
    <thead><tr><th class="r">STT</th><th>Mã</th><th>Ngày lập</th><th>Nội dung</th><th>Hình thức</th><th class="r">Số tiền (đ)</th><th>Trạng thái</th><th>Ngày duyệt</th><th>Đang chờ</th></tr></thead>
    <tbody>${rows.map((r,i) => `<tr class="click" data-open="${r.id}" tabindex="0">
      <td class="r faint">${i+1}</td><td class="code">${esc(r.code||'—')}</td><td class="small num" style="white-space:nowrap">${dt(r.createdAt)}</td>
      <td><div class="ttl">${esc(r.title)}</div><div class="small faint">${esc(user(r.requester).name)} · ${esc(dept(r.dept).name)} · ${esc(r.cat)}${r.direct?' · kế toán lập trực tiếp':''}</div></td>
      <td class="small muted">${PAYM[pmOf(r)]}</td>
      <td class="r" title="${money(r.amount)}">${fmt(r.amount)}</td>
      <td>${pill(r.status)}${invTag(r)}</td>
      <td class="small num" style="white-space:nowrap">${r.approvedAt ? dt(r.approvedAt) : '<span class="faint">—</span>'}</td>
      <td class="small muted">${waitLabel(r)}</td></tr>`).join('')}</tbody>
    <tfoot><tr><td></td><td colspan="4">${rows.length} phiếu</td><td class="r">${fmt(rows.reduce((a,r)=>a+r.amount,0))}</td><td colspan="3"></td></tr></tfoot></table></div>`
  : `<div class="card empty"><b>Không có phiếu nào khớp</b>Đổi bộ lọc hoặc bấm “Tạo đề xuất chi”.</div>`}`;
}
function reqExportRows(ids){
  const head = ['STT','Mã phiếu','Ngày lập','Người đề xuất','Phòng','Nhóm chi phí','Hạng mục','Nội dung','Người nhận','MST','Hình thức TT','Chứng từ','Số hoá đơn','Trước thuế','Tiền thuế','Tổng hoá đơn','Số tiền đề nghị','Trạng thái','Ngày duyệt','Đợt chi','Hẹn thanh toán','Ngày chi','UNC','Hoá đơn bổ sung'];
  return [head].concat(ids.map((id,i) => { const r = req(id), p = payeeOf(r), b = batch(r.batchId) || S.batches.find(x => x.lines && x.lines[id]);
    return [i+1, r.code||'', dt(r.createdAt), user(r.requester).name, dept(r.dept).name, catGroup(r.cat), r.cat, r.title, p.name, p.tax||'', PAYM[pmOf(r)], EVIDENCE[r.evidence][0], invNos(r).join(', '),
      linesSum(r.lines,'pre'), linesSum(r.lines,'tax'), linesSum(r.lines,'total'), r.amount, STATUS[r.status][0], r.approvedAt ? dt(r.approvedAt) : '', b ? b.code : '', b ? dOnly(b.payDate+'T00:00') : '', r.paidAt ? dt(r.paidAt) : '', (r.unc||[]).join(', '),
      r.evidence !== 'later' ? '' : r.invDone ? 'Đã bổ sung ' + dOnly(r.invDone.at) : r.status === 'paid' ? 'Chờ, hạn ' + dOnly(r.invDue) : 'Sau khi chi']; }));
}

/* ===================== Form tạo / sửa ===================== */
const blankLine = rate => ({inv:'', desc:'', pre:0, rate:rate || '10', tax:0, total:0});
function formFrom(r){
  const isNew = !S.reqs.includes(r); const lines = r.lines && r.lines.length ? clone(r.lines) : [blankLine()];
  const sum = linesSum(lines,'total');
  return {id:r.id, isNew, type:r.type||'opex', title:r.title, amount:r.amount, amountManual: !!r.amount && r.amount !== sum, dept:r.dept, cat:isNew ? ((S.cats.find(g => g[1].length) || [0,['Khác']])[1][0]) : r.cat, evidence:r.evidence, payMethod:r.payMethod||'transfer',
    payeeType:r.payee.type === 'other' ? 'vendor' : r.payee.type, vendorId:r.payee.vendorId||'v1', outlay:r.payee.outlay||'transfer', other: r.payee.type === 'other' ? {name:r.payee.name, acc:r.payee.acc, bank:r.payee.bank} : {name:'', acc:'', bank:''},
    needBy:r.needBy||'', settleDue: r.settleDue ? localDate(r.settleDue) : '', note:r.note||'', files:clone(r.files), lines, status:r.status};
}
function linesTable(f){
  return `<div class="tw"><table class="lines"><thead><tr><th style="width:104px">Số hoá đơn</th><th>Diễn giải</th><th class="r" style="width:118px">Trước thuế</th><th style="width:78px">Thuế suất</th><th class="r" style="width:104px">Tiền thuế</th><th class="r" style="width:118px">Thành tiền</th><th style="width:34px"></th></tr></thead>
    <tbody>${f.lines.map((l,i) => `<tr>
      <td><input class="in cell" data-ln="${i}" data-k="inv" value="${esc(l.inv)}" aria-label="Số hoá đơn dòng ${i+1}" autocomplete="off"></td>
      <td><input class="in cell" data-ln="${i}" data-k="desc" value="${esc(l.desc)}" aria-label="Diễn giải dòng ${i+1}" placeholder="${i ? '' : 'VD: phí di chuyển'}" autocomplete="off"></td>
      <td><input class="in cell r" data-ln="${i}" data-k="pre" data-money inputmode="numeric" value="${l.pre?fmt(l.pre):''}" aria-label="Trước thuế dòng ${i+1}" autocomplete="off"></td>
      <td><select class="in cell" data-ln="${i}" data-k="rate" aria-label="Thuế suất dòng ${i+1}">${Object.entries(RATES).map(([k,v]) => `<option value="${k}" ${l.rate===k?'selected':''}>${v}</option>`).join('')}</select></td>
      <td><input class="in cell r" data-ln="${i}" data-k="tax" data-money inputmode="numeric" value="${l.tax?fmt(l.tax):''}" aria-label="Tiền thuế dòng ${i+1}" autocomplete="off"></td>
      <td><input class="in cell r" data-ln="${i}" data-k="total" data-money inputmode="numeric" value="${l.total?fmt(l.total):''}" aria-label="Thành tiền dòng ${i+1}" autocomplete="off"></td>
      <td>${f.lines.length > 1 ? `<button type="button" class="copy" data-rmline="${i}" aria-label="Xoá dòng ${i+1}">${icon('x')}</button>` : ''}</td></tr>`).join('')}</tbody>
    <tfoot><tr><td colspan="2">Tổng ${f.lines.length > 1 ? f.lines.length + ' dòng' : ''}</td><td class="r" id="sumPre">${fmt(linesSum(f.lines,'pre'))}</td><td></td><td class="r" id="sumTax">${fmt(linesSum(f.lines,'tax'))}</td><td class="r" id="sumTot">${fmt(linesSum(f.lines,'total'))}</td><td></td></tr></tfoot></table></div>
    <div class="row" style="margin-top:8px"><label class="btn sec sm" style="position:relative">${icon('upload',14)} Upload hoá đơn<input type="file" multiple accept=".xml,.pdf,image/*" data-invfiles style="position:absolute;opacity:0;width:1px;height:1px"></label><button type="button" class="btn ghost sm" data-action="addline">${icon('plus',14)} Thêm dòng</button><button type="button" class="btn ghost sm" data-action="pastelines">Dán nhiều dòng từ Excel</button><button type="button" class="btn ghost sm" data-action="xmldemo">Thử đọc hoá đơn XML mẫu</button></div>
    <span class="hint">Nhập Thành tiền và thuế suất để tự tách trước thuế, hoặc nhập Trước thuế. Tiền thuế sửa tay được để khớp từng đồng với hoá đơn. Mỗi mức thuế hoặc mỗi hoá đơn một dòng; bấm “Upload hoá đơn”: file XML thì bảng tự điền, file PDF hoặc ảnh thì đính kèm và nhập số theo hoá đơn.</span>`;
}
function vendorMatches(q){ q = noAccent(q || '').toLowerCase().trim(); if (!q) return [];
  return S.vendors.filter(x => !x.deleted && noAccent([x.name, x.tax, x.acc, x.accName, x.bank].join(' ')).toLowerCase().includes(q)); }
function vendorResults(){ const q = UI.vq || ''; if (!q.trim()) return ''; const m = vendorMatches(q);
  return m.length ? `<div class="vlist">${m.slice(0,6).map(x => `<button type="button" data-vpick="${x.id}"><b>${esc(x.name)}</b><span>${x.tax ? 'MST ' + esc(x.tax) + ' · ' : ''}${esc(x.bank)} ${esc(x.acc)}${x.verified ? '' : ' · chưa xác minh'}</span></button>`).join('')}${m.length > 6 ? `<div class="small faint" style="padding:6px 10px">Còn ${m.length-6} kết quả, gõ thêm để thu hẹp.</div>` : ''}</div>`
    : `<div class="small faint" style="padding:6px 2px">Không tìm thấy “${esc(q)}”. Bấm “NCC mới” để thêm.</div>`; }
function amountHint(f){ const sum = linesSum(f.lines,'total');
  if (f.type === 'refund') return f.amount ? short(f.amount) : '';
  if (f.evidence === 'advance') return f.amount ? 'Tạm ứng, quyết toán chứng từ sau.' : 'Nhập số tiền tạm ứng.';
  if (!sum) return 'Điền bảng tiền và thuế ở trên.';
  if (!f.amount || f.amount === sum) return 'Bằng tổng thành tiền.';
  if (f.amount < sum) return `Thanh toán một phần: ${Math.round(f.amount/sum*100)}% của ${fmt(sum)} đ, còn lại ${fmt(sum - f.amount)} đ.`;
  return `Lớn hơn tổng thành tiền ${fmt(f.amount - sum)} đ.`; }
function viewForm(){
  const f = UI.form; const E = UI.errors; const r = req(f.id); const editing = r && !f.isNew;
  let ret = null;
  if (r && r.status === 'returned') { const s = r.rounds[r.rounds.length-1].chain.find(x => x.status === 'returned'); const by = Object.keys(s.decisions).find(k => s.decisions[k].act === 'return'); ret = {by:s.decisions[by].by || by, note:s.decisions[by].note}; }
  const v = vendor(f.vendorId); const refund = f.type === 'refund';
  return `<button class="crumb" ${editing ? `data-open="${r.id}"` : 'data-go="list"'}>‹ ${editing ? esc(r.code||'Nháp') : 'Đề xuất chi'}</button>
  ${head('Đề xuất chi', f.isNew ? 'Tạo đề xuất chi' : f.status === 'returned' ? 'Sửa và gửi lại' : 'Tiếp tục bản nháp', `Người đề xuất ${esc(user(editing ? r.requester : S.me).name)}`)}
  ${ret ? `<div class="alert warn" style="margin-bottom:16px"><b>${esc(user(ret.by).name)} trả lại:</b> ${esc(ret.note)}<div class="small muted" style="margin-top:2px">Gửi lại mở vòng duyệt mới theo cấu hình hiện tại, lịch sử vòng trước vẫn giữ.</div></div>` : ''}
  <div class="g75 wide">
    <form class="card stack" id="reqForm" novalidate>
      <div class="fgrid">
        <div class="f full"><span class="flabel">Loại đề xuất</span>
          <div class="choice two" role="radiogroup" aria-label="Loại đề xuất">
            <label><input type="radio" name="type" value="opex" ${!refund?'checked':''}>Chi phí hoạt động<span>Mua hàng, dịch vụ, hoàn ứng</span></label>
            <label><input type="radio" name="type" value="refund" ${refund?'checked':''}>Hoàn tiền<span>Khách chuyển nhầm, hoàn lại cho khách</span></label>
          </div></div>
        <label class="f full">${refund ? 'Nội dung hoàn tiền' : 'Nội dung chi'}<input class="in ${E.title?'err':''}" name="title" value="${esc(f.title)}" placeholder="${refund ? 'VD: Hoàn tiền khách ABC chuyển nhầm ngày 10/11' : 'VD: Văn phòng phẩm quý 4/2026'}" autocomplete="off">${E.title?`<span class="ferr">${E.title}</span>`:''}</label>
        <label class="f">Phòng chịu chi phí<select class="in" name="dept">${S.depts.map(d => `<option value="${d.id}" ${d.id===f.dept?'selected':''}>${esc(d.name)}</option>`).join('')}</select></label>
        ${refund ? `<label class="f">Hạng mục<input class="in" value="${REFUND_CAT}" disabled><span class="hint">Không tính vào chi phí.</span></label>`
          : `<label class="f">Hạng mục chi phí<select class="in" name="cat">${catOptions(f.cat)}</select><span class="hint">Kế toán có thể phân loại lại sau.</span></label>`}
        <label class="f">Hình thức thanh toán mong muốn<select class="in ${E.pm?'err':''}" name="payMethod" ${f.payeeType==='me' && !refund ? 'disabled' : ''}>${Object.entries(PAYM).filter(([k]) => !refund || k === 'transfer').map(([k,x]) => `<option value="${k}" ${(f.payeeType==='me' && !refund ? 'transfer' : f.payMethod)===k?'selected':''}>${x}</option>`).join('')}</select>${E.pm?`<span class="ferr">${E.pm}</span>`:f.payeeType==='me' && !refund ? '<span class="hint">Hoàn ứng luôn chuyển khoản về STK của bạn.</span>':''}</label>
        <label class="f">Cần chi trước ngày<input class="in" type="date" name="needBy" value="${esc(f.needBy)}"></label>
        ${f.evidence === 'advance' ? `<label class="f">Hạn quyết toán tạm ứng<input class="in ${E.settleDue?'err':''}" type="date" name="settleDue" value="${esc(f.settleDue||'')}" min="${localDate()}">${E.settleDue ? `<span class="ferr">${E.settleDue}</span>` : '<span class="hint">Chọn ngày bạn sẽ nộp hoá đơn chi thực tế. Quá ngày này sẽ được nhắc.</span>'}</label>` : ''}
        ${refund ? `
        <label class="f full">Tên người nhận hoàn tiền<input class="in ${E.other?'err':''}" name="oname" value="${esc(f.other.name)}" autocomplete="off">${E.other?`<span class="ferr">${E.other}</span>`:''}</label>
        <label class="f">Số tài khoản<input class="in" name="oacc" value="${esc(f.other.acc)}" autocomplete="off" data-acc><span class="hint">Chấp nhận số và chữ.</span></label>
        <label class="f">Ngân hàng<input class="in" name="obank" value="${esc(f.other.bank)}" autocomplete="off"></label>
        <label class="f">Số tiền hoàn<div class="money"><input class="in ${E.amount?'err':''}" name="amount" inputmode="numeric" value="${f.amount?fmt(f.amount):''}" placeholder="0" data-money autocomplete="off"><span>đ</span></div>${E.amount?`<span class="ferr">${E.amount}</span>`:`<span class="hint" id="amtHint">${esc(amountHint(f))}</span>`}</label>
        <div class="f"><span class="flabel">Căn cứ</span><span class="hint" style="font-size:12.5px">Đính kèm sao kê thể hiện khoản tiền khách chuyển nhầm vào công ty.</span></div>`
        : `
        <div class="f full"><span class="flabel">Thanh toán cho</span>
          <div class="choice two" role="radiogroup" aria-label="Thanh toán cho">
            <label><input type="radio" name="payeeType" value="vendor" ${f.payeeType==='vendor'?'checked':''}>Nhà cung cấp<span>Công ty trả thẳng cho NCC</span></label>
            <label><input type="radio" name="payeeType" value="me" ${f.payeeType==='me'?'checked':''}>Hoàn ứng cho tôi<span>Tôi đã tự trả trước</span></label>
          </div></div>
        ${f.payeeType==='vendor'
          ? `<div class="f full"><span class="flabel">Nhà cung cấp</span><input type="hidden" name="vendorId" value="${esc(f.vendorId)}">
              ${v ? `<div class="vsel"><div><b>${esc(v.name)}</b><div class="small faint">${v.tax ? 'MST ' + esc(v.tax) + ' · ' : ''}${esc(v.bank)} ${esc(v.acc)} · ${esc(v.accName || '')}</div></div>${v.verified ? '<span class="pill st-approved">STK đã xác minh</span>' : '<span class="pill st-returned">STK chưa xác minh</span>'}</div>` : ''}
              <div class="row" style="flex-wrap:nowrap"><input class="in" type="search" data-vq placeholder="Tìm nhà cung cấp khác theo tên, MST hoặc số tài khoản" value="${esc(UI.vq||'')}" autocomplete="off" aria-label="Tìm nhà cung cấp"><button type="button" class="btn sec" data-action="newvendor">${icon('plus',16)} NCC mới</button></div>
              <div id="vres">${vendorResults()}</div>${v && !v.verified ? '<span class="hint">STK chưa xác minh: vẫn gửi duyệt được, kế toán xác minh trước khi đưa vào đợt chi.</span>' : ''}</div>`
          : `<label class="f full">Tôi đã trả bằng<select class="in ${E.outlay?'err':''}" name="outlay">${Object.entries(OUTLAY).map(([k,x]) => `<option value="${k}" ${f.outlay===k?'selected':''}>${x}</option>`).join('')}</select>${E.outlay?`<span class="ferr">${E.outlay}</span>`:''}</label>`}
        <div class="f full"><span class="flabel">Chứng từ hiện có</span>
          <div class="choice five" role="radiogroup" aria-label="Chứng từ">${Object.entries(EVIDENCE).map(([k,x]) => `<label><input type="radio" name="evidence" value="${k}" ${f.evidence===k?'checked':''}>${x[0]}<span>${x[1]}</span></label>`).join('')}</div>
          ${f.evidence==='later' ? `<span class="hint">Vẫn nhập đủ số tiền và thuế theo ĐNTT. Sau khi chi, hệ thống nhắc bạn và kế toán bổ sung hoá đơn trong ${INV_DAYS} ngày.</span>` : ''}${f.evidence==='advance' ? '<span class="hint">Tạm ứng không cần hoá đơn hay bảng tiền thuế. Chỉ nhập số tiền đề nghị và nội dung; chứng từ quyết toán sau.</span>' : ''}</div>
        ${f.evidence === 'advance' ? '' : `<div class="f full"><span class="flabel">Tiền và thuế</span>${linesTable(f)}${E.lines?`<span class="ferr">${E.lines}</span>`:''}</div>`}
        <label class="f">Số tiền đề nghị thanh toán lần này<div class="money"><input class="in ${E.amount?'err':''}" name="amount" inputmode="numeric" value="${f.amount?fmt(f.amount):''}" placeholder="0" data-money autocomplete="off"><span>đ</span></div>${E.amount?`<span class="ferr">${E.amount}</span>`:''}<span class="hint" id="amtHint">${esc(amountHint(f))}</span></label>
        <div class="f"><span class="flabel">Thanh toán nhiều đợt</span><span class="hint" style="font-size:12.5px">Đặt cọc hoặc trả theo đợt: giữ bảng trên theo tổng giá trị, sửa ô bên trái thành số tiền của đợt này.</span></div>`}
        <div class="f full"><span class="flabel">Đính kèm</span>
          ${f.files.length ? `<div class="flist">${f.files.map((x,i) => `<div class="frow">${icon('file',15)}<span class="fname">${esc(fname(x.name))}</span><select class="in" data-fkind="${i}" aria-label="Loại file ${esc(x.name)}">${Object.entries(FKIND).map(([k,n]) => `<option value="${k}" ${x.kind===k?'selected':''}>${n}</option>`).join('')}</select><button type="button" class="copy" data-rmfile="${i}" aria-label="Bỏ ${esc(x.name)}">${icon('x')}</button></div>`).join('')}</div>` : ''}
          <div class="row"><label class="btn sec sm" style="position:relative">${icon('clip',14)} Thêm file<input type="file" multiple data-files style="position:absolute;opacity:0;width:1px;height:1px"></label><span class="hint">Chọn loại cho từng file. Mọi loại chứng từ đều đính kèm được, kể cả khi chưa có hoá đơn.</span></div>
          ${E.files?`<span class="ferr">${E.files}</span>`:''}</div>
        <label class="f full">Ghi chú cho người duyệt<textarea class="in" name="note" placeholder="Mục đích, căn cứ, người tham dự">${esc(f.note)}</textarea></label>
      </div>
      <div class="row" style="border-top:1px solid var(--line);padding-top:16px">
        ${editing && f.status !== 'returned' ? `<button type="button" class="btn ghost" data-action="deldraft">Xoá nháp</button>` : ''}
        <span class="sp"></span>
        ${f.status !== 'returned' ? `<button type="button" class="btn sec" data-action="savedraft">Lưu nháp</button>` : ''}
        <button type="submit" class="btn pri">${f.status==='returned' ? 'Gửi lại để duyệt' : 'Gửi duyệt'}</button>
      </div>
    </form>
    <div class="card sticky" id="preview" style="position:sticky;top:24px"></div>
  </div>`;
}
function readForm(){
  const fm = $('#reqForm'); if (!fm) return; const fd = new FormData(fm); const f = UI.form;
  f.type = fd.get('type') || 'opex'; f.title = (fd.get('title')||'').trim(); f.dept = fd.get('dept'); if (fd.get('cat')) f.cat = fd.get('cat');
  f.needBy = fd.get('needBy')||''; f.settleDue = fd.get('settleDue')||''; f.note = (fd.get('note')||'').trim();
  if (fd.get('evidence')) f.evidence = fd.get('evidence'); if (fd.get('payMethod')) f.payMethod = fd.get('payMethod');
  if (fd.get('payeeType')) f.payeeType = fd.get('payeeType');
  if (fd.get('vendorId')) f.vendorId = fd.get('vendorId'); if (fd.get('outlay')) f.outlay = fd.get('outlay');
  if (fd.has('oname')) f.other = {name:(fd.get('oname')||'').trim(), acc:(fd.get('oacc')||'').trim(), bank:(fd.get('obank')||'').trim()};
  if (f.evidence === 'advance') f.lines = []; else if (!f.lines.length) f.lines = [blankLine()];
  const amt = digits(fd.get('amount')); const sum = linesSum(f.lines,'total');
  if (f.type === 'refund') { f.amount = amt; f.amountManual = true; }
  else { f.amountManual = !!amt && amt !== sum; f.amount = amt || sum; }
}
function lineInput(t){
  const f = UI.form; const i = +t.dataset.ln, k = t.dataset.k; const l = f.lines[i]; if (!l) return;
  if (k === 'inv' || k === 'desc') l[k] = t.value;
  else if (k === 'rate') { l.rate = t.value; if (l.anchor === 'total' && l.total) { l.pre = Math.round(l.total / (1 + rateNum(l.rate))); l.tax = l.total - l.pre; } else if (l.pre) { l.tax = Math.round(l.pre * rateNum(l.rate)); l.total = l.pre + l.tax; } }
  else { const v = digits(t.value); l[k] = v;
    if (k === 'pre') { l.anchor = 'pre'; l.tax = Math.round(v * rateNum(l.rate)); l.total = v + l.tax; }
    else if (k === 'total') { l.anchor = 'total'; l.pre = Math.round(v / (1 + rateNum(l.rate))); l.tax = v - l.pre; }
    else if (k === 'tax') { if (l.anchor === 'total') l.pre = l.total - v; else l.total = l.pre + v; } } // sửa tay tiền thuế: giữ nguyên ô người dùng vừa nhập (tổng hoặc trước thuế)
  ['pre','tax','total'].forEach(x => { if (x === k) return; const el = $(`[data-ln="${i}"][data-k="${x}"]`); if (el) el.value = l[x] ? fmt(l[x]) : ''; });
  syncTotals();
}
function syncTotals(){
  const f = UI.form; const sum = linesSum(f.lines,'total');
  const set = (id, v) => { const el = $('#' + id); if (el) el.textContent = fmt(v); };
  set('sumPre', linesSum(f.lines,'pre')); set('sumTax', linesSum(f.lines,'tax')); set('sumTot', sum);
  if (!f.amountManual) { f.amount = sum; const a = $('#reqForm [name=amount]'); if (a) a.value = sum ? fmt(sum) : ''; }
  const h = $('#amtHint'); if (h) h.textContent = amountHint(f);
}
function formRequester(){ const r = req(UI.form.id); return r && !UI.form.isNew ? r.requester : S.me; }
function cashBlocked(f){ const cash = f.payeeType === 'me' ? f.outlay === 'cash' : f.payMethod === 'cash';
  return f.type !== 'refund' && cash && f.evidence === 'vat' && linesSum(f.lines,'total') >= S.config.cashLimit; }
function updatePreview(){
  const box = $('#preview'); if (!box) return; const f = UI.form; const sum = linesSum(f.lines,'total');
  if (!f.amount) { box.innerHTML = `<div class="ch"><h2 class="card-title">Chuỗi duyệt dự kiến</h2></div><p class="small faint" style="margin:0">Nhập số tiền để xem ai sẽ duyệt.</p>`; return; }
  const res = buildChain({amount:f.amount, dept:f.dept, requester:formRequester(), evidence:f.evidence}, S.config);
  box.innerHTML = `<div class="ch"><h2 class="card-title">Chuỗi duyệt dự kiến</h2><span class="card-note">${esc(ruleLabel(res.rule))} · ${esc(bandRange(res.band, res.rule))}</span></div>
    <div class="kv" style="grid-template-columns:1fr 1fr;margin-bottom:16px"><div><span>Đề nghị thanh toán</span><b class="num">${money(f.amount)}</b></div>${f.type !== 'refund' && f.evidence !== 'advance' ? `<div><span>Trong đó thuế</span><b class="num">${money(sum ? Math.round(linesSum(f.lines,'tax') * f.amount / sum) : 0)}</b></div>` : ''}</div>
    ${chainHtml(res.steps.map(s => Object.assign({}, s, {status:'waiting', decisions:{}})), true)}
    ${res.problems.length ? `<div class="alert neg" style="margin-top:16px"><b>Chưa gửi được.</b> ${esc(res.problems.join(' '))} Nhờ Quản trị bổ sung người duyệt.</div>` : ''}
    ${cashBlocked(f) ? `<div class="alert neg" style="margin-top:16px">Hoá đơn VAT từ ${short(S.config.cashLimit)} phải trả không dùng tiền mặt để được khấu trừ.</div>` : ''}
    <p class="small faint" style="margin:16px 0 0">Đủ chữ ký thì phiếu sang Chờ thanh toán: kế toán lập đợt và hẹn ngày chi, người duyệt đợt duyệt, kế toán chi rồi đính kèm UNC. Bạn theo dõi được ngày hẹn và tải UNC ngay trên phiếu.</p>`;
}
function validate(f){
  const E = {}; const kinds = f.files.map(x => x.kind); const sum = linesSum(f.lines,'total');
  if (!f.title || f.title.length < 6) E.title = 'Ghi rõ nội dung, tối thiểu 6 ký tự.';
  if (!f.amount) E.amount = 'Nhập số tiền lớn hơn 0.';
  if (f.type !== 'refund' && f.evidence === 'advance' && (!f.settleDue || f.settleDue < localDate())) E.settleDue = 'Chọn hạn quyết toán, không được ở quá khứ.';
  if (f.type === 'refund') {
    if (!f.other.name || !f.other.acc || !f.other.bank) E.other = 'Nhập đủ tên người nhận, số tài khoản và ngân hàng.';
    if (!kinds.includes('statement')) E.files = 'Đính kèm sao kê ngân hàng và chọn loại “Sao kê ngân hàng”.';
  } else {
    if (f.evidence === 'advance') { /* tạm ứng: không cần bảng tiền, hoá đơn hay file */ } else if (!sum) E.lines = 'Nhập ít nhất một dòng tiền.';
    else if (f.amount > sum) E.amount = `Số tiền đề nghị lớn hơn tổng thành tiền ${fmt(sum)} đ.`;
    if (f.evidence === 'vat' && f.lines.some(l => l.total && !l.inv.trim())) E.lines = 'Điền số hoá đơn cho từng dòng.';
    if (cashBlocked(f)) E[f.payeeType === 'me' ? 'outlay' : 'pm'] = `Hoá đơn VAT từ ${short(S.config.cashLimit)} không được trả bằng tiền mặt.`;
    if (f.evidence === 'vat' && !kinds.includes('vat')) E.files = 'Đính kèm hoá đơn và chọn loại “Hoá đơn VAT”.';
    else if (f.evidence === 'sale' && !kinds.includes('sale')) E.files = 'Đính kèm hoá đơn và chọn loại “Hoá đơn bán hàng”.';
    else if (f.evidence === 'bill' && !f.files.length) E.files = 'Đính kèm bill hoặc biên nhận.';
    else if (f.evidence === 'later' && !kinds.some(k => ['dntt','contract','quote'].includes(k))) E.files = 'Chưa có hoá đơn thì đính kèm ĐNTT, hợp đồng hoặc báo giá.';
  }
  const res = buildChain({amount:f.amount, dept:f.dept, requester:formRequester(), evidence:f.evidence}, S.config);
  if (res.problems.length) E.chain = res.problems.join(' ');
  return E;
}
function applyForm(r, f){
  const refund = f.type === 'refund';
  if (!refund && f.evidence === 'advance' && f.settleDue) r.settleDue = new Date(f.settleDue + 'T23:59:59').getTime(); else if (f.evidence !== 'advance') delete r.settleDue;
  Object.assign(r, {type:f.type, title:f.title || 'Nháp chưa đặt tên', amount:f.amount, dept:f.dept, cat: refund ? REFUND_CAT : f.cat, evidence: refund ? 'bill' : f.evidence, needBy:f.needBy, note:f.note, files:clone(f.files),
    lines: refund ? [{inv:'', desc:f.title, pre:f.amount, rate:'kct', tax:0, total:f.amount}] : clone(f.lines.filter(l => l.total || l.inv || l.desc)),
    payMethod: refund || f.payeeType === 'me' ? 'transfer' : f.payMethod,
    payee: refund ? Object.assign({type:'other'}, f.other) : f.payeeType === 'me' ? {type:'me', outlay:f.outlay} : {type:'vendor', vendorId:f.vendorId}});
}
async function submitForm(){
  readForm(); const f = UI.form; const E = validate(f); UI.errors = E;
  if (Object.keys(E).length) { renderMain(); toast(E.chain || 'Kiểm tra lại các ô báo đỏ.'); return; }
  if (!(await flushUploads(f.files.map(x => x.name)))) return;
  let r = req(f.id);
  if (!r || f.isNew) { r = newReq({id:f.id, requester:S.me, lines:[]}); S.reqs.push(r); }
  applyForm(r, f);
  if (f.status === 'returned') log(r, S.me, 'Sửa nội dung theo yêu cầu bổ sung');
  submitReq(r, S.me);
  S.guideDone[f.status === 'returned' ? 'resubmit' : 'create'] = true;
  save(); toast(`Đã gửi ${r.code}, chờ ${currentStep(r).label.toLowerCase()}`); go('detail', r.id);
}
function saveDraft(){
  readForm(); const f = UI.form; let r = req(f.id);
  if (!r || f.isNew) { r = newReq({id:f.id, requester:S.me, lines:[]}); S.reqs.push(r); log(r, S.me, 'Tạo nháp'); }
  applyForm(r, f); save(); toast('Đã lưu nháp'); go('detail', r.id);
}
/* Đọc hoá đơn điện tử XML (TT78): số HĐ, người bán, bảng tổng hợp theo thuế suất. */
function parseInvoiceXml(text){
  try {
    const doc = new DOMParser().parseFromString(text, 'application/xml'); if (doc.getElementsByTagName('parsererror').length) return null;
    const g = (el, tag) => { const n = el.getElementsByTagName(tag)[0]; return n ? n.textContent.trim() : ''; };
    const inv = g(doc,'SHDon'); const nb = doc.getElementsByTagName('NBan')[0];
    const seller = nb ? g(nb,'Ten') : ''; const sellerTax = nb ? g(nb,'MST') : '';
    const num = s => Math.round(Number(String(s).replace(/,/g,'')) || 0);
    const rateKey = s => { const m = String(s).match(/(\d+(?:\.\d+)?)\s*%/); if (m) { const k = String(Number(m[1])); return RATES[k] ? k : 'kct'; } return 'kct'; };
    let lines = [...doc.getElementsByTagName('LTSuat')].map(n => ({inv, desc:seller, rate:rateKey(g(n,'TSuat')), pre:num(g(n,'ThTien')), tax:num(g(n,'TThue'))}));
    if (!lines.length) { const pre = num(g(doc,'TgTCThue')), tax = num(g(doc,'TgTThue')); if (pre || tax) lines = [{inv, desc:seller, rate: pre ? rateKey(Math.round(tax/pre*100) + '%') : 'kct', pre, tax}]; }
    if (!inv && !lines.length) return null;
    lines.forEach(l => l.total = l.pre + l.tax);
    return {inv, seller, sellerTax, lines, total:num(g(doc,'TgTTTBSo'))};
  } catch (e) { return null; }
}
function applyInvoice(p, fileName){
  const f = UI.form; const empty = f.lines.every(l => !l.total && !l.inv);
  f.lines = (empty ? [] : f.lines).concat(p.lines.map(l => Object.assign(blankLine(), l)));
  if (f.evidence !== 'vat' && f.evidence !== 'sale') f.evidence = 'vat';
  const v = p.sellerTax && S.vendors.find(x => !x.deleted && (x.tax||'').replace(/\D/g,'') === p.sellerTax.replace(/\D/g,'')); if (v) { f.payeeType = 'vendor'; f.vendorId = v.id; }
  if (!f.amountManual) f.amount = linesSum(f.lines,'total');
  toast(`Đã đọc hoá đơn ${p.inv || ''} từ ${fileName}: ${p.lines.length} mức thuế, tổng ${fmt(linesSum(p.lines,'total'))} đ${v ? ', khớp NCC theo MST' : ''}`);
}
const SAMPLE_XML = `<HDon><DLHDon><TTChung><KHHDon>C26TAA</KHHDon><SHDon>00214586</SHDon><NLap>2026-07-30</NLap></TTChung><NDHDon><NBan><Ten>CÔNG TY TNHH THỂ THAO MẪU</Ten><MST>0318507000</MST></NBan><TToan><THTTLTSuat><LTSuat><TSuat>8%</TSuat><ThTien>2812963</ThTien><TThue>225037</TThue></LTSuat></THTTLTSuat><TgTCThue>2812963</TgTCThue><TgTThue>225037</TgTThue><TgTTTBSo>3038000</TgTTTBSo></TToan></NDHDon></DLHDon></HDon>`;
function parsePaste(text){
  const out = []; const n = s => digits(String(s||'').split(/[,.]\d{1,2}$/)[0]);
  text.split(/\r?\n/).forEach(row => { if (!row.trim()) return; const c = row.split(/\t|;/).map(x => x.trim()); if (c.length < 2) return;
    let [inv, desc, pre, rate, tax, total] = c; if (c.length === 2) { desc = ''; pre = c[1]; }
    const rk = /kct|khong/i.test(noAccent(rate||'')) ? 'kct' : (String(rate||'').match(/\d+/) || ['10'])[0]; const l = blankLine(RATES[rk] ? rk : '10');
    l.inv = inv; l.desc = desc || ''; l.pre = n(pre); if (!l.pre && !n(total)) return;
    if (l.pre) { l.tax = tax !== undefined && tax !== '' ? n(tax) : Math.round(l.pre * rateNum(l.rate)); l.total = l.pre + l.tax; }
    else { l.total = n(total); l.pre = Math.round(l.total / (1 + rateNum(l.rate))); l.tax = l.total - l.pre; }
    out.push(l); });
  return out;
}

/* ===================== Chờ tôi duyệt ===================== */
function viewInbox(){
  const u = me(); const items = inboxFor(u.id).sort((a,b) => a.submittedAt - b.submittedAt);
  const sel = items.filter(r => UI.inboxSel[r.id]);
  return `${head('Phê duyệt','Chờ tôi duyệt', items.length ? `${items.length} khoản, ${short(items.reduce((a,r)=>a+r.amount,0))}, cũ nhất ở trên.` : '',
    items.length ? `<div class="row"><button class="btn sec" data-action="selall">${sel.length === items.length ? 'Bỏ chọn' : 'Chọn tất cả'}</button><button class="btn pri" data-action="bulk" ${sel.length?'':'disabled'}>${sel.length ? `Duyệt ${sel.length} khoản` : 'Duyệt khoản đã chọn'}</button></div>` : '')}
  ${items.length ? `<div class="card">${items.map(r => { const s = currentStep(r); const rd = r.rounds[r.rounds.length-1];
    const p = actorsFor(r, u.id)[0];
    const others = s.approvers.filter(id => id !== u.id && id !== p);
    const meta = [r.code, user(r.requester).name, dept(r.dept).name, s.label, rd.no>1 ? `vòng ${rd.no}` : '', r.evidence==='bill' ? 'không hoá đơn' : r.evidence==='later' ? 'chưa có hoá đơn' : r.evidence==='advance' ? 'tạm ứng' : '', PAYM[pmOf(r)], 'lập ' + dt(r.createdAt)].filter(Boolean).join(' · ');
    return `<div class="li">
      <span class="faint small num" style="width:18px;text-align:right">${items.indexOf(r)+1}</span><input type="checkbox" class="cb" aria-label="Chọn ${esc(r.code)}" data-sel="${r.id}" ${UI.inboxSel[r.id]?'checked':''}>
      <div class="grow"><div class="t">${esc(r.title)}</div><div class="m">${esc(meta)}</div>${p !== u.id ? `<div class="m deleg">Duyệt thay ${esc(user(p).name)}</div>` : ''}
        ${others.length ? `<div class="m">${s.mode==='all' ? 'Cần thêm chữ ký của ' : 'Cùng nhận, một người duyệt là đủ: '}${esc(vnJoin(others.map(id => user(id).name)))}</div>` : ''}</div>
      <span class="amt">${money(r.amount)}</span>
      <button class="btn ghost sm" data-open="${r.id}">Xem</button><button class="btn sec sm" data-quick="${r.id}">Duyệt</button></div>`; }).join('')}</div>`
  : `<div class="card empty"><b>Không có khoản nào chờ bạn</b>${inConfig(u.id) || principalsOf(u.id).length ? 'Khi một đề xuất tới lượt bạn, nó xuất hiện ở đây.' : 'Vai này không có bước duyệt nào. Đổi sang Lê Thu Hà để duyệt.'}</div>`}`;
}


/* ===================== Chi tiết đề xuất ===================== */
function chainHtml(chain, preview){
  return `<div class="chain">${chain.map((s,i) => `<div class="step ${s.status}">
      <div class="dot">${s.status==='approved'?'✓':s.status==='rejected'?'✕':s.status==='returned'?'↩':s.key==='NI'?'+':i+1}</div>
      <div><div class="sh"><b>${esc(s.label)}</b>${s.approvers.length>1?`<span class="meta">${s.mode==='all'?'tất cả cùng duyệt':'một người bất kỳ'}</span>`:''}${s.status==='current'?'<span class="pill st-pending">Đang chờ</span>':''}</div>
        <div class="ppl">${s.approvers.length ? s.approvers.map(id => { const d = s.decisions[id];
          const dg = !d && (s.status === 'current' || preview) ? delegatesOf(id) : [];
          return `<div class="person">${avatar(id,'xs')}<span class="nm">${esc(user(id)?.name||'?')}<small>${esc(user(id)?.title||'')}${dg.length ? ` · ủy quyền cho ${esc(vnJoin(dg.map(g => user(g.to).name)))}` : ''}</small></span>
          ${d ? `<span class="dec ${d.act==='approve'?'ok':d.act==='reject'?'no':'ret'}">${d.act==='approve'?'Đã duyệt':d.act==='reject'?'Từ chối':'Trả lại'} ${dt(d.at)}${d.by && d.by !== id ? `<br><span class="deleg">${esc(user(d.by).name)} duyệt thay</span>` : ''}</span>` : !preview && s.status==='approved' && s.mode==='any' ? '<span class="dec">Không cần</span>' : ''}</div>
          ${d && d.note ? `<div class="pnote">${esc(d.note)}</div>` : ''}`; }).join('') : '<span class="small" style="color:var(--critical)">Chưa có người duyệt</span>'}</div>
        ${s.notes.map(n => `<div class="snote">${esc(n)}</div>`).join('')}
      </div></div>`).join('')}</div>`;
}
const fileChips = (files, dl, r) => files.map(x => { const n = typeof x === 'string' ? x : x.name; const k = typeof x === 'string' ? '' : FKIND[x.kind];
  return `<span class="chip file">${icon('file',13)} ${esc(fname(n))}${k ? `<span class="lvn">${k}</span>` : ''}<button type="button" data-dlunc="${esc(n)}" data-r="${r ? r.id : ''}" aria-label="Tải ${esc(fname(n))}" title="Tải về">${icon('upload',12)}</button></span>`; }).join('');
function viewDetail(){
  const r = req(UI.id); if (!r) return `<div class="card empty"><b>Không tìm thấy phiếu</b></div>`;
  const u = me(); const p = payeeOf(r); const mine = r.requester === u.id;
  const s = currentStep(r); const b = batch(r.batchId) || S.batches.find(x => x.lines && x.lines[r.id]);
  const rounds = [...r.rounds].reverse(); const E = UI.errors; const sum = linesSum(r.lines,'total');
  let actions = '';
  const actors = actorsFor(r, u.id);
  if (actors.length) actions = `<div class="card"><div class="ch"><h2 class="card-title">${actors[0] !== u.id && actors.length === 1 ? `Bạn duyệt thay ${esc(user(actors[0]).name)}, ${esc(s.label.toLowerCase())}` : `Bạn duyệt ${esc(s.label.toLowerCase())}`}</h2><span class="card-note">${s.approvers.length>1 ? (s.mode==='all' ? `cần cả ${s.approvers.length} người` : `một trong ${s.approvers.length} người là đủ`) : 'người duyệt duy nhất'}</span></div>
      ${actors.length > 1 ? `<label class="f" style="margin-bottom:12px">Duyệt với tư cách<select class="in" id="decAs">${actors.map(a => `<option value="${a}">${a === u.id ? 'Chính tôi' : 'Duyệt thay ' + esc(user(a).name)}</option>`).join('')}</select></label>` : `<input type="hidden" id="decAs" value="${actors[0]}">`}
      ${actors[0] !== u.id && actors.length === 1 ? `<p class="small muted" style="margin:0 0 12px">Theo ủy quyền của ${esc(user(actors[0]).name)}. Nhật ký sẽ ghi bạn duyệt thay.</p>` : ''}
      <label class="f">Ghi chú<textarea class="in ${E.note?'err':''}" id="decNote" placeholder="Bắt buộc khi từ chối hoặc trả lại"></textarea>${E.note?`<span class="ferr">${E.note}</span>`:''}</label>
      <div class="row" style="margin-top:12px"><button class="btn pri" data-decide="approve">Duyệt</button><button class="btn sec" data-decide="return">Trả lại bổ sung</button><span class="sp"></span><button class="btn danger" data-decide="reject">Từ chối</button></div></div>`;
  else if (mine && (r.status === 'returned' || r.status === 'draft')) actions = `<div class="alert warn"><b>${r.status==='returned'?'Phiếu cần bạn bổ sung.':'Bản nháp chưa gửi.'}</b> ${r.status==='returned'?'Sửa theo ghi chú của người duyệt rồi gửi lại.':'Hoàn thiện rồi gửi để bắt đầu duyệt.'}
      <div class="row" style="margin-top:10px"><button class="btn pri sm" data-go="edit" data-id="${r.id}">${r.status==='returned'?'Sửa và gửi lại':'Tiếp tục soạn'}</button></div></div>`;
  else if (r.status === 'pending' && s && Object.values(s.decisions).some(d => d.by === u.id)) actions = `<div class="alert ok">Bạn đã duyệt bước này.${s.mode==='all'?' Đang chờ những người còn lại.':''}</div>`;
  else if (r.status === 'pending' && mine) actions = `<div class="alert info">Đang chờ ${esc(vnJoin(waitingOn(r).map(id=>user(id).name), s.mode==='all'?' và ':' hoặc '))} duyệt ${esc(s.label.toLowerCase())}.</div>`;
  else if (r.status === 'approved') actions = `<div class="alert ok">Đã đủ chữ ký, đang chờ thanh toán. ${u.acct ? '<button class="linkbtn" data-go="batches">Đưa vào đợt chi</button>' : 'Kế toán sẽ lập đợt chi và hẹn ngày thanh toán.'}</div>`;
  else if (r.status === 'in_batch' && b) actions = `<div class="alert info">Trong đợt ${esc(b.code)}, <b>hẹn thanh toán ${dOnly(b.payDate+'T00:00')}</b> (${dueText(b.payDate)}). ${esc(batchNext(b).split(',')[0])}. ${seesBatches(u) ? `<button class="linkbtn" data-batch="${b.id}">Mở đợt chi</button>` : ''}</div>`;
  else if (r.status === 'cancelled') actions = `<div class="alert neg"><b>Lệnh đã huỷ${r.cancelledBy ? ' bởi ' + esc(user(r.cancelledBy).name) : ''}.</b> ${esc(r.cancelReason || '')}</div>`;
  const adv = advState(r);
  if (adv === 'waiting' || adv === 'overdue') { if (mine || u.acct) actions += `<div class="alert ${adv==='overdue'?'neg':'warn'}"><b>${adv==='overdue'?'Quá hạn quyết toán':'Chờ quyết toán tạm ứng'}, hạn ${dOnly(r.settleDue)} (${dueText(localDate(r.settleDue))}).</b> Khoản đã chi tạm ứng, cần nộp hoá đơn chi thực tế.<div class="row" style="margin-top:10px">${mine ? `<button class="btn pri sm" data-settle="${r.id}">Quyết toán tạm ứng</button>` : ''}${u.acct ? `<button class="btn sec sm" data-editsettle="${r.id}">Sửa hạn quyết toán</button>` : ''}</div></div>`; }
  else if (adv === 'submitted') actions += `<div class="alert info">Đang duyệt phiếu quyết toán ${esc(req(r.settle.reqId).code)}.</div>`;
  else if (adv === 'refund_waiting') actions += `<div class="alert warn"><b>Chờ nhận hoàn ứng ${money(r.settle.refund.amount)}.</b> Chi thực tế thấp hơn tạm ứng, ${esc(user(r.requester).name)} hoàn lại công ty.${u.acct ? `<div class="row" style="margin-top:10px"><button class="btn pri sm" data-refundsettle="${r.id}">Đã nhận tiền hoàn</button></div>` : ''}</div>`;
  if (invPending(r)) actions += `<div class="alert warn"><b>Chờ bổ sung hoá đơn, hạn ${dOnly(r.invDue)} (${dueText(localDate(r.invDue))}).</b> Khoản này đã chi khi chưa có hoá đơn.
      <div class="row" style="margin-top:10px">${mine || u.acct ? `<button class="btn pri sm" data-addinv="${r.id}">Bổ sung hoá đơn</button>` : ''}${u.acct && !mine ? `<button class="btn sec sm" data-remindinv="${r.id}">Nhắc người đề xuất</button>` : ''}</div></div>`;
  if (r.refundBack) actions += `<div class="alert ${r.refundBack.status==='received'?'ok':'warn'}"><b>${r.refundBack.status==='received' ? `NCC đã hoàn ${money(r.refundBack.amount)} ngày ${dOnly(r.refundBack.receivedAt)}.` : `Chuyển nhầm, chờ NCC hoàn ${money(r.refundBack.amount)}.`}</b> ${esc(r.refundBack.note||'')}
      ${r.refundBack.status==='waiting' && u.acct ? `<div class="row" style="margin-top:10px"><button class="btn pri sm" data-refundgot="${r.id}">Đã nhận tiền hoàn</button></div>` : ''}</div>`;

  const paySt = r.status === 'paid' ? 'done' : (r.status === 'approved' || (b && b.status !== 'review')) ? 'current' : 'waiting';
  const payTxt = r.status === 'paid' ? `${PAYM[pmOf(r)]}, đã chi ${dt(r.paidAt)}` : r.status === 'approved' ? 'Chờ kế toán lập đợt chi và hẹn ngày' : b ? `Đợt ${b.code} ${b.status === 'review' ? 'chờ duyệt' : 'đã duyệt'}, hẹn thanh toán ${dOnly(b.payDate+'T00:00')} (${dueText(b.payDate)})` : '';
  const payStep = ['approved','in_batch','paid'].includes(r.status) ? `<div class="chain" style="margin-top:18px"><div class="step ${paySt}"><div class="dot">${r.status==='paid'?'✓':'₫'}</div><div><div class="sh"><b>Thanh toán</b></div><div class="small faint">${esc(payTxt)}</div></div></div>
    ${r.evidence==='later' ? `<div class="step ${r.invDone?'done':r.status==='paid'?'current':'waiting'}"><div class="dot">${r.invDone?'✓':'HĐ'}</div><div><div class="sh"><b>Bổ sung hoá đơn</b></div><div class="small faint">${r.invDone ? `${esc(user(r.invDone.by).name)} bổ sung ${dt(r.invDone.at)}` : r.status==='paid' ? `Hạn ${dOnly(r.invDue)} (${dueText(localDate(r.invDue))})` : `Trong ${INV_DAYS} ngày sau khi chi`}</div></div></div>` : ''}</div>` : '';
  const more = [];
  if (canCancelOrder(r, u)) more.push(`<button class="btn danger sm" data-cancelorder="${r.id}">Huỷ lệnh</button>`);
  if (u.acct && r.status === 'paid' && !r.refundBack && r.payee.type === 'vendor') more.push(`<button class="btn ghost sm" data-wrongpay="${r.id}">Báo chuyển nhầm, chờ NCC hoàn</button>`);
  if (u.acct && r.payee.type === 'vendor' && !p.ok) more.push(`<button class="btn sec sm" data-verify="${p.vendorId}">Xác minh STK nhà cung cấp</button>`);

  return `<button class="crumb" data-go="list">‹ Đề xuất chi</button>
  <div class="head"><div><p class="eyebrow">${esc(r.code||'Nháp')}${r.type==='refund'?' · Hoàn tiền':''}${r.direct?' · Kế toán lập trực tiếp':''}</p><h1>${esc(r.title)}</h1><p>${esc(user(r.requester).name)}, ${esc(dept(r.dept).name)} · lập ${dt(r.createdAt)} &nbsp;${pill(r.status)}${invTag(r)}</p></div>
    <div style="text-align:right"><div class="amount-big">${money(r.amount)}</div><div class="small faint">${sum && r.amount < sum ? `đợt này, trên tổng ${money(sum)}` : esc(PAYM[pmOf(r)])}</div></div></div>
  <div class="g75 wide">
    <div class="stack">
      ${actions}
      <div class="card"><div class="ch"><h2 class="card-title">Thông tin chi</h2><div class="row">${more.join('')}</div></div>
        <div class="kv">
          <div><span>Thanh toán cho</span><b>${esc(p.name)}${p.note?` <span class="faint">(${p.note})</span>`:''}</b>${p.tax?`<div class="small faint">MST ${esc(p.tax)}</div>`:''}</div>
          <div><span>Tài khoản nhận</span><b>${esc(p.bank+' '+p.acc)}</b>${r.payee.type==='vendor'&&!p.ok?'<div class="small" style="color:var(--watch)">STK chưa xác minh</div>':`<div class="small faint">${esc(p.accName)}</div>`}</div>
          <div><span>Hình thức thanh toán</span><b>${PAYM[pmOf(r)]}</b>${r.payee.type==='me'?`<div class="small faint">Đã ứng bằng ${OUTLAY[r.payee.outlay||'transfer'].toLowerCase()}</div>`:''}</div>
          <div><span>Hạng mục chi phí</span>${u.acct && r.type !== 'refund' && r.status !== 'cancelled' ? `<select class="in" style="height:32px;margin-top:2px" data-recat="${r.id}" aria-label="Phân loại lại hạng mục">${catOptions(r.cat)}</select>` : `<b>${esc(r.cat)}</b>`}<div class="small faint">${esc(catGroup(r.cat))}</div></div>
          <div><span>Chứng từ</span><b>${r.type==='refund' ? 'Sao kê' : EVIDENCE[r.evidence][0]}</b></div>
          <div><span>Cần chi trước</span><b>${r.needBy ? dOnly(r.needBy) : 'Không yêu cầu'}</b></div>
        </div>
        ${r.note ? `<div class="pnote" style="margin:16px 0 0">${esc(r.note)}</div>` : ''}
        ${r.type !== 'refund' && r.lines.length ? `<div class="sect"><b class="small">Tiền và thuế</b><div class="tw" style="margin-top:8px"><table><thead><tr><th class="r">STT</th><th>Số hoá đơn</th><th>Diễn giải</th><th class="r">Trước thuế</th><th>Thuế suất</th><th class="r">Tiền thuế</th><th class="r">Thành tiền</th></tr></thead>
          <tbody>${r.lines.map((l,i) => `<tr><td class="r faint">${i+1}</td><td class="num">${esc(l.inv) || '<span class="faint">Chưa có</span>'}</td><td class="small" style="min-width:160px">${esc(l.desc)}</td><td class="r">${fmt(l.pre)}</td><td>${RATES[l.rate]}</td><td class="r">${fmt(l.tax)}</td><td class="r">${fmt(l.total)}</td></tr>`).join('')}</tbody>
          <tfoot><tr><td></td><td colspan="2">Tổng</td><td class="r">${fmt(linesSum(r.lines,'pre'))}</td><td></td><td class="r">${fmt(linesSum(r.lines,'tax'))}</td><td class="r">${fmt(sum)}</td></tr>
          ${r.amount < sum ? `<tr><td></td><td colspan="5" style="border-top:0;font-weight:500">Đề nghị thanh toán đợt này (${Math.round(r.amount/sum*100)}%)</td><td class="r" style="border-top:0">${fmt(r.amount)}</td></tr>` : ''}</tfoot></table></div></div>` : ''}
        <div class="sect"><b class="small">Chứng từ đính kèm</b><div class="chips" style="margin-top:8px">${r.files.length ? fileChips(r.files) : '<span class="small faint">Chưa có file</span>'}</div></div>
        ${r.unc && r.unc.length ? `<div class="sect"><b class="small">UNC, bằng chứng đã chi</b><div class="chips" style="margin-top:8px">${fileChips(r.unc, true, r)}</div><div class="small faint" style="margin-top:6px">Bấm mũi tên trên file để tải về.</div></div>` : ''}
      </div>
      <div class="card"><div class="ch"><h2 class="card-title">Nhật ký</h2><span class="card-note">${r.log.length} sự kiện</span></div>
        ${[...r.log].reverse().map(l => `<div class="li">${l.by ? avatar(l.by,'xs') : '<span class="av xs" aria-hidden="true">·</span>'}<div class="grow small"><b style="font-weight:500">${l.by?esc(user(l.by).name):'Hệ thống'}</b> <span class="muted">${esc(l.text)}</span></div><span class="small faint">${dt(l.at)}</span></div>`).join('')}
        <div class="row" style="margin-top:14px"><input class="in" id="cmt" placeholder="Viết bình luận" style="flex:1" aria-label="Bình luận"><button class="btn sec" data-action="comment">Gửi</button></div>
      </div>
    </div>
    <div class="stack">
      ${rounds.length ? rounds.map((rd,i) => `<div class="card" ${i?'style="opacity:.75"':''}><div class="ch"><h2 class="card-title">${i===0?'Tiến trình':'Vòng '+rd.no+' trước đó'}</h2><span class="card-note">${rounds.length>1&&i===0?`Vòng ${rd.no} · `:''}${rd.rule ? esc(rd.rule) + ' · ' : ''}${esc(rd.band)} · v${rd.cfgVersion}</span></div>
        ${chainHtml(rd.chain)}${i===0 ? payStep : ''}</div>`).join('')
      : `<div class="card"><div class="ch"><h2 class="card-title">Tiến trình</h2></div>${r.direct ? `<p class="small faint" style="margin:0">Kế toán lập trực tiếp, không qua chuỗi duyệt đề xuất. Người duyệt đợt chi là chốt duyệt duy nhất.</p>${payStep}` : '<p class="small faint" style="margin:0">Phiếu chưa gửi nên chưa có chuỗi duyệt.</p>'}</div>`}
    </div>
  </div>`;
}

/* ===================== Đợt chi — danh sách ===================== */
function queueRows(){ const f = UI.qflt; return S.reqs.filter(r => r.status === 'approved' && (!f.pm || pmOf(r) === f.pm)).sort((a,b) => (a.approvedAt||0) - (b.approvedAt||0)); }
function viewBatches(){
  const u = me(); const approved = queueRows();
  const sel = approved.filter(r => UI.batchSel[r.id]);
  const order = {review:0, approved:1, paying:1, paid:2, returned:3, cancelled:4};
  const bs = S.batches.filter(b => inRange(b.createdAt, UI.bflt)).sort((a,b) => order[a.status]-order[b.status] || b.createdAt-a.createdAt);
  if (!UI.payDateNew) UI.payDateNew = localDate();
  return `${head('Thanh toán','Đợt chi','Kế toán lập đợt và hẹn ngày thanh toán, người duyệt đợt duyệt, kế toán chi tiền rồi đính kèm UNC.')}
  <div class="stack">
    ${u.acct ? `<div class="card"><div class="ch"><h2 class="card-title">Chờ thanh toán</h2><div class="row"><span class="card-note">${approved.length} khoản, ${short(approved.reduce((a,r)=>a+r.amount,0))}</span>
        <select class="in" style="width:160px;height:30px" data-flt="qflt.pm" aria-label="Lọc hình thức">${['', ...Object.keys(PAYM)].map(k => `<option value="${k}" ${UI.qflt.pm===k?'selected':''}>${k ? PAYM[k] : 'Mọi hình thức'}</option>`).join('')}</select>
        <button class="btn sec sm" data-action="direct">${icon('plus',14)} Khoản chi trực tiếp</button></div></div>
      ${UI.errors.batch ? `<div class="alert neg" style="margin-bottom:12px">${esc(UI.errors.batch)}</div>` : ''}
      ${approved.length ? `<div class="tw"><table><thead><tr><th style="width:40px"><input type="checkbox" class="cb" data-bselall="1" aria-label="Chọn tất cả" ${sel.length === approved.length ? 'checked' : ''}></th><th class="r">STT</th><th>Mã</th><th>Ngày duyệt</th><th>Nội dung</th><th>Người nhận</th><th>Hình thức</th><th class="r">Số tiền (đ)</th></tr></thead><tbody>
        ${approved.map((r,i) => { const p = payeeOf(r); return `<tr><td><input type="checkbox" class="cb" aria-label="Chọn ${esc(r.code)}" data-bsel="${r.id}" ${UI.batchSel[r.id]?'checked':''}></td><td class="r faint">${i+1}</td>
          <td class="code"><button class="linkbtn" data-open="${r.id}">${esc(r.code)}</button></td><td class="small num" style="white-space:nowrap">${dt(r.approvedAt)}</td><td><div class="ttl">${esc(r.title)}</div><div class="small faint">${esc(user(r.requester).name)}${r.direct?' · lập trực tiếp, chưa qua đề xuất':''}${r.needBy?` · cần trước ${dOnly(r.needBy)}`:''}</div></td>
          <td class="small">${esc(p.name)}${p.note?` <span class="faint">(${p.note})</span>`:''}${p.ok?'':`<div style="color:var(--watch)">STK chưa xác minh <button class="linkbtn" data-verify="${p.vendorId}">Xác minh</button></div>`}</td><td class="small muted">${PAYM[pmOf(r)]}</td><td class="r">${fmt(r.amount)}</td></tr>`; }).join('')}</tbody></table></div>
        <div class="row" style="margin-top:14px"><span class="small faint">${sel.length ? `Đã chọn ${sel.length} khoản, ${money(sel.reduce((a,r)=>a+r.amount,0))}` : 'Chọn các khoản muốn gộp vào một đợt.'}</span><span class="sp"></span>
          <label class="fdate">Hẹn ngày thanh toán<input class="in" type="date" id="payDateNew" value="${UI.payDateNew}" min="${localDate()}"></label>
          <button class="btn pri" data-action="mkbatch" ${sel.length?'':'disabled'}>Lập đợt và trình duyệt</button></div>`
      : `<div class="empty"><b>Không có khoản nào chờ thanh toán</b>Đề xuất duyệt xong các bước sẽ xuất hiện ở đây. Kế toán cũng tự thêm được khoản chi trực tiếp.</div>`}</div>` : ''}
    <div class="card"><div class="ch"><h2 class="card-title">Các đợt chi</h2><span class="card-note">Chuỗi duyệt theo tổng tiền đợt, cấu hình ở mục Cấu hình</span></div>
      ${filterBar('bflt', {extra:`<span class="sp"></span><button class="btn sec sm" data-action="xbatches">Xuất Excel</button>`})}
      ${bs.length ? `<div class="tw"><table><thead><tr><th class="r">STT</th><th>Mã</th><th>Ngày giờ lập</th><th>Hẹn thanh toán</th><th class="r">Số khoản</th><th class="r">Tổng (đ)</th><th>Trạng thái</th><th>Bước tiếp theo</th></tr></thead><tbody>
        ${bs.map((b,i) => `<tr class="click" data-batch="${b.id}" tabindex="0"><td class="r faint">${i+1}</td><td class="code">${esc(b.code)}</td><td class="small num" style="white-space:nowrap">${dt(b.createdAt)}<div class="faint">${esc(user(b.createdBy).name)}</div></td>
          <td class="small num">${dOnly(b.payDate+'T00:00')}${['approved','paying','review'].includes(b.status) ? `<div class="${dayDiff(b.payDate) < 0 ? '' : 'faint'}" ${dayDiff(b.payDate) < 0 ? 'style="color:var(--critical)"' : ''}>${dueText(b.payDate)}</div>` : ''}</td>
          <td class="r">${b.reqIds.length}</td><td class="r">${fmt(batchTotal(b))}</td><td>${bpill(b.status)}</td><td class="small muted">${esc(batchNext(b))}</td></tr>`).join('')}
      </tbody></table></div>` : `<div class="empty"><b>Không có đợt chi nào khớp</b></div>`}</div>
  </div>`;
}

/* ===================== Đợt chi — chi tiết ===================== */
function bankRow(b, id, i){ const r = req(id), p = payeeOf(r); return [i+1, p.accName, String(p.acc).toUpperCase(), bankText(p.bank), r.amount, memoOf(r), r.code, uncName(b, id)]; }
const BANK_HEAD = ['STT','TEN NGUOI THU HUONG','SO TAI KHOAN','NGAN HANG','SO TIEN','NOI DUNG','MA PHIEU','TEN FILE UNC'];
const transferIds = b => b.reqIds.filter(id => pmOf(req(id)) === 'transfer');
function viewBatch(){
  const b = batch(UI.id); if (!b) return `<div class="card empty"><b>Không tìm thấy đợt chi</b></div>`;
  const u = me(); const E = UI.errors; const total = batchTotal(b); const paid = paidCount(b);
  const open = b.status === 'approved' || b.status === 'paying';
  const transferring = open && u.acct;
  const canEditLines = (u.acct || u.admin) && ['review','approved','paying'].includes(b.status);
  const pf = UI.bpm || '';
  const ids = b.reqIds.filter(id => !pf || pmOf(req(id)) === pf);
  const unpaidCk = b.reqIds.filter(id => !b.lines[id].paidAt && pmOf(req(id)) === 'transfer');
  const sel = unpaidCk.filter(id => UI.paySel[id]);
  const bact = batchActors(b, u.id); const directN = b.reqIds.filter(id => req(id).direct).length;
  let action = '';
  if (bact.length) action = `<div class="card"><div class="ch"><h2 class="card-title">${bact[0] !== u.id ? `Bạn duyệt thay ${esc(user(bact[0]).name)}` : 'Bạn duyệt đợt chi này'}</h2><span class="card-note">${b.reqIds.length} khoản, ${money(total)}</span></div>
      <p class="small muted" style="margin:0 0 12px">${directN ? `${directN} khoản do kế toán lập trực tiếp, chưa qua đề xuất: bạn là chốt duyệt duy nhất của các khoản đó. ` : 'Mỗi khoản đã qua đủ các bước duyệt. '}Đây là chốt cuối trước khi tiền rời công ty.</p>
      <label class="f">Ghi chú<textarea class="in ${E.bnote?'err':''}" id="bNote" placeholder="Bắt buộc khi trả lại kế toán"></textarea>${E.bnote?`<span class="ferr">${E.bnote}</span>`:''}</label>
      ${bact.length > 1 ? `<label class="f" style="margin-top:12px">Duyệt với tư cách<select class="in" id="bAs">${bact.map(a => `<option value="${a}">${a === u.id ? 'Chính tôi' : 'Duyệt thay ' + esc(user(a).name)}</option>`).join('')}</select></label>` : `<input type="hidden" id="bAs" value="${bact[0]}">`}
      <div class="row" style="margin-top:12px"><button class="btn pri" data-bapprove="${b.id}">Duyệt đợt chi</button><button class="btn sec" data-breturn="${b.id}">Trả lại kế toán</button></div></div>`;
  else if (b.status === 'review') action = `<div class="alert info">${esc(batchNext(b))}.${b.createdBy === u.id && isSigner(u.id) ? ' Bạn lập đợt này nên không tự duyệt.' : ''}</div>`;
  else if (transferring) action = `<div class="alert ${dayDiff(b.payDate) < 0 ? 'warn' : 'ok'}"><b>${esc(user(b.approvedBy).name)}${b.approvedFor ? ' (duyệt thay ' + esc(user(b.approvedFor).name) + ')' : ''} đã duyệt ${dt(b.approvedAt)}.</b>${b.approveNote ? ` “${esc(b.approveNote)}”` : ''} Hẹn thanh toán ${dOnly(b.payDate+'T00:00')} (${dueText(b.payDate)}). Chuyển khoản xác nhận được hàng loạt, Visa và tiền mặt xác nhận từng khoản.</div>`;
  else if (open) action = `<div class="alert info">Đã duyệt, kế toán chi tiền ${paid}/${b.reqIds.length} khoản, hẹn ${dOnly(b.payDate+'T00:00')}.</div>`;
  else if (b.status === 'paid') action = `<div class="alert ok">Đã chi xong ${b.reqIds.length} khoản lúc ${dt(b.paidAt)}.</div>`;
  else if (b.status === 'returned') action = `<div class="alert warn"><b>${esc(user(b.returnedBy).name)}${b.returnedFor ? ' (thay ' + esc(user(b.returnedFor).name) + ')' : ''} trả lại:</b> ${esc(b.returnNote)} Các khoản đã về lại chờ thanh toán.</div>`;
  else if (b.status === 'cancelled') action = `<div class="alert neg"><b>Đợt đã huỷ.</b> ${esc(b.cancelReason || '')}</div>`;

  const copyBtn = (text, label) => `<button class="copy" data-copy="${esc(text)}" aria-label="Sao chép ${label}" title="Sao chép ${label}">${icon('copy')}</button>`;
  const rows = ids.map(id => { const r = req(id); const p = payeeOf(r); const ln = b.lines[id]; const memo = memoOf(r); const pm = pmOf(r); const i = b.reqIds.indexOf(id);
    return `<tr>
      ${transferring ? `<td>${ln.paidAt || pm !== 'transfer' ? '' : `<input type="checkbox" class="cb" aria-label="Chọn ${esc(r.code)}" data-psel="${id}" ${UI.paySel[id]?'checked':''}>`}</td>` : ''}
      <td class="r faint">${i+1}</td>
      <td style="min-width:210px"><div class="ttl">${esc(p.accName)}</div><div class="small faint"><button class="linkbtn" data-open="${r.id}">${esc(r.code)}</button> ${esc(r.title)}${r.direct?' · lập trực tiếp':''}</div></td>
      <td class="small" style="white-space:nowrap">${esc(p.bank)}<div class="num">${esc(p.acc)}${transferring && pm === 'transfer' ? copyBtn(p.acc,'số tài khoản') : ''}</div><div class="faint">${PAYM[pm]}</div></td>
      <td class="r" style="white-space:nowrap">${fmt(r.amount)}${transferring && pm === 'transfer' ? copyBtn(String(r.amount),'số tiền') : ''}</td>
      <td class="small muted" style="min-width:200px;max-width:300px">${pm === 'transfer' ? esc(memo) + (transferring ? copyBtn(memo,'nội dung chuyển khoản') : '') : '<span class="faint">Không chuyển khoản</span>'}</td>
      <td>${ln.paidAt ? `<div class="paid-mark">Đã chi ${dt(ln.paidAt)}</div><div class="chips" style="margin-top:4px">${fileChips(ln.unc, true, r)}</div>`
        : transferring ? `<button class="btn sec sm" data-payone="${id}">${icon('upload',14)} ${pm === 'transfer' ? 'Đã chuyển' : pm === 'visa' ? 'Đã chi thẻ' : 'Đã chi tiền mặt'}</button>` : '<span class="small faint">Chưa chi</span>'}</td>
      <td style="white-space:nowrap;text-align:right">${pm === 'transfer' ? `<button class="copy" data-copyrow="${id}" aria-label="Sao chép dòng ${i+1}" title="Sao chép cả dòng">${icon('copy')}</button>` : ''}${canEditLines && !ln.paidAt ? `<button class="btn ghost sm" data-dropline="${id}" data-kind="removed">Gỡ</button><button class="btn ghost sm" data-dropline="${id}" data-kind="cancelled" style="color:var(--critical)">Huỷ</button>` : ''}</td></tr>`; }).join('');
  const steps = [
    {st:'done', t:'Kế toán lập đợt', m:`${user(b.createdBy).name}, ${dt(b.createdAt)}`},
    {st: b.status==='review' ? 'current' : b.status==='returned' ? 'returned' : b.status==='cancelled' && !b.approvedBy ? 'rejected' : 'done', t:'Duyệt đợt chi', m: b.approvedBy ? `${user(b.approvedBy).name}${b.approvedFor ? ' thay ' + user(b.approvedFor).name : ''}, ${dt(b.approvedAt)}` : b.status==='returned' ? `Trả lại bởi ${user(b.returnedBy).name}` : batchChainText(b)},
    {st: b.status==='paid' ? 'done' : open ? 'current' : 'waiting', t:'Chi tiền và đính kèm UNC', m:`${paid}/${b.reqIds.length} khoản, hẹn ${dOnly(b.payDate+'T00:00')}${b.paidAt?', xong '+dt(b.paidAt):''}`}
  ];
  const cnt = k => b.reqIds.filter(id => pmOf(req(id)) === k).length;
  return `<button class="crumb" data-go="batches">‹ Đợt chi</button>
  <div class="head"><div><p class="eyebrow">Đợt chi</p><h1>${esc(b.code)}</h1><p>${b.reqIds.length} khoản · lập bởi ${esc(user(b.createdBy).name)} ${dt(b.createdAt)} &nbsp;${bpill(b.status)}</p></div>
    <div style="text-align:right"><div class="amount-big">${money(total)}</div><div class="small faint">${paid ? `đã chi ${money(b.reqIds.filter(id=>b.lines[id].paidAt).reduce((a,id)=>a+req(id).amount,0))}` : 'tổng đợt'}</div></div></div>
  <div class="stack">
    <div class="g75">
      <div class="stack">${action || '<span></span>'}</div>
      <div class="card"><div class="ch"><h2 class="card-title">Tiến trình</h2>${u.acct && ['review','approved','paying'].includes(b.status) ? `<label class="fdate">Hẹn thanh toán<input class="in" type="date" value="${b.payDate}" data-bdate="${b.id}" style="height:30px"></label>` : ''}</div>
        <div class="chain">${steps.map((s,i) => `<div class="step ${s.st}"><div class="dot">${s.st==='done'?'✓':s.st==='returned'?'↩':s.st==='rejected'?'✕':i+1}</div><div><div class="sh"><b>${esc(s.t)}</b></div><div class="small faint">${esc(s.m)}</div></div></div>`).join('')}</div>
        ${u.acct && ['review','approved'].includes(b.status) && !paid ? `<div class="row" style="margin-top:14px"><button class="btn danger sm" data-cancelbatch="${b.id}">Huỷ cả đợt</button></div>` : ''}
      </div>
    </div>
    <div class="card"><div class="ch"><h2 class="card-title">Danh sách khoản chi</h2>
        <div class="row"><div class="seg" role="group" aria-label="Lọc hình thức">${[['','Tất cả',b.reqIds.length],['transfer','Chuyển khoản',cnt('transfer')],['visa','Visa',cnt('visa')],['cash','Tiền mặt',cnt('cash')]].filter(x => !x[0] || x[2]).map(x => `<button type="button" class="${pf===x[0]?'on':''}" data-bpm="${x[0]}" aria-pressed="${pf===x[0]}">${x[1]} ${x[2]}</button>`).join('')}</div></div></div>
      <div class="row" style="margin-bottom:12px">
        <button class="btn sec sm" data-copysheet="${b.id}">${icon('copy',14)} Sao chép bảng</button><button class="btn sec sm" data-xbank="${b.id}">Xuất Excel ngân hàng</button>${paid ? `<button class="btn sec sm" data-zunc="${b.id}">Tải ${paid} UNC (.zip)</button>` : ''}
        <span class="sp"></span>${transferring && unpaidCk.length ? `<button class="btn sec sm" data-selck="${b.id}">${sel.length === unpaidCk.length ? 'Bỏ chọn' : `Chọn cả ${unpaidCk.length} khoản chuyển khoản`}</button><button class="btn pri sm" data-paysel="${b.id}" ${sel.length?'':'disabled'}>${sel.length ? `Xác nhận đã chuyển ${sel.length} khoản` : 'Xác nhận đã chuyển khoản'}</button>` : ''}</div>
      <div class="tw"><table><thead><tr>${transferring?`<th style="width:40px"><input type="checkbox" class="cb" data-pselall="${b.id}" aria-label="Chọn tất cả khoản chuyển khoản" ${unpaidCk.length && sel.length === unpaidCk.length ? 'checked' : ''} ${unpaidCk.length ? '' : 'disabled'}></th>`:''}<th class="r">STT</th><th>Người thụ hưởng</th><th>Ngân hàng, STK, hình thức</th><th class="r">Số tiền (đ)</th><th>Nội dung chuyển khoản</th><th>UNC</th><th></th></tr></thead>
        <tbody>${rows || `<tr><td colspan="8" class="small faint" style="text-align:center;padding:20px">Không có khoản nào theo bộ lọc.</td></tr>`}</tbody><tfoot><tr>${transferring?'<td></td>':''}<td></td><td colspan="2">Tổng ${ids.length} khoản${pf ? ' ' + PAYM[pf].toLowerCase() : ''}</td><td class="r">${fmt(ids.reduce((a,id)=>a+req(id).amount,0))}</td><td colspan="3"></td></tr></tfoot></table></div>
      <p class="small faint" style="margin:10px 0 0">Tên thụ hưởng và nội dung theo chuẩn ngân hàng: in hoa, không dấu. File Excel có cột “TEN FILE UNC” cho từng giao dịch; đặt tên UNC theo cột đó thì khi upload hàng loạt hệ thống tự ghép đúng khoản.</p>
      ${b.removed && b.removed.length ? `<div class="sect"><b class="small">Khoản đã gỡ hoặc huỷ khỏi đợt</b>${b.removed.map(x => `<div class="li"><div class="grow small"><button class="linkbtn" data-open="${x.id}">${esc(req(x.id).code)}</button> ${esc(req(x.id).title)} <span class="faint">· ${x.kind==='cancelled'?'huỷ lệnh':'gỡ, về lại chờ thanh toán'} · ${esc(x.reason)}</span></div><span class="small faint">${esc(user(x.by).name)}, ${dt(x.at)}</span><span class="amt">${fmt(req(x.id).amount)}</span></div>`).join('')}</div>` : ''}
    </div>
    <div class="card"><div class="ch"><h2 class="card-title">Lịch sử đợt chi</h2><span class="card-note">${(b.log||[]).length} sự kiện</span></div>
      ${[...(b.log||[])].reverse().map(l => `<div class="li">${avatar(l.by,'xs')}<div class="grow small"><b style="font-weight:500">${esc(user(l.by).name)}</b> <span class="muted">${esc(l.text)}</span></div><span class="small faint">${dt(l.at)}</span></div>`).join('')}</div>
  </div>`;
}

function payModal(bid, ids){
  UI.payDraft = {bid, ids, files:[], map:{}, why:{}, mode: ids.length > 1 ? 'each' : 'one', date:localDate(), err:''};
  drawPayModal();
}
/* Gợi ý ghép UNC với khoản theo tên file: mã phiếu, STK người nhận, số tiền, tên người nhận. Bản thật đọc nội dung UNC (OCR/parse PDF) để lấy STK và số tiền. */
function uncMatch(name, r){
  const flat = noAccent(fname(name)).toLowerCase().replace(/[^a-z0-9]/g,'');
  const nums = (fname(name).match(/\d[\d.,]*/g) || []).map(x => x.replace(/[.,]/g,''));
  const p = payeeOf(r);
  if (flat.includes(r.code.toLowerCase().replace(/[^a-z0-9]/g,''))) return ['mã phiếu', 4];
  const acc = String(p.acc).replace(/\D/g,''); if (acc.length >= 6 && nums.some(n => n.includes(acc))) return ['STK', 3];
  if (nums.some(n => n === String(r.amount))) return ['số tiền', 2];
  const words = noAccent(p.name).toLowerCase().split(/\s+/).filter(w => w.length > 1);
  const tail = words.slice(-2).join(''); if (tail.length >= 5 && flat.includes(tail)) return ['tên người nhận', 1];
  return null;
}
/* Ghép theo độ tin cậy cao trước (mã phiếu > STK > số tiền > tên). Một file khớp ngang nhau với nhiều khoản (cùng NCC, cùng số tiền) thì không tự ghép. */
function autoMapUnc(){
  const P = UI.payDraft; if (P.mode !== 'each') return 0;
  const used = new Set(Object.values(P.map)); const pairs = [];
  P.ids.forEach(id => { if (P.map[id] != null) return; P.files.forEach((f,i) => { if (used.has(i)) return; const m = uncMatch(f, req(id)); if (m) pairs.push({id, i, why:m[0], sc:m[1]}); }); });
  pairs.sort((a,b) => b.sc - a.sc); let n = 0; P.ambig = P.ambig || {};
  pairs.forEach(pr => {
    if (P.map[pr.id] != null || used.has(pr.i)) return;
    const rivals = pairs.filter(x => x.i === pr.i && x.sc === pr.sc && x.id !== pr.id && P.map[x.id] == null);
    const rivalsF = pairs.filter(x => x.id === pr.id && x.sc === pr.sc && x.i !== pr.i && !used.has(x.i));
    if (rivals.length || rivalsF.length) { P.ambig[pr.id] = pr.why; return; }
    P.map[pr.id] = pr.i; P.why[pr.id] = pr.why; delete P.ambig[pr.id]; used.add(pr.i); n++;
  });
  return n;
}
function addUncFiles(names){
  const P = UI.payDraft; names.forEach(n => { if (!P.files.includes(n)) P.files.push(n); });
  if (P.ids.length > 1 && P.files.length === 1 && P.mode === 'each' && !Object.keys(P.map).length) P.mode = 'one';
  if (P.ids.length > 1 && P.files.length > 1 && P.mode === 'one') P.mode = 'each';
  autoMapUnc(); P.err = '';
}
function payIssues(){
  const P = UI.payDraft; const out = [];
  if (!P.files.length) return ['Đính kèm ít nhất một file UNC.'];
  if (P.mode === 'each') {
    const miss = P.ids.filter(id => P.map[id] == null); if (miss.length) out.push(`${miss.length} khoản chưa gán UNC: ${miss.map(id => req(id).code).join(', ')}.`);
    const usedIdx = new Set(Object.values(P.map)); const idle = P.files.filter((f,i) => !usedIdx.has(i)); if (idle.length) out.push(`File chưa gán khoản nào: ${idle.join(', ')}. Gán vào khoản hoặc bỏ file.`);
  }
  return out;
}
function drawPayModal(){
  const P = UI.payDraft; const b = batch(P.bid); const sum = P.ids.reduce((a,id)=>a+req(id).amount,0);
  const multi = P.ids.length > 1;
  const count = {}; Object.values(P.map).forEach(i => count[i] = (count[i]||0) + 1);
  const mapped = P.ids.filter(id => P.map[id] != null).length;
  const mapTable = multi && P.mode === 'each' && P.files.length ? `
    <div class="tw"><table><thead><tr><th>Khoản</th><th class="r">Số tiền (đ)</th><th>UNC</th></tr></thead><tbody>
    ${P.ids.map(id => { const r = req(id); const p = payeeOf(r); const i = P.map[id];
      return `<tr><td><div class="ttl">${esc(p.name)}</div><div class="small faint">${esc(r.code)} · ${esc(p.bank)} ${esc(p.acc)}</div></td>
        <td class="r" style="white-space:nowrap">${fmt(r.amount)}</td>
        <td style="min-width:220px"><select class="in ${i==null?'err':''}" style="height:34px" data-umap="${id}" aria-label="UNC cho ${esc(r.code)}"><option value="">Chọn file UNC</option>${P.files.map((f,fi) => `<option value="${fi}" ${fi===i?'selected':''}>${esc(f)}</option>`).join('')}</select>
          <div class="small ${i==null?'':'faint'}" style="margin-top:3px;${i==null?'color:var(--critical)':''}">${i==null ? (P.ambig && P.ambig[id] ? `Nhiều file cùng khớp theo ${P.ambig[id]}, chọn thủ công` : 'Chưa gán') : P.why[id] ? 'Tự khớp theo ' + P.why[id] : 'Chọn thủ công'}${i!=null && count[i] > 1 ? ` · file này dùng cho ${count[i]} khoản` : ''}</div></td></tr>`; }).join('')}
    </tbody></table></div>` : '';
  const iss = payIssues();
  $('#layer').innerHTML = `<div class="modal-bg" data-close="1"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="ph" style="width:min(${multi?760:560}px,100%)">
    <div class="modal-h"><h3 id="ph">${!multi && pmOf(req(P.ids[0])) === 'visa' ? 'Xác nhận đã chi thẻ Visa' : !multi && pmOf(req(P.ids[0])) === 'cash' ? 'Xác nhận đã chi tiền mặt' : 'Xác nhận đã chuyển khoản'}</h3><p>${!multi ? `${esc(req(P.ids[0]).code)} · ${esc(payeeOf(req(P.ids[0])).name)}` : `${P.ids.length} khoản trong đợt ${esc(b.code)}`} · ${money(sum)}</p></div>
    <div class="modal-b stack">
      <label class="drop">${icon('upload',20)}<b style="font-weight:500;color:var(--ink)">${multi ? 'Kéo thả hoặc chọn nhiều file UNC cùng lúc' : pmOf(req(P.ids[0])) === 'transfer' ? 'Đính kèm UNC' : 'Đính kèm sao kê thẻ hoặc phiếu chi'}</b><span class="small faint">PDF hoặc ảnh. ${multi ? 'Hệ thống tự ghép file với khoản theo tên file: đúng cột “TEN FILE UNC” trong Excel, hoặc có mã phiếu, STK, số tiền, tên người nhận.' : 'Uỷ nhiệm chi hoặc sao kê ngân hàng.'}</span><input type="file" multiple data-unc></label>
      ${multi && !P.files.length ? `<div class="row"><span class="small faint" style="flex:1">Chưa có file thật để thử?</span><button type="button" class="btn ghost sm" data-uncdemo="1">Thêm ${P.ids.length} file UNC mẫu</button></div>` : ''}
      ${P.files.length ? `<div class="chips">${P.files.map((n,i)=>`<span class="chip file ${multi && P.mode==='each' && !count[i] ? 'warnc' : ''}" ${multi && P.mode==='each' && !count[i] ? 'title="Chưa gán khoản nào"' : ''}>${esc(fname(n))}<button type="button" data-rmunc="${i}" aria-label="Bỏ ${esc(fname(n))}">${icon('x',12)}</button></span>`).join('')}</div>` : ''}
      ${multi && P.files.length ? `<div class="row"><div class="seg" role="group" aria-label="Cách gắn UNC"><button type="button" class="${P.mode==='each'?'on':''}" data-umode="each" aria-pressed="${P.mode==='each'}">Mỗi khoản một UNC</button><button type="button" class="${P.mode==='one'?'on':''}" data-umode="one" aria-pressed="${P.mode==='one'}">Dùng chung cho cả ${P.ids.length} khoản</button></div>
        <span class="sp"></span>${P.mode==='each' ? `<span class="small faint">${mapped}/${P.ids.length} khoản đã gán</span>${mapped < P.ids.length && P.files.length >= P.ids.length ? '<button type="button" class="btn ghost sm" data-uncorder="1">Gán theo thứ tự file</button>' : ''}` : ''}</div>` : ''}
      ${mapTable}
      ${multi && P.mode==='one' && P.files.length ? `<p class="small faint" style="margin:0">Cả ${P.files.length} file sẽ gắn vào từng khoản trong ${P.ids.length} khoản, phù hợp khi ngân hàng trả một UNC cho cả lệnh chuyển theo lô.</p>` : ''}
      ${P.err ? `<div class="alert neg">${P.err}</div>` : ''}
      <label class="f" style="max-width:220px">Ngày chuyển<input class="in" type="date" id="payDate" value="${P.date}"></label>
    </div>
    <div class="modal-f"><span class="small faint" style="flex:1;align-self:center">${P.files.length && iss.length ? esc(iss[0]) : ''}</span><button class="btn sec" data-close="1">Huỷ</button><button class="btn pri" data-action="confirmpay" ${P.files.length && !iss.length ? '' : 'disabled'}>Xác nhận đã chi</button></div></div></div>`;
}

/* ===================== Cấu hình ===================== */
function inConfig(id, C){ C = C || S.config; return C.rules.some(rl => rl.bands.some(b => b.levels.some(l => l.approvers.includes(id)))) || C.fallback.includes(id) || (C.noInvoice.on && C.noInvoice.approvers.includes(id)) || !!(C.batch && C.batch.bands.some(b => b.levels.some(l => l.approvers.includes(id)))); }
function viewConfig(){
  if (!UI.draft) UI.draft = {config: clone(S.config)};
  const D = UI.draft; const admin = me().admin;
  const tabs = admin ? [['request','Đề xuất chi'],['batch','Đợt chi'],['cats','Hạng mục chi phí'],['vendors','NCC / người nhận']]
    : [['cats','Hạng mục chi phí'],['vendors','NCC / người nhận']];
  const tab = tabs.some(t => t[0] === UI.cfgTab) ? UI.cfgTab : tabs[0][0]; UI.cfgTab = tab;
  const body = tab === 'request' ? `${cfgMatrix(D)}${cfgRules(D)}` : tab === 'batch' ? cfgBatch(D) : tab === 'cats' ? cfgCats() : vendorsDirHtml();
  const approval = tab === 'request' || tab === 'batch';
  return `${head('Quản trị','Cấu hình', admin ? `Đang áp dụng v${S.config.version}, cập nhật ${dt(S.config.updatedAt)} bởi ${esc(user(S.config.updatedBy)?.name||'—')}. Thay đổi phê duyệt chỉ áp cho đề xuất gửi sau khi lưu.` : 'Hạng mục chi phí và danh bạ NCC: thay đổi có hiệu lực ngay.')}
  <div class="tabs" role="tablist">${tabs.map(t => `<button role="tab" aria-selected="${tab===t[0]}" class="${tab===t[0]?'on':''}" data-cfgtab="${t[0]}">${t[1]}</button>`).join('')}</div>
  <div class="${approval ? 'g75' : ''}">
    <div class="stack">${body}</div>
    ${approval ? `<div class="stack sticky" id="cfgSide" style="position:sticky;top:24px"><div class="card" id="sim" ${tab==='request'?'':'hidden'}></div><div id="issues">${issuesHtml(D)}</div></div>` : ''}
  </div>
  <div id="savebarHost">${approval ? savebarHtml(D) : ''}</div>`;
}
function isDirty(D){ return JSON.stringify(D.config) !== JSON.stringify(S.config); }
function issuesHtml(D){ const is = cfgIssues(D); if (!is.length) return '';
  return `<div class="alert ${is.some(i=>i.err)?'neg':'warn'}"><b>${is.some(i=>i.err) ? 'Cần sửa trước khi lưu' : 'Lưu ý'}</b><ul>${is.map(i=>`<li>${esc(i.t)}</li>`).join('')}</ul></div>`; }
function savebarHtml(D){ if (!isDirty(D)) return ''; const bad = cfgIssues(D).some(i=>i.err);
  return `<div class="savebar" role="region" aria-label="Thay đổi chưa lưu"><span style="flex:1">Có thay đổi chưa lưu${bad?', còn lỗi cần sửa':''}.</span>
    <button class="btn ghost sm" data-action="cfgreset">Bỏ thay đổi</button><button class="btn pri sm" data-action="cfgsave" ${bad?'disabled':''}>Lưu thành v${S.config.version+1}</button></div>`; }
function personChips(list, path, pool){
  const avail = (pool ? pool.map(id => user(id)).filter(Boolean) : activeUsers()).filter(u => u.active && !u.deleted && !list.includes(u.id));
  return `<div class="chips">${list.map(id => { const u = user(id); const bad = !u || !u.active || u.deleted;
    return `<span class="chip ${bad?'warnc':''}" title="${bad?'Người này đã ngừng hoạt động':''}">${avatar(id,'xs')}${esc(u?.name||'?')}<span class="lvn">${esc(u?.title||'')}</span><button type="button" data-rm="${path}" data-uid="${id}" aria-label="Bỏ ${esc(u?.name||'')}">${icon('x',12)}</button></span>`; }).join('')}
    ${avail.length ? `<select class="add" data-add="${path}" aria-label="Thêm người duyệt"><option value="">+ Thêm người</option>${avail.map(u => `<option value="${u.id}">${esc(u.name)} · ${esc(u.title)}</option>`).join('')}</select>` : '<span class="small faint">Đã chọn hết</span>'}</div>`;
}
const allBands = C => C.rules.flatMap(r => r.bands).concat(C.batch ? C.batch.bands : []);
const findBand = id => { const B0 = UI.draft.config.batch; const bb = B0 && B0.bands.find(x => x.id === id); if (bb) return {rule:B0, band:bb}; for (const r of UI.draft.config.rules) { const b = r.bands.find(x => x.id === id); if (b) return {rule:r, band:b}; } return {}; };
function ruleSummary(rl){
  return sortedBands(rl).map(b => `<div class="rsum"><span class="rb">${esc(bandRange(b, rl).replace(/^./, x => x.toUpperCase()))}</span><span>${b.levels.map((lv,li) => `<b>Cấp ${li+1}:</b> ${lv.approvers.length ? esc(lv.approvers.map(id => (user(id)||{}).name || '?').join(lv.mode === 'all' ? ' và ' : ', ')) : '<span style="color:var(--critical)">chưa chọn</span>'}`).join('<br>')}</span></div>`).join('');
}
function scopeChips(rl){
  const usedD = UI.draft.config.rules.flatMap(r => r.depts), usedP = UI.draft.config.rules.flatMap(r => r.people);
  const dAvail = S.depts.filter(d => !usedD.includes(d.id)), pAvail = activeUsers().filter(u => !usedP.includes(u.id));
  return `<div class="fgrid">
    <div class="f"><span class="flabel">Phòng ban</span><div class="chips">${rl.depts.map(id => `<span class="chip" style="padding-left:10px">${esc((dept(id)||{}).name||id)}<button type="button" data-rmscope="${rl.id}" data-k="depts" data-v="${id}" aria-label="Bỏ phòng">${icon('x',12)}</button></span>`).join('')}
      ${dAvail.length ? `<select class="add" data-addscope="${rl.id}" data-k="depts" aria-label="Thêm phòng ban"><option value="">+ Thêm phòng</option>${dAvail.map(d => `<option value="${d.id}">${esc(d.name)}</option>`).join('')}</select>` : ''}</div><span class="hint">Tính theo phòng chịu chi phí trên đề xuất.</span></div>
    <div class="f"><span class="flabel">Người đề xuất cụ thể</span><div class="chips">${rl.people.map(id => `<span class="chip">${avatar(id,'xs')}${esc((user(id)||{}).name||'?')}<button type="button" data-rmscope="${rl.id}" data-k="people" data-v="${id}" aria-label="Bỏ người">${icon('x',12)}</button></span>`).join('')}
      ${pAvail.length ? `<select class="add" data-addscope="${rl.id}" data-k="people" aria-label="Thêm người đề xuất"><option value="">+ Thêm người</option>${pAvail.map(u => `<option value="${u.id}">${esc(u.name)} · ${esc(u.title)}</option>`).join('')}</select>` : ''}</div><span class="hint">Ưu tiên hơn quy định theo phòng.</span></div>
  </div>`;
}
function bandsEditor(rl, pool){
  const bs = sortedBands(rl);
  return bs.map((b,bi) => { const lo = bi>0 ? bs[bi-1].max : 0; const last = b.max == null;
    return `<div class="band">
      <div class="band-h"><div class="range"><span class="bno">Mốc ${bi+1}</span>${lo ? `trên ${fmt(lo)} đ` : 'từ 0 đ'}<span class="faint">đến</span>
        ${last ? '<span>không giới hạn</span>' : `<div class="money"><input class="in" data-max="${b.id}" inputmode="numeric" value="${fmt(b.max)}" data-money aria-label="Mức trần mốc ${bi+1}"><span>đ</span></div>`}</div>
        <div class="row" style="margin-left:auto;flex-wrap:nowrap"><span class="small faint">Cấp duyệt</span>
        <div class="seg" role="group" aria-label="Số cấp duyệt mốc ${bi+1}">${[1,2,3].map(n => `<button type="button" class="${b.levels.length===n?'on':''}" data-nlv="${b.id}" data-n="${n}" aria-pressed="${b.levels.length===n}">${n} cấp</button>`).join('')}</div>
        ${bs.length>1 && !last ? `<button class="btn ghost sm" data-rmband="${b.id}" aria-label="Xoá mốc ${bi+1}">${icon('x',14)}</button>` : ''}</div></div>
      <div class="band-b">${b.levels.map((lv,li) => `<div class="lvl">
        <div class="lvl-n"><b>Cấp ${li+1}</b><span>${li===0 ? 'duyệt đầu tiên' : 'sau khi cấp ' + li + ' xong'}</span></div>
        <div>${personChips(lv.approvers, `${b.id}.${li}`, pool)}
          ${lv.approvers.length>1 ? `<div class="lvl-opts"><label>Khi nhiều người <select data-mode="${b.id}.${li}"><option value="any" ${lv.mode==='any'?'selected':''}>Một người bất kỳ duyệt là đủ</option><option value="all" ${lv.mode==='all'?'selected':''}>Tất cả phải cùng duyệt</option></select></label></div>` : ''}</div></div>`).join('')}</div></div>`; }).join('');
}
function cfgMatrix(D){
  const C = D.config; const rules = [...C.rules].sort((a,b) => (a.def?1:0) - (b.def?1:0));
  return `<p class="small muted" style="margin:0">Mỗi quy định trả lời ba câu theo thứ tự: áp dụng cho ai, mốc tiền nào, mấy cấp duyệt và ai duyệt. Đề xuất khớp quy định theo người trước, rồi theo phòng, cuối cùng là mặc định.</p>
  ${rules.map(rl => { const open = UI.cfgRule === rl.id;
    return `<div class="rule ${open?'open':''}">
      <button type="button" class="rule-h" data-ruleopen="${rl.id}" aria-expanded="${open}">
        <div class="rule-s"><span class="small faint">${rl.def ? 'Áp dụng cho' : rl.people.length && !rl.depts.length ? 'Người đề xuất' : 'Phòng ban'}</span><b>${rl.def ? 'Mọi đề xuất còn lại' : esc(ruleLabel(rl))}</b>${rl.def ? '<span class="small faint">Quy định mặc định, không xoá được</span>' : ''}</div>
        <div class="rule-b">${ruleSummary(rl)}</div><span class="chev">${icon('chev',16)}</span></button>
      ${open ? `<div class="rule-e stack">
        <div><div class="rstep"><i>1</i>Áp dụng cho</div>${rl.def ? '<p class="small muted" style="margin:0">Mọi đề xuất không khớp quy định nào khác theo người hoặc theo phòng.</p>' : scopeChips(rl)}</div>
        <div><div class="rstep"><i>2</i>Mốc tiền</div><div class="stack" style="gap:12px">${bandsEditor(rl)}<div><button class="btn sec sm" data-addband="${rl.id}">${icon('plus',14)} Thêm mốc tiền</button></div></div>
          <p class="small faint" style="margin:10px 0 0"><i class="rstep-i">3</i> Ở mỗi mốc chọn 1, 2 hoặc 3 cấp duyệt và người duyệt từng cấp.</p></div>
        ${rl.def ? '' : `<div><button class="btn danger sm" data-rmrule="${rl.id}">Xoá quy định</button></div>`}
      </div>` : ''}</div>`; }).join('')}
  <div><button class="btn sec" data-action="addrule">${icon('plus',16)} Thêm quy định</button></div>`;
}
function catUse(c){ return S.reqs.filter(r => r.cat === c).length; }
function cfgCats(){
  return `<p class="small muted" style="margin:0">Hạng mục người đề xuất chọn khi tạo phiếu, gom theo nhóm để báo cáo. Kế toán và Quản trị được thêm, sửa tên, xoá. Thay đổi có hiệu lực ngay, không cần lưu.</p>
  ${S.cats.map((g,gi) => `<div class="card"><div class="ch"><div class="row"><h2 class="card-title">${esc(g[0])}</h2><button type="button" class="copy" data-editgroup="${gi}" aria-label="Sửa tên nhóm ${esc(g[0])}" title="Sửa tên nhóm">${icon('edit',14)}</button></div><div class="row"><span class="card-note">${g[1].length} hạng mục</span>${g[1].length ? '' : `<button class="btn ghost sm" data-rmgroup="${gi}" style="color:var(--critical)">Xoá nhóm</button>`}</div></div>
      <div class="chips">${g[1].map((c,ci) => { const n = catUse(c); return `<span class="chip" style="padding-left:10px">${esc(c)}${n ? `<span class="lvn">${n} phiếu</span>` : ''}<button type="button" data-editcat="${gi}.${ci}" aria-label="Sửa ${esc(c)}" title="Sửa tên">${icon('edit',12)}</button><button type="button" data-rmcat="${gi}.${ci}" aria-label="Xoá ${esc(c)}">${icon('x',12)}</button></span>`; }).join('') || '<span class="small faint">Chưa có hạng mục</span>'}</div>
      <form class="row catadd" data-g="${gi}" style="margin-top:12px"><input class="in" name="name" placeholder="Tên hạng mục mới" autocomplete="off" style="flex:1;max-width:320px;height:34px" aria-label="Tên hạng mục mới trong nhóm ${esc(g[0])}"><button class="btn sec sm" type="submit">${icon('plus',14)} Thêm</button></form></div>`).join('')}
  <div class="card"><div class="ch"><h2 class="card-title">Thêm nhóm chi phí</h2></div>
    <form class="row" id="groupAdd"><input class="in" name="name" placeholder="Tên nhóm, VD: Thuế và lệ phí" autocomplete="off" style="flex:1;max-width:320px;height:34px" aria-label="Tên nhóm mới"><button class="btn sec sm" type="submit">${icon('plus',14)} Thêm nhóm</button></form>
    <p class="small faint" style="margin:10px 0 0">“${REFUND_CAT}” là hạng mục hệ thống cho đề xuất hoàn tiền, không tính vào chi phí. Đổi tên hạng mục sẽ đổi tên trên toàn bộ phiếu đã lập; xoá hạng mục không đổi các phiếu cũ.</p></div>`;
}
function cfgRules(D){
  const C = D.config; const sg = signers();
  return `<div class="card"><div class="ch"><h2 class="card-title">Chống tự duyệt</h2></div>
      <label class="toggle"><input type="checkbox" data-rule="blockSelf" ${C.blockSelf?'checked':''}><div><b>Không ai duyệt khoản của chính mình</b><span>Người đề xuất bị bỏ khỏi bước mình đứng tên, kể cả khi được ủy quyền; bước đó không còn ai thì người dự phòng duyệt thay.</span></div></label>
      <div class="sect"><b>Người duyệt dự phòng</b><p class="hint">Dùng khi một bước trống sau khi lọc: phòng chưa có người duyệt, trùng người, tự duyệt.</p>${personChips(C.fallback,'fallback')}</div></div>
    <div class="card"><div class="ch"><h2 class="card-title">Chi không hoá đơn</h2></div>
      <label class="toggle"><input type="checkbox" data-rule="noInvoice" ${C.noInvoice.on?'checked':''}><div><b>Thêm một bước kiểm soát</b><span>Nối sau các bước của khung tiền; bỏ qua nếu người kiểm soát đã có trong chuỗi.</span></div></label>
      ${C.noInvoice.on ? personChips(C.noInvoice.approvers,'noInvoice') : ''}</div>
    <div class="card"><div class="ch"><h2 class="card-title">Tiền mặt và đợt chi</h2></div>
      <label class="f" style="max-width:320px">Hoá đơn VAT từ mức này phải trả không dùng tiền mặt<div class="money"><input class="in" data-rule="cashLimit" inputmode="numeric" value="${fmt(C.cashLimit)}" data-money><span>đ</span></div><span class="hint">Ngưỡng 5 triệu theo Luật Thuế GTGT 2024, kế toán trưởng xác nhận lại.</span></label>
      <div class="sect"><b>Người duyệt đợt chi</b><p class="hint">Ai được chọn làm người duyệt đợt chi là quyền ở trang Người dùng. Chuỗi duyệt đợt cấu hình ở tab Đợt chi.</p>
        <div class="chips">${sg.length ? sg.map(id => `<span class="chip" style="padding-right:10px">${avatar(id,'xs')}${esc(user(id).name)}</span>`).join('') : '<span class="small" style="color:var(--critical)">Chưa có ai</span>'}<button class="btn ghost sm" data-go="users">Sửa ở Người dùng</button></div></div></div>`;
}
function cfgIssues(D){
  const C = D.config; const out = []; const seenD = {}, seenP = {};
  C.rules.forEach(rl => { const nm = `Quy định ${ruleLabel(rl)}`; const bs = sortedBands(rl);
    if (!rl.def && !rl.depts.length && !rl.people.length) out.push({err:true, t:`${nm}: chọn ít nhất một phòng ban hoặc một người đề xuất.`});
    rl.depts.forEach(id => { if (seenD[id]) out.push({err:true, t:`Phòng ${(dept(id)||{}).name} nằm trong hai quy định.`}); seenD[id] = 1; });
    rl.people.forEach(id => { if (seenP[id]) out.push({err:true, t:`${(user(id)||{}).name} nằm trong hai quy định.`}); seenP[id] = 1; });
    bs.forEach((b,i) => {
      if (b.max != null && (!b.max || b.max <= 0)) out.push({err:true, t:`${nm}, mốc ${i+1}: mức trần phải lớn hơn 0.`});
      if (i>0 && b.max != null && bs[i-1].max != null && b.max <= bs[i-1].max) out.push({err:true, t:`${nm}, mốc ${i+1}: mức trần phải lớn hơn mốc trước.`});
      b.levels.forEach((lv,li) => {
        if (!lv.approvers.length) out.push({err:true, t:`${nm}, mốc ${i+1}, cấp ${li+1}: chưa chọn người duyệt.`});
        lv.approvers.forEach(id => { const u = user(id); if (!u || !u.active || u.deleted) out.push({err:true, t:`${nm}, mốc ${i+1}, cấp ${li+1}: ${u?u.name:'người dùng'} đã ngừng hoạt động.`}); });
        if (lv.mode === 'all' && lv.approvers.length > 3) out.push({err:false, t:`${nm}, mốc ${i+1}, cấp ${li+1}: ${lv.approvers.length} người cùng phải duyệt dễ làm phiếu treo lâu.`});
        if (li > 0 && b.levels[li-1].approvers.length === lv.approvers.length && lv.approvers.every(x => b.levels[li-1].approvers.includes(x))) out.push({err:false, t:`${nm}, mốc ${i+1}: cấp ${li} và ${li+1} cùng người, cấp sau sẽ chuyển sang người dự phòng.`});
        if (rl.people.length && !rl.depts.length && rl.people.every(p => lv.approvers.includes(p)) && lv.approvers.length === rl.people.length) out.push({err:false, t:`${nm}, mốc ${i+1}, cấp ${li+1}: người duyệt trùng người đề xuất, cấp này sẽ chuyển sang người dự phòng.`});
      });
    }); });
  const bB = C.batch; if (bB) { const bs = sortedBands(bB);
    bs.forEach((b,i) => {
      if (b.max != null && (!b.max || b.max <= 0)) out.push({err:true, t:`Đợt chi, mốc ${i+1}: mức trần phải lớn hơn 0.`});
      if (i>0 && b.max != null && bs[i-1].max != null && b.max <= bs[i-1].max) out.push({err:true, t:`Đợt chi, mốc ${i+1}: mức trần phải lớn hơn mốc trước.`});
      b.levels.forEach((lv,li) => {
        if (!lv.approvers.length) out.push({err:true, t:`Đợt chi, mốc ${i+1}, cấp ${li+1}: chưa chọn người duyệt.`});
        lv.approvers.forEach(id => { const u = user(id); if (!u || !u.active || u.deleted || !u.signer) out.push({err:true, t:`Đợt chi, mốc ${i+1}, cấp ${li+1}: ${u ? u.name : 'người dùng'} không còn quyền duyệt đợt chi.`}); });
      }); });
    if (bs.length && bs[bs.length-1].max != null) out.push({err:false, t:'Mốc cuối của đợt chi nên để không giới hạn, để mọi đợt đều có chuỗi duyệt.'});
  }
  if (C.noInvoice.on && !C.noInvoice.approvers.length) out.push({err:true, t:'Chi không hoá đơn: chưa chọn người kiểm soát.'});
  if (!signers().length) out.push({err:true, t:'Chưa có ai được quyền duyệt đợt chi (trang Người dùng).'});
  if (!C.fallback.length) out.push({err:false, t:'Chưa có người duyệt dự phòng.'});
  return out;
}
function updateSim(){
  const box = $('#sim'); if (!box) return; const D = UI.draft; const sm = UI.sim;
  const res = buildChain({amount:sm.amount, dept:sm.dept, requester:sm.requester, evidence:sm.evidence}, D.config);
  box.innerHTML = `<div class="ch"><h2 class="card-title">Thử chuỗi duyệt</h2><span class="card-note">trên cấu hình đang sửa</span></div>
    <div class="fgrid">
      <label class="f">Người đề xuất<select class="in" data-sim="requester">${activeUsers().map(u => `<option value="${u.id}" ${u.id===sm.requester?'selected':''}>${esc(u.name)}</option>`).join('')}</select></label>
      <label class="f">Phòng chịu chi phí<select class="in" data-sim="dept">${S.depts.map(d => `<option value="${d.id}" ${d.id===sm.dept?'selected':''}>${esc(d.name)}</option>`).join('')}</select></label>
      <label class="f">Số tiền<div class="money"><input class="in" data-sim="amount" inputmode="numeric" value="${fmt(sm.amount)}" data-money><span>đ</span></div></label>
      <label class="f">Chứng từ<select class="in" data-sim="evidence">${Object.entries(EVIDENCE).map(([k,v]) => `<option value="${k}" ${k===sm.evidence?'selected':''}>${v[0]}</option>`).join('')}</select></label>
    </div>
    <p class="small faint" style="margin:16px 0 12px">Khớp quy định <b style="color:var(--ink);font-weight:500">${esc(ruleLabel(res.rule))}</b>, mốc ${esc(bandRange(res.band, res.rule))}, ${res.band.levels.length} cấp</p>
    ${chainHtml(res.steps.map(s => Object.assign({}, s, {status:'waiting', decisions:{}})), true)}
    ${res.problems.length ? `<div class="alert neg" style="margin-top:14px">${res.problems.map(esc).join('<br>')}</div>` : ''}`;
}
function cfgPath(path){ const C = UI.draft.config;
  if (path === 'fallback') return C.fallback; if (path === 'noInvoice') return C.noInvoice.approvers;
  const [bid, li] = path.split('.'); return allBands(C).find(b => b.id === bid).levels[+li]; }
function listAt(path){ const x = cfgPath(path); return Array.isArray(x) ? x : x.approvers; }
function refreshCfgSide(){ updateSim(); const D = UI.draft; $('#issues').innerHTML = issuesHtml(D); $('#savebarHost').innerHTML = savebarHtml(D); }

/* ===================== Cấu hình: đợt chi, đổi tên hạng mục, danh bạ NCC ===================== */
IC.edit = '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>';
function cfgBatch(D){
  const B = D.config.batch; const pool = signers();
  return `<p class="small muted" style="margin:0">Đợt chi do kế toán lập đi qua chuỗi duyệt riêng, chọn theo tổng tiền của đợt. Chỉ chọn được người có quyền “Duyệt đợt chi” ở trang Người dùng. Người lập đợt không tự duyệt đợt của mình.</p>
  <div class="rule open"><div class="rule-e stack" style="border-top:0">
    <div><div class="rstep"><i>1</i>Mốc theo tổng tiền đợt</div><div class="stack" style="gap:12px">${bandsEditor(B, pool)}<div><button class="btn sec sm" data-addband="batch">${icon('plus',14)} Thêm mốc tiền</button></div></div></div>
    <p class="small faint" style="margin:0">Ở mỗi mốc chọn 1, 2 hoặc 3 cấp duyệt và người duyệt từng cấp. Một người chỉ ký một cấp trong cùng đợt.</p>
  </div></div>`;
}
function renameModal(kind, gi, ci){
  const cur = kind === 'cat' ? S.cats[gi][1][ci] : S.cats[gi][0];
  UI.renameKind = {kind, gi, ci};
  modal(`<form class="modal" id="renameForm" role="dialog" aria-modal="true" aria-labelledby="rnh" novalidate>
    <div class="modal-h"><h3 id="rnh">${kind === 'cat' ? 'Sửa hạng mục' : 'Sửa tên nhóm'}</h3><p>${kind === 'cat' ? 'Tên mới áp dụng cho cả các phiếu đã lập có hạng mục này.' : 'Tên nhóm dùng để gom báo cáo và lọc.'}</p></div>
    <div class="modal-b"><label class="f">Tên mới<input class="in" name="name" value="${esc(cur)}" autocomplete="off"></label><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Lưu</button></div></form>`, 460);
}
function formPrefill(sel){
  const f = $(sel); if (!f) return {}; const out = {};
  new FormData(f).forEach((v,k) => { if (k !== 'files') out[k] = String(v); });
  return out;
}
function vendorStats(v){
  const rs = S.reqs.filter(r => r.payee && r.payee.type === 'vendor' && r.payee.vendorId === v.id);
  const paid = rs.filter(r => r.status === 'paid');
  return {n: rs.length, paidN: paid.length, paidAmt: paid.reduce((a,r) => a + r.amount, 0), last: paid.reduce((m,r) => Math.max(m, r.paidAt || 0), 0) || null};
}
function vendorsDirHtml(){
  const q = noAccent((UI.vdq || '').trim()).toLowerCase();
  const list = S.vendors.filter(v => !v.deleted && (!q || noAccent([v.name, v.tax, v.acc, v.accName, v.bank].join(' ')).toLowerCase().includes(q)));
  const unv = S.vendors.filter(v => !v.deleted && !v.verified).length;
  return `<div class="card" style="padding:16px 18px">
    <div class="row" style="margin-bottom:12px"><input class="in" type="search" data-vdq value="${esc(UI.vdq || '')}" placeholder="Tìm theo tên, MST, số tài khoản" autocomplete="off" aria-label="Tìm NCC" style="flex:1 1 260px;min-width:220px;max-width:380px">
      <span class="small faint">${list.length} NCC${unv ? `, ${unv} STK chưa xác minh` : ''}</span><span class="sp"></span><button class="btn ghost sm" data-vtemplate="1">${icon('download',14)} Tải file mẫu</button><button class="btn sec sm" data-vexport="1">${icon('download',14)} Xuất danh bạ</button><label class="btn sec sm" style="position:relative">${icon('upload',14)} Nhập từ file<input type="file" accept=".xlsx,.xls,.csv" data-vimport style="position:absolute;opacity:0;width:1px;height:1px"></label><button class="btn pri sm" data-vnew="1">${icon('plus',14)} Thêm NCC</button></div>
    <div class="tw"><table class="vdir"><thead><tr><th>Tên NCC / người nhận</th><th>MST</th><th>Ngân hàng</th><th>Số tài khoản</th><th>Trạng thái</th><th class="r">Số phiếu</th><th class="r">Đã chi (đ)</th><th></th></tr></thead><tbody>
      ${list.map(v => { const st = vendorStats(v); return `<tr class="click" data-vview="${v.id}" tabindex="0"><td><div class="ttl">${esc(v.name)}</div><div class="small faint">${esc(v.accName || '')}</div></td><td class="small num">${esc(v.tax || '—')}</td><td class="small">${esc(v.bank)}</td><td class="small num">${esc(v.acc)}</td><td>${v.verified ? '<span class="pill st-approved">Đã xác minh</span>' : '<span class="pill st-returned">Chưa xác minh</span>'}</td><td class="r">${st.n}</td><td class="r">${fmt(st.paidAmt)}</td><td class="act">${!v.verified && me().acct ? `<button class="btn ghost sm" data-verify="${v.id}">Xác minh</button>` : ''}<button class="btn ghost sm" data-vedit="${v.id}">Sửa</button><button class="btn ghost sm" data-vdel="${v.id}" style="color:var(--critical)">Xoá</button></td></tr>`; }).join('') || '<tr><td colspan="8" class="small faint" style="text-align:center;padding:20px">Không có NCC nào khớp.</td></tr>'}
    </tbody></table></div>
    <p class="small faint" style="margin:10px 0 0">NCC thêm ở đề xuất chi hoặc khoản chi trực tiếp đều lưu ở đây, và ngược lại. Đổi số tài khoản sẽ đưa NCC về trạng thái chưa xác minh.</p></div>`;
}
function vendorDetailModal(id){
  const v = vendor(id); if (!v) return; const st = vendorStats(v);
  const rs = S.reqs.filter(r => r.payee && r.payee.vendorId === id).sort((a,b) => b.createdAt - a.createdAt).slice(0,10);
  modal(`<div class="modal" role="dialog" aria-modal="true" aria-labelledby="vdh" style="width:min(640px,100%)">
    <div class="modal-h"><h3 id="vdh">${esc(v.name)}</h3><p>${v.verified ? 'STK đã xác minh' : 'STK chưa xác minh'}${v.tax ? ` · MST ${esc(v.tax)}` : ''}</p></div>
    <div class="modal-b stack">
      <div class="kv"><div><span>Ngân hàng</span><b>${esc(v.bank)}</b></div><div><span>Số tài khoản</span><b class="num">${esc(v.acc)}</b></div><div><span>Tên tài khoản</span><b>${esc(v.accName || '—')}</b></div></div>
      <div class="kv"><div><span>Số phiếu</span><b>${st.n}</b></div><div><span>Đã chi</span><b>${money(st.paidAmt)}</b></div><div><span>Chi gần nhất</span><b>${st.last ? dOnly(st.last) : '—'}</b></div></div>
      ${rs.length ? `<div class="tw"><table><thead><tr><th>Mã</th><th>Nội dung</th><th class="r">Số tiền (đ)</th><th>Trạng thái</th></tr></thead><tbody>${rs.map(r => `<tr><td class="code">${esc(r.code || 'Nháp')}</td><td class="small">${esc(r.title)}</td><td class="r">${fmt(r.amount)}</td><td>${pill(r.status)}</td></tr>`).join('')}</tbody></table></div>` : '<p class="small faint" style="margin:0">Chưa có phiếu nào gửi tới NCC này.</p>'}
    </div>
    <div class="modal-f"><button class="btn sec" data-vedit="${v.id}">Sửa thông tin</button><button class="btn pri" data-close="1">Đóng</button></div></div>`);
}
function vendorDelete(id){
  const v = vendor(id); if (!v) return;
  const open = S.reqs.filter(r => r.payee && r.payee.vendorId === id && ['draft','pending','returned','approved','in_batch'].includes(r.status));
  if (open.length) { toast(`Chưa xoá được ${v.name}: còn ${open.length} phiếu đang xử lý (${open.map(r => r.code || 'nháp').join(', ')}).`); return; }
  askConfirm(`Xoá ${v.name} khỏi danh bạ?`, 'NCC sẽ không còn xuất hiện khi chọn người nhận. Phiếu và đợt chi cũ vẫn giữ tên NCC.', 'Xoá NCC', () => { v.deleted = true; save(); toast(`Đã xoá ${v.name}`); renderMain(); });
}
document.addEventListener('click', e => {
  if (e.target.closest('button, a, input, select, textarea, label')) return;
  const tr = e.target.closest('tr[data-vview]'); if (tr) vendorDetailModal(tr.dataset.vview);
});

/* ===================== Tạm ứng: quyết toán ===================== */
IC.download = '<path d="M12 4v12M7 11l5 5 5-5M4 20h16"/>';
const ADV_LABEL = {waiting:'Chờ quyết toán', overdue:'Quá hạn', submitted:'Đang duyệt quyết toán', refund_waiting:'Chờ nhận hoàn ứng', done:'Đã đóng'};
function advState(r){
  if (r.evidence !== 'advance' || r.status !== 'paid') return null;
  const s = r.settle; const due = r.settleDue && now() > r.settleDue ? 'overdue' : 'waiting';
  if (s && s.reqId) { const q = req(s.reqId); if (q && q.status === 'paid') return 'done'; if (q && ['rejected','cancelled'].includes(q.status)) return due; return 'submitted'; }
  if (s && s.refund) return s.refund.status === 'received' ? 'done' : 'refund_waiting';
  if (s && s.done) return 'done';
  return due;
}
const settleSpent = () => (UI.settleRows || []).reduce((a,x) => a + digits(x.total), 0);
function settleSummary(r, spent){
  const d = spent - r.amount;
  if (!spent) return 'Nhập hoá đơn để tính chênh lệch với số đã tạm ứng.';
  if (d > 0) return `Chi thực tế ${money(spent)}, chi thêm ${money(d)} so với tạm ứng. Phiếu chênh lệch đi qua duyệt và chi như đề xuất thường.`;
  if (d < 0) return `Chi thực tế ${money(spent)}, bạn cần hoàn lại ${money(-d)} cho công ty. Kế toán xác nhận khi nhận được tiền.`;
  return `Chi thực tế ${money(spent)}, khớp đúng tạm ứng. Không phát sinh chênh lệch.`;
}
function settleModal(id){ UI.settleFor = id; UI.settleRows = [{inv:'', desc:'', total:'', tax:''}]; UI.settleFiles = []; renderSettle(); }
function renderSettle(){
  const r = req(UI.settleFor); if (!r) return; const rows = UI.settleRows;
  modal(`<form class="modal" id="settleForm" role="dialog" aria-modal="true" aria-labelledby="sth" novalidate>
    <div class="modal-h"><h3 id="sth">Quyết toán tạm ứng</h3><p>${esc(r.code)} · ${esc(r.title)} · đã tạm ứng ${money(r.amount)}</p></div>
    <div class="modal-b stack">
      <div class="tw"><table class="lines"><thead><tr><th style="width:130px">Số hoá đơn</th><th>Diễn giải</th><th class="r" style="width:130px">Thành tiền</th><th class="r" style="width:120px">Tiền thuế</th><th style="width:34px"></th></tr></thead><tbody>
        ${rows.map((x,i) => `<tr><td><input class="in cell" data-sr="${i}" data-sk="inv" value="${esc(x.inv)}" aria-label="Số hoá đơn dòng ${i+1}" autocomplete="off"></td><td><input class="in cell" data-sr="${i}" data-sk="desc" value="${esc(x.desc)}" aria-label="Diễn giải dòng ${i+1}" autocomplete="off"></td><td><input class="in cell r" data-sr="${i}" data-sk="total" data-money inputmode="numeric" value="${x.total ? fmt(digits(x.total)) : ''}" aria-label="Thành tiền dòng ${i+1}" autocomplete="off"></td><td><input class="in cell r" data-sr="${i}" data-sk="tax" data-money inputmode="numeric" value="${x.tax ? fmt(digits(x.tax)) : ''}" aria-label="Tiền thuế dòng ${i+1}" autocomplete="off"></td><td>${rows.length > 1 ? `<button type="button" class="copy" data-srowrm="${i}" aria-label="Xoá dòng ${i+1}">${icon('x')}</button>` : ''}</td></tr>`).join('')}
      </tbody></table></div>
      <div class="row"><button type="button" class="btn ghost sm" data-srowadd="1">${icon('plus',14)} Thêm dòng</button><label class="btn sec sm" style="position:relative">${icon('clip',14)} Đính kèm hoá đơn<input type="file" multiple accept=".pdf,.xml,image/*" data-sfiles style="position:absolute;opacity:0;width:1px;height:1px"></label></div>
      ${UI.settleFiles.length ? `<div class="chips">${UI.settleFiles.map(n => `<span class="chip file">${icon('file',13)} ${esc(fname(n))}</span>`).join('')}</div>` : ''}
      <div class="alert info" id="setSum">${settleSummary(r, settleSpent())}</div>
      <p class="small faint" style="margin:0">Tiền thuế không bắt buộc. Nếu chi thêm, phiếu chênh lệch sẽ đi qua chuỗi duyệt giống đề xuất thường.</p>
      <div id="mErr"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Gửi quyết toán</button></div></form>`, 680);
}
function submitSettle(adv, rows, files, by){
  const spent = rows.reduce((a,x) => a + digits(x.total), 0);
  const lines = rows.map(x => { const total = digits(x.total), tax = Math.min(digits(x.tax), total), pre = total - tax; const rk = tax && pre ? String(Math.round(tax / pre * 100)) : 'kct';
    return {inv:String(x.inv).trim(), desc:String(x.desc || '').trim() || adv.title, pre, rate: RATES[rk] ? rk : 'kct', tax, total}; });
  const fl = files.map(n => ({name:n, kind:'vat'}));
  const diff = spent - adv.amount;
  adv.settleLines = lines; adv.settleFiles = files.slice();
  log(adv, by, `Gửi quyết toán tạm ứng, chi thực tế ${fmt(spent)} đ`);
  if (diff > 0) {
    const r = newReq({title:`Quyết toán tạm ứng ${adv.code}: ${adv.title}`.slice(0,140), amount:diff, dept:adv.dept, cat:adv.cat, evidence:'vat', payMethod:'transfer', payee:{type:'me', outlay:'transfer'}, files:fl, lines, requester:by, note:`Quyết toán ${adv.code}. Tạm ứng ${fmt(adv.amount)} đ, chi thực tế ${fmt(spent)} đ.`});
    r.code = nextCode('A'); r.settleOf = adv.id; S.reqs.push(r); submitReq(r, by);
    adv.settle = {reqId:r.id, at:now()}; log(adv, by, `Đã gửi phiếu ${r.code} chi thêm ${fmt(diff)} đ`);
    return `Đã gửi phiếu ${r.code}, chi thêm ${fmt(diff)} đ`;
  }
  if (diff < 0) { adv.settle = {refund:{amount:-diff, status:'waiting'}, at:now()}; log(adv, by, `Hoàn lại ${fmt(-diff)} đ cho công ty, chờ kế toán xác nhận`); return `Đã gửi quyết toán, chờ hoàn ${fmt(-diff)} đ`; }
  adv.settle = {done:true, at:now()}; log(adv, by, 'Quyết toán khớp tạm ứng, đóng khoản'); return 'Đã quyết toán, khớp tạm ứng';
}
function refundReceived(r){ r.settle.refund.status = 'received'; r.settle.refund.receivedAt = now(); log(r, S.me, `Đã nhận hoàn ứng ${fmt(r.settle.refund.amount)} đ`); }

/* ===================== Danh bạ NCC: file mẫu, xuất, nhập ===================== */
const VCOLS = ['Mã NCC (để trống nếu thêm mới)','Tên NCC / người nhận','Mã số thuế','Ngân hàng','Số tài khoản','Tên tài khoản (IN HOA, không dấu)','Ghi chú'];
function vendorTemplate(){
  exportXlsx('Mau_nhap_danh_ba_NCC.xlsx', [
    {name:'Mau NCC', rows:[VCOLS, ['', 'VÍ DỤ, XOÁ DÒNG NÀY: Công ty TNHH Ví Dụ', '0101234567', 'Vietcombank', '0011223344', 'CONG TY TNHH VI DU', 'Ghi chú tuỳ chọn']]},
    {name:'Huong dan', rows:[['Hướng dẫn nhập danh bạ NCC / người nhận'], [''],
      ['1. Điền dữ liệu từ dòng 2 trở đi trong sheet “Mau NCC”. Xoá dòng ví dụ trước khi nhập.'],
      ['2. Mã NCC: để trống để thêm mới. Điền mã có sẵn (lấy từ file xuất) để cập nhật đúng NCC đó.'],
      ['3. Không có mã: hệ thống khớp theo Ngân hàng + Số tài khoản. Khớp thì cập nhật, không khớp thì thêm mới.'],
      ['4. Số tài khoản: số và chữ, không dấu cách. Nên định dạng cột là Văn bản để không mất số 0 đầu.'],
      ['5. Tên tài khoản: IN HOA không dấu, đúng như trên ngân hàng.'],
      ['6. NCC nhập từ file luôn ở trạng thái chưa xác minh. Kế toán xác minh trước khi đưa vào đợt chi.'],
      ['7. Dòng lỗi được báo trước khi nhập. Các dòng hợp lệ vẫn nhập được.']]}
  ]);
}
function vendorExport(){
  const list = S.vendors.filter(v => !v.deleted);
  exportXlsx(`Danh_ba_NCC_${localDate()}.xlsx`, [{name:'Danh ba NCC', rows:[VCOLS.slice(0,6).concat(['Đã xác minh','Số phiếu','Đã chi (đ)'])].concat(list.map(v => { const st = vendorStats(v); return [v.id, v.name, v.tax || '', v.bank, v.acc, v.accName || '', v.verified ? 'Có' : 'Không', st.n, st.paidAmt]; }))}]);
}
function vendorRowCheck(v, selfId){
  if (v.name.length < 3) return 'Tên NCC thiếu hoặc quá ngắn';
  if (!v.bank) return 'Thiếu ngân hàng';
  if (!v.acc) return 'Thiếu số tài khoản';
  if (!/^[0-9A-Z]{4,24}$/.test(v.acc)) return 'Số tài khoản không hợp lệ (4 đến 24 ký tự số hoặc chữ)';
  if (!v.accName) return 'Thiếu tên tài khoản';
  if (v.tax && !/^\d{10}(-?\d{3})?$/.test(v.tax)) return 'Mã số thuế không hợp lệ';
  if (S.vendors.some(x => !x.deleted && x.id !== selfId && x.acc === v.acc && x.bank.toLowerCase() === v.bank.toLowerCase())) return 'Số tài khoản đã có trong danh bạ';
  return null;
}
function parseVendorRows(wb){
  const sn = wb.SheetNames.find(n => /ncc/i.test(n)) || wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, defval:'', raw:false});
  const hi = rows.findIndex(r => r.some(c => /^ten ncc/.test(noAccent(String(c)).toLowerCase().trim())));
  if (hi < 0) return {error:'Không thấy dòng tiêu đề có cột “Tên NCC”. Hãy dùng file mẫu của hệ thống.'};
  const head = rows[hi].map(c => noAccent(String(c)).toLowerCase().trim());
  const col = re => head.findIndex(h => re.test(h));
  const C = {id:col(/^ma ncc/), name:col(/^ten ncc/), tax:col(/ma so thue/), bank:col(/ngan hang/), acc:col(/so tai khoan/), accName:col(/ten tai khoan/)};
  if (['name','bank','acc','accName'].some(k => C[k] < 0)) return {error:'File thiếu cột bắt buộc: Tên NCC, Ngân hàng, Số tài khoản, Tên tài khoản.'};
  const out = []; const seen = new Set();
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i]; const get = k => C[k] >= 0 ? String(r[C[k]] ?? '').trim() : '';
    const name = get('name'); if (!name && !get('acc')) continue;
    const line = i + 1;
    if (/^vi du/i.test(noAccent(name))) { out.push({line, name, action:'skip', msg:'dòng ví dụ'}); continue; }
    const v = {name, tax:get('tax').replace(/\s/g,''), bank:get('bank'), acc:get('acc').replace(/\s/g,'').toUpperCase(), accName:bankText(get('accName'))};
    const id = get('id');
    const existing = id ? S.vendors.find(x => x.id === id && !x.deleted) : S.vendors.find(x => !x.deleted && x.acc === v.acc && x.bank.toLowerCase() === v.bank.toLowerCase());
    if (id && !existing) { out.push({line, name, action:'error', msg:`Mã NCC ${id} không có trong danh bạ`}); continue; }
    const err = vendorRowCheck(v, existing ? existing.id : null);
    if (err) { out.push({line, name, action:'error', msg:err}); continue; }
    const key = v.bank.toLowerCase() + '|' + v.acc;
    if (seen.has(key)) { out.push({line, name, action:'error', msg:'Trùng số tài khoản với dòng khác trong file'}); continue; }
    seen.add(key); out.push({line, name, action: existing ? 'update' : 'new', id: existing ? existing.id : null, v});
  }
  if (!out.length) return {error:'Không có dòng dữ liệu nào trong file.'};
  return {rows: out};
}
function vimportModal(fname){
  const p = UI.vimp; const good = p.rows.filter(x => x.action === 'new' || x.action === 'update');
  const cnt = a => p.rows.filter(x => x.action === a).length; const nErr = cnt('error');
  modal(`<div class="modal" role="dialog" aria-modal="true" aria-labelledby="vih" style="width:min(680px,100%)">
    <div class="modal-h"><h3 id="vih">Nhập danh bạ từ ${esc(fname)}</h3><p>Kiểm tra trước khi nhập. Dòng lỗi bị bỏ qua, các dòng hợp lệ vẫn được nhập.</p></div>
    <div class="modal-b stack">
      <div class="kv"><div><span>Thêm mới</span><b>${cnt('new')}</b></div><div><span>Cập nhật</span><b>${cnt('update')}</b></div><div><span>Lỗi, bỏ qua</span><b style="color:${nErr ? 'var(--critical)' : 'inherit'}">${nErr}</b></div></div>
      ${cnt('skip') ? `<p class="small faint" style="margin:0">${cnt('skip')} dòng ví dụ được bỏ qua.</p>` : ''}
      ${nErr ? `<div class="tw" style="max-height:260px;overflow:auto"><table><thead><tr><th class="r">Dòng</th><th>Tên</th><th>Lỗi</th></tr></thead><tbody>${p.rows.filter(x => x.action === 'error').slice(0,100).map(x => `<tr><td class="r faint">${x.line}</td><td class="small">${esc(x.name)}</td><td class="small" style="color:var(--critical)">${esc(x.msg)}</td></tr>`).join('')}</tbody></table></div>` : ''}
      ${cnt('update') ? `<p class="small muted" style="margin:0">Sẽ cập nhật: ${p.rows.filter(x => x.action === 'update').slice(0,5).map(x => esc(x.name)).join(', ')}${cnt('update') > 5 ? '…' : ''}. Đổi số tài khoản sẽ đưa NCC về chưa xác minh.</p>` : ''}
    </div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="button" class="btn pri" data-vimportgo="1" ${good.length ? '' : 'disabled'}>Nhập ${good.length} dòng</button></div></div>`);
}
function applyVendorImport(){
  const p = UI.vimp; if (!p) return; let nN = 0, nU = 0;
  p.rows.forEach(x => {
    if (x.action === 'new') { S.vendors.push(Object.assign({id:uid('v'), verified:false, createdBy:S.me, deleted:false}, x.v)); nN++; }
    else if (x.action === 'update') { const old = vendor(x.id); const changed = old.bank !== x.v.bank || old.acc !== x.v.acc || old.accName !== x.v.accName; Object.assign(old, x.v); if (changed) { old.verified = false; old.verifiedBy = null; } nU++; }
  });
  UI.vimp = null; save(); closeModal(); toast(`Đã thêm ${nN} NCC, cập nhật ${nU} NCC. STK mới chờ kế toán xác minh.`); renderMain();
}
document.addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
  if (d.srowadd) { UI.settleRows.push({inv:'', desc:'', total:'', tax:''}); renderSettle(); return; }
  if (d.srowrm != null) { UI.settleRows.splice(+d.srowrm, 1); renderSettle(); return; }
  if (d.settle) { settleModal(d.settle); return; }
  if (d.refundsettle) { refundReceived(req(d.refundsettle)); save(); toast('Đã ghi nhận nhận hoàn ứng'); render(); return; }
  if (d.vimportgo) { applyVendorImport(); return; }
  if (d.vtemplate) { vendorTemplate(); return; }
  if (d.vexport) { vendorExport(); return; }
  if (d.editsettle) { dueModal(d.editsettle); return; }
});
document.addEventListener('change', async e => {
  const t = e.target; const d = t.dataset;
  if (d.sfiles != null && t.files) { [...t.files].forEach(f => { const k = addBlob(f); if (k) UI.settleFiles.push(k); }); t.value = ''; renderSettle(); return; }
  if (d.vimport != null && t.files && t.files[0]) { const f = t.files[0]; t.value = '';
    if (typeof XLSX === 'undefined') { toast('Chưa tải được thư viện Excel, thử lại sau vài giây.'); return; }
    try { const wb = XLSX.read(await f.arrayBuffer(), {type:'array'}); const p = parseVendorRows(wb);
      if (p.error) { toast(p.error); return; }
      UI.vimp = p; vimportModal(f.name); }
    catch (err) { toast('Không đọc được file. Dùng file .xlsx hoặc .csv theo mẫu.'); }
  }
});
document.addEventListener('input', e => {
  const t = e.target;
  if (t.dataset.sr != null && UI.settleRows) { const x = UI.settleRows[+t.dataset.sr]; if (!x) return; x[t.dataset.sk] = t.value;
    const box = $('#setSum'); if (box) box.innerHTML = settleSummary(req(UI.settleFor), settleSpent()); }
});

/* ===================== Xác nhận ===================== */
let confirmFn = null;
function askConfirm(title, body, yes, fn){
  confirmFn = fn;
  $('#layer').innerHTML = `<div class="modal-bg" data-close="1"><div class="modal" role="alertdialog" aria-modal="true" aria-labelledby="ch" style="width:min(440px,100%)">
    <div class="modal-h"><h3 id="ch">${esc(title)}</h3><p>${esc(body)}</p></div><div style="height:16px"></div>
    <div class="modal-f"><button class="btn sec" data-close="1">Giữ nguyên</button><button class="btn danger" data-action="confirmyes">${esc(yes)}</button></div></div></div>`;
  $('#layer .btn.sec').focus();
}

/* ===================== Sự kiện ===================== */
function markGuide(r){ const ch = r.rounds[r.rounds.length-1].chain; ch.forEach(x => Object.entries(x.decisions).forEach(([p,d]) => { if (d.by === S.me) { if (p !== S.me) S.guideDone.delegated = true; else S.guideDone.approve = true; } })); }
function copyText(txt, msg){ try { navigator.clipboard.writeText(txt).then(() => toast(msg), () => toast('Trình duyệt chặn sao chép')); } catch(err) { toast('Trình duyệt chặn sao chép'); } }

document.addEventListener('click', e => {
  const t = e.target.closest('button, [data-open], [data-batch], [data-close], .li[data-go]'); if (!t) return;
  const d = t.dataset;
  if (d.close != null && !t.matches('button')) { if (e.target === t) closeModal(); return; }
  if (d.close != null) { closeModal(); return; }
  if (d.go) { if (d.tab) UI.listTab = d.tab; go(d.go, d.id); return; }
  if (d.open) { go('detail', d.open); return; }
  if (d.batch) { go('batch', d.batch); return; }
  if (d.guide) { S.me = d.guide; save(); toast(`Đang dùng vai ${user(d.guide).name}`); go(d.gv); return; }
  if (d.listtab) { UI.listTab = d.listtab; renderMain(); return; }
  if (d.cfgtab) { UI.cfgTab = d.cfgtab; renderMain(); return; }
  if (d.quick) { const r = req(d.quick); act(r, S.me, 'approve', '', null, actorsFor(r, S.me)[0]); markGuide(r); save(); toast(`Đã duyệt ${r.code}${r.status==='approved'?', đủ chữ ký':''}`); delete UI.inboxSel[r.id]; render(); return; }
  if (d.decide) {
    const r = req(UI.id); const note = ($('#decNote')?.value || '').trim();
    if (d.decide !== 'approve' && note.length < 5) { UI.errors = {note: d.decide==='reject' ? 'Ghi lý do từ chối, tối thiểu 5 ký tự.' : 'Ghi rõ cần bổ sung gì, tối thiểu 5 ký tự.'}; renderMain(); $('#decNote').focus(); return; }
    act(r, S.me, d.decide, note, null, $('#decAs')?.value || S.me); markGuide(r); save();
    toast(d.decide==='approve' ? `Đã duyệt ${r.code}${r.status==='approved'?', đủ chữ ký':''}` : d.decide==='reject' ? `Đã từ chối ${r.code}` : `Đã trả lại ${r.code} cho ${user(r.requester).name}`);
    const next = inboxFor(S.me)[0]; if (next && d.decide==='approve') go('detail', next.id); else { UI.errors = {}; render(); }
    return;
  }
  if (d.rmfile != null) { readForm(); UI.form.files.splice(+d.rmfile,1); renderMain(); return; }
  if (d.rmunc != null) { const P = UI.payDraft; const k = +d.rmunc; P.files.splice(k,1);
    Object.keys(P.map).forEach(id => { if (P.map[id] === k) { delete P.map[id]; delete P.why[id]; } else if (P.map[id] > k) P.map[id]--; });
    autoMapUnc(); drawPayModal(); return; }
  if (d.umode) { const P = UI.payDraft; P.date = $('#payDate')?.value || P.date; P.mode = d.umode; autoMapUnc(); drawPayModal(); return; }
  if (d.uncorder) { const P = UI.payDraft; P.date = $('#payDate')?.value || P.date; const used = new Set(Object.values(P.map)); let fi = 0;
    P.ids.forEach(id => { if (P.map[id] != null) return; while (fi < P.files.length && used.has(fi)) fi++; if (fi < P.files.length) { P.map[id] = fi; P.why[id] = ''; used.add(fi); fi++; } }); drawPayModal(); return; }
  if (d.uncdemo) { const P = UI.payDraft; P.date = $('#payDate')?.value || P.date;
    const names = P.ids.map((id,i) => { const r = req(id), p = payeeOf(r); return [uncName(batch(P.bid), id), `MSB_CK_${String(p.acc).replace(/\D/g,'')}.pdf`, `UNC_${r.amount}.pdf`, `Uy_nhiem_chi_${noAccent(p.name).split(' ').slice(-2).join('_')}.pdf`][i % 4]; });
    names.push('Sao_ke_MSB_12-11.pdf');
    addUncFiles(names.reverse()); drawPayModal(); return; }
  if (d.bapprove) { const b = batch(d.bapprove); approveBatch(b, S.me, ($('#bNote')?.value||'').trim(), null, $('#bAs')?.value || S.me); S.guideDone.sign = true; save(); toast(b.status === 'approved' ? `Đã duyệt ${b.code}, kế toán có thể chuyển tiền` : `Đã duyệt bước của ${b.code}, chờ người duyệt tiếp theo`); render(); return; }
  if (d.breturn) { const b = batch(d.breturn); const note = ($('#bNote')?.value||'').trim();
    if (note.length < 5) { UI.errors = {bnote:'Ghi lý do trả lại, tối thiểu 5 ký tự.'}; renderMain(); $('#bNote').focus(); return; }
    returnBatch(b, S.me, note, $('#bAs')?.value || S.me); save(); toast(`Đã trả lại ${b.code} cho kế toán`); render(); return; }
  if (d.payone) { payModal(UI.id, [d.payone]); return; }
  if (d.paysel) { const b = batch(d.paysel); payModal(b.id, b.reqIds.filter(id => !b.lines[id].paidAt && UI.paySel[id] && pmOf(req(id)) === 'transfer')); return; }
  if (d.payall) { const b = batch(d.payall); payModal(b.id, b.reqIds.filter(id => !b.lines[id].paidAt)); return; }
  if (d.copy != null) { copyText(d.copy, 'Đã sao chép'); return; }
  if (d.copysheet) { const b = batch(d.copysheet); const ck = transferIds(b);
    copyText([BANK_HEAD.join('\t')].concat(ck.map((id,i) => bankRow(b, id, i).join('\t'))).join('\n'), `Đã sao chép ${ck.length} dòng chuyển khoản, dán được vào Excel`); return; }
  if (d.nlv) { const {rule, band:b} = findBand(d.nlv); const n = +d.n;
    while (b.levels.length < n) { const i = b.levels.length; const tpl = rule.bands.map(x => x.levels[i]).find(Boolean); b.levels.push(tpl ? clone(tpl) : {approvers:[], mode:'any', sameDept:false}); }
    b.levels.length = n; renderMain(); return; }
  if (d.rmband) { const {rule} = findBand(d.rmband); rule.bands = rule.bands.filter(b => b.id !== d.rmband); renderMain(); return; }
  if (d.ruleopen) { UI.cfgRule = UI.cfgRule === d.ruleopen ? null : d.ruleopen; renderMain(); return; }
  if (d.rmrule) { const C = UI.draft.config; C.rules = C.rules.filter(r => r.id !== d.rmrule); UI.cfgRule = null; renderMain(); return; }
  if (d.rmscope) { const rl = UI.draft.config.rules.find(r => r.id === d.rmscope); rl[d.k] = rl[d.k].filter(x => x !== d.v); renderMain(); return; }
  if (d.addband) { const rl = d.addband === 'batch' ? UI.draft.config.batch : UI.draft.config.rules.find(r => r.id === d.addband); const fin = rl.bands.filter(b => b.max != null).map(b => b.max); const top = fin.length ? Math.max(...fin) : 0;
    const lastB = rl.bands.find(b => b.max == null); rl.bands.push({id:uid('b'), max: top ? top*2 : 10000000, levels: clone(lastB ? lastB.levels : [{approvers:[],mode:'any'}])}); renderMain(); toast('Đã thêm mốc tiền, chỉnh mức trần và người duyệt'); return; }
  if (d.editcat) { const [gi, ci] = d.editcat.split('.').map(Number); renameModal('cat', gi, ci); return; }
  if (d.editgroup != null) { renameModal('group', +d.editgroup); return; }
  if (d.rmcat) { const [gi, ci] = d.rmcat.split('.').map(Number); const name = S.cats[gi][1][ci]; const n = catUse(name);
    const doIt = () => { S.cats[gi][1].splice(ci,1); save(); toast(`Đã xoá hạng mục ${name}`); renderMain(); };
    if (n) askConfirm(`Xoá hạng mục ${name}?`, `${n} phiếu đang dùng hạng mục này sẽ giữ tên cũ. Hạng mục không còn trong danh sách chọn.`, 'Xoá hạng mục', doIt); else doIt(); return; }
  if (d.rmgroup != null) { const g = S.cats[+d.rmgroup]; if (g && !g[1].length) { S.cats.splice(+d.rmgroup,1); save(); toast(`Đã xoá nhóm ${g[0]}`); renderMain(); } return; }
  if (d.rm) { const l = listAt(d.rm); const i = l.indexOf(d.uid); if (i >= 0) l.splice(i,1); const lv = cfgPath(d.rm); if (!Array.isArray(lv) && lv.approvers.length < 2) lv.mode = 'any'; renderMain(); return; }
  const a = d.action;
  if (a === 'confirmyes') { const fn = confirmFn; confirmFn = null; closeModal(); if (fn) fn(); return; }
  if (a === 'confirmpay') { confirmPay(); return; }
  if (a === 'reset') { askConfirm('Khôi phục dữ liệu mẫu?', 'Mọi phiếu, đợt chi và cấu hình đã đổi trên trình duyệt này sẽ mất.', 'Khôi phục', () => { seed(); UI.inboxSel = {}; UI.batchSel = {}; go('home'); toast('Đã khôi phục dữ liệu mẫu'); }); return; }
  if (a === 'savedraft') { saveDraft(); return; }
  if (a === 'deldraft') { const r = req(UI.form.id); S.reqs = S.reqs.filter(x => x !== r); save(); toast('Đã xoá nháp'); go('list'); return; }
  if (a === 'cancel') { const r = req(UI.id); askConfirm(`Huỷ ${r.code||'bản nháp'}?`, 'Phiếu dừng ở trạng thái Đã huỷ và rời khỏi hàng chờ duyệt.', 'Huỷ phiếu', () => { cancelReq(r, S.me); save(); toast('Đã huỷ phiếu'); render(); }); return; }
  if (a === 'comment') { const v = ($('#cmt').value||'').trim(); if (!v) return; const r = req(UI.id); log(r, S.me, `bình luận: ${v}`); save(); renderMain(); return; }
  if (a === 'selall') { const items = inboxFor(S.me); const all = items.every(r => UI.inboxSel[r.id]); items.forEach(r => UI.inboxSel[r.id] = !all); renderMain(); return; }
  if (a === 'bulk') { const ids = Object.keys(UI.inboxSel).filter(k => UI.inboxSel[k]); let n = 0; ids.forEach(id => { const r = req(id); if (r && canAct(r, S.me)) { act(r, S.me, 'approve', 'Duyệt hàng loạt', null, actorsFor(r, S.me)[0]); markGuide(r); n++; } }); UI.inboxSel = {}; save(); toast(`Đã duyệt ${n} khoản`); render(); return; }
  if (a === 'mkbatch') {
    const sel = S.reqs.filter(r => r.status === 'approved' && UI.batchSel[r.id]);
    const bad = sel.filter(r => !payeeOf(r).ok);
    if (bad.length) { UI.errors = {batch:`${bad.map(r=>r.code).join(', ')} có STK nhà cung cấp chưa xác minh. Bỏ chọn hoặc xác minh STK trước khi lập đợt.`}; renderMain(); return; }
    const pd = $('#payDateNew')?.value || UI.payDateNew || localDate(); if (pd < localDate()) { UI.errors = {batch:'Ngày hẹn thanh toán không được ở quá khứ.'}; renderMain(); return; }
    const plan = buildBatchChain(sel.reduce((a,r) => a + r.amount, 0), S.me);
    if (plan.problems.length) { UI.errors = {batch: `${plan.problems.join(' ')} Nhờ Quản trị bổ sung người duyệt đợt chi ở Cấu hình.`}; renderMain(); return; }
    const b = createBatch(sel.map(r => r.id), S.me, null, pd); UI.batchSel = {}; UI.payDateNew = ''; S.guideDone.batch = true; save(); toast(`Đã lập ${b.code}, hẹn thanh toán ${dOnly(b.payDate+'T00:00')}, chờ duyệt đợt`); go('batch', b.id); return; }
  if (a === 'addrule') { const C = UI.draft.config; const def = C.rules.find(r => r.def) || C.rules[0];
    const nr = {id:uid('r'), depts:[], people:[], bands: clone(def.bands).map(b => Object.assign(b, {id:uid('b')}))}; C.rules.push(nr); UI.cfgRule = nr.id; renderMain(); toast('Đã thêm quy định, chọn phòng ban hoặc người đề xuất'); return; }
  if (a === 'cfgreset') { UI.draft = {config:clone(S.config)}; renderMain(); return; }
  if (a === 'cfgsave') {
    if (cfgIssues(UI.draft).some(i => i.err)) return;
    const C = UI.draft.config; C.version = S.config.version + 1; C.updatedAt = now(); C.updatedBy = S.me;
    S.config = C; S.guideDone.config = true; save(); UI.draft = {config:clone(S.config)};
    toast(`Đã lưu cấu hình v${S.config.version}, áp cho đề xuất gửi từ bây giờ`); render(); return; }
});

document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && $('#layer .modal-bg')) closeModal();
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('tr[data-open], tr[data-batch], tr[data-vview], .li[role="button"]')) { e.preventDefault(); e.target.click(); }
});

document.addEventListener('input', e => {
  const t = e.target;
  if (t.matches('[data-money]')) {
    const pos = t.value.length - t.selectionStart; const v = digits(t.value); t.value = v ? fmt(v) : '';
    try { const p = Math.max(0, t.value.length - pos); t.setSelectionRange(p,p); } catch(err) {}
  }
  if (t.closest('#reqForm')) { if (t.dataset.ln != null) lineInput(t); else { readForm(); syncTotals(); } updatePreview(); return; }
  if (t.dataset.sim && t.tagName === 'INPUT') { UI.sim.amount = digits(t.value); updateSim(); const n = $('#sim [data-sim="amount"]'); if (n) { n.focus(); try { n.setSelectionRange(n.value.length, n.value.length); } catch(err){} } return; }
  if (t.dataset.max) { findBand(t.dataset.max).band.max = digits(t.value); refreshCfgSide(); return; }
  if (t.dataset.rule === 'cashLimit') { UI.draft.config.cashLimit = digits(t.value); refreshCfgSide(); return; }
});

document.addEventListener('change', e => {
  const t = e.target; const d = t.dataset;
  if (d.change === 'persona') { setPersona(t.value); return; }
  if (d.unc != null) { UI.payDraft.date = $('#payDate')?.value || UI.payDraft.date; addUncFiles([...t.files].map(f => addBlob(f)).filter(Boolean)); drawPayModal(); return; }
  if (d.umap) { const P = UI.payDraft; P.date = $('#payDate')?.value || P.date; if (t.value === '') { delete P.map[d.umap]; } else { P.map[d.umap] = +t.value; } P.why[d.umap] = ''; drawPayModal(); return; }
  if (t.closest('#reqForm') && ['payeeType','vendorId','type','evidence','payMethod','outlay'].includes(t.name)) { readForm(); const f = UI.form;
    if (t.name === 'evidence' && (f.evidence === 'bill' || f.evidence === 'sale')) { f.lines.forEach(l => { l.rate = 'kct'; l.tax = 0; l.pre = l.total; }); }
    if (t.name === 'type') { UI.errors = {}; if (f.type === 'refund') { f.amountManual = true; } else { f.amountManual = false; f.amount = linesSum(f.lines,'total'); } }
    renderMain(); return; }
  if (d.sel) { UI.inboxSel[d.sel] = t.checked; renderMain(); return; }
  if (d.bsel) { UI.batchSel[d.bsel] = t.checked; UI.errors = {}; renderMain(); return; }
  if (d.psel) { UI.paySel[d.psel] = t.checked; renderMain(); return; }
  if (d.sim && t.tagName === 'SELECT') { UI.sim[d.sim] = t.value; updateSim(); return; }
  if (d.add) { if (!t.value) return; listAt(d.add).push(t.value); renderMain(); return; }
  if (d.addscope) { if (!t.value) return; UI.draft.config.rules.find(r => r.id === d.addscope)[d.k].push(t.value); renderMain(); return; }
  if (d.mode) { cfgPath(d.mode).mode = t.value; renderMain(); return; }
  if (d.same) { cfgPath(d.same).sameDept = t.checked; renderMain(); return; }
  if (d.rule === 'blockSelf') { UI.draft.config.blockSelf = t.checked; renderMain(); return; }
  if (d.rule === 'noInvoice') { UI.draft.config.noInvoice.on = t.checked; renderMain(); return; }
  if (d.max) { renderMain(); return; }
});

['dragover','drop'].forEach(type => document.addEventListener(type, e => {
  const z = e.target.closest && e.target.closest('.drop'); if (!z || !UI.payDraft) return;
  e.preventDefault();
  if (type === 'drop' && e.dataTransfer) { UI.payDraft.date = $('#payDate')?.value || UI.payDraft.date; addUncFiles([...e.dataTransfer.files].map(f => addBlob(f)).filter(Boolean)); drawPayModal(); }
}));

document.addEventListener('submit', e => {
  e.preventDefault();
  if (e.target.id === 'reqForm') { submitForm(); return; }
  if (e.target.id === 'userForm') { saveUserForm(e.target); return; }
  if (e.target.classList.contains('catadd') || e.target.id === 'groupAdd') { const name = (new FormData(e.target).get('name') || '').toString().trim(); if (name.length < 2) { toast('Nhập tên, tối thiểu 2 ký tự.'); return; }
    const dup = name === REFUND_CAT || S.cats.some(g => g[0].toLowerCase() === name.toLowerCase() || g[1].some(c => c.toLowerCase() === name.toLowerCase())); if (dup) { toast(`“${name}” đã có.`); return; }
    if (e.target.id === 'groupAdd') S.cats.push([name, []]); else S.cats[+e.target.dataset.g][1].push(name); save(); toast(`Đã thêm ${name}`); renderMain(); return; }
  if (e.target.id === 'delegForm') { saveDelegForm(e.target); return; }
});

/* ===================== Người dùng ===================== */
function permText(u){ const p = []; if (u.acct) p.push('Kế toán'); if (u.signer) p.push('Duyệt đợt chi'); if (u.admin) p.push('Quản trị'); if (u.dash) p.push('Dashboard'); return p.join(', '); }
function configSlots(id){
  const C = S.config; const out = [];
  C.rules.forEach(rl => { const by = {}; sortedBands(rl).forEach((b,i) => b.levels.forEach((l,li) => { if (l.approvers.includes(id)) (by[li+1] = by[li+1] || []).push(i+1); }));
    Object.entries(by).forEach(([st, ks]) => out.push(`${ruleLabel(rl)}: bước ${st} (mốc ${ks.join(', ')})`)); });
  if (C.fallback.includes(id)) out.push('dự phòng'); if (C.noInvoice.on && C.noInvoice.approvers.includes(id)) out.push('kiểm soát không hoá đơn');
  if (C.batch) sortedBands(C.batch).forEach((b,i) => b.levels.forEach((l,li) => { if (l.approvers.includes(id)) out.push(`Đợt chi: cấp ${li+1}, mốc ${i+1}`); }));
  return out;
}
function viewUsers(){
  const list = S.users.filter(u => !u.deleted);
  return `${head('Quản trị','Người dùng', `${list.filter(u=>u.active).length} người đang hoạt động. Quyền duyệt đề xuất chọn theo từng bước ở Cấu hình phê duyệt.`, `<button class="btn pri" data-unew="1">${icon('plus',16)} Thêm người dùng</button>`)}
  <div class="tw"><table>
    <thead><tr><th>Người dùng</th><th>Phòng</th><th>Quyền</th><th>Trong luồng duyệt</th><th>Trạng thái</th><th></th></tr></thead>
    <tbody>${list.map(u => { const sl = configSlots(u.id); return `<tr>
      <td><div class="row" style="flex-wrap:nowrap">${avatar(u.id,'sm')}<div><div class="ttl">${esc(u.name)}</div><div class="small faint">${esc(u.title)}</div></div></div></td>
      <td class="small muted">${esc(dept(u.dept)?.name||'—')}</td>
      <td class="small">${esc(permText(u)) || '<span class="faint">—</span>'}</td>
      <td class="small muted">${sl.length ? esc(sl.join(', ')) : '<span class="faint">—</span>'}</td>
      <td>${u.active ? '<span class="pill st-approved">Hoạt động</span>' : '<span class="pill">Ngừng</span>'}</td>
      <td style="text-align:right;white-space:nowrap"><button class="btn ghost sm" data-uedit="${u.id}">Sửa</button><button class="btn ghost sm" data-udel="${u.id}" style="color:var(--critical)">Xoá</button></td></tr>`; }).join('')}</tbody></table></div>
  <p class="small faint" style="margin-top:12px">“Trong luồng duyệt” cho biết người này đang được chọn ở khung tiền và bước nào. Muốn xoá người đang trong luồng, gỡ họ ở Cấu hình phê duyệt trước.</p>`;
}
function userModal(id){
  const u = id ? user(id) : {name:'', title:'', dept:me().dept, acct:false, signer:false, admin:false, active:true};
  UI.uedit = id || null;
  const perm = (k, t, d) => `<label class="toggle" style="padding:6px 0"><input type="checkbox" name="${k}" ${u[k]?'checked':''}><div><b>${t}</b><span>${d}</span></div></label>`;
  $('#layer').innerHTML = `<div class="modal-bg" data-close="1"><form class="modal" id="userForm" role="dialog" aria-modal="true" aria-labelledby="uh" novalidate>
    <div class="modal-h"><h3 id="uh">${id ? 'Sửa người dùng' : 'Thêm người dùng'}</h3>${id ? `<p>${esc(u.name)}</p>` : ''}</div>
    <div class="modal-b stack">
      <div class="fgrid">
        <label class="f full">Họ và tên<input class="in" name="name" value="${esc(u.name)}" autocomplete="off" required></label>
        <label class="f">Chức danh<input class="in" name="title" value="${esc(u.title)}" autocomplete="off"></label>
        <label class="f">Email đăng nhập<input class="in" name="email" type="email" value="${esc(u.email||'')}" autocomplete="off"><span class="hint">Dùng đúng email này để đăng nhập.</span></label>
        <label class="f">Phòng<select class="in" name="dept">${S.depts.map(d => `<option value="${d.id}" ${d.id===u.dept?'selected':''}>${esc(d.name)}</option>`).join('')}</select></label>
      </div>
      <div><span class="flabel">Phân quyền</span>
        <div class="checks">${perm('acct','Kế toán','Lập đợt chi, xem thông tin chuyển khoản, đánh dấu đã chuyển và đính kèm UNC.')}
        ${perm('signer','Duyệt đợt chi','Duyệt đợt chi trước khi kế toán chuyển tiền.')}
        ${perm('admin','Quản trị','Quản lý người dùng và cấu hình phê duyệt.')}
        ${perm('dash','Dashboard quản trị','Xem dashboard quản trị: dòng tiền, chi theo phòng và hạng mục, rủi ro. Quản trị luôn xem được.')}
        ${id ? perm('active','Đang hoạt động','Bỏ chọn để tạm khoá: người này không đăng nhập và không nhận việc duyệt.') : ''}</div>
        <p class="small faint" style="margin:6px 0 0">Ai cũng tạo được đề xuất chi. Người duyệt đề xuất được chọn theo từng bước ở Cấu hình phê duyệt.</p></div>
      <div id="uErr"></div>
    </div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">${id ? 'Lưu thay đổi' : 'Thêm người dùng'}</button></div></form></div>`;
  $('#userForm [name=name]').focus();
}
function saveUserForm(f){
  const fd = new FormData(f); const id = UI.uedit;
  const v = {name:(fd.get('name')||'').trim(), title:(fd.get('title')||'').trim(), email:(fd.get('email')||'').trim().toLowerCase(), dept:fd.get('dept'), acct:!!fd.get('acct'), signer:!!fd.get('signer'), admin:!!fd.get('admin'), dash:!!fd.get('dash'), active: id ? !!fd.get('active') : true};
  const errs = [];
  if (v.name.length < 2) errs.push('Nhập họ và tên.');
  if (!v.title) errs.push('Nhập chức danh.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.email)) errs.push('Nhập email đăng nhập hợp lệ.');
  if (S.users.some(u => !u.deleted && u.id !== id && (u.email||'') === v.email)) errs.push('Email này đã gắn với người dùng khác.');
  const others = S.users.filter(u => !u.deleted && u.id !== id);
  if (id && !v.admin && !others.some(u => u.admin && u.active)) errs.push('Phải còn ít nhất một người có quyền Quản trị.');
  if (id === S.me && !v.active) errs.push('Không tự khoá tài khoản đang dùng.');
  if (id === S.me && !v.admin && user(id).admin) errs.push('Không tự bỏ quyền Quản trị của tài khoản đang dùng.');
  if (errs.length) { $('#uErr').innerHTML = `<div class="alert neg">${errs.map(esc).join('<br>')}</div>`; return; }
  if (id) { const u = user(id); const was = u.active; Object.assign(u, v); toast(`Đã lưu ${u.name}${was && !v.active && inConfig(id) ? '. Người này còn trong luồng duyệt, kiểm tra Cấu hình phê duyệt.' : ''}`); }
  else { S.users.push(Object.assign({id:uid('u'), deleted:false}, v)); toast(`Đã thêm ${v.name}`); }
  S.guideDone.users = true; save(); closeModal(); render();
}
function deleteBlockers(id){
  const u = user(id); const out = [];
  if (id === S.me) out.push('Không tự xoá tài khoản đang dùng.');
  const sl = configSlots(id); if (sl.length) out.push(`Đang trong luồng duyệt (${sl.join(', ')}). Gỡ ở Cấu hình phê duyệt trước.`);
  const waiting = S.reqs.filter(r => r.status === 'pending' && waitingOn(r).includes(id)); if (waiting.length) out.push(`Đang có phiếu chờ người này duyệt: ${waiting.map(r => r.code).join(', ')}.`);
  const dg = (S.delegations||[]).filter(g => (g.from === id || g.to === id) && delegActive(g)); if (dg.length) out.push('Có ủy quyền đang hiệu lực liên quan tới người này. Thu hồi trước.');
  if (u.admin && !S.users.some(x => x.id !== id && !x.deleted && x.active && x.admin)) out.push('Là người Quản trị cuối cùng.');
  if (u.signer && !S.users.some(x => x.id !== id && !x.deleted && x.active && x.signer)) out.push('Là người duyệt đợt chi duy nhất.');
  if (S.batches.some(b => b.status === 'review' && b.createdBy === id)) out.push('Có đợt chi do người này lập đang chờ duyệt.');
  return out;
}
function deleteUser(id){
  const u = user(id); const bl = deleteBlockers(id);
  if (bl.length) {
    $('#layer').innerHTML = `<div class="modal-bg" data-close="1"><div class="modal" role="alertdialog" aria-modal="true" aria-labelledby="dh" style="width:min(480px,100%)">
      <div class="modal-h"><h3 id="dh">Chưa xoá được ${esc(u.name)}</h3><p>Xử lý các mục sau rồi xoá lại. Nếu chỉ cần chặn tạm, bỏ chọn “Đang hoạt động” trong Sửa.</p></div>
      <div class="modal-b"><ul style="margin:0;padding-left:18px">${bl.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>
      <div class="modal-f"><button class="btn sec" data-close="1">Đóng</button></div></div></div>`;
    return;
  }
  askConfirm(`Xoá ${u.name}?`, 'Người này không còn đăng nhập và không xuất hiện trong danh sách chọn. Lịch sử phiếu và nhật ký vẫn giữ tên.', 'Xoá người dùng', () => { u.deleted = true; u.active = false; save(); toast(`Đã xoá ${u.name}`); render(); });
}

/* ===================== Ủy quyền duyệt ===================== */
function delegPeriod(g){ return !g.start && !g.end ? 'Không giới hạn' : `${g.start ? dOnly(g.start + 'T00:00') : '…'} – ${g.end ? dOnly(g.end + 'T00:00') : 'không giới hạn'}`; }
function onBehalfLog(){
  const out = [];
  S.reqs.forEach(r => r.rounds.forEach(rd => rd.chain.forEach(s => Object.entries(s.decisions).forEach(([p,d]) => { if (d.by && d.by !== p) out.push({at:d.at, by:d.by, forId:p, open:r.id, what:`${r.code} · ${s.label}`, act:d.act==='approve'?'Duyệt':d.act==='reject'?'Từ chối':'Trả lại'}); }))));
  S.batches.forEach(b => { if (b.approvedFor) out.push({at:b.approvedAt, by:b.approvedBy, forId:b.approvedFor, batch:b.id, what:`Đợt ${b.code}`, act:'Duyệt đợt'}); if (b.returnedFor) out.push({at:b.returnedAt, by:b.returnedBy, forId:b.returnedFor, batch:b.id, what:`Đợt ${b.code}`, act:'Trả lại đợt'}); });
  return out.sort((a,b) => b.at - a.at);
}
function delegTable(rows, cols){
  if (!rows.length) return '<div class="empty" style="padding:20px"><b>Chưa có</b></div>';
  return `<div class="tw"><table><thead><tr>${cols.from?'<th>Người ủy quyền</th>':''}${cols.to?'<th>Người được ủy quyền</th>':''}<th>Thời hạn</th><th>Lý do</th><th>Trạng thái</th><th></th></tr></thead><tbody>
    ${rows.map(g => { const st = delegState(g); return `<tr>
      ${cols.from?`<td><div class="row" style="flex-wrap:nowrap">${avatar(g.from,'xs')}${esc(user(g.from).name)}</div></td>`:''}
      ${cols.to?`<td><div class="row" style="flex-wrap:nowrap">${avatar(g.to,'xs')}${esc(user(g.to).name)}</div></td>`:''}
      <td class="small">${esc(delegPeriod(g))}</td><td class="small muted">${esc(g.note||'—')}</td>
      <td><span class="pill ${st[1]}">${st[0]}</span>${g.revokedAt?`<div class="small faint">${dt(g.revokedAt)}</div>`:''}</td>
      <td style="text-align:right">${!g.revokedAt && (g.from === S.me || me().admin) && st[0] !== 'Hết hạn' ? `<button class="btn ghost sm" data-revoke="${g.id}">Thu hồi</button>` : ''}</td></tr>`; }).join('')}</tbody></table></div>`;
}
function viewDelegate(){
  const u = me(); const G = S.delegations || []; const dg = UI.dg;
  const mine = G.filter(g => g.from === u.id).sort((a,b) => b.createdAt - a.createdAt);
  const toMe = G.filter(g => g.to === u.id).sort((a,b) => b.createdAt - a.createdAt);
  const pr = principalsOf(u.id);
  const pendingFor = S.reqs.filter(r => actorsFor(r, u.id).some(p => p !== u.id)).length + S.batches.filter(b => batchActors(b, u.id).some(p => p !== u.id)).length;
  const logs = onBehalfLog().filter(x => u.admin || x.by === u.id || x.forId === u.id);
  const fromId = dg.from || u.id; const fromU = user(fromId);
  const hasRole = inConfig(fromId) || isSigner(fromId);
  return `${head('Phê duyệt','Ủy quyền duyệt','Khi vắng mặt, người được ủy quyền duyệt thay các bước và đợt chi đang giao cho bạn. Mọi lần duyệt thay đều ghi rõ duyệt thay ai.')}
  <div class="stack">
    ${pr.length ? `<div class="alert info">Bạn đang được ủy quyền duyệt thay ${esc(vnJoin(pr.map(id => user(id).name)))}. ${pendingFor ? `Có ${pendingFor} việc chờ trong <button class="linkbtn" data-go="inbox">Chờ tôi duyệt</button>${pendingFor && S.batches.some(b => batchActors(b,u.id).some(p=>p!==u.id)) ? ' và Đợt chi' : ''}.` : 'Hiện chưa có việc chờ.'}</div>` : ''}
    <div class="g75">
      <div class="card"><div class="ch"><h2 class="card-title">${u.admin && fromId !== u.id ? 'Tạo ủy quyền thay người khác' : 'Ủy quyền cho người khác'}</h2></div>
        <form id="delegForm" class="stack" novalidate>
          <div class="fgrid">
            ${u.admin ? `<label class="f">Người ủy quyền<select class="in" name="from" data-dg="from">${activeUsers().map(x => `<option value="${x.id}" ${x.id===fromId?'selected':''}>${x.id===u.id?'Tôi — ':''}${esc(x.name)}</option>`).join('')}</select><span class="hint">Quản trị tạo được ủy quyền cho người đang vắng.</span></label>` : ''}
            <label class="f ${u.admin?'':'full'}">Người được duyệt thay<select class="in" name="to" data-dg="to"><option value="">Chọn người</option>${activeUsers().filter(x => x.id !== fromId).map(x => `<option value="${x.id}" ${x.id===dg.to?'selected':''}>${esc(x.name)} · ${esc(x.title)}</option>`).join('')}</select></label>
            <div class="f full"><span class="flabel">Thời hạn</span>
              <div class="choice two" role="radiogroup" aria-label="Thời hạn">
                <label><input type="radio" name="mode" value="none" data-dg="mode" ${dg.mode==='none'?'checked':''}>Không giới hạn<span>Hiệu lực tới khi thu hồi</span></label>
                <label><input type="radio" name="mode" value="range" data-dg="mode" ${dg.mode==='range'?'checked':''}>Trong khoảng thời gian<span>Tự hết hiệu lực sau ngày kết thúc</span></label>
              </div></div>
            ${dg.mode === 'range' ? `<label class="f">Từ ngày<input class="in" type="date" name="start" value="${esc(dg.start)}"></label><label class="f">Đến hết ngày<input class="in" type="date" name="end" value="${esc(dg.end)}"></label>` : ''}
            <label class="f full">Lý do<input class="in" name="note" value="${esc(dg.note)}" placeholder="VD: Nghỉ phép, công tác" autocomplete="off"></label>
          </div>
          ${!hasRole ? `<p class="small faint" style="margin:0">${fromId===u.id?'Bạn':esc(fromU.name)} hiện không có bước duyệt hay quyền duyệt đợt nào, ủy quyền chỉ có tác dụng khi được giao duyệt.</p>` : ''}
          ${dg.err ? `<div class="alert neg">${esc(dg.err)}</div>` : ''}
          <div class="row"><span class="small faint" style="flex:1">Người được ủy quyền không duyệt thay được phiếu do chính họ đề xuất, và không ký hai lần trong cùng một phiếu.</span><button class="btn pri" type="submit">Ủy quyền</button></div>
        </form></div>
      <div class="card"><div class="ch"><h2 class="card-title">Được ủy quyền cho tôi</h2><span class="card-note">${toMe.filter(g=>delegActive(g)).length} đang hiệu lực</span></div>${delegTable(toMe, {from:true})}</div>
    </div>
    <div class="card"><div class="ch"><h2 class="card-title">Tôi đã ủy quyền</h2><span class="card-note">${mine.length} lần</span></div>${delegTable(mine, {to:true})}</div>
    ${u.admin ? `<div class="card"><div class="ch"><h2 class="card-title">Toàn bộ ủy quyền trong công ty</h2><span class="card-note">${G.filter(g=>delegActive(g)).length} đang hiệu lực</span></div>${delegTable([...G].sort((a,b)=>b.createdAt-a.createdAt), {from:true, to:true})}</div>` : ''}
    <div class="card"><div class="ch"><h2 class="card-title">Nhật ký duyệt thay</h2><span class="card-note">${u.admin ? 'toàn công ty' : 'liên quan tới bạn'}</span></div>
      ${logs.length ? `<div class="tw"><table><thead><tr><th>Thời gian</th><th>Người duyệt</th><th>Duyệt thay</th><th>Thao tác</th><th>Phiếu / đợt</th></tr></thead><tbody>
        ${logs.map(x => `<tr class="click" ${x.open?`data-open="${x.open}"`:`data-batch="${x.batch}"`} tabindex="0"><td class="small num">${dt(x.at)}</td><td>${esc(user(x.by).name)}</td><td>${esc(user(x.forId).name)}</td><td class="small">${x.act}</td><td class="small muted">${esc(x.what)}</td></tr>`).join('')}</tbody></table></div>`
      : '<div class="empty" style="padding:20px"><b>Chưa có lần duyệt thay nào</b></div>'}</div>
  </div>`;
}
function readDelegForm(){ const f = $('#delegForm'); if (!f) return; const fd = new FormData(f); const dg = UI.dg;
  if (fd.get('from')) dg.from = fd.get('from'); dg.to = fd.get('to') || ''; dg.mode = fd.get('mode') || 'none';
  if (fd.has('start')) dg.start = fd.get('start'); if (fd.has('end')) dg.end = fd.get('end'); dg.note = (fd.get('note')||'').trim(); }
function saveDelegForm(){
  readDelegForm(); const dg = UI.dg; const from = dg.from || S.me; let err = '';
  if (!dg.to) err = 'Chọn người được duyệt thay.';
  else if (dg.to === from) err = 'Không ủy quyền cho chính mình.';
  else if (dg.mode === 'range' && (!dg.start || !dg.end)) err = 'Chọn ngày bắt đầu và ngày kết thúc.';
  else if (dg.mode === 'range' && dg.end < dg.start) err = 'Ngày kết thúc phải sau ngày bắt đầu.';
  else if (dg.mode === 'range' && dg.end < localDate()) err = 'Ngày kết thúc đã qua.';
  else if (S.delegations.some(g => g.from === from && g.to === dg.to && !g.revokedAt && !(g.end && g.end < localDate()))) err = `Đã có ủy quyền còn hiệu lực cho ${user(dg.to).name}. Thu hồi trước nếu muốn đổi thời hạn.`;
  if (err) { dg.err = err; renderMain(); return; }
  S.delegations.push({id:uid('g'), from, to:dg.to, start: dg.mode==='range' ? dg.start : null, end: dg.mode==='range' ? dg.end : null, note:dg.note, createdAt:now(), createdBy:S.me, revokedAt:null});
  S.guideDone.mkdeleg = true; save();
  toast(`${user(dg.to).name} được duyệt thay ${from === S.me ? 'bạn' : user(from).name}${dg.mode==='range' ? ' tới hết ' + dOnly(dg.end+'T00:00') : ', không giới hạn'}`);
  UI.dg = {from:S.me, to:'', mode:'none', start:localDate(), end:'', note:'', err:''}; render();
}

document.addEventListener('click', e => {
  const t = e.target.closest('button'); if (!t) return; const d = t.dataset;
  if (d.unew) { userModal(null); return; }
  if (d.uedit) { userModal(d.uedit); return; }
  if (d.udel) { deleteUser(d.udel); return; }
  if (d.revoke) { const g = S.delegations.find(x => x.id === d.revoke);
    askConfirm('Thu hồi ủy quyền?', `${user(g.to).name} sẽ không duyệt thay ${user(g.from).name} nữa. Những lần đã duyệt thay vẫn giữ trong nhật ký.`, 'Thu hồi', () => { g.revokedAt = now(); g.revokedBy = S.me; save(); toast('Đã thu hồi ủy quyền'); render(); }); return; }
});
document.addEventListener('change', e => {
  const t = e.target; if (!t.dataset.dg) return;
  readDelegForm(); UI.dg.err = ''; if (t.dataset.dg === 'from' && UI.dg.to === UI.dg.from) UI.dg.to = ''; renderMain();
});

/* ===================== Hộp thoại nghiệp vụ ===================== */
function modal(html, width){ $('#layer').innerHTML = `<div class="modal-bg" data-close="1">${html.replace('<form ', `<form style="width:min(${width||520}px,100%)" `)}</div>`; const first = $('#layer input:not([type=hidden]):not([type=file]), #layer textarea, #layer select'); if (first) first.focus(); }
let reasonFn = null;
function askReason(title, body, yes, fn){
  reasonFn = fn;
  modal(`<form class="modal" id="reasonForm" role="dialog" aria-modal="true" aria-labelledby="rh" novalidate>
    <div class="modal-h"><h3 id="rh">${esc(title)}</h3><p>${esc(body)}</p></div>
    <div class="modal-b"><label class="f">Lý do<textarea class="in" name="reason" placeholder="Bắt buộc, tối thiểu 5 ký tự"></textarea><span class="ferr" id="reasonErr"></span></label></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Giữ nguyên</button><button type="submit" class="btn danger">${esc(yes)}</button></div></form>`, 460);
}
function vendorModal(id, o){
  o = o || {ctx:'form'}; const v = (id && vendor(id)) || {name:'', tax:'', bank:'', acc:'', accName:''};
  UI.vctx = {id: id || null, ctx: o.ctx || 'form'};
  modal(`<form class="modal" id="vendorForm" role="dialog" aria-modal="true" aria-labelledby="vh" novalidate>
    <div class="modal-h"><h3 id="vh">${id ? 'Sửa nhà cung cấp' : 'Thêm nhà cung cấp'}</h3><p>${id ? 'Đổi ngân hàng, số hoặc tên tài khoản sẽ đưa NCC về trạng thái chưa xác minh.' : 'STK mới ở trạng thái chưa xác minh cho tới khi kế toán kiểm tra.'}</p></div>
    <div class="modal-b"><div class="fgrid">
      <label class="f full">Tên nhà cung cấp<input class="in" name="name" autocomplete="off" value="${esc(v.name)}"></label>
      <label class="f">Mã số thuế<input class="in" name="tax" autocomplete="off" inputmode="numeric" value="${esc(v.tax || '')}"></label>
      <label class="f">Ngân hàng<input class="in" name="bank" autocomplete="off" placeholder="VD: Vietcombank" value="${esc(v.bank || '')}"></label>
      <label class="f">Số tài khoản<input class="in" name="acc" autocomplete="off" data-acc value="${esc(v.acc || '')}"><span class="hint">Số và chữ, không dấu cách.</span></label>
      <label class="f">Tên tài khoản<input class="in" name="accName" autocomplete="off" placeholder="IN HOA KHONG DAU" data-upper value="${esc(v.accName || '')}"></label>
    </div><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">${id ? 'Lưu thay đổi' : 'Thêm nhà cung cấp'}</button></div></form>`, 560);
}
function directModal(pre){
  pre = pre || {}; const act = S.vendors.filter(x => !x.deleted);
  modal(`<form class="modal" id="directForm" role="dialog" aria-modal="true" aria-labelledby="dh" novalidate>
    <div class="modal-h"><h3 id="dh">Khoản chi trực tiếp</h3><p>Kế toán lập thẳng vào hàng chờ thanh toán, không qua đề xuất. Người duyệt đợt chi là chốt duyệt duy nhất.</p></div>
    <div class="modal-b"><div class="fgrid">
      <label class="f full">Nội dung chi<input class="in" name="title" value="${esc(pre.title || '')}" autocomplete="off" placeholder="VD: Phí duy trì tài khoản tháng 11"></label>
      <div class="f full"><span class="flabel">Nhà cung cấp</span>
        <div class="row" style="flex-wrap:nowrap"><input class="in" type="search" data-vfilter placeholder="Tìm theo tên, MST hoặc số tài khoản" autocomplete="off" style="flex:1"><button type="button" class="btn sec" data-action="dnewvendor">${icon('plus',16)} NCC mới</button></div>
        <select class="in" name="vendorId" style="margin-top:6px">${act.map(x => `<option value="${x.id}" ${x.id === pre.vendorId ? 'selected' : ''}>${esc(x.name)} — ${esc(x.bank)} ${esc(x.acc)}${x.verified ? '' : ' (chưa xác minh)'}</option>`).join('')}</select></div>
      <label class="f">Hạng mục chi phí<select class="in" name="cat">${catOptions(pre.cat || '')}</select></label>
      <label class="f">Phòng chịu chi phí<select class="in" name="dept">${S.depts.map(d => `<option value="${d.id}" ${d.id === (pre.dept || me().dept) ? 'selected' : ''}>${esc(d.name)}</option>`).join('')}</select></label>
      <label class="f">Thành tiền, gồm VAT<div class="money"><input class="in" name="total" inputmode="numeric" data-money autocomplete="off" value="${esc(pre.total || '')}"><span>đ</span></div></label>
      <label class="f">Trong đó tiền thuế<div class="money"><input class="in" name="tax" inputmode="numeric" data-money autocomplete="off" value="${esc(pre.tax || '')}"><span>đ</span></div></label>
      <label class="f">Số hoá đơn<input class="in" name="inv" autocomplete="off" value="${esc(pre.inv || '')}"></label>
      <label class="f">Hình thức thanh toán<select class="in" name="payMethod">${Object.entries(PAYM).map(([k,x]) => `<option value="${k}" ${k === (pre.payMethod || 'transfer') ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label class="f full">Chứng từ<input class="in" type="file" name="files" multiple style="padding:6px 10px;height:auto"></label>
    </div><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Thêm vào chờ thanh toán</button></div></form>`, 600);
}
function invModal(id){
  const r = req(id); UI.invFor = id;
  modal(`<form class="modal" id="invForm" role="dialog" aria-modal="true" aria-labelledby="ih" novalidate>
    <div class="modal-h"><h3 id="ih">Bổ sung hoá đơn</h3><p>${esc(r.code)} · ${esc(payeeOf(r).name)} · tổng ${money(linesSum(r.lines,'total'))}</p></div>
    <div class="modal-b"><div class="fgrid">
      <label class="f">Số hoá đơn<input class="in" name="inv" autocomplete="off" placeholder="Nhiều số cách nhau dấu phẩy"></label>
      <label class="f">Tiền thuế trên hoá đơn<div class="money"><input class="in" name="tax" inputmode="numeric" data-money value="${fmt(linesSum(r.lines,'tax'))}" autocomplete="off"><span>đ</span></div><span class="hint">Sửa nếu lệch với ĐNTT.</span></label>
      <label class="f full">File hoá đơn (PDF, XML)<input class="in" type="file" name="files" multiple style="padding:6px 10px;height:auto"></label>
    </div><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Lưu hoá đơn</button></div></form>`, 520);
}
function pasteModal(){
  modal(`<form class="modal" id="pasteForm" role="dialog" aria-modal="true" aria-labelledby="pth" novalidate>
    <div class="modal-h"><h3 id="pth">Dán nhiều dòng từ Excel</h3><p>Mỗi dòng một hoá đơn hoặc một mức thuế. Thứ tự cột: Số hoá đơn, Diễn giải, Trước thuế, Thuế suất, Tiền thuế. Bỏ trống tiền thuế thì hệ thống tự tính.</p></div>
    <div class="modal-b"><textarea class="in" name="text" style="min-height:180px;font-variant-numeric:tabular-nums" placeholder="0001234&#9;Phí di chuyển&#9;1.200.000&#9;8%&#9;96.000&#10;0001234&#9;Phí dịch vụ&#9;300.000&#9;10%&#9;30.000"></textarea><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Thêm các dòng</button></div></form>`, 640);
}
function wrongModal(id){
  const r = req(id); UI.wrongFor = id;
  modal(`<form class="modal" id="wrongForm" role="dialog" aria-modal="true" aria-labelledby="wh" novalidate>
    <div class="modal-h"><h3 id="wh">Báo chuyển nhầm, chờ NCC hoàn</h3><p>${esc(r.code)} · đã chi ${money(r.amount)} cho ${esc(payeeOf(r).name)}</p></div>
    <div class="modal-b"><div class="fgrid">
      <label class="f">Số tiền NCC phải hoàn<div class="money"><input class="in" name="amount" inputmode="numeric" data-money value="${fmt(r.amount)}" autocomplete="off"><span>đ</span></div></label>
      <label class="f full">Diễn giải<textarea class="in" name="note" placeholder="VD: Chuyển trùng 2 lần, NCC xác nhận hoàn trong tuần"></textarea></label>
    </div><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Ghi nhận chờ hoàn</button></div></form>`, 500);
}
const mErr = t => { const e = $('#mErr'); if (e) e.innerHTML = `<div class="alert neg">${esc(t)}</div>`; };

/* ===================== Lịch sử và báo cáo ===================== */
function historyRows(){
  const u = me(); const f = UI.rflt; const q = noAccent(f.q || '').toLowerCase(); const out = []; const scoped = f.cat || f.who;
  visibleReqs().filter(r => catMatch(r, f.cat) && (!f.who || r.requester === f.who)).forEach(r => r.log.forEach(l => out.push({at:l.at, by:l.by, code:r.code || 'Nháp', what:r.title, text:l.text, open:r.id, amt:r.amount})));
  if (seesBatches(u) && !scoped) S.batches.forEach(b => (b.log||[]).forEach(l => out.push({at:l.at, by:l.by, code:b.code, what:`Đợt chi ${b.reqIds.length} khoản`, text:l.text, batch:b.id, amt:batchTotal(b)})));
  return out.filter(x => inRange(x.at, f) && (!q || noAccent([x.code, x.what, x.text, x.by ? user(x.by).name : 'Hệ thống'].join(' ')).toLowerCase().includes(q))).sort((a,b) => b.at - a.at);
}
function advGroup(adv, keyFn){
  const m = {};
  adv.forEach(r => {
    const k = keyFn(r); const g = (m[k] = m[k] || {k, n:0, amt:0, pend:0, over:0, rwait:0, rdone:0});
    g.n++; g.amt += r.amount; const st = advState(r);
    if (['waiting','submitted','refund_waiting'].includes(st)) g.pend++;
    if (st === 'overdue') g.over++;
    if (r.settle && r.settle.refund) { if (r.settle.refund.status === 'received') g.rdone += r.settle.refund.amount; else g.rwait += r.settle.refund.amount; }
  });
  return Object.values(m).sort((a,b) => b.amt - a.amt);
}
function advTable(title, rows){
  const label = title.replace('Theo ', '');
  return `<div class="card"><div class="ch"><h2 class="card-title">${title}</h2><span class="card-note">${rows.length} dòng</span></div>${rows.length ? `<div class="tw"><table><thead><tr><th>${esc(label.charAt(0).toUpperCase() + label.slice(1))}</th><th class="r">Số khoản</th><th class="r">Đã tạm ứng (đ)</th><th class="r">Chưa đóng</th><th class="r">Hoàn thừa chờ (đ)</th><th class="r">Hoàn thừa đã nhận (đ)</th></tr></thead><tbody>${rows.map(g => `<tr><td class="small">${esc(g.k)}</td><td class="r">${g.n}</td><td class="r">${fmt(g.amt)}</td><td class="r" style="${g.over ? 'color:var(--critical)' : ''}">${g.pend}${g.over ? ` · ${g.over} quá hạn` : ''}</td><td class="r" style="${g.rwait ? 'color:var(--watch)' : ''}">${fmt(g.rwait)}</td><td class="r">${fmt(g.rdone)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="empty" style="padding:16px"><b>Chưa có dữ liệu</b></div>'}</div>`;
}
function advGroupRows(list, label){
  return [[label, 'Số khoản', 'Đã tạm ứng', 'Chưa đóng', 'Quá hạn', 'Hoàn thừa chờ', 'Hoàn thừa đã nhận']].concat(list.map(g => [g.k, g.n, g.amt, g.pend, g.over, g.rwait, g.rdone]));
}
function dueModal(id){
  const r = req(id); if (!r) return; UI.dueFor = id;
  modal(`<form class="modal" id="dueForm" role="dialog" aria-modal="true" aria-labelledby="duh" novalidate>
    <div class="modal-h"><h3 id="duh">Sửa hạn quyết toán</h3><p>${esc(r.code)} · ${esc(r.title)}. Hạn hiện tại ${r.settleDue ? dOnly(r.settleDue) : '—'}.</p></div>
    <div class="modal-b"><label class="f">Hạn mới<input class="in" type="date" name="settleDue" min="${localDate()}" value="${r.settleDue ? localDate(r.settleDue) : ''}"></label><span class="hint">Thay đổi được ghi vào nhật ký phiếu.</span><div id="mErr" style="margin-top:12px"></div></div>
    <div class="modal-f"><button type="button" class="btn sec" data-close="1">Huỷ</button><button type="submit" class="btn pri">Lưu hạn mới</button></div></form>`, 460);
}
function advSummary(rs){
  const adv = rs.filter(r => r.evidence === 'advance');
  const st = r => advState(r);
  const sumOf = (list, fn) => list.reduce((a,r) => a + fn(r), 0);
  const settledOf = r => r.settle && r.settle.reqId ? req(r.settle.reqId) : null;
  const spentOf = r => sumOf(r.settleLines || [], l => l.total || 0);
  return {
    n: adv.length, amt: sumOf(adv, r => r.amount),
    paidN: adv.filter(r => r.status === 'paid').length,
    done: adv.filter(r => st(r) === 'done').length,
    pendingN: adv.filter(r => ['waiting','submitted','refund_waiting'].includes(st(r))).length,
    overdueN: adv.filter(r => st(r) === 'overdue').length,
    spent: sumOf(adv.filter(r => r.settleLines), spentOf),
    extra: sumOf(adv.filter(r => settledOf(r) && settledOf(r).status !== 'cancelled' && settledOf(r).status !== 'rejected'), r => settledOf(r).amount),
    refundDone: sumOf(adv.filter(r => r.settle && r.settle.refund && r.settle.refund.status === 'received'), r => r.settle.refund.amount),
    refundWait: sumOf(adv.filter(r => st(r) === 'refund_waiting'), r => r.settle.refund.amount),
    refundWaitN: adv.filter(r => st(r) === 'refund_waiting').length,
    rows: adv,
    byPerson: advGroup(adv, r => user(r.requester).name),
    byDept: advGroup(adv, r => dept(r.dept).name)
  };
}
function summaryData(){
  const f = UI.rflt; const rs = visibleReqs().filter(r => r.status !== 'draft' && inRange(r.createdAt, f) && catMatch(r, f.cat) && (!f.who || r.requester === f.who));
  const grp = (keyFn) => { const m = {}; rs.forEach(r => { const k = keyFn(r); (m[k] = m[k] || {k, n:0, amt:0, paid:0}); m[k].n++; m[k].amt += r.amount; if (r.status === 'paid') m[k].paid += r.amount; }); return Object.values(m).sort((a,b) => b.amt - a.amt); };
  const done = rs.filter(r => r.approvedAt && r.submittedAt && !r.direct);
  const paid = rs.filter(r => r.status === 'paid');
  return {rs, paid, adv: advSummary(rs),
    kpi:{made:rs.length, approved:rs.filter(r => r.approvedAt).length, paidN:paid.length, paidAmt:paid.reduce((a,r)=>a+r.amount,0), closed:rs.filter(r => ['rejected','cancelled'].includes(r.status)).length,
      avgH: done.length ? done.reduce((a,r) => a + (r.approvedAt - r.submittedAt), 0) / done.length / H : null, invWait: rs.filter(invPending).length, tax: paid.reduce((a,r) => { const s = linesSum(r.lines,'total'); return a + (s ? Math.round(linesSum(r.lines,'tax') * r.amount / s) : 0); }, 0)},
    tables:[
      ['Theo trạng thái', grp(r => STATUS[r.status][0])], ['Theo nhóm chi phí', grp(r => catGroup(r.cat))], ['Theo hạng mục', grp(r => r.cat)], ['Theo phòng', grp(r => dept(r.dept).name)],
      ['Theo hình thức thanh toán', grp(r => PAYM[pmOf(r)])], ['Theo người nhận', grp(r => payeeOf(r).name)], ['Theo người đề xuất', grp(r => user(r.requester).name)]
    ]};
}
function viewReports(){
  const tab = UI.rtab || 'history'; const f = UI.rflt;
  const tabs = `<div class="tabs" role="tablist"><button role="tab" class="${tab==='history'?'on':''}" data-rtab="history">Lịch sử thao tác</button><button role="tab" class="${tab==='summary'?'on':''}" data-rtab="summary">Báo cáo tổng hợp</button></div>`;
  if (tab === 'history') {
    const rows = historyRows(); const show = rows.slice(0,300);
    return `${head('Theo dõi','Lịch sử và báo cáo','Mọi thao tác lập, duyệt, huỷ, chi trên phiếu và đợt chi ở một chỗ. Tìm chung, lọc theo ngày, xuất Excel.')}${tabs}
    ${filterBar('rflt', {q:'Tìm mã phiếu, mã đợt, người thực hiện, nội dung thao tác', cat:true, who:true, extra:`<span class="sp"></span><button class="btn sec sm" data-action="xhistory">Xuất Excel</button>`})}
    ${rows.length ? `<div class="tw"><table><thead><tr><th class="r">STT</th><th>Thời gian</th><th>Mã</th><th>Người thực hiện</th><th>Thao tác</th><th>Phiếu, đợt</th><th class="r">Số tiền (đ)</th></tr></thead><tbody>
      ${show.map((x,i) => `<tr class="click" ${x.open?`data-open="${x.open}"`:`data-batch="${x.batch}"`} tabindex="0"><td class="r faint">${i+1}</td><td class="small num" style="white-space:nowrap">${dt(x.at)}</td><td class="code">${esc(x.code)}</td><td class="small">${x.by ? esc(user(x.by).name) : 'Hệ thống'}</td><td class="small">${esc(x.text)}</td><td class="small muted">${esc(x.what)}</td><td class="r">${fmt(x.amt)}</td></tr>`).join('')}
      </tbody></table></div>${rows.length > show.length ? `<p class="small faint">Hiển thị ${show.length}/${rows.length} dòng mới nhất. Thu hẹp bộ lọc hoặc xuất Excel để xem đủ.</p>` : `<p class="small faint" style="margin-top:10px">${rows.length} thao tác.</p>`}`
    : `<div class="card empty"><b>Không có thao tác nào khớp</b>Đổi từ khoá hoặc khoảng ngày.</div>`}`;
  }
  const d = summaryData(); const k = d.kpi; const tot = d.rs.reduce((a,r)=>a+r.amount,0) || 1;
  const kp = (l,v,s) => `<div class="kpi static"><span class="l">${l}</span><span class="v">${v}</span><span class="s">${s}</span></div>`;
  return `${head('Theo dõi','Lịch sử và báo cáo','Số liệu tính trên các phiếu lập trong khoảng ngày đang lọc.')}${tabs}
  ${filterBar('rflt', {cat:true, who:true, extra:`<span class="sp"></span><button class="btn sec sm" data-action="xsummary">Xuất Excel</button>`})}
  <div class="stack">
    <div class="g4">${kp('Lệnh đã lập', k.made, `${short(tot === 1 && !d.rs.length ? 0 : d.rs.reduce((a,r)=>a+r.amount,0))} đề nghị`)}${kp('Đã duyệt xong', k.approved, k.avgH == null ? 'Chưa đủ dữ liệu thời gian duyệt' : `Duyệt trung bình ${k.avgH < 24 ? k.avgH.toFixed(1).replace('.',',') + ' giờ' : (k.avgH/24).toFixed(1).replace('.',',') + ' ngày'}`)}${kp('Đã chi', k.paidN, `${short(k.paidAmt)}, thuế ${short(k.tax)}`)}${kp('Từ chối, huỷ', k.closed, `${k.invWait} khoản chờ hoá đơn`)}</div>
    ${d.adv.n ? `<div class="card" style="margin-bottom:16px"><div class="ch"><h2 class="card-title">Tạm ứng và hoàn ứng</h2><span class="card-note">${d.adv.n} khoản tạm ứng trong kỳ</span></div>
      <div class="tw"><table><thead><tr><th>Chỉ tiêu</th><th class="r">Số khoản</th><th class="r">Số tiền (đ)</th></tr></thead><tbody>
        <tr><td>Đã tạm ứng (đã chi)</td><td class="r">${d.adv.paidN}</td><td class="r">${fmt(d.adv.amt)}</td></tr>
        <tr><td>Đã quyết toán</td><td class="r">${d.adv.done}</td><td class="r">${fmt(d.adv.spent)}</td></tr>
        <tr><td>Chờ quyết toán hoặc đang duyệt</td><td class="r">${d.adv.pendingN}</td><td class="r">—</td></tr>
        <tr><td>Quá hạn quyết toán</td><td class="r" style="${d.adv.overdueN ? 'color:var(--critical)' : ''}">${d.adv.overdueN}</td><td class="r">—</td></tr>
        <tr><td>Chi thêm sau quyết toán</td><td class="r">—</td><td class="r">${fmt(d.adv.extra)}</td></tr>
        <tr><td>Hoàn ứng thừa đã nhận</td><td class="r">—</td><td class="r">${fmt(d.adv.refundDone)}</td></tr>
        <tr><td>Hoàn ứng thừa chờ nhận</td><td class="r">${d.adv.refundWaitN}</td><td class="r" style="${d.adv.refundWait ? 'color:var(--watch)' : ''}">${fmt(d.adv.refundWait)}</td></tr>
      </tbody></table></div>
      <p class="small faint" style="margin:10px 0 0">Hoàn ứng thừa là phần người đề xuất đã nhận tạm ứng nhưng chi thực tế thấp hơn, và phải trả lại công ty. Phần chi thêm là phần công ty chi bổ sung qua phiếu chênh lệch.</p></div>
      <div class="g2" style="margin-top:16px">${advTable('Theo người đề xuất', d.adv.byPerson)}${advTable('Theo phòng', d.adv.byDept)}</div>` : ''}
    <div class="g2">${d.tables.map(([name, rows]) => `<div class="card"><div class="ch"><h2 class="card-title">${name}</h2><span class="card-note">${rows.length} dòng</span></div>
      ${rows.length ? `<div class="tw"><table><thead><tr><th class="r">STT</th><th>${name.replace('Theo ','').replace(/^./, c => c.toUpperCase())}</th><th class="r">Số lệnh</th><th class="r">Đề nghị (đ)</th><th class="r">Đã chi (đ)</th><th class="r">Tỷ trọng</th></tr></thead><tbody>
        ${rows.slice(0,12).map((x,i) => `<tr><td class="r faint">${i+1}</td><td class="small">${esc(x.k)}</td><td class="r">${x.n}</td><td class="r">${fmt(x.amt)}</td><td class="r">${fmt(x.paid)}</td><td class="r">${(x.amt/tot*100).toFixed(1).replace('.',',')}%</td></tr>`).join('')}</tbody></table></div>${rows.length > 12 ? `<p class="small faint" style="margin:8px 0 0">Còn ${rows.length-12} dòng trong file Excel.</p>` : ''}` : '<div class="empty" style="padding:16px"><b>Chưa có dữ liệu</b></div>'}</div>`).join('')}</div>
  </div>`;
}

/* ===================== Sự kiện mở rộng ===================== */
function refocus(sel){ const el = $(sel); if (el) { el.focus(); try { const n = el.value.length; el.setSelectionRange(n,n); } catch (e) {} } }
document.addEventListener('click', e => {
  const t = e.target.closest('button'); if (!t) return; const d = t.dataset; const a = d.action;
  if (d.fltclear) { UI[d.fltclear] = {q:'', from:'', to:'', pm:'', cat:'', who:''}; renderMain(); return; }
  if (d.vpick) { readForm(); UI.form.vendorId = d.vpick; UI.form.payeeType = 'vendor'; UI.vq = ''; renderMain(); return; }
  if (d.rtab) { UI.rtab = d.rtab; renderMain(); return; }
  if (d.bpm != null) { UI.bpm = d.bpm; renderMain(); return; }
  if (d.rmline != null) { readForm(); UI.form.lines.splice(+d.rmline,1); if (!UI.form.amountManual) UI.form.amount = linesSum(UI.form.lines,'total'); renderMain(); return; }
  if (a === 'addline') { readForm(); const f = UI.form; f.lines.push(blankLine(f.evidence === 'bill' || f.evidence === 'sale' ? 'kct' : (f.lines[f.lines.length-1] || {}).rate)); renderMain(); refocus(`[data-ln="${f.lines.length-1}"][data-k="inv"]`); return; }
  if (a === 'pastelines') { readForm(); pasteModal(); return; }
  if (a === 'xmldemo') { readForm(); const p = parseInvoiceXml(SAMPLE_XML); const f = UI.form; applyInvoice(p, 'HD_00214586.xml'); const dk = addBlob(new File([SAMPLE_XML], 'HD_00214586.xml', {type:'text/xml'})); if (dk) f.files.push({name:dk, kind:'vat'}); renderMain(); return; }
  if (a === 'newvendor') { readForm(); vendorModal(null, {ctx:'form'}); return; }
  if (a === 'dnewvendor') { UI.dprefill = formPrefill('#directForm'); vendorModal(null, {ctx:'direct'}); return; }
  if (d.vnew) { vendorModal(null, {ctx:'dir'}); return; }
  if (d.vedit) { vendorModal(d.vedit, {ctx:'dir'}); return; }
  if (d.vdel) { vendorDelete(d.vdel); return; }
  if (a === 'direct') { directModal(); return; }
  if (d.verify) { const v = vendor(d.verify); v.verified = true; v.verifiedBy = S.me; S.reqs.forEach(r => { if (r.payee.vendorId === v.id && ['pending','approved','in_batch'].includes(r.status)) log(r, S.me, `Xác minh STK ${v.bank} ${v.acc} của ${v.name}`); }); save(); toast(`Đã xác minh STK của ${v.name}`); UI.errors = {}; render(); return; }
  if (d.dlunc) { downloadKey(d.dlunc); return; }
  if (d.addinv) { invModal(d.addinv); return; }
  if (d.remindinv) { const r = req(d.remindinv); log(r, S.me, `Nhắc ${user(r.requester).name} bổ sung hoá đơn, hạn ${dOnly(r.invDue)}`); r.invReminded = now(); save(); toast(`Đã nhắc ${user(r.requester).name}`); render(); return; }
  if (d.refundgot) { const r = req(d.refundgot); r.refundBack.status = 'received'; r.refundBack.receivedAt = now(); log(r, S.me, `Đã nhận NCC hoàn ${fmt(r.refundBack.amount)} đ`); save(); toast('Đã ghi nhận tiền hoàn'); render(); return; }
  if (d.wrongpay) { wrongModal(d.wrongpay); return; }
  if (d.cancelorder) { const r = req(d.cancelorder); askReason(`Huỷ lệnh ${r.code || 'nháp'}?`, r.status === 'in_batch' ? 'Khoản này sẽ được rút khỏi đợt chi và dừng ở trạng thái Đã huỷ.' : 'Lệnh dừng ở trạng thái Đã huỷ, không chi nữa.', 'Huỷ lệnh', reason => { cancelOrder(r, S.me, reason); toast(`Đã huỷ ${r.code}`); }); return; }
  if (d.dropline) { const b = batch(UI.id); const r = req(d.dropline); const kill = d.kind === 'cancelled';
    askReason(kill ? `Huỷ ${r.code} khỏi đợt?` : `Gỡ ${r.code} khỏi đợt?`, kill ? 'Lệnh bị huỷ hẳn, không chi nữa.' : 'Khoản về lại hàng chờ thanh toán để đưa vào đợt khác.', kill ? 'Huỷ khoản' : 'Gỡ khỏi đợt', reason => {
      if (kill) cancelOrder(r, S.me, reason); else { dropLine(b, r.id, S.me, reason, 'removed'); r.status = 'approved'; r.batchId = null; log(r, S.me, `Gỡ khỏi đợt ${b.code}: ${reason}. Về lại chờ thanh toán`); }
      toast(kill ? `Đã huỷ ${r.code}` : `Đã gỡ ${r.code} khỏi ${b.code}`); }); return; }
  if (d.cancelbatch) { const b = batch(d.cancelbatch); askReason(`Huỷ đợt ${b.code}?`, 'Các khoản chưa chi về lại hàng chờ thanh toán.', 'Huỷ đợt', reason => { cancelBatch(b, S.me, reason); toast(`Đã huỷ ${b.code}`); }); return; }
  if (d.copyrow) { const b = batch(UI.id); copyText(bankRow(b, d.copyrow, b.reqIds.indexOf(d.copyrow)).join('\t'), 'Đã sao chép dòng'); return; }
  if (d.xbank) { const b = batch(d.xbank); const ck = transferIds(b);
    exportXlsx(`${b.code}_import_ngan_hang.xlsx`, [
      {name:'CHUYEN KHOAN', rows:[BANK_HEAD].concat(ck.map((id,i) => { const row = bankRow(b, id, i); return row; }))},
      {name:'TAT CA KHOAN', rows:[['STT','MA PHIEU','NGUOI THU HUONG','HINH THUC','SO TIEN','TRANG THAI','NGAY CHI','TEN FILE UNC','UNC DA DINH KEM']].concat(b.reqIds.map((id,i) => { const r = req(id), ln = b.lines[id]; return [i+1, r.code, payeeOf(r).accName, bankText(PAYM[pmOf(r)]), r.amount, ln.paidAt ? 'DA CHI' : 'CHUA CHI', ln.paidAt ? dt(ln.paidAt) : '', uncName(b,id), ln.unc.join(', ')]; }))}
    ]); return; }
  if (d.zunc) { const b = batch(d.zunc); const items = [];
    b.reqIds.forEach(id => { const ln = b.lines[id]; if (!ln.paidAt) return; ln.unc.forEach((n,k) => { const base = uncName(b,id); const ext = (fname(n).match(/\.[a-z0-9]+$/i) || ['.pdf'])[0]; items.push({name: base.replace(/\.pdf$/, (k ? '_' + (k+1) : '') + ext), key: n}); }); });
    zipKeys(`UNC_${b.code}.zip`, items); return; }
  if (d.selck) { const b = batch(d.selck); const ck = b.reqIds.filter(id => !b.lines[id].paidAt && pmOf(req(id)) === 'transfer'); const all = ck.every(id => UI.paySel[id]); ck.forEach(id => UI.paySel[id] = !all); renderMain(); return; }
  if (a === 'xlist') { exportXlsx(`De_xuat_chi_${localDate()}.xlsx`, [{name:'De xuat chi', rows:reqExportRows(UI.listShown || [])}]); return; }
  if (a === 'zuncs') { const items = []; (UI.listShown || []).forEach(id => { const r = req(id); (r.unc || []).forEach((n,k) => items.push({name: `${r.code}_${k ? (k+1) + '_' : ''}${fname(n)}`, key: n})); }); zipKeys(`UNC_de_xuat_${localDate()}.zip`, items); return; }
  if (a === 'xbatches') { const bs = S.batches.filter(b => inRange(b.createdAt, UI.bflt));
    exportXlsx(`Dot_chi_${localDate()}.xlsx`, [{name:'Dot chi', rows:[['STT','Mã đợt','Ngày giờ lập','Người lập','Hẹn thanh toán','Số khoản','Tổng tiền','Trạng thái','Người duyệt','Duyệt thay','Ngày duyệt','Ngày chi xong']].concat(bs.map((b,i) => [i+1, b.code, dt(b.createdAt), user(b.createdBy).name, dOnly(b.payDate+'T00:00'), b.reqIds.length, batchTotal(b), BSTATUS[b.status][0], b.approvedBy ? user(b.approvedBy).name : '', b.approvedFor ? user(b.approvedFor).name : '', b.approvedAt ? dt(b.approvedAt) : '', b.paidAt ? dt(b.paidAt) : '']))},
      {name:'Chi tiet khoan', rows:[['Mã đợt','STT','Mã phiếu','Nội dung','Người thụ hưởng','Hình thức','Số tiền','Ngày chi','UNC']].concat(bs.flatMap(b => b.reqIds.map((id,i) => { const r = req(id), ln = b.lines[id]; return [b.code, i+1, r.code, r.title, payeeOf(r).name, PAYM[pmOf(r)], r.amount, ln.paidAt ? dt(ln.paidAt) : '', ln.unc.join(', ')]; })))}]); return; }
  if (a === 'xhistory') { exportXlsx(`Lich_su_thao_tac_${localDate()}.xlsx`, [{name:'Lich su', rows:[['STT','Thời gian','Mã','Người thực hiện','Thao tác','Phiếu, đợt','Số tiền']].concat(historyRows().map((x,i) => [i+1, dt(x.at), x.code, x.by ? user(x.by).name : 'Hệ thống', x.text, x.what, x.amt]))}]); return; }
  if (a === 'xsummary') { const d2 = summaryData(); const k = d2.kpi;
    exportXlsx(`Bao_cao_tong_hop_${localDate()}.xlsx`, [{name:'Tong quan', rows:[['Chỉ tiêu','Giá trị'],['Từ ngày', UI.rflt.from || 'Tất cả'],['Đến ngày', UI.rflt.to || 'Tất cả'],['Lệnh đã lập', k.made],['Đã duyệt xong', k.approved],['Đã chi (số lệnh)', k.paidN],['Đã chi (số tiền)', k.paidAmt],['Trong đó thuế', k.tax],['Từ chối, huỷ', k.closed],['Chờ hoá đơn', k.invWait],['Thời gian duyệt trung bình (giờ)', k.avgH == null ? '' : Math.round(k.avgH*10)/10]]}]
      .concat(d2.tables.map(([name, rows]) => ({name, rows:[['STT', name.replace('Theo ',''), 'Số lệnh','Đề nghị','Đã chi']].concat(rows.map((x,i) => [i+1, x.k, x.n, x.amt, x.paid]))})))
      .concat([{name:'Chi tiet phieu', rows:reqExportRows(d2.rs.map(r => r.id))}, {name:'Tam ung', rows:[['Mã phiếu','Nội dung','Người đề xuất','Phòng','Số đã tạm ứng','Hạn quyết toán','Trạng thái quyết toán','Số chi thực tế','Chênh lệch (chi thêm + / hoàn thừa -)','Hoàn thừa phải trả','Ngày nhận hoàn']].concat(d2.adv.rows.map(r => { const sp = r.settleLines ? linesSum(r.settleLines,'total') : 0; const rf = r.settle && r.settle.refund; return [r.code, r.title, user(r.requester).name, dept(r.dept).name, r.amount, r.settleDue ? dOnly(r.settleDue) : '', ADV_LABEL[advState(r)] || '', sp || '', sp ? sp - r.amount : '', rf ? rf.amount : '', rf && rf.receivedAt ? dt(rf.receivedAt) : '']; }))}, {name:'Tam ung theo nguoi', rows:advGroupRows(d2.adv.byPerson, 'Người đề xuất')}, {name:'Tam ung theo phong', rows:advGroupRows(d2.adv.byDept, 'Phòng')}])); return; }
});
document.addEventListener('change', async e => {
  const t = e.target; const d = t.dataset;
  if (d.flt && t.type !== 'search') { const [k, f] = d.flt.split('.'); UI[k][f] = t.value; if (k === 'qflt') UI.batchSel = {}; renderMain(); return; }
  if (d.fkind != null) { UI.form.files[+d.fkind].kind = t.value; return; }
  if (d.recat) { const r = req(d.recat); const old = r.cat; r.cat = t.value; log(r, S.me, `Phân loại lại hạng mục: ${old} → ${r.cat}`); save(); toast('Đã phân loại lại hạng mục'); renderMain(); return; }
  if (d.bdate) { const b = batch(d.bdate); if (!t.value) return; const old = b.payDate; b.payDate = t.value; blog(b, S.me, `Đổi ngày hẹn thanh toán ${dOnly(old+'T00:00')} → ${dOnly(b.payDate+'T00:00')}`); b.reqIds.forEach(id => log(req(id), S.me, `Hẹn thanh toán đổi sang ${dOnly(b.payDate+'T00:00')}`)); save(); toast('Đã đổi ngày hẹn thanh toán'); render(); return; }
  if (d.bselall) { queueRows().forEach(r => UI.batchSel[r.id] = t.checked); UI.errors = {}; renderMain(); return; }
  if (d.pselall) { const b = batch(d.pselall); b.reqIds.forEach(id => { if (!b.lines[id].paidAt && pmOf(req(id)) === 'transfer') UI.paySel[id] = t.checked; }); renderMain(); return; }
  if (t.id === 'payDateNew') { UI.payDateNew = t.value; return; }
  if ((d.files != null || d.invfiles != null) && t.closest('#reqForm')) {
    readForm(); const f = UI.form; let parsed = 0; const inv = d.invfiles != null; let other = 0;
    for (const file of [...t.files]) {
      const key = addBlob(file); if (key) f.files.push({name:key, kind: inv ? (f.evidence === 'sale' ? 'sale' : 'vat') : guessKind(file.name)});
      if (inv && !/\.xml$/i.test(file.name)) other++;
      if (/\.xml$/i.test(file.name)) { try { const p = parseInvoiceXml(await file.text()); if (p && p.lines.length) { applyInvoice(p, file.name); parsed++; } else toast(`Không đọc được ${file.name}, nhập tay bảng tiền và thuế.`); } catch (err) {} }
    }
    if (inv && f.evidence !== 'vat' && f.evidence !== 'sale') f.evidence = 'vat';
    if (inv && other && !parsed) toast('Đã đính kèm hoá đơn. File PDF hoặc ảnh chưa tự đọc được, nhập số tiền và thuế theo hoá đơn.');
    renderMain(); return; }
});
document.addEventListener('input', e => {
  const t = e.target; const d = t.dataset;
  if (d.acc != null) { const v = t.value.replace(/[^0-9a-zA-Z]/g,'').toUpperCase(); if (v !== t.value) t.value = v; }
  if (d.upper != null) { const p = t.selectionStart; const nv = noAccent(t.value).toUpperCase().replace(/[^A-Z0-9 .,\-\/]/g,''); if (nv !== t.value) { t.value = nv; try { t.setSelectionRange(p,p); } catch (err) {} } }
  if (d.vq != null) { UI.vq = t.value; const box = $('#vres'); if (box) box.innerHTML = vendorResults(); }
  if (d.vdq != null) { UI.vdq = t.value; renderMain(); refocus('[data-vdq]'); return; }
  if (d.vfilter != null) { const sel = t.closest('form').querySelector('[name=vendorId]'); const m = t.value.trim() ? vendorMatches(t.value) : S.vendors.filter(x => !x.deleted);
    sel.innerHTML = m.length ? m.map(x => `<option value="${x.id}">${esc(x.name)} — ${esc(x.bank)} ${esc(x.acc)}${x.verified?'':' (chưa xác minh)'}</option>`).join('') : '<option value="">Không tìm thấy</option>'; }
  if (d.flt && t.type === 'search') { const [k, f] = d.flt.split('.'); UI[k][f] = t.value; renderMain(); refocus(`[data-flt="${d.flt}"]`); }
});
document.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.dataset && e.target.dataset.vq != null) { e.preventDefault(); const m = vendorMatches(e.target.value); if (m.length) { readForm(); UI.form.vendorId = m[0].id; UI.vq = ''; renderMain(); } } });
document.addEventListener('submit', async e => {
  const f = e.target; const fd = new FormData(f); const g = k => (fd.get(k) || '').toString().trim();
  if (f.id === 'dueForm') { const r = req(UI.dueFor); const v = g('settleDue');
    if (!r) return; if (!v || v < localDate()) return mErr('Chọn hạn quyết toán, không được ở quá khứ.');
    const old = r.settleDue; r.settleDue = new Date(v + 'T23:59:59').getTime();
    log(r, S.me, `Đổi hạn quyết toán tạm ứng: ${dOnly(old)} → ${dOnly(r.settleDue)}`); save(); closeModal(); toast('Đã đổi hạn quyết toán'); renderMain(); return; }
  if (f.id === 'settleForm') { const adv = req(UI.settleFor); if (!adv) return;
    const rows = UI.settleRows.filter(x => digits(x.total) || String(x.inv || '').trim());
    if (!rows.length || rows.some(x => !digits(x.total))) return mErr('Mỗi dòng cần thành tiền lớn hơn 0.');
    if (rows.some(x => !String(x.inv || '').trim())) return mErr('Điền số hoá đơn cho từng dòng.');
    if (!UI.settleFiles.length) return mErr('Đính kèm ít nhất một hoá đơn hoặc chứng từ chi.');
    if (!(await flushUploads(UI.settleFiles))) return;
    const msg = submitSettle(adv, rows, UI.settleFiles.slice(), S.me); save(); closeModal(); toast(msg); renderMain(); return; }
  if (f.id === 'renameForm') { const k = UI.renameKind || {}; const name = g('name');
    if (name.length < 2) return mErr('Nhập tên, tối thiểu 2 ký tự.');
    const taken = list => list.some(x => x.toLowerCase() === name.toLowerCase());
    if (k.kind === 'cat') { const old = S.cats[k.gi][1][k.ci];
      if (name !== old && (name === REFUND_CAT || taken(S.cats.flatMap(x => [x[0], ...x[1]])))) return mErr(`“${name}” đã có.`);
      S.cats[k.gi][1][k.ci] = name; S.reqs.forEach(r => { if (r.cat === old) r.cat = name; });
      save(); closeModal(); toast(`Đã đổi “${old}” thành “${name}”`); renderMain(); return; }
    const oldG = S.cats[k.gi][0]; if (name !== oldG && taken(S.cats.map(x => x[0]))) return mErr(`Nhóm “${name}” đã có.`);
    S.cats[k.gi][0] = name; save(); closeModal(); toast(`Đã đổi nhóm thành “${name}”`); renderMain(); return; }
  if (f.id === 'reasonForm') { const reason = g('reason'); if (reason.length < 5) { $('#reasonErr').textContent = 'Ghi lý do, tối thiểu 5 ký tự.'; return; }
    const fn = reasonFn; reasonFn = null; closeModal(); if (fn) fn(reason); save(); if (UI.view === 'batch' && batch(UI.id) && batch(UI.id).status === 'cancelled' && !batch(UI.id).reqIds.length) go('batches'); else render(); return; }
  if (f.id === 'vendorForm') { const ctx = UI.vctx || {id:null, ctx:'form'};
    const v = {name:g('name'), tax:g('tax'), bank:g('bank'), acc:g('acc'), accName:bankText(g('accName'))};
    if (v.name.length < 3 || !v.bank || !v.acc || !v.accName) return mErr('Nhập đủ tên NCC, ngân hàng, số tài khoản và tên tài khoản.');
    if (!/^[0-9A-Z]{4,24}$/.test(v.acc)) return mErr('Số tài khoản gồm 4 đến 24 ký tự số hoặc chữ.');
    if (v.tax && !/^\d{10}(-?\d{3})?$/.test(v.tax)) return mErr('Mã số thuế gồm 10 số, hoặc 10 số kèm 3 số chi nhánh.');
    if (S.vendors.some(x => !x.deleted && x.id !== ctx.id && x.acc === v.acc && x.bank.toLowerCase() === v.bank.toLowerCase())) return mErr('Số tài khoản này đã có trong danh bạ NCC.');
    let vid;
    if (ctx.id) { const old = vendor(ctx.id); const changed = old.bank !== v.bank || old.acc !== v.acc || old.accName !== v.accName; Object.assign(old, v);
      if (changed) { old.verified = false; old.verifiedBy = null; S.reqs.forEach(r => { if (r.payee && r.payee.vendorId === old.id && ['pending','approved','in_batch'].includes(r.status)) log(r, S.me, `STK của ${old.name} đổi, cần xác minh lại`); }); }
      vid = old.id; }
    else { const nv = Object.assign({id:uid('v'), verified:false, createdBy:S.me, deleted:false}, v); S.vendors.push(nv); vid = nv.id; }
    save(); closeModal();
    if (ctx.ctx === 'form') { if (UI.form) { UI.form.payeeType = 'vendor'; UI.form.vendorId = vid; } toast(`Đã thêm ${v.name}, STK chờ kế toán xác minh`); renderMain(); }
    else if (ctx.ctx === 'direct') { directModal(Object.assign({}, UI.dprefill || {}, {vendorId: vid})); toast(ctx.id ? `Đã lưu ${v.name}` : `Đã thêm ${v.name} vào danh bạ NCC`); }
    else { toast(ctx.id ? `Đã lưu ${v.name}` : `Đã thêm ${v.name}, STK chờ kế toán xác minh`); renderMain(); }
    return; }
  if (f.id === 'directForm') { const total = digits(g('total')), tax = digits(g('tax')); const title = g('title');
    if (title.length < 6) return mErr('Ghi rõ nội dung chi, tối thiểu 6 ký tự.'); if (!g('vendorId')) return mErr('Chọn nhà cung cấp.'); if (!total) return mErr('Nhập thành tiền.'); if (tax >= total) return mErr('Tiền thuế phải nhỏ hơn thành tiền.');
    const names = [...(f.querySelector('[name=files]').files || [])].map(x => addBlob(x)).filter(Boolean);
    const pre = total - tax; const rk = tax ? String(Math.round(tax / pre * 100)) : 'kct';
    if (!(await flushUploads(names))) return;
    const r = newReq({title, amount:total, dept:g('dept'), cat:g('cat'), evidence: g('inv') ? 'vat' : 'later', payMethod:g('payMethod'), payee:{type:'vendor', vendorId:g('vendorId')}, files:names, requester:S.me, direct:true,
      lines:[{inv:g('inv'), desc:title, pre, rate:RATES[rk] ? rk : '10', tax, total}]});
    r.code = nextCode('A'); r.status = 'approved'; r.approvedAt = now(); r.submittedAt = now(); log(r, S.me, 'Kế toán lập khoản chi trực tiếp, không qua đề xuất'); S.reqs.push(r); UI.batchSel[r.id] = true;
    save(); closeModal(); toast(`Đã thêm ${r.code} vào chờ thanh toán`); render(); return; }
  if (f.id === 'invForm') { const r = req(UI.invFor); const names = [...(f.querySelector('[name=files]').files || [])].map(x => addBlob(x)).filter(Boolean); const inv = g('inv');
    if (!inv) return mErr('Nhập số hoá đơn.'); if (!names.length) return mErr('Đính kèm file hoá đơn.');
    if (!(await flushUploads(names))) return;
    const tax = digits(g('tax')); const nos = inv.split(',').map(x => x.trim()).filter(Boolean);
    r.lines.forEach((l,i) => { if (!l.inv) l.inv = nos[Math.min(i, nos.length-1)]; });
    const cur = linesSum(r.lines,'tax'); if (tax !== cur && r.lines.length) { const l = r.lines[r.lines.length-1]; l.tax += tax - cur; l.pre = l.total - l.tax; }
    names.forEach(n => r.files.push({name:n, kind:'vat'})); r.invDone = {at:now(), by:S.me, inv:nos, files:names};
    log(r, S.me, `Bổ sung hoá đơn ${nos.join(', ')}${tax !== cur ? `, chỉnh tiền thuế ${fmt(cur)} → ${fmt(tax)} đ` : ''}`); save(); closeModal(); toast('Đã bổ sung hoá đơn'); render(); return; }
  if (f.id === 'pasteForm') { const ls = parsePaste(fd.get('text') || ''); if (!ls.length) return mErr('Không đọc được dòng nào. Kiểm tra thứ tự cột và dấu tab giữa các cột.');
    const fm = UI.form; if (fm.lines.every(l => !l.total && !l.inv)) fm.lines = []; fm.lines = fm.lines.concat(ls); if (!fm.amountManual) fm.amount = linesSum(fm.lines,'total');
    closeModal(); toast(`Đã thêm ${ls.length} dòng, tổng ${fmt(linesSum(ls,'total'))} đ`); renderMain(); return; }
  if (f.id === 'wrongForm') { const r = req(UI.wrongFor); const amount = digits(g('amount')); if (!amount || amount > r.amount) return mErr('Số tiền hoàn phải lớn hơn 0 và không quá số đã chi.'); if (g('note').length < 5) return mErr('Ghi diễn giải, tối thiểu 5 ký tự.');
    r.refundBack = {status:'waiting', amount, note:g('note'), at:now(), by:S.me}; log(r, S.me, `Báo chuyển nhầm, chờ NCC hoàn ${fmt(amount)} đ: ${g('note')}`); save(); closeModal(); toast('Đã ghi nhận chờ NCC hoàn'); render(); return; }
});

/* ===================== Khởi động ===================== */


/* ===================== Lưu trữ và đăng nhập (Supabase) ===================== */
let SESSION_EMAIL = null, STATE_VERSION = 0, saveTimer = null, pushing = false, dirty = false;
const $root = () => document.getElementById('main');

/* Trạng thái mới cho công ty: chỉ có cấu trúc mặc định, không có người dùng, NCC hay phiếu mẫu. */
function freshState(){
  return {
    me: null, guideDone: {},
    depts: [{id:'OPS',name:'Vận hành'},{id:'SAL',name:'Kinh doanh'},{id:'FIN',name:'Tài chính – Kế toán'},{id:'TEC',name:'Công nghệ'},{id:'BOD',name:'Ban điều hành'}],
    users: [], vendors: [], cats: JSON.parse(JSON.stringify(CATS0)), userBanks: {},
    config: {
      version: 1, updatedAt: now(), updatedBy: null,
      rules: [{id:'rDEF', def:true, depts:[], people:[], bands:[
        {id:'bDEF1', max:5000000, levels:[{approvers:[], mode:'any'}]},
        {id:'bDEF2', max:null, levels:[{approvers:[], mode:'any'}]}]}],
      batch: {bands:[{id:'bB1', max:null, levels:[{approvers:[], mode:'any'}]}]},
      blockSelf: true, fallback: [], noInvoice: {on:true, approvers:[]}, cashLimit: 5000000
    },
    reqs: [], batches: [], delegations: [], seq: {A:0, PB:0}
  };
}
function normalizeState(x){
  x.guideDone = x.guideDone || {}; x.fileNames = x.fileNames || {}; x.users = x.users || []; x.vendors = x.vendors || []; x.reqs = x.reqs || [];
  x.batches = x.batches || []; x.delegations = x.delegations || []; x.userBanks = x.userBanks || {}; x.seq = x.seq || {A:0, PB:0};
  x.config.batch = x.config.batch || {bands:[{id:'bB1', max:null, levels:[{approvers:[], mode:'any'}]}]};
  x.config.fallback = x.config.fallback || []; x.config.noInvoice = x.config.noInvoice || {on:true, approvers:[]};
  return x;
}
function fatal(msg){
  $root().innerHTML = `<div class="card" style="max-width:520px;margin:48px auto"><h1 style="margin:0 0 8px;font-size:18px">Không mở được ứng dụng</h1><p class="muted" style="margin:0 0 16px">${esc(msg)}</p><button class="btn sec" data-action="logout">Đăng xuất</button></div>`;
}
function renderLogin(msg){
  document.body.classList.add('auth');
  const side = document.getElementById('side'); if (side) side.innerHTML = '';
  $root().innerHTML = `<div class="auth-wrap"><div class="auth-card">
    <div class="auth-brand"><span class="logo logo-lg" role="img" aria-label="Xperise"></span><span class="auth-app">Spend Ops · Đề xuất chi và phê duyệt</span></div>
    <h1>Đăng nhập</h1>
    <p class="sub">Dùng tài khoản do quản trị hệ thống cấp cho bạn.</p>
    <form id="loginForm" class="stack" novalidate>
      <label class="f">Email<input class="in" name="email" type="email" required autocomplete="username" placeholder="ten@xperise.com"></label>
      <label class="f">Mật khẩu<input class="in" name="password" type="password" required autocomplete="current-password"></label>
      ${msg ? `<div class="alert neg">${esc(msg)}</div>` : ''}
      <button class="btn pri" type="submit">Đăng nhập</button>
    </form>
    <p class="auth-foot">Không có đăng ký công khai. Quên mật khẩu? Liên hệ quản trị hệ thống.</p>
  </div></div>`;
}
function renderNoAccess(){
  fatal(`Tài khoản ${SESSION_EMAIL} chưa được thêm vào hệ thống. Nhờ quản trị thêm email này ở mục Người dùng.`);
}

async function loadState(){
  const { data, error } = await SB.from('app_state').select('data, version').eq('id', 'main').maybeSingle();
  if (error) { fatal('Không tải được dữ liệu từ máy chủ: ' + error.message); return false; }
  if (!data) {
    S = freshState();
    const ins = await SB.from('app_state').insert({ id:'main', data:S, version:1 });
    if (ins.error) { fatal('Không tạo được dữ liệu ban đầu: ' + ins.error.message); return false; }
    STATE_VERSION = 1; return true;
  }
  S = normalizeState(data.data); STATE_VERSION = data.version; return true;
}

/* Tài khoản đầu tiên: khi chưa có người dùng nào, người đăng nhập đầu tiên thành quản trị và được gán vào mọi bước duyệt. */
function bootstrapAdmin(email){
  const id = uid('u');
  S.users.push({id, name: email.split('@')[0], title:'Quản trị hệ thống', dept:'TEC', acct:true, signer:true, admin:true, active:true, email});
  S.config.rules.forEach(rl => rl.bands.forEach(b => b.levels.forEach(l => { l.approvers = [id]; })));
  S.config.batch.bands.forEach(b => b.levels.forEach(l => { l.approvers = [id]; }));
  S.config.fallback = [id]; S.config.noInvoice.approvers = [id]; S.config.updatedBy = id; S.config.updatedAt = now();
  save();
  return S.users[S.users.length - 1];
}

async function boot(){
  if (!SB) { fatal('Chưa cấu hình Supabase. Thêm VITE_SUPABASE_URL và VITE_SUPABASE_ANON_KEY (xem README).'); return; }
  const { data: { session } } = await SB.auth.getSession();
  if (!session) { renderLogin(); return; }
  SESSION_EMAIL = (session.user.email || '').toLowerCase();
  if (!(await loadState())) return;
  let u = S.users.find(x => (x.email || '').toLowerCase() === SESSION_EMAIL && !x.deleted);
  if (!u && S.users.length === 0) u = bootstrapAdmin(SESSION_EMAIL);
  if (!u) { renderNoAccess(); return; }
  if (!u.active) { renderNoAccess(); return; }
  S.me = u.id;
  render();
  if (S.users.length === 1 && u.admin) toast('Đã tạo tài khoản quản trị đầu tiên. Thêm người dùng và kiểm tra cấu hình phê duyệt.');
}

/* Ghi lên Supabase có kiểm soát phiên bản: nếu có người khác vừa ghi, không ghi đè mà báo tải lại. */
async function pushState(){
  if (!S || !SB) return;
  if (pushing) { saveTimer = setTimeout(pushState, 500); return; }
  pushing = true;
  try {
    const payload = JSON.parse(JSON.stringify(S));
    const next = STATE_VERSION + 1;
    const { data, error } = await SB.from('app_state')
      .update({ data: payload, version: next, updated_at: new Date().toISOString() })
      .eq('id', 'main').eq('version', STATE_VERSION).select('version');
    if (error) throw error;
    if (!data || data.length === 0) {
      dirty = true;
      toast('Dữ liệu vừa được người khác cập nhật. Tải lại trang trước khi làm tiếp, để không ghi đè thay đổi của họ.');
    } else { STATE_VERSION = data[0].version; dirty = false; }
  } catch (err) {
    toast('Chưa lưu được lên máy chủ, hệ thống sẽ thử lại.');
    saveTimer = setTimeout(pushState, 3000);
  } finally { pushing = false; }
}
window.addEventListener('beforeunload', e => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });

document.addEventListener('submit', async e => {
  if (e.target.id !== 'loginForm') return;
  e.preventDefault();
  const fd = new FormData(e.target);
  const email = String(fd.get('email') || '').trim(), password = String(fd.get('password') || '');
  if (!email || !password) { renderLogin('Nhập email và mật khẩu.'); return; }
  const { error } = await SB.auth.signInWithPassword({ email, password });
  if (error) { renderLogin('Email hoặc mật khẩu chưa đúng.'); return; }
  boot();
});
document.addEventListener('click', async e => {
  const b = e.target.closest('button[data-action="logout"]'); if (!b) return;
  if (SB) await SB.auth.signOut();
  location.reload();
});

/* ===================== Dashboard quản trị (chỉ người được phân quyền) ===================== */
const canDash = u => !!u && !!(u.admin || u.dash);
function dashRange(p){
  const d = new Date();
  const start = p === 'month' ? new Date(d.getFullYear(), d.getMonth(), 1).getTime()
    : new Date(d.getFullYear(), d.getMonth(), d.getDate() - ((d.getDay() + 6) % 7)).getTime();
  return {start, label: p === 'month' ? 'tháng này' : 'tuần này'};
}
function dashData(){
  const D = UI.dash || {period:'week', dept:''}; const {start} = dashRange(D.period);
  const inDept = r => !D.dept || r.dept === D.dept;
  const live = S.reqs.filter(r => r.status !== 'draft' && inDept(r));
  const created = live.filter(r => (r.createdAt || 0) >= start && !['rejected','cancelled'].includes(r.status));
  const sum = (list, fn) => list.reduce((a,r) => a + (fn ? fn(r) : r.amount), 0);
  const paidP = S.reqs.filter(r => r.status === 'paid' && (r.paidAt || 0) >= start && inDept(r));
  const in7 = now() + 7*24*H;
  const toPay7 = live.filter(r => r.status === 'in_batch' && r.batchId && batch(r.batchId) && new Date(batch(r.batchId).payDate + 'T23:59:59').getTime() <= in7);
  const pending = live.filter(r => r.status === 'pending');
  const hasInv = r => r.evidence === 'vat' || r.evidence === 'sale';
  const grp = (list, keyFn) => {
    const m = {}; list.forEach(r => { const k = keyFn(r); const g = (m[k] = m[k] || {k, n:0, amt:0, has:0}); g.n++; g.amt += r.amount; if (hasInv(r)) g.has += r.amount; });
    return Object.values(m).sort((a,b) => b.amt - a.amt);
  };
  const createdAmt = sum(created);
  const doneP = live.filter(r => r.approvedAt && r.submittedAt && r.approvedAt >= start);
  const advOpen = S.reqs.filter(r => r.evidence === 'advance' && ['waiting','overdue','submitted','refund_waiting'].includes(advState(r)) && inDept(r));
  const invOver = S.reqs.filter(r => invPending(r) && r.invDue && now() > r.invDue && inDept(r));
  const invAll = S.reqs.filter(r => invPending(r) && inDept(r));
  return {
    period: D.period, label: dashRange(D.period).label, dept: D.dept,
    paid: {n: paidP.length, amt: sum(paidP)},
    due7: {n: toPay7.length, amt: sum(toPay7)},
    pending: {n: pending.length, amt: sum(pending)},
    created: {n: created.length, amt: createdAmt, hasAmt: sum(created, r => hasInv(r) ? r.amount : 0)},
    byDept: grp(created, r => (dept(r.dept) || {name:'Chưa gán phòng'}).name).map(g => ({...g, share: createdAmt ? g.amt / createdAmt : 0, noInv: g.amt - g.has})),
    byCat: grp(created, r => r.cat).map(g => ({...g, cov: g.amt ? g.has / g.amt : 0})),
    risk: {
      invOver: {n: invOver.length, amt: sum(invOver)},
      invAll: {n: invAll.length, amt: sum(invAll)},
      advOver: S.reqs.filter(r => advState(r) === 'overdue' && inDept(r)).length,
      advOpen: {n: advOpen.length, amt: sum(advOpen)},
      refundWait: sum(S.reqs.filter(r => r.settle && r.settle.refund && r.settle.refund.status === 'waiting' && inDept(r)), r => r.settle.refund.amount),
      vendorsUnv: S.vendors.filter(v => !v.deleted && !v.verified).length,
      returned: live.filter(r => r.status === 'returned').length,
      speedD: doneP.length ? doneP.reduce((a,r) => a + (r.approvedAt - r.submittedAt), 0) / doneP.length / (24*H) : null,
      speedN: doneP.length,
      covPct: createdAmt ? sum(created, r => hasInv(r) ? r.amount : 0) / createdAmt : null
    }
  };
}
function dashCard(title, big, note, tone){
  return `<div class="card" style="padding:16px 18px"><p class="eyebrow" style="margin:0 0 6px">${title}</p><div style="font-size:20px;font-weight:600;${tone === 'bad' ? 'color:var(--critical)' : ''}">${big}</div>${note ? `<p class="small muted" style="margin:6px 0 0">${note}</p>` : ''}</div>`;
}
function viewDash(){
  const D = UI.dash = UI.dash || {period:'week', dept:''};
  const k = dashData();
  const share = v => (v * 100).toFixed(1) + '%';
  const bar = (v, c) => `<div style="height:6px;background:var(--line,#eee);border-radius:3px;overflow:hidden"><i style="display:block;height:100%;width:${Math.min(100, v*100).toFixed(1)}%;background:${c||'var(--accent,#0f766e)'}"></i></div>`;
  const tabBtn = (v, t) => `<button role="tab" aria-selected="${D.period===v}" class="${D.period===v?'on':''}" data-dperiod="${v}">${t}</button>`;
  return `${head('Quản trị', 'Dashboard quản trị', `Dòng tiền, phân bổ chi phí và rủi ro, ${k.label}. Chỉ người được phân quyền xem.`,
    `<div class="row"><div class="tabs" role="tablist" style="margin:0">${tabBtn('week','Tuần')}${tabBtn('month','Tháng')}</div><button class="btn sec sm" data-dexport="1">${icon('download',14)} Xuất Excel</button></div>`)}
  <div class="card" style="padding:12px 16px;margin-bottom:16px"><div class="row"><span class="small faint">Phòng chịu chi phí</span><select class="in" data-ddept style="max-width:240px"><option value="">Tất cả phòng</option>${S.depts.map(d => `<option value="${d.id}" ${D.dept===d.id?'selected':''}>${esc(d.name)}</option>`).join('')}</select></div></div>

  <p class="eyebrow" style="margin:0 0 8px">Dòng tiền</p>
  <div class="g3" style="margin-bottom:16px">
    ${dashCard(`Đã chi ${k.label}`, money(k.paid.amt), `${k.paid.n} khoản đã chuyển`)}
    ${dashCard('Sẽ chi trong 7 ngày tới', money(k.due7.amt), `${k.due7.n} khoản trong đợt chi, đã duyệt`)}
    ${dashCard('Đang chờ duyệt', money(k.pending.amt), `${k.pending.n} đề xuất đang chờ người duyệt`)}
  </div>
  <div class="g3" style="margin-bottom:16px">
    ${dashCard(`Đề xuất mới ${k.label}`, money(k.created.amt), `${k.created.n} đề xuất, không tính đã huỷ và từ chối`)}
    ${dashCard('Tỉ lệ có hoá đơn', k.risk.covPct == null ? '—' : share(k.risk.covPct), 'Tính theo giá trị đề xuất mới trong kỳ')}
    ${dashCard('Tốc độ duyệt', k.risk.speedD == null ? '—' : k.risk.speedD.toFixed(1) + ' ngày', k.risk.speedN ? `Trung bình từ gửi đến duyệt, ${k.risk.speedN} đề xuất trong kỳ` : 'Chưa có đề xuất duyệt xong trong kỳ')}
  </div>

  <div class="card" style="margin-bottom:16px"><div class="ch"><h2 class="card-title">Chi theo phòng chịu chi phí · ${k.label}</h2><span class="card-note">${k.byDept.length} phòng</span></div>
    ${k.byDept.length ? `<div class="tw"><table><thead><tr><th>Phòng</th><th class="r">Số lệnh</th><th class="r">Giá trị (đ)</th><th style="width:22%">Tỉ trọng</th><th class="r">Có hoá đơn (đ)</th><th class="r">Chưa có hoá đơn (đ)</th></tr></thead><tbody>
    ${k.byDept.map(g => `<tr><td>${esc(g.k)}</td><td class="r">${g.n}</td><td class="r">${fmt(g.amt)}</td><td>${bar(g.share)}<span class="small faint">${share(g.share)}</span></td><td class="r">${fmt(g.has)}</td><td class="r" style="${g.noInv ? 'color:var(--critical)' : ''}">${g.noInv ? fmt(g.noInv) : '—'}</td></tr>`).join('')}
    </tbody></table></div>` : '<div class="empty" style="padding:16px"><b>Chưa có đề xuất trong kỳ</b></div>'}
  </div>

  <div class="card" style="margin-bottom:16px"><div class="ch"><h2 class="card-title">Chi theo hạng mục · ${k.label}</h2><span class="card-note">${k.byCat.length} hạng mục</span></div>
    ${k.byCat.length ? `<div class="tw"><table><thead><tr><th>Hạng mục</th><th class="r">Số lệnh</th><th class="r">Giá trị (đ)</th><th style="width:26%">Phủ hoá đơn</th></tr></thead><tbody>
    ${k.byCat.map(g => `<tr><td>${esc(g.k)}</td><td class="r">${g.n}</td><td class="r">${fmt(g.amt)}</td><td>${bar(g.cov, g.cov < 0.5 ? 'var(--critical)' : undefined)}<span class="small faint">${share(g.cov)}</span></td></tr>`).join('')}
    </tbody></table></div>` : '<div class="empty" style="padding:16px"><b>Chưa có đề xuất trong kỳ</b></div>'}
  </div>

  <p class="eyebrow" style="margin:0 0 8px">Rủi ro cần xử lý</p>
  <div class="g3" style="margin-bottom:16px">
    ${dashCard('Quá hạn bổ sung hoá đơn', `${k.risk.invOver.n} khoản`, `Tổng ${money(k.risk.invOver.amt)}. Khoản chi chưa có hoá đơn quá hạn nộp`, k.risk.invOver.n ? 'bad' : '')}
    ${dashCard('Tạm ứng quá hạn quyết toán', `${k.risk.advOver} khoản`, `Đang mở ${k.risk.advOpen.n} khoản tạm ứng, ${money(k.risk.advOpen.amt)}`, k.risk.advOver ? 'bad' : '')}
    ${dashCard('Hoàn ứng thừa chờ nhận', money(k.risk.refundWait), 'Người đề xuất cần hoàn lại công ty, kế toán chưa xác nhận')}
  </div>
  <div class="g3">
    ${dashCard('NCC chưa xác minh tài khoản', `${k.risk.vendorsUnv}`, 'Xác minh trong danh bạ NCC trước khi đưa vào đợt chi', k.risk.vendorsUnv ? 'bad' : '')}
    ${dashCard('Đề xuất đang bị trả lại', `${k.risk.returned}`, 'Người đề xuất cần bổ sung hoặc sửa')}
    ${dashCard('Chi phí đang chờ bổ sung hoá đơn', money(k.risk.invAll.amt), `${k.risk.invAll.n} khoản đã chi, chưa đủ hoá đơn (toàn bộ thời gian)`)}
  </div>`;
}
document.addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
  if (d.dperiod) { UI.dash = Object.assign({period:'week', dept:''}, UI.dash, {period: d.dperiod}); renderMain(); return; }
  if (d.dexport) {
    const k = dashData(); const pct = v => v == null ? '' : Math.round(v*1000)/10;
    exportXlsx(`Dashboard_quan_tri_${localDate()}.xlsx`, [
      {name:'Phong', rows:[['Phòng','Số lệnh','Giá trị (đ)','Tỉ trọng (%)','Có hoá đơn (đ)','Chưa có hoá đơn (đ)']].concat(k.byDept.map(g => [g.k, g.n, g.amt, pct(g.share), g.has, g.noInv]))},
      {name:'Hang muc', rows:[['Hạng mục','Số lệnh','Giá trị (đ)','Phủ hoá đơn (%)']].concat(k.byCat.map(g => [g.k, g.n, g.amt, pct(g.cov)]))},
      {name:'Rui ro', rows:[['Chỉ tiêu','Giá trị'],['Đã chi trong kỳ (đ)', k.paid.amt],['Sẽ chi 7 ngày tới (đ)', k.due7.amt],['Đang chờ duyệt (đ)', k.pending.amt],['Quá hạn bổ sung hoá đơn (khoản)', k.risk.invOver.n],['Tạm ứng quá hạn (khoản)', k.risk.advOver],['Hoàn ứng thừa chờ nhận (đ)', k.risk.refundWait],['NCC chưa xác minh', k.risk.vendorsUnv],['Đề xuất bị trả lại', k.risk.returned],['Tốc độ duyệt TB (ngày)', k.risk.speedD == null ? '' : Math.round(k.risk.speedD*10)/10]]}
    ]);
    return;
  }
});
document.addEventListener('change', e => {
  if (e.target.dataset.ddept != null) { UI.dash = Object.assign({period:'week', dept:''}, UI.dash, {dept: e.target.value}); renderMain(); }
});

/* Khởi động sau khi mọi khối đã khai báo. */
boot();

async function confirmPay(){ const P = UI.payDraft; P.date = $('#payDate').value || P.date;
    const iss = payIssues(); if (iss.length) { P.err = iss.map(esc).join('<br>'); drawPayModal(); return; }
    if (!(await flushUploads(P.files))) return;
    const b = batch(P.bid); const at = new Date(P.date + 'T' + new Date().toTimeString().slice(0,8)).getTime() || now();
    payLines(b, P.ids, S.me, P.mode === 'each' && P.ids.length > 1 ? id => [P.files[P.map[id]]] : P.files, at); P.ids.forEach(id => delete UI.paySel[id]); S.guideDone.pay = true; save(); closeModal();
    toast(b.status === 'paid' ? `Đã chuyển xong đợt ${b.code}` : `Đã đánh dấu ${P.ids.length} khoản đã chuyển`); render(); return; }
