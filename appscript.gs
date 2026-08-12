/************************************************************
 * 1. Existing doPost — DO NOT MODIFY (kept exactly as given)
 ************************************************************/
function doPost(e) {
  try {
    Logger.log("Incoming request:");
    Logger.log(e.postData ? e.postData.contents : "NO POST DATA RECEIVED");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sheet1");

    const data = JSON.parse(e.postData.contents);

    const row = [
      data.timestamp || "",
      data.eventTitle || "",
      data.registrationId || "",
      data.fullName || "",
      data.email || "",
      data.phone || "",
      data.isMember || "",
      data.htsMemberId || "",
      data.foodPreference || "",
      data.adults || "",
      data.kids || "",
      data.attendeeNames || "",
      "", // attendstatus column initially empty
      "" // paymentStatus column initially empty
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

//
const SHEET_NAME="Sheet1";
const ADMIN_PASSWORD="hts2025";

function doGet(e){
  try{
    const p=e.parameter||{},m=p.mode||"";
    if(m==="verify")return verify(p);
    if(m==="updatePayment")return payment(p);
    if(m==="dashboard")return dashboard(p);
    return json({status:"ERROR",message:"Invalid mode"});
  }catch(err){
    return json({status:"ERROR",message:String(err)});
  }
}

// VERIFY + CHECK-IN
function verify(p){
  if(String(p.password||"")!==ADMIN_PASSWORD)
    return json({status:"AUTH_FAILED",message:"Incorrect password"});

  const id=norm(p.registrationId),email=normEmail(p.email);
  if(!id&&!email)return json({status:"ERROR",message:"Registration ID or email required"});

  const d=data();
  if(!d)return json({status:"ERROR",message:"Sheet not found"});

  let x=id?find(d.rows,"registrationId",id):null;
  if(!x&&email)x=find(d.rows,"email",email);

  if(!x) return json({status:"ERROR",message:"Registration not found",email:email});

  const r = x.obj, row = x.index + 2, out = {
    registrationId: val(r, "registrationId"),
    email: val(r, "email"),
    fullName: val(r, "fullName"),
    phone: val(r, "phone"),
    isMember: val(r, "isMember"),
    htsMemberId: val(r, "htsMemberId"),
    foodPreference: val(r, "foodPreference"),
    adults: val(r, "adults"),
    kids: val(r, "kids"),
    attendeeNames: val(r, "attendeeNames"),
    paymentStatus: val(r, "paymentStatus"),
    attendStatus: val(r, "attendStatus")
  };

  const col=header(d.headers,"attendStatus");
  if(!col)return json({...out,status:"ERROR",message:"attendStatus column not found"});

  const cell=d.sheet.getRange(row,col),old=String(cell.getDisplayValue()).trim();

  if(old.toLowerCase()==="ok")
    return json({...out,status:"Already updated",message:"Attendance already updated",attendStatus:"OK"});

  try{
    cell.setValue("OK");
    SpreadsheetApp.flush();
  }catch(err){
    return json({...out,status:"ERROR",message:"Attendance update failed: "+err});
  }

  const now=String(cell.getDisplayValue()).trim();
  if(now.toLowerCase()!=="ok")
    return json({...out,status:"ERROR",message:"Attendance update failed",attendStatus:now});

  return json({...out,status:"Updated successfully",message:"Attendance updated successfully",attendStatus:"OK"});
}

// UPDATE PAYMENT
function payment(p){
  if(String(p.password||"")!==ADMIN_PASSWORD)
    return json({status:"AUTH_FAILED",message:"Incorrect password"});

  const id=norm(p.registrationId);
  if(!id)return json({status:"ERROR",message:"Registration ID required"});

  const d=data();
  if(!d)return json({status:"ERROR",message:"Sheet not found"});

  const x=find(d.rows,"registrationId",id);
  if(!x)return json({status:"ERROR",message:"Registration not found",registrationId:id});
  const r = x.obj, row = x.index + 2, out = {
    registrationId: val(r, "registrationId"),
    email: val(r, "email"),
    fullName: val(r, "fullName"),
    attendeeNames: val(r, "attendeeNames"),
    paymentStatus: val(r, "paymentStatus")
  };

  const col=header(d.headers,"paymentStatus");
  if(!col)return json({...out,status:"ERROR",message:"paymentStatus column not found"});

  const cell=d.sheet.getRange(row,col),old=String(cell.getDisplayValue()).trim();

  if(old.toLowerCase()==="ok")
    return json({...out,status:"Payment already updated",message:"Payment already updated",paymentStatus:"OK"});

  try{
    cell.setValue("OK");
    SpreadsheetApp.flush();
  }catch(err){
    return json({...out,status:"ERROR",message:"Payment update failed: "+err});
  }

  const now=String(cell.getDisplayValue()).trim();
  if(now.toLowerCase()!=="ok")
    return json({...out,status:"ERROR",message:"Payment update failed",paymentStatus:now});

  return json({...out,status:"Payment updated successfully",message:"Payment updated successfully",paymentStatus:"OK"});
}

// SHEET DATA
function data(){
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if(!s)return null;

  const v=s.getDataRange().getDisplayValues();
  const h=v[0].map(x=>String(x).trim().toLowerCase());
  const rows=v.slice(1).map((a,i)=>{
    const o={};h.forEach((x,j)=>o[x]=String(a[j]||"").trim());
    return{index:i,obj:o};
  });

  return{sheet:s,headers:h,rows:rows};
}

// FIND RECORD
function find(rows,key,value){
  const target = key === "email" ? normEmail(value) : norm(value);
  return rows.find(x => {
    const fieldVal = val(x.obj, key);
    const v = key === "email" ? normEmail(fieldVal) : norm(fieldVal);
    return v === target;
  }) || null;
}

function norm(v){
  return String(v||"").trim().toLowerCase().replace(/\s+/g,"");
}

function normEmail(v){
  return String(v||"").trim().toLowerCase();
}

function canonicalKey(s){
  return String(s||"").toString().trim().toLowerCase().replace(/[^a-z0-9]/g,"");
}

function header(h,name){
  const target = canonicalKey(name);
  for (let i = 0; i < h.length; i++) {
    if (canonicalKey(h[i]) === target) return i+1;
  }
  return null;
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== DASHBOARD ====================

function dashboard(){
  const s=getData();
  if(!s)return json({status:"ERROR",message:"Sheet not found: "+SHEET_NAME});

  const g={
    member:newGroup(),
    guest:newGroup(),
    nonMember:newGroup(),
    all:new Set()
  };

  s.rows.forEach(x=>{
    const r=x.obj;
    const id=String(val(r,"registrationId")||("row_"+x.index)).trim();
    g.all.add(id);

    const type=String(val(r,"isMember")||"").trim().toLowerCase();
    const t=type==="yes"||type==="true"||type==="y"
      ?g.member
      :type==="guest"?g.guest:g.nonMember;

    t.ids.add(id);
    t.adults+=num(val(r,"adults"));

    const kids=String(val(r,"kids")||"").toLowerCase();
    if(kids&&kids!=="none")t.kids+=num(kids);

    const food=String(val(r,"foodPreference")||"").toLowerCase();
    if(food.includes("non"))t.nonveg++;
    else if(food.includes("veg"))t.veg++;

    if(String(val(r,"attendStatus")||"").toLowerCase()==="ok")t.attend++;
    if(String(val(r,"paymentStatus")||"").toLowerCase()==="ok")t.paid++;
  });

  const member=groupOut(g.member),guest=groupOut(g.guest),nonMember=groupOut(g.nonMember);

  const total={
    registrations:g.all.size,
    adults:member.adults+guest.adults+nonMember.adults,
    kids:member.kids+guest.kids+nonMember.kids,
    veg:member.veg+guest.veg+nonMember.veg,
    nonveg:member.nonveg+guest.nonveg+nonMember.nonveg,
    attend:member.attend+guest.attend+nonMember.attend,
    paid:member.paid+guest.paid+nonMember.paid
  };

  return json({member,memberGuest:guest,nonMember,total});
}

// ==================== DATA HELPERS ====================

function getData(){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if(!sheet)return null;

  const v=sheet.getDataRange().getValues();
  if(!v.length)return null;

  const headers=v[0].map(x=>String(x).trim().toLowerCase());
  const rows=v.slice(1).map((row,i)=>{
    const obj={};
    headers.forEach((h,j)=>obj[h]=row[j]);
    return {index:i,obj};
  });

  return {sheet,headers,rows};
}

function find(rows,field,value){
  const target = String(value).trim().toLowerCase();
  return rows.find(x=>{
    const fv = val(x.obj, field);
    if(field==="email") return normEmail(fv)===target;
    return norm(fv)===target;
  }) || null;
}

function header(headers,name){
  const target = canonicalKey(name);
  if (Array.isArray(headers)) {
    for (let i = 0; i < headers.length; i++) {
      if (canonicalKey(headers[i]) === target) return i + 1;
    }
    return null;
  }
  for (const k in headers) {
    if (canonicalKey(k) === target) return headers[k];
  }
  return null;
}

function val(obj,key){
  const target = canonicalKey(key);
  for (const k in obj) {
    if (canonicalKey(k) === target) return obj[k] ?? "";
  }
  return "";
}

function num(v){
  const n=Number(v);
  return Number.isFinite(n)?Math.round(n):0;
}

// ==================== DASHBOARD HELPERS ====================

function newGroup(){
  return {
    ids:new Set(),
    adults:0,
    kids:0,
    veg:0,
    nonveg:0,
    attend:0,
    paid:0
  };
}

function groupOut(g){
  return {
    registrations:g.ids.size,
    adults:g.adults,
    kids:g.kids,
    veg:g.veg,
    nonveg:g.nonveg,
    attend:g.attend,
    paid:g.paid
  };
}

// ==================== RESPONSE ====================

function json(data){
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== DASHBOARD HANDLER ==========
function handleDashboard() {
  const sheetInfo = getSheetAndRows();
  if (!sheetInfo) {
    return jsonResponse({
      member: zeroGroup(),
      nonMember: zeroGroup(),
      memberGuest: zeroGroup(),
      total: zeroGroup()
    });
  }

  const { rows } = sheetInfo;

  const member = { registrationsSet: new Set(), adults: 0, kids: 0, veg: 0, nonveg: 0, attend: 0, paid: 0 };
  const nonMember = { registrationsSet: new Set(), adults: 0, kids: 0, veg: 0, nonveg: 0, attend: 0, paid: 0 };
  const memberGuest = { registrationsSet: new Set(), adults: 0, kids: 0, veg: 0, nonveg: 0, attend: 0, paid: 0 };
  const allIds = new Set();

  rows.forEach(rowObj => {
    const regIdVal = getVal(rowObj, 'registrationId');
    const regId = (regIdVal !== undefined && regIdVal !== null && String(regIdVal).trim() !== '') ?
      String(regIdVal).trim() : ('row_' + Math.random().toString(36).slice(2, 8));

    allIds.add(regId);

    const isMemberVal = (getVal(rowObj, 'isMember') || '').toString().trim().toLowerCase();
    const isMemberFlag = (isMemberVal === 'yes' || isMemberVal === 'true' || isMemberVal === 'y');
    const isGuestFlag = (isMemberVal === 'guest');

    const target = isMemberFlag ? member : (isGuestFlag ? memberGuest : nonMember);

    target.registrationsSet.add(regId);

    target.adults += safeNum(getVal(rowObj, 'adults'));

    const kidsVal = (getVal(rowObj, 'kids') || '').toString().trim().toLowerCase();
    if (kidsVal && kidsVal !== 'none') {
      target.kids += safeNum(getVal(rowObj, 'kids'));
    }

    const food = (getVal(rowObj, 'foodPreference') || '').toString().toLowerCase();
    if (food.includes('non')) {
      target.nonveg += 1;
    } else if (food.includes('veg')) {
      target.veg += 1;
    }

    if ((getVal(rowObj, 'attendStatus') || '').toString().trim().toLowerCase() === 'ok') {
      target.attend += 1;
    }

    if ((getVal(rowObj, 'paymentStatus') || '').toString().trim().toLowerCase() === 'ok') {
      target.paid += 1;
    }
  });

  const memberOut = {
    registrations: member.registrationsSet.size,
    adults: member.adults,
    kids: member.kids,
    veg: member.veg,
    nonveg: member.nonveg,
    attend: member.attend,
    paid: member.paid
  };
  const nonMemberOut = {
    registrations: nonMember.registrationsSet.size,
    adults: nonMember.adults,
    kids: nonMember.kids,
    veg: nonMember.veg,
    nonveg: nonMember.nonveg,
    attend: nonMember.attend,
    paid: nonMember.paid
  };
  const memberGuestOut = {
    registrations: memberGuest.registrationsSet.size,
    adults: memberGuest.adults,
    kids: memberGuest.kids,
    veg: memberGuest.veg,
    nonveg: memberGuest.nonveg,
    attend: memberGuest.attend,
    paid: memberGuest.paid
  };
  const totalOut = {
    registrations: allIds.size,
    adults: memberOut.adults + nonMemberOut.adults + memberGuestOut.adults,
    kids: memberOut.kids + nonMemberOut.kids + memberGuestOut.kids,
    veg: memberOut.veg + nonMemberOut.veg + memberGuestOut.veg,
    nonveg: memberOut.nonveg + nonMemberOut.nonveg + memberGuestOut.nonveg,
    attend: memberOut.attend + nonMemberOut.attend + memberGuestOut.attend,
    paid: memberOut.paid + nonMemberOut.paid + memberGuestOut.paid
  };

  return jsonResponse({
    member: memberOut,
    nonMember: nonMemberOut,
    memberGuest: memberGuestOut,
    total: totalOut
  });
}

// ========== UTILITIES ==========

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getSheetAndRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { sheet: sheet, headers: {}, rows: [] };
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRow = values[0].map(h => (h || '').toString().trim());
  const headersMap = {}; // exact header text -> column index (1-based)
  for (let i = 0; i < headerRow.length; i++) {
    const key = headerRow[i] || ('Column' + (i + 1));
    headersMap[key] = i + 1;
  }

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = {};
    for (let c = 0; c < headerRow.length; c++) {
      const key = headerRow[c] || ('Column' + (c + 1));
      row[key] = values[r][c];
    }
    rows.push(row);
  }

  return { sheet: sheet, headers: headersMap, rows: rows };
}

// Case-insensitive value getter: looks up a field on a row object regardless
// of the exact casing used in the sheet's header row.
function getVal(rowObj, name) {
  const target = canonicalKey(name);
  for (const k in rowObj) {
    if (canonicalKey(k) === target) return rowObj[k];
  }
  return undefined;
}

// Case-insensitive header -> column index lookup.
function findHeaderCol(headers, name) {
  const target = canonicalKey(name);
  for (const k in headers) {
    if (canonicalKey(k) === target) return headers[k];
  }
  return null;
}

// Find a row by a given canonical field name (case-insensitive) and value.
function findRow(rows, fieldName, searchValue) {
  if (!rows || rows.length === 0) return null;
  const target = String(searchValue).trim().toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const v = getVal(rows[i], fieldName);
    if (v === null || v === undefined) continue;
    if (String(v).trim().toLowerCase() === target) {
      return { rowIndex: i + 1, rowObj: rows[i] }; // 1-based relative to data rows
    }
  }
  return null;
}

function safeNum(v) {
  const n = Number(v);
  return Number.isFinite(n) ? Math.round(n) : 0;
}

function zeroGroup() {
  return { registrations: 0, adults: 0, kids: 0, veg: 0, nonveg: 0, attend: 0, paid: 0 };
}