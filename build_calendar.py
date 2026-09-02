import win32com.client
from datetime import datetime, date
from collections import defaultdict
import json, calendar as cal_module

# ─── อ่านข้อมูลจาก Excel ───
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r'C:\Users\X\Desktop\TL_Schedule.xlsx')
ws = wb.Sheets('9月2026')

records = []
for row_idx in range(6, ws.UsedRange.Rows.Count + 1):
    b=ws.Cells(row_idx,2).Value; c=ws.Cells(row_idx,3).Value
    d=ws.Cells(row_idx,4).Value; e=ws.Cells(row_idx,5).Value
    f=ws.Cells(row_idx,6).Value; g=ws.Cells(row_idx,7).Value
    h=ws.Cells(row_idx,8).Value; i_=ws.Cells(row_idx,9).Value
    j=ws.Cells(row_idx,10).Value; k=ws.Cells(row_idx,11).Value
    l=ws.Cells(row_idx,12).Value; m=ws.Cells(row_idx,13).Value
    if c is None and f is None: continue
    def tdo(v):
        if v is None: return None
        if isinstance(v, datetime): return v.strftime('%Y-%m-%d')
        if hasattr(v,'year'):
            try: return f'{v.year:04d}-{v.month:02d}-{v.day:02d}'
            except: return None
        return None
    def ts(v):
        if v is None: return None
        return str(v).strip()
    def tf(v):
        if v is None: return None
        try: return float(v)
        except: return None
    records.append({
        'date':tdo(b),'qNo':ts(c),'customer':ts(f),
        'pcs':tf(d),'kg':tf(e),
        'arrival':tdo(g),'ship':tdo(h),
        'wrapShip':tdo(i_),'wrapArr':tdo(j),
        'ndkShip':tdo(k),'ndkArr':tdo(l),
        'memo':ts(m),
    })
excel.Quit()

monthTh=['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม']
monthShort=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']

def serR(r):
    def s(v):
        if isinstance(v,datetime): return v.strftime('%Y-%m-%d')
        return v
    return {k:s(v) for k,v in r.items()}
records_json=[serR(r) for r in records]
records_json_str=json.dumps(records_json,ensure_ascii=False)

today=date.today()
today_str=today.strftime('%Y-%m-%d')

# สร้างรายการวันนี้
today_events=[]
for r in records:
    if not r or (not r['qNo'] and not r['customer']): continue
    hasData=r['arrival'] or r['ship'] or r['wrapShip'] or r['wrapArr'] or r['ndkShip'] or r['ndkArr'] or r['memo']
    if not hasData: continue
    if r['arrival']==today_str: today_events.append({'type':'arrival','icon':'📥','label':'入荷','date':r['arrival'],'r':r})
    if r['ship']==today_str: today_events.append({'type':'ship','icon':'📦','label':'出荷','date':r['ship'],'r':r})
    if r['wrapShip']==today_str: today_events.append({'type':'wrap','icon':'📦','label':'ﾗｯﾌﾟ出荷','date':r['wrapShip'],'r':r})
    if r['wrapArr']==today_str: today_events.append({'type':'wrap','icon':'📥','label':'ﾗｯﾌﾟ入荷','date':r['wrapArr'],'r':r})
    if r['ndkShip']==today_str: today_events.append({'type':'ndk','icon':'📦','label':'NDK出荷','date':r['ndkShip'],'r':r})
    if r['ndkArr']==today_str: today_events.append({'type':'ndk','icon':'📥','label':'NDK入荷','date':r['ndkArr'],'r':r})
    if r['memo']: today_events.append({'type':'memo','icon':'📝','label':'備考','date':r['date'],'r':r})

priority={'ship':0,'ndk':1,'wrap':2,'arrival':3,'memo':4}
today_events_sorted=sorted(today_events,key=lambda e:priority.get(e['type'],5))

# สร้าง HTML สำหรับรายการวันนี้
today_list_html=''
if today_events_sorted:
    for ev in today_events_sorted:
        r=ev['r']
        bc={'ship':'bar-ship','ndk':'bar-ndk','wrap':'bar-wrap','arrival':'bar-arrival','memo':'bar-memo'}.get(ev['type'],'bar-other')
        today_list_html+=f'''
    <div class="evt-bar {ev['type']} {bc}" onclick="event.stopPropagation();openEdit('{r['qNo']}','{r['customer'] or ''}')">
      <span class="evt-bar-icon">{ev['icon']}</span>
      <div class="evt-bar-main"><span class="evt-bar-qno">{r['qNo']}</span><span class="evt-bar-cust">{r['customer'] or ''}</span></div>
      <div class="evt-bar-right"><span class="evt-bar-label">{ev['label']}</span><span class="evt-bar-date">{ev['date']}</span></div>
      <button class="evt-bar-del" onclick="event.stopPropagation();confirmDel('{r['qNo']}','{r['customer'] or ''}')" title="ลบ">✕</button>
    </div>'''
else:
    today_list_html='<div class="empty-today">วันนี้ไม่มีกำหนดการ</div>'

today_label=f'{today.day} {monthTh[today.month-1]} {today.year}'

# สร้าง day cards
days_in_month=cal_module.monthrange(today.year,today.month)[1]
first_weekday=cal_module.monthrange(today.year,today.month)[0]
START_OFFSET=first_weekday

def get_events_for_date(y,m,d):
    ts_=f'{y:04d}-{m+1:02d}-{d:02d}'
    results=[]
    for r in records:
        if not r or (not r['qNo'] and not r['customer']): continue
        hasData=r['arrival'] or r['ship'] or r['wrapShip'] or r['wrapArr'] or r['ndkShip'] or r['ndkArr'] or r['memo']
        if not hasData: continue
        if r['arrival']==ts_: results.append({'type':'arrival','icon':'📥','label':'入荷','date':r['arrival'],'r':r})
        if r['ship']==ts_: results.append({'type':'ship','icon':'📦','label':'出荷','date':r['ship'],'r':r})
        if r['wrapShip']==ts_: results.append({'type':'wrap','icon':'📦','label':'ﾗｯﾌﾟ出荷','date':r['wrapShip'],'r':r})
        if r['wrapArr']==ts_: results.append({'type':'wrap','icon':'📥','label':'ﾗｯﾌﾟ入荷','date':r['wrapArr'],'r':r})
        if r['ndkShip']==ts_: results.append({'type':'ndk','icon':'📦','label':'NDK出荷','date':r['ndkShip'],'r':r})
        if r['ndkArr']==ts_: results.append({'type':'ndk','icon':'📥','label':'NDK入荷','date':r['ndkArr'],'r':r})
        if r['memo']: results.append({'type':'memo','icon':'📝','label':'備考','date':r['date'],'r':r})
    return results

def get_bg(evts):
    has={}
    for e in evts: has[e['type']]=True
    if has.get('ship') and has.get('arrival'): return 'bg-mixed'
    if has.get('ship'): return 'bg-ship'
    if has.get('ndk'): return 'bg-ndk'
    if has.get('wrap'): return 'bg-wrap'
    if has.get('arrival'): return 'bg-arrival'
    if has.get('memo'): return 'bg-memo'
    return 'bg-empty'

day_cards_html=''
for i in range(14):
    day_num=i-START_OFFSET+1
    y,m,d=today.year,today.month,day_num
    if day_num<1:
        m=today.month-1
        if m<0: m=11; y=today.year-1
        d=cal_module.monthrange(y,m)[1]+day_num
    elif day_num>days_in_month:
        m=today.month+1
        if m>11: m=0; y=today.year+1
        d=day_num-days_in_month
    events=get_events_for_date(y,m,d)
    is_today=(y==today.year and m==today.month and d==today.day)
    day_label=f'{d} {monthShort[m]}'
    if is_today: day_label=f'🔴 {day_label}'
    bg=get_bg(events)
    evts_html=''
    if events:
        pr={'ship':0,'ndk':1,'wrap':2,'arrival':3,'memo':4}
        sorted_evts=sorted(events,key=lambda e:pr.get(e['type'],5))
        for ev in sorted_evts[:3]:
            r=ev['r']
            cls=f'evt-{ev["type"]}'
            icon=ev['icon']
            label=ev['label']
            date_disp=ev['date'] if ev['date'] else '—'
            evts_html+=f'''
            <div class="{cls}" onclick="event.stopPropagation();openEdit('{r["qNo"]}','{r["customer"] or ""}')" title="{r["qNo"]}">
              <span class="evt-icon-xs">{icon}</span>
              <span class="evt-qno-xs">{r["qNo"]}</span>
              <span class="evt-cust-xs">{r["customer"] or ""}</span>
              <span class="evt-right-xs"><span class="evt-label-xs">{label}</span> <span class="evt-date-xs">{date_disp}</span></span>
            </div>'''
        if len(sorted_evts)>3:
            evts_html+=f'<div class="evt-more-xs">+{len(sorted_evts)-3} รายการ</div>'
    else:
        evts_html='<div class="evt-empty-xs">—</div>'
    today_cls=' today' if is_today else ''
    count_label=f'{len(events)} รายการ' if events else '—'
    day_cards_html+=f'''
    <div class="day-card {bg}{today_cls}" onclick="showDay('{f"{y:04d}-{m+1:02d}-{d:02d}"}')">
      <div class="day-head"><span class="day-date">{day_label}</span><span class="day-count">{count_label}</span></div>
      <div class="day-evts">{evts_html}</div>
    </div>'''

# ─── HTML Template ───
HTML = r'''<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TL Calendar — Siam Yuken</title>
<style>
  :root { --fg: var(--foreground); --muted: var(--muted-foreground); --accent: var(--accent); --border: var(--border); --card: var(--card); }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; color: var(--fg); background: var(--bg); padding: 14px; min-height: 100vh; font-size: 13px; }
  h1 { font-size: 1.1rem; font-weight: 600; margin-bottom: 1px; }
  .sub { color: var(--muted); font-size: 0.7rem; margin-bottom: 8px; }
  .nav-bar { display: flex; align-items: center; gap: 4px; margin-bottom: 7px; flex-wrap: wrap; }
  .nav-btn { padding: 3px 7px; border-radius: 4px; border: 1px solid var(--border); background: var(--card); color: var(--fg); cursor: pointer; font-size: 0.7rem; transition: all 0.1s; }
  .nav-btn:hover { background: rgba(128,128,128,0.07); }
  .nav-btn.primary { background: var(--accent); color: #fff; border-color: var(--accent); }
  .nav-btn.primary:hover { opacity: 0.88; }
  .month-label { font-size: 0.8rem; font-weight: 600; min-width: 120px; text-align: center; padding: 3px 7px; border-radius: 4px; background: var(--card); border: 1px solid var(--border); }
  .spacer { flex: 1; }
  .today-section { margin-bottom: 8px; }
  .today-header { display: flex; align-items: center; gap: 6px; margin-bottom: 3px; }
  .today-badge { background: var(--accent); color: #fff; padding: 2px 7px; border-radius: 4px; font-size: 0.7rem; font-weight: 600; }
  .today-count { font-size: 0.65rem; color: var(--muted); }
  .evt-bar { display: flex; align-items: center; gap: 5px; padding: 3px 5px; border-radius: 4px; cursor: pointer; transition: all 0.08s; position: relative; font-size: 0.76rem; }
  .evt-bar:hover { filter: brightness(0.95); }
  .evt-bar-icon { font-size: 0.78rem; width: 14px; text-align: center; flex-shrink: 0; }
  .evt-bar-main { flex: 1; display: flex; flex-direction: column; min-width: 0; }
  .evt-bar-qno { font-family: 'Courier New', monospace; font-weight: 600; font-size: 0.7rem; }
  .evt-bar-cust { font-size: 0.62rem; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .evt-bar-right { text-align: right; flex-shrink: 0; line-height: 1.3; }
  .evt-bar-label { font-size: 0.65rem; font-weight: 600; }
  .evt-bar-date { font-size: 0.58rem; color: var(--muted); }
  .evt-bar-del { position: absolute; right: 1px; top: 1px; width: 15px; height: 15px; border: none; border-radius: 50%; background: rgba(239,68,68,0.75); color: #fff; font-size: 0.5rem; cursor: pointer; line-height: 1; padding: 0; opacity: 0; transition: opacity 0.12s; }
  .evt-bar:hover .evt-bar-del { opacity: 1; }
  .evt-bar-del:hover { background: #ef4444; }
  .empty-today { color: var(--muted); text-align: center; padding: 6px; font-size: 0.72rem; font-style: italic; }
  .bar-ship { background: rgba(16,185,129,0.08); border-left: 2.5px solid #10b981; }
  .bar-ndk { background: rgba(239,68,68,0.07); border-left: 2.5px solid #f472b6; }
  .bar-wrap { background: rgba(245,158,11,0.09); border-left: 2.5px solid #f59e0b; }
  .bar-arrival { background: rgba(59,130,246,0.08); border-left: 2.5px solid #3b82f6; }
  .bar-memo { background: rgba(107,114,128,0.09); border-left: 2.5px solid #6b7280; }
  .bar-other { background: rgba(128,128,128,0.05); border-left: 2.5px solid #9ca3af; }
  .cal-sec { margin-bottom: 6px; }
  .cal-sec-title { font-size: 0.65rem; color: var(--muted); margin-bottom: 3px; padding-bottom: 2px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 3px; }
  .cal-sec-title::before { content: '📅'; font-size: 0.65rem; }
  .cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }
  .cal-hdr { text-align: center; font-size: 0.58rem; color: var(--muted); font-weight: 600; padding: 1px 0; border-bottom: 1px solid var(--border); }
  .day-card { border: 1px solid var(--border); border-radius: 3px; padding: 2px 3px; cursor: pointer; transition: all 0.1s; min-height: 42px; background: var(--card); }
  .day-card:hover { filter: brightness(0.96); transform: translateY(-1px); }
  .day-card.today { border: 2px solid var(--accent); background: rgba(128,128,128,0.03); }
  .day-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 1px; }
  .day-date { font-size: 0.58rem; font-weight: 600; }
  .day-card.today .day-date { color: var(--accent); }
  .day-count { font-size: 0.5rem; color: var(--muted); background: rgba(128,128,128,0.08); padding: 0 2px; border-radius: 4px; }
  .day-evts { display: flex; flex-direction: column; gap: 1px; }
  .evt-row-small { display: flex; align-items: center; gap: 2px; padding: 1px 2px; border-radius: 2px; font-size: 0.52rem; cursor: pointer; }
  .evt-row-small:hover { filter: brightness(0.88); }
  .evt-arrival { background: rgba(59,130,246,0.1); }
  .evt-ship { background: rgba(16,185,129,0.1); }
  .evt-wrap { background: rgba(245,158,11,0.1); }
  .evt-ndk { background: rgba(239,68,68,0.1); }
  .evt-memo { background: rgba(107,114,128,0.1); }
  .evt-icon-xs { font-size: 0.48rem; width: 6px; text-align: center; flex-shrink: 0; }
  .evt-qno-xs { font-family: 'Courier New', monospace; font-weight: 600; font-size: 0.48rem; }
  .evt-cust-xs { font-size: 0.45rem; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .evt-right-xs { text-align: right; flex-shrink: 0; line-height: 1.2; }
  .evt-label-xs { font-size: 0.45rem; opacity: 0.7; }
  .evt-date-xs { font-size: 0.43rem; color: var(--muted); }
  .evt-more-xs { font-size: 0.45rem; color: var(--muted); text-align: center; padding: 0 1px; font-style: italic; }
  .evt-empty-xs { font-size: 0.45rem; color: var(--muted); text-align: center; padding: 2px; font-style: italic; }
  .bg-ship { background: rgba(16,185,129,0.025); }
  .bg-arrival { background: rgba(59,130,246,0.025); }
  .bg-ndk { background: rgba(239,68,68,0.025); }
  .bg-wrap { background: rgba(245,158,11,0.025); }
  .bg-memo { background: rgba(107,114,128,0.025); }
  .bg-mixed { background: rgba(16,185,129,0.02); }
  .bg-empty { background: rgba(128,128,128,0.02); }
  .dp { margin-top: 5px; padding: 6px 8px; border: 1px solid var(--border); border-radius: 4px; background: var(--card); display: none; max-width: 420px; }
  .dp.visible { display: block; }
  .dp-hdr { display: flex; align-items: center; justify-content: space-between; margin-bottom: 3px; }
  .dp-title { font-size: 0.78rem; font-weight: 600; }
  .dp-close { border: none; background: transparent; color: var(--muted); cursor: pointer; font-size: 0.85rem; padding: 0 2px; }
  .dp-close:hover { color: var(--fg); }
  .info-tbl { width: 100%; border-collapse: collapse; font-size: 0.62rem; }
  .info-tbl td { padding: 1px 3px; border-bottom: 1px solid rgba(128,128,128,0.06); }
  .info-tbl td:first-child { color: var(--muted); width: 32%; }
  .info-tbl td:last-child { font-weight: 500; }
  .mb { display: none; position: fixed; top:0; left:0; right:0; bottom:0; background: rgba(0,0,0,0.4); z-index: 1000; justify-content: center; align-items: center; }
  .mb.visible { display: flex; }
  .modal { background: var(--card); border: 1px solid var(--border); border-radius: 5px; padding: 10px 12px; width: 90%; max-width: 380px; }
  .modal h3 { font-size: 0.8rem; margin-bottom: 8px; }
  .modal .m-close { float: right; border: none; background: transparent; color: var(--muted); cursor: pointer; font-size: 0.85rem; padding: 0 2px; }
  .modal .m-close:hover { color: var(--fg); }
  .fg { margin-bottom: 5px; }
  .fg label { display: block; font-size: 0.6rem; color: var(--muted); margin-bottom: 1px; font-weight: 500; }
  .fg input { width: 100%; border: 1px solid var(--border); border-radius: 3px; padding: 3px 5px; font-size: 0.7rem; background: var(--bg); color: var(--fg); font-family: inherit; }
  .fg input:focus { outline: 2px solid var(--accent); border-color: transparent; }
  .fr { display: grid; grid-template-columns: 1fr 1fr; gap: 5px; }
  .ma { display: flex; gap: 4px; justify-content: flex-end; margin-top: 8px; padding-top: 6px; border-top: 1px solid var(--border); }
  .btn { padding: 3px 8px; border-radius: 3px; border: none; cursor: pointer; font-size: 0.65rem; font-weight: 500; }
  .btn-primary { background: var(--accent); color: #fff; }
  .btn-primary:hover { opacity: 0.85; }
  .btn-secondary { background: var(--card); color: var(--fg); border: 1px solid var(--border); }
  .btn-secondary:hover { background: rgba(128,128,128,0.07); }
  .btn-danger { background: #ef4444; color: #fff; }
  .btn-danger:hover { opacity: 0.85; }
  .cb { display: none; position: fixed; top:0; left:0; right:0; bottom:0; background: rgba(0,0,0,0.4); z-index: 2000; justify-content: center; align-items: center; }
  .cb.visible { display: flex; }
  .cbox { background: var(--card); border: 1px solid var(--border); border-radius: 4px; padding: 8px 12px; width: 90%; max-width: 260px; text-align: center; }
  .cbox h4 { font-size: 0.78rem; margin-bottom: 1px; }
  .cbox p { font-size: 0.68rem; color: var(--muted); margin-bottom: 6px; }
  .cbox .ca { display: flex; gap: 4px; justify-content: center; }
  .footer { margin-top: 8px; padding-top: 4px; border-top: 1px solid var(--border); text-align: center; font-size: 0.55rem; color: var(--muted); }
</style>
</head>
<body>

<h1>📅 TL Calendar — Siam Yuken</h1>
<div class="sub">入荷出荷確認表 · คลิกแถบเพื่อแก้ไข · ปุ่ม ✕ เพื่อลบ</div>

<div class="nav-bar">
  <button class="nav-btn" id="prevMonth">◀</button>
  <span class="month-label" id="monthLabel">__MONTH_LABEL__</span>
  <button class="nav-btn" id="nextMonth">▶</button>
  <button class="nav-btn primary" id="todayBtn">📌 วันนี้</button>
  <span class="spacer"></span>
  <button class="nav-btn primary" id="addBtn" style="background:var(--accent);color:#fff;">➕ เพิ่ม</button>
  <button class="nav-btn" id="exportBtn">💾 CSV</button>
  <button class="nav-btn" id="resetBtn" style="color:#ef4444;">↺ รีเซ็ต</button>
</div>

<div class="today-section">
  <div class="today-header">
    <span class="today-badge">📌 วันนี้ __TODAY_LABEL__</span>
    <span class="today-count" id="todayCount">กำลังโหลด...</span>
  </div>
  <div id="todayList">
    __TODAY_LIST__
  </div>
</div>

<div class="cal-sec">
  <div class="cal-sec-title">14 วันข้างหน้า — คลิกวันเพื่อดูรายละเอียด</div>
  <div class="cal-grid" id="calHeader">
    <div class="cal-hdr">จ.</div><div class="cal-hdr">อ.</div><div class="cal-hdr">พ.</div>
    <div class="cal-hdr">พฤ.</div><div class="cal-hdr">ศ.</div><div class="cal-hdr">ส.</div><div class="cal-hdr">อา.</div>
  </div>
  <div class="cal-grid" id="dayGrid">
    __DAY_CARDS__
  </div>
</div>

<div class="dp" id="detailPanel"></div>

<div class="mb" id="modalBackdrop">
  <div class="modal">
    <button class="m-close" onclick="closeModal()">✕</button>
    <h3 id="modalTitle">➕ เพิ่มรายการใหม่</h3>
    <input type="hidden" id="editId">
    <div class="fg"><label>Quotation No. *</label><input type="text" id="f_qNo" placeholder="Q2609T001"></div>
    <div class="fg"><label>お客様 (Customer) *</label><input type="text" id="f_customer" placeholder="บริษัท ลูกค้า"></div>
    <div class="fr">
      <div class="fg"><label>入荷日</label><input type="date" id="f_arrival"></div>
      <div class="fg"><label>出荷日</label><input type="date" id="f_ship"></div>
    </div>
    <div class="fr">
      <div class="fg"><label>ﾗｯﾌﾟ出荷</label><input type="date" id="f_wrapShip"></div>
      <div class="fg"><label>ﾗｯﾌﾟ入荷</label><input type="date" id="f_wrapArr"></div>
    </div>
    <div class="fr">
      <div class="fg"><label>NDK出荷</label><input type="date" id="f_ndkShip"></div>
      <div class="fg"><label>NDK入荷</label><input type="date" id="f_ndkArr"></div>
    </div>
    <div class="fr">
      <div class="fg"><label>PCS</label><input type="number" id="f_pcs" step="0.01" min="0"></div>
      <div class="fg"><label>KG</label><input type="number" id="f_kg" step="0.001" min="0"></div>
    </div>
    <div class="fg"><label>備考 (Memo)</label><input type="text" id="f_memo" placeholder="หมายเหตุ"></div>
    <div class="ma">
      <button class="btn btn-danger" id="deleteBtn" style="display:none;" onclick="deleteCur()">🗑 ลบ</button>
      <button class="btn btn-secondary" onclick="closeModal()">ยกเลิก</button>
      <button class="btn btn-primary" onclick="saveRec()">💾 บันทึก</button>
    </div>
  </div>
</div>

<div class="cb" id="confirmBackdrop">
  <div class="cbox">
    <h4>⚠ ยืนยันการลบ</h4>
    <p id="confirmMsg">คุณแน่ใจไหมว่าจะลบรายการนี้?</p>
    <div class="ca">
      <button class="btn btn-secondary" onclick="closeConfirm()">ยกเลิก</button>
      <button class="btn btn-danger" id="confirmDelBtn">🗑 ลบ</button>
    </div>
  </div>
</div>

<div class="footer">
  TL Calendar — Siam Yuken · 入荷出荷確認表<br>
  ข้อมูลจาก TL_Schedule.xlsx · กดรีเซ็ตเพื่อกลับเป็นข้อมูลเดิม
</div>

<script>
const RECORDS = __RECORDS_JSON__;
const MONTHS=["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
const MONTHS_SHORT=["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];

let curYear=__CUR_YEAR__, curMonth=__CUR_MONTH__;
let currentEditId=null, currentCustomer=null;

let records = JSON.parse(localStorage.getItem('tl_cal_v5') || JSON.stringify(RECORDS));

function saveAll(){localStorage.setItem('tl_cal_v5', JSON.stringify(records));}

function pad(n){return String(n).padStart(2,'0');}
function makeDate(y,m,d){return y+'-'+pad(m+1)+'-'+pad(d);}

function getAllEvents(y,m,d){
  const ts=makeDate(y,m,d);
  const ret=[];
  for(const r of records){
    if(!r||(!r.qNo&&!r.customer)) continue;
    const has=r.arrival||r.ship||r.wrapShip||r.wrapArr||r.ndkShip||r.ndkArr||r.memo;
    if(!has) continue;
    if(r.arrival===ts) ret.push({t:'arrival',icon:'📥',label:'入荷',date:r.arrival,r});
    if(r.ship===ts) ret.push({t:'ship',icon:'📦',label:'出荷',date:r.ship,r});
    if(r.wrapShip===ts) ret.push({t:'wrap',icon:'📦',label:'ﾗｯﾌﾟ出荷',date:r.wrapShip,r});
    if(r.wrapArr===ts) ret.push({t:'wrap',icon:'📥',label:'ﾗｯﾌﾟ入荷',date:r.wrapArr,r});
    if(r.ndkShip===ts) ret.push({t:'ndk',icon:'📦',label:'NDK出荷',date:r.ndkShip,r});
    if(r.ndkArr===ts) ret.push({t:'ndk',icon:'📥',label:'NDK入荷',date:r.ndkArr,r});
    if(r.memo) ret.push({t:'memo',icon:'📝',label:'備考',date:r.date,r});
  }
  return ret;
}

function getDayBg(evts){
  const has={};
  for(const e of evts) has[e.t]=1;
  if(has.ship&&has.arrival) return 'bg-mixed';
  if(has.ship) return 'bg-ship';
  if(has.ndk) return 'bg-ndk';
  if(has.wrap) return 'bg-wrap';
  if(has.arrival) return 'bg-arrival';
  if(has.memo) return 'bg-memo';
  return 'bg-empty';
}

function renderCalendar(){
  document.getElementById('monthLabel').textContent=MONTHS[curMonth]+' '+curYear;
  const grid=document.getElementById('dayGrid');
  const t=new Date();
  const tY=t.getFullYear(), tM=t.getMonth(), tD=t.getDate();
  const daysInMon=new Date(curYear,curMonth+1,0).getDate();
  const firstWd=new Date(curYear,curMonth,1).getDay();
  const startOff=firstWd===0?6:firstWd-1;
  let html='';
  for(let i=0;i<14;i++){
    let d=i-startOff+1, y=curYear, m=curMonth;
    if(d<1){m--;if(m<0){m=11;y--;}d=new Date(y,m+1,0).getDate()+d;}
    else if(d>daysInMon){m++;if(m>11){m=0;y++;}d=d-daysInMon;}
    const ts2=makeDate(y,m,d);
    const evts=getAllEvents(y,m,d);
    const isT=(y===tY&&m===tM&&d===tD);
    const dayLabel=d+' '+MONTHS_SHORT[m];
    const bg=getDayBg(evts);
    const priority={ship:0,ndk:1,wrap:2,arrival:3,memo:4};
    const sorted=evts.sort((a,b)=>(priority[a.t]||5)-(priority[b.t]||5));
    let evtsHtml='';
    if(sorted.length>0){
      for(const e of sorted.slice(0,3)){
        const cls='evt-'+e.t;
        evtsHtml+=`<div class="${cls}" onclick="event.stopPropagation();openEdit('${e.r.qNo}','${e.r.customer||''}')" title="${e.r.qNo}">
          <span class="evt-icon-xs">${e.icon}</span>
          <span class="evt-qno-xs">${e.r.qNo}</span>
          <span class="evt-cust-xs">${e.r.customer||''}</span>
          <span class="evt-right-xs"><span class="evt-label-xs">${e.label}</span> <span class="evt-date-xs">${e.date||'—'}</span></span>
        </div>`;
      }
      if(sorted.length>3) evtsHtml+=`<div class="evt-more-xs">+${sorted.length-3}</div>`;
    } else {
      evtsHtml='<div class="evt-empty-xs">—</div>';
    }
    const tClass=isT?' today':'';
    const countLabel=sorted.length>0?sorted.length+' รายการ':'—';
    html+=`<div class="day-card ${bg}${tClass}" onclick="showDay('${ts2}')">
      <div class="day-head"><span class="day-date">${dayLabel}</span><span class="day-count">${countLabel}</span></div>
      <div class="day-evts">${evtsHtml}</div>
    </div>`;
  }
  grid.innerHTML=html;
}

function buildTodayList(){
  const t=new Date();
  const ts=makeDate(t.getFullYear(),t.getMonth(),t.getDate());
  const todays=[];
  for(const r of records){
    if(!r||(!r.qNo&&!r.customer)) continue;
    if(r.arrival===ts) todays.push({t:'arrival',icon:'📥',label:'入荷',date:r.arrival,r});
    if(r.ship===ts) todays.push({t:'ship',icon:'📦',label:'出荷',date:r.ship,r});
    if(r.wrapShip===ts) todays.push({t:'wrap',icon:'📦',label:'ﾗｯﾌﾟ出荷',date:r.wrapShip,r});
    if(r.wrapArr===ts) todays.push({t:'wrap',icon:'📥',label:'ﾗｯﾌﾟ入荷',date:r.wrapArr,r});
    if(r.ndkShip===ts) todays.push({t:'ndk',icon:'📦',label:'NDK出荷',date:r.ndkShip,r});
    if(r.ndkArr===ts) todays.push({t:'ndk',icon:'📥',label:'NDK入荷',date:r.ndkArr,r});
    if(r.memo&&r.date===ts) todays.push({t:'memo',icon:'📝',label:'備考',date:r.date,r});
  }
  const priority={ship:0,ndk:1,wrap:2,arrival:3,memo:4};
  todays.sort((a,b)=>(priority[a.t]||5)-(priority[b.t]||5));
  let html='';
  if(todays.length>0){
    for(const e of todays){
      const bc={
        'ship':'bar-ship','ndk':'bar-ndk','wrap':'bar-wrap','arrival':'bar-arrival','memo':'bar-memo'
      }[e.t]||'bar-other';
      html+=`<div class="evt-bar ${e.t} ${bc}" onclick="event.stopPropagation();openEdit('${e.r.qNo}','${e.r.customer||''}')">
        <span class="evt-bar-icon">${e.icon}</span>
        <div class="evt-bar-main"><span class="evt-bar-qno">${e.r.qNo}</span><span class="evt-bar-cust">${e.r.customer||''}</span></div>
        <div class="evt-bar-right"><span class="evt-bar-label">${e.label}</span><span class="evt-bar-date">${e.date}</span></div>
        <button class="evt-bar-del" onclick="event.stopPropagation();confirmDel('${e.r.qNo}','${e.r.customer||''}')" title="ลบ">✕</button>
      </div>`;
    }
  } else {
    html='<div class="empty-today">วันนี้ไม่มีกำหนดการ</div>';
  }
  document.getElementById('todayList').innerHTML=html;
  document.getElementById('todayCount').textContent=todays.length+' รายการในวันนี้';
}

function showDay(ts){
  const panel=document.getElementById('detailPanel');
  const evts=[];
  for(const r of records){
    if(!r||(!r.qNo&&!r.customer)) continue;
    if(r.arrival===ts) evts.push({t:'arrival',icon:'📥',label:'入荷',r});
    if(r.ship===ts) evts.push({t:'ship',icon:'📦',label:'出荷',r});
    if(r.wrapShip===ts) evts.push({t:'wrap',icon:'📦',label:'ﾗｯﾌﾟ出荷',r});
    if(r.wrapArr===ts) evts.push({t:'wrap',icon:'📥',label:'ﾗｯﾌﾟ入荷',r});
    if(r.ndkShip===ts) evts.push({t:'ndk',icon:'📦',label:'NDK出荷',r});
    if(r.ndkArr===ts) evts.push({t:'ndk',icon:'📥',label:'NDK入荷',r});
    if(r.memo) evts.push({t:'memo',icon:'📝',label:'備考',r});
  }
  if(evts.length===0){
    panel.innerHTML=`<div style="font-size:0.72rem;color:var(--muted);">📅 ${ts} — ไม่มีกำหนดการ</div>`;
    panel.classList.add('visible');
    return;
  }
  const byCust={};
  for(const e of evts){
    const cust=e.r.customer||'(ไม่ระบุ)';
    if(!byCust[cust]) byCust[cust]=[];
    byCust[cust].push(e);
  }
  let html=`<div class="dp-hdr"><span style="font-weight:600;">📅 ${ts}</span><button class="dp-close" onclick="document.getElementById('detailPanel').classList.remove('visible')">✕</button></div>`;
  for(const cust in byCust){
    const evts2=byCust[cust];
    const priority2={ship:0,ndk:1,wrap:2,arrival:3,memo:4};
    evts2.sort((a,b)=>(priority2[a.t]||5)-(priority2[b.t]||5));
    html+=`<div style="margin-top:4px;padding-top:4px;border-top:1px solid rgba(128,128,128,0.1);">
      <div style="font-size:0.68rem;font-weight:600;color:var(--muted);">🏭 ${cust} (${evts2.length})</div>
      <table class="info-tbl">`;
    for(const e of evts2){
      const r=e.r;
      html+=`<tr><td>Quotation No.</td><td><strong>${r.qNo}</strong></td></tr>
        <tr><td>お客様 (Customer)</td><td>${r.customer||'—'}</td></tr>
        <tr><td>出荷日 (Ship Date)</td><td>${r.ship||'—'}</td></tr>
        <tr><td>NDK出荷日</td><td>${r.ndkShip||'—'}</td></tr>
        <tr><td>NDK入荷日</td><td>${r.ndkArr||'—'}</td></tr>
        <tr><td>入荷日</td><td>${r.arrival||'—'}</td></tr>
        <tr><td>ﾗｯﾌﾟ出荷日</td><td>${r.wrapShip||'—'}</td></tr>
        <tr><td>ﾗｯﾌﾟ入荷日</td><td>${r.wrapArr||'—'}</td></tr>
        <tr><td>PCS</td><td>${r.pcs!==null?r.pcs:'—'}</td></tr>
        <tr><td>KG</td><td>${r.kg!==null?r.kg:'—'}</td></tr>`;
      if(r.memo) html+=`<tr><td>備考</td><td>${r.memo}</td></tr>`;
      html+=`<tr style="border-bottom:none;"><td colspan="2" style="height:1px;"></td></tr>`;
    }
    html+=`</table></div>`;
  }
  panel.innerHTML=html;
  panel.classList.add('visible');
}

function openEdit(qNo,customer){
  const r=records.find(rec=>rec.qNo===qNo&&rec.customer===customer);
  if(!r) return;
  currentEditId=qNo; currentCustomer=customer;
  document.getElementById('modalTitle').textContent='✏️ แก้ไขรายการ';
  document.getElementById('editId').value=qNo;
  document.getElementById('f_qNo').value=r.qNo||'';
  document.getElementById('f_customer').value=r.customer||'';
  document.getElementById('f_arrival').value=r.arrival||'';
  document.getElementById('f_ship').value=r.ship||'';
  document.getElementById('f_wrapShip').value=r.wrapShip||'';
  document.getElementById('f_wrapArr').value=r.wrapArr||'';
  document.getElementById('f_ndkShip').value=r.ndkShip||'';
  document.getElementById('f_ndkArr').value=r.ndkArr||'';
  document.getElementById('f_pcs').value=r.pcs!==null?r.pcs:'';
  document.getElementById('f_kg').value=r.kg!==null?r.kg:'';
  document.getElementById('f_memo').value=r.memo||'';
  document.getElementById('deleteBtn').style.display='block';
  openModal();
}
function openAdd(){
  currentEditId=null; currentCustomer=null;
  document.getElementById('modalTitle').textContent='➕ เพิ่มรายการใหม่';
  document.getElementById('editId').value='';
  for(const id of ['f_qNo','f_customer','f_arrival','f_ship','f_wrapShip','f_wrapArr','f_ndkShip','f_ndkArr','f_pcs','f_kg','f_memo'])
    document.getElementById(id).value='';
  document.getElementById('deleteBtn').style.display='none';
  openModal();
}
function openModal(){document.getElementById('modalBackdrop').classList.add('visible');}
function closeModal(){document.getElementById('modalBackdrop').classList.remove('visible');currentEditId=null;currentCustomer=null;}
function closeConfirm(){document.getElementById('confirmBackdrop').classList.remove('visible');}

function saveRec(){
  const qNo=document.getElementById('f_qNo').value.trim();
  const customer=document.getElementById('f_customer').value.trim();
  if(!qNo){alert('กรุณาระบุ Quotation No.');return;}
  if(!customer){alert('กรุณาระบุ Customer');return;}
  const rec={
    date:'',
    qNo:qNo,
    customer:customer,
    pcs:parseFloat(document.getElementById('f_pcs').value)||null,
    kg:parseFloat(document.getElementById('f_kg').value)||null,
    arrival:document.getElementById('f_arrival').value||null,
    ship:document.getElementById('f_ship').value||null,
    wrapShip:document.getElementById('f_wrapShip').value||null,
    wrapArr:document.getElementById('f_wrapArr').value||null,
    ndkShip:document.getElementById('f_ndkShip').value||null,
    ndkArr:document.getElementById('f_ndkArr').value||null,
    memo:document.getElementById('f_memo').value.trim()||null,
  };
  if(currentEditId){
    const idx=records.findIndex(r=>r.qNo===currentEditId&&r.customer===currentCustomer);
    if(idx>=0) records[idx]=rec; else records.push(rec);
  } else records.push(rec);
  saveAll();
  closeModal();
  buildTodayList();
  renderCalendar();
}

function confirmDel(qNo,customer){
  document.getElementById('confirmMsg').textContent=`ลบ ${qNo} (${customer})?`;
  document.getElementById('confirmDelBtn').onclick=function(){
    const idx=records.findIndex(r=>r.qNo===qNo&&r.customer===customer);
    if(idx>=0) records.splice(idx,1);
    saveAll();
    closeConfirm();
    buildTodayList();
    renderCalendar();
  };
  document.getElementById('confirmBackdrop').classList.add('visible');
}
function deleteCur(){
  if(!currentEditId) return;
  confirmDel(currentEditId,currentCustomer);
}

function exportCSV(){
  let csv='Quotation No.,Customer,入荷日,出荷日,ﾗｯﾌﾟ出荷日,ﾗｯﾌﾟ入荷日,NDK出荷日,NDK入荷日,PCS,KG,備考\n';
  for(const r of records){
    csv+='"'+r.qNo+'","'+(r.customer||'')+'","'+(r.arrival||'')+'","'+(r.ship||'')+'","'+(r.wrapShip||'')+'","'+(r.wrapArr||'')+'","'+(r.ndkShip||'')+'","'+(r.ndkArr||'')+'","'+(r.pcs!==null?r.pcs:'')+'","'+(r.kg!==null?r.kg:'')+'","'+(r.memo||'')+'"'+'\n';
  }
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  const link=document.createElement('a');
  link.href=URL.createObjectURL(blob);
  link.download='TL_Calendar_'+new Date().toISOString().split('T')[0]+'.csv';
  link.click();
  URL.revokeObjectURL(link.href);
}

function resetAll(){
  if(confirm('🚫 ล้างการแก้ไขทั้งหมด และกลับไปใช้ข้อมูลต้นฉบับจาก Excel?')){
    localStorage.removeItem('tl_cal_v5');
    records=JSON.parse(JSON.stringify(RECORDS));
    buildTodayList();
    renderCalendar();
  }
}

document.getElementById('prevMonth').onclick=function(){curMonth--;if(curMonth<0){curMonth=11;curYear--;}renderCalendar();};
document.getElementById('nextMonth').onclick=function(){curMonth++;if(curMonth>11){curMonth=0;curYear++;}renderCalendar();};
document.getElementById('todayBtn').onclick=function(){const t=new Date();curYear=t.getFullYear();curMonth=t.getMonth();renderCalendar();};
document.getElementById('addBtn').onclick=openAdd;
document.getElementById('exportBtn').onclick=exportCSV;
document.getElementById('resetBtn').onclick=resetAll;
document.getElementById('modalBackdrop').onclick=function(e){if(e.target.id==='modalBackdrop')closeModal();};
document.getElementById('confirmBackdrop').onclick=function(e){if(e.target.id==='confirmBackdrop')closeConfirm();};

buildTodayList();
renderCalendar();
</script>
</body>
</html>'''

# แทนที่ placeholder
full_html = HTML
full_html = full_html.replace('__MONTH_LABEL__', f'{monthTh[today.month-1]} {today.year}')
full_html = full_html.replace('__TODAY_LABEL__', today_label)
full_html = full_html.replace('__TODAY_LIST__', today_list_html)
full_html = full_html.replace('__DAY_CARDS__', day_cards_html)
full_html = full_html.replace('__CUR_YEAR__', str(today.year))
full_html = full_html.replace('__CUR_MONTH__', str(today.month))
full_html = full_html.replace('__RECORDS_JSON__', records_json_str)

with open(r'C:\Users\X\Desktop\TL_Calendar.html', 'w', encoding='utf-8') as f:
    f.write(full_html)

today_by_type=defaultdict(int)
for ev in today_events_sorted: today_by_type[ev['type']]+=1

print(f'✅ Done! {len(records)} records')
print(f'📅 {monthTh[today.month-1]} {today.year}')
print(f'📌 วันนี้ {today.day} {monthTh[today.month-1]} {today.year}')
print(f'📦 出荷: {today_by_type.get("ship",0)} · 📥 入荷: {today_by_type.get("arrival",0)} · 🏭 NDK: {today_by_type.get("ndk",0)} · 📋 ﾗｯﾌﾟ: {today_by_type.get("wrap",0)}')
print(f'✨ ออกแบบใหม่: แถบรายการวันนี้ + Calendar Grid + แก้ไขได้')
