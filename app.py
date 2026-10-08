import io, re, copy, time
from collections import defaultdict
from datetime import datetime, date
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import ExcelWriter
from zipfile import ZipFile, ZIP_DEFLATED

st.set_page_config(page_title='後方数値分析ツール生成', layout='wide')
st.title('後方数値分析ツール生成')
st.caption('v7：Excel保存の圧縮負荷を軽減。3シート・数式・書式・A1選択は維持します。')

FONT_NAME = 'Meiryo UI'
KEEP_SHEETS = ['コストデータ', '後方数値データ(加工版)', '媒体コードマスタver3']

def norm(v):
    if v is None: return ''
    return str(v).replace('\n','').replace('\r','').replace('　',' ').strip()

def compact(v):
    s=norm(v).lower()
    return re.sub(r'[\s()（）_・!！]+','',s).replace('配信','')

def month_key(v):
    if isinstance(v,(datetime,date)): return v.strftime('%Y/%m')
    if isinstance(v,(int,float)):
        # Excel serial date
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(v).strftime('%Y/%m')
        except Exception: return ''
    s=norm(v)
    for fmt in ('%Y/%m/%d','%Y-%m-%d','%Y/%m','%Y-%m'):
        try: return datetime.strptime(s,fmt).strftime('%Y/%m')
        except ValueError: pass
    return ''

def age_decade(v):
    try: return int(float(v))//10*10
    except Exception: return None

def report_sheet_for(wb, item2):
    target = 'YDO' if norm(item2) == 'Yahoo!ダイレクトオファー' else norm(item2)
    t=compact(target)
    for sn in wb.sheetnames:
        if t and t in compact(sn): return wb[sn]
    return None

def parse_report_sheet_aggregated(ws):
    """Read only needed values and aggregate immediately.
    Returns (normal_by_month, ydo_by_month, months).
    """
    header_row=None; headers={}
    # read_only worksheets are much faster when iterated rather than random cell access
    for ridx,row in enumerate(ws.iter_rows(min_row=1,max_row=20,values_only=True), start=1):
        vals=[norm(v) for v in row[:20]]
        if '配信日' in vals and any('Cost(Gross)' in x for x in vals):
            header_row=ridx
            headers={norm(v):i for i,v in enumerate(row) if v is not None}
            break
    if not header_row: return defaultdict(float), defaultdict(float), set()
    c_date=headers.get('配信日'); c_seg=headers.get('セグメント別')
    c_cost=next((i for h,i in headers.items() if 'Cost(Gross)' in h),None)
    c_code=next((i for h,i in headers.items() if '媒体コード' in h),None)
    if c_date is None or c_seg is None or c_cost is None:
        return defaultdict(float), defaultdict(float), set()
    normal=defaultdict(float); ydo=defaultdict(float); months=set(); parent=''
    max_idx=max(c_date,c_seg,c_cost,c_code if c_code is not None else 0)
    for row in ws.iter_rows(min_row=header_row+1,max_col=max_idx+1,values_only=True):
        dt=row[c_date]; seg=norm(row[c_seg]); m=month_key(dt)
        if not seg or not m: continue
        months.add(m)
        code=norm(row[c_code]) if c_code is not None else ''
        ma=re.match(r'^(\d{2})\s*[-～~]',seg)
        if not ma and not code and compact(seg) not in ('total','合計'):
            parent=seg
        try: cost=float(row[c_cost] or 0)
        except Exception: cost=0.0
        normal[(m,compact(seg))] += cost
        if ma:
            ydo[(m,compact(parent),int(ma.group(1)))] += cost
    return normal,ydo,months

def build_cost_sheet(template_wb, report_wb):
    ws=template_wb['コストデータ']
    # Map requested media names to actual report sheets once.
    media_to_sheet={}
    wanted={norm(ws.cell(r,2).value) for r in range(2,ws.max_row+1) if ws.cell(r,2).value}
    for item2 in wanted:
        target='YDO' if item2=='Yahoo!ダイレクトオファー' else item2
        ct=compact(target)
        media_to_sheet[item2]=next((sn for sn in report_wb.sheetnames if ct and ct in compact(sn)),None)

    aggs={}; months=set()
    # Parse only sheets actually referenced by cost data.
    for sn in {x for x in media_to_sheet.values() if x}:
        normal,ydo,ms=parse_report_sheet_aggregated(report_wb[sn])
        aggs[sn]=(normal,ydo); months.update(ms)
    months=sorted(months)
    if not months: raise ValueError('ファイル2から配信年月を取得できませんでした。')

    old_max=ws.max_column; new_max=5+len(months)
    for c in range(6,max(old_max,new_max)+1):
        ws.cell(1,c).value = months[c-6] if c <= new_max else None
        if c>new_max:
            for r in range(2,ws.max_row+1): ws.cell(r,c).value=None
    base_width=ws.column_dimensions['F'].width or 12
    for c in range(6,new_max+1):
        ws.column_dimensions[get_column_letter(c)].width=base_width
        if c!=6: ws.cell(1,c)._style=copy.copy(ws['F1']._style)

    unmatched=set()
    for r in range(2,ws.max_row+1):
        item2=norm(ws.cell(r,2).value); sn=media_to_sheet.get(item2)
        if not sn: continue
        normal,ydo=aggs[sn]
        item3=norm(ws.cell(r,3).value); seg=norm(ws.cell(r,4).value); age=age_decade(ws.cell(r,5).value)
        if item2=='Yahoo!ダイレクトオファー':
            # Parent labels may contain extra words/parentheses, so resolve matching parents once per row.
            candidates={k[1] for k in ydo.keys() if (not item3 or compact(item3) in k[1]) and (not seg or compact(seg) in k[1])}
            for j,m in enumerate(months,6):
                ws.cell(r,j).value=sum(ydo.get((m,p,age),0.0) for p in candidates) if age is not None else 0
        else:
            wanted_seg=compact(seg or item3)
            for j,m in enumerate(months,6):
                if wanted_seg:
                    ws.cell(r,j).value=sum(v for (mm,ss),v in normal.items() if mm==m and wanted_seg in ss and ss not in ('total','合計'))
                else:
                    ws.cell(r,j).value=sum(v for (mm,ss),v in normal.items() if mm==m)
    return months, unmatched

def load_master_inputs(files):
    rows=[]
    for f in files:
        wb=load_workbook(io.BytesIO(f.getvalue()),read_only=True,data_only=True)
        ws=wb[wb.sheetnames[0]]
        for vals in ws.iter_rows(min_row=2,min_col=1,max_col=7,values_only=True):
            if any(v is not None for v in vals): rows.append(vals)
        wb.close()
    return rows

def build_master_sheet(template_wb, master_files):
    ws=template_wb['媒体コードマスタver3']
    # Capture existing D:G / I formulas by media code before replacing rows.
    formula_map={}
    for r in range(2,ws.max_row+1):
        code=norm(ws.cell(r,2).value)
        if code:
            formula_map[code]=tuple(ws.cell(r,c).value for c in (4,5,6,7,9)) + (r,)
    rows=load_master_inputs(master_files)
    old_max=ws.max_row
    # Clear A:L data only, preserve header and formatting skeleton.
    for r in range(2,max(old_max,len(rows)+1)+1):
        for c in range(1,13): ws.cell(r,c).value=None
    missing_formula=0
    for idx,vals in enumerate(rows,start=2):
        a,b,c,d,e,f,g=vals
        ws.cell(idx,1).value=a; ws.cell(idx,2).value=b; ws.cell(idx,3).value=c
        ws.cell(idx,8).value=d
        ws.cell(idx,10).value=e; ws.cell(idx,11).value=f; ws.cell(idx,12).value=g
        old=formula_map.get(norm(b))
        if old:
            for col,val in zip((4,5,6,7,9),old[:5]):
                if isinstance(val,str) and val.startswith('='):
                    try: ws.cell(idx,col).value=Translator(val,origin=f'{get_column_letter(col)}{old[5]}').translate_formula(f'{get_column_letter(col)}{idx}')
                    except Exception: ws.cell(idx,col).value=re.sub(r'(?<=[A-Z])\d+',str(idx),val)
                else: ws.cell(idx,col).value=val
        else:
            # Generic fallback only when the uploaded media code does not exist in the template.
            ws.cell(idx,4).value=f'=IFERROR(TEXTBEFORE(TEXTAFTER(C{idx},"_",1),"_"),"")'
            ws.cell(idx,5).value=f'=IFERROR(TEXTBEFORE(TEXTAFTER(C{idx},"_",2),"_"),"")'
            # F/G/I have no safe generic rule; leave them blank rather than inventing a formula.
            ws.cell(idx,9).value=f'=IFERROR(TEXTBEFORE(TEXTAFTER(C{idx},"_",3),"_"),"")'
            missing_formula+=1
    # Remove unused rows physically.
    if ws.max_row > len(rows)+1: ws.delete_rows(len(rows)+2, ws.max_row-(len(rows)+1))
    # The source template contains hidden rows; output all master rows visibly.
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].hidden = False
    return len(rows),missing_formula

def derive_master_media(ws):
    """Map media code -> media label. Prefer L (uploaded G/menu code), then parsed C."""
    out={}
    for r in range(2,ws.max_row+1):
        code=norm(ws.cell(r,2).value); label=norm(ws.cell(r,12).value)
        if not label:
            parts=norm(ws.cell(r,3).value).split('_')
            label=parts[2] if len(parts)>2 else ''
        if code: out[code]=label
    return out

def build_backend_sheet(template_wb, raw_file):
    ws=template_wb['後方数値データ(加工版)']
    raw_wb=load_workbook(io.BytesIO(raw_file.getvalue()),read_only=True,data_only=True)
    raw=raw_wb[raw_wb.sheetnames[0]]
    header_row=next(raw.iter_rows(min_row=1,max_row=1,min_col=1,max_col=30,values_only=True))
    expected=[ws.cell(1,c).value for c in range(1,31)]
    if [norm(x) for x in header_row] != [norm(x) for x in expected]:
        raw_wb.close(); raise ValueError('ファイル4のA:ADヘッダーがテンプレートと一致しません。')

    media_map=derive_master_media(template_wb['媒体コードマスタver3'])
    cost_ws=template_wb['コストデータ']
    month_cols={norm(cost_ws.cell(1,c).value):c for c in range(6,cost_ws.max_column+1) if cost_ws.cell(1,c).value}
    media_month_cost=defaultdict(float)
    for r in range(2,cost_ws.max_row+1):
        media=norm(cost_ws.cell(r,2).value)
        if not media: continue
        for m,c in month_cols.items():
            try: media_month_cost[(compact(media),m)] += float(cost_ws.cell(r,c).value or 0)
            except Exception: pass

    # Stream input in read-only mode; build compact derived metadata in one pass.
    raw_rows=[]; derived=[]; counts=defaultdict(int)
    for vals in raw.iter_rows(min_row=2,min_col=1,max_col=30,values_only=True):
        if not any(v is not None for v in vals): continue
        vals=tuple(vals); raw_rows.append(vals)
        code=norm(vals[2]); media=media_map.get(code,''); m=month_key(vals[1]); age=age_decade(vals[11])
        derived.append((media,m,age))
        if media and m: counts[(compact(media),m)] += 1
    raw_wb.close()

    # Keep header formatting, remove old body once, then append FINAL 35-column rows directly.
    # Final columns = A:AD + AG(media) + AK(cost) + AL(age) + AM(month) + AN(count).
    # Fast clear: delete the previous body directly, avoiding expensive row shifting.
    # Preserve header cells, worksheet dimensions, formatting and metadata.
    if ws.max_row > 1:
        ws._cells = {key: cell for key, cell in ws._cells.items() if key[0] == 1}
        for row_idx in list(ws.row_dimensions):
            if row_idx > 1:
                del ws.row_dimensions[row_idx]
    for vals,(media,m,age) in zip(raw_rows,derived):
        cnt=counts.get((compact(media),m),0)
        total=media_month_cost.get((compact(media),m),0.0)
        ws.append(vals + (media, total/cnt if cnt else 0, age, m, cnt))

    # Rewrite retained calculated headers because unwanted columns never get created.
    retained_headers = ['媒体','コスト','年齢グループ','申込月','YDO行数判定']
    for c,h in enumerate(retained_headers,start=31): ws.cell(1,c).value=h
    return len(raw_rows)

def finish_sheet_settings(wb):
    # Existing template formatting is retained. Avoid restyling millions of cells.
    # Remove worksheet filters as requested and keep sensible template widths.
    for ws in wb.worksheets:
        ws.auto_filter.ref = None
        ws.sheet_view.selection[0].activeCell = 'A1'
        ws.sheet_view.selection[0].sqref = 'A1'
        ws.sheet_view.topLeftCell = 'A1'
        if getattr(ws, 'tables', None):
            for table in ws.tables.values():
                try: table.autoFilter = None
                except Exception: pass
        for c in range(1, ws.max_column + 1):
            letter = get_column_letter(c)
            current = ws.column_dimensions[letter].width
            if current and current > 45:
                ws.column_dimensions[letter].width = 45

def save_fast(wb, out):
    # openpyxlの標準保存と同じExcelWriterを使い、ZIP圧縮レベルのみ変更。
    # Excelの内容や数式は変えず、CPU負荷を軽減する。
    with ZipFile(out, 'w', ZIP_DEFLATED, allowZip64=True, compresslevel=1) as archive:
        ExcelWriter(wb, archive).write_data()


def generate(template_file, report_file=None, master_files=None, raw_file=None, progress=None):
    def step(pct, msg):
        if progress: progress(pct, msg)

    step(5, 'テンプレートを読み込んでいます…')
    wb=load_workbook(io.BytesIO(template_file.getvalue()),data_only=False)
    try:
        wb.calculation.fullCalcOnLoad = False
        wb.calculation.forceFullCalc = False
        wb.calculation.calcMode = 'manual'
    except Exception:
        pass
    # Make newly created/un-styled cells use Meiryo UI without restyling millions of cells.
    if wb._fonts:
        base = wb._fonts[0]
        wb._fonts[0] = Font(name=FONT_NAME, size=base.sz or 11, bold=base.b, italic=base.i,
                            vertAlign=base.vertAlign, underline=base.u, strike=base.strike, color=base.color)
    months=[]; unmatched=set(); master_n=None; raw_n=None; missing_formula=0
    timings={}
    t0=time.perf_counter()

    if report_file is not None:
        step(15, 'メール広告レポートを集計しています…')
        report=load_workbook(io.BytesIO(report_file.getvalue()),read_only=True,data_only=True)
        tx=time.perf_counter(); months,unmatched=build_cost_sheet(wb,report); report.close(); timings['コスト集計']=time.perf_counter()-tx
    else:
        step(30, 'ファイル2未指定：コストデータを維持します…')

    if master_files:
        step(40, '媒体コードマスタを更新しています…')
        tx=time.perf_counter(); master_n,missing_formula=build_master_sheet(wb,master_files); timings['媒体コード']=time.perf_counter()-tx
    else:
        step(50, 'ファイル3未指定：媒体コードマスタを維持します…')

    if raw_file is not None:
        step(58, '後方数値ローデータを取り込んでいます…')
        tx=time.perf_counter(); raw_n=build_backend_sheet(wb,raw_file); timings['後方数値']=time.perf_counter()-tx
    else:
        step(72, 'ファイル4未指定：後方数値データを維持します…')

    step(78, '出力シートを整えています…')
    for sn in list(wb.sheetnames):
        if sn not in KEEP_SHEETS: del wb[sn]
    finish_sheet_settings(wb)

    step(88, 'Excelを書き出しています…')
    tx=time.perf_counter(); out=io.BytesIO(); save_fast(wb, out); out.seek(0); timings['Excel保存']=time.perf_counter()-tx
    timings['合計']=time.perf_counter()-t0
    step(100, '生成完了！')
    return out,months,master_n,raw_n,missing_formula,unmatched,timings

with st.form('files'):
    col1, col2 = st.columns(2)
    with col1:
        f1=st.file_uploader('ファイル1：後方数値分析ツール（必須）',type=['xlsx'])
        f3=st.file_uploader('ファイル3：媒体コード元データ（複数選択可）',type=['xlsx'],accept_multiple_files=True)
    with col2:
        f2=st.file_uploader('ファイル2：メール広告レポート（任意）',type=['xlsx'])
        f4=st.file_uploader('ファイル4：後方数値データ ローデータ（任意）',type=['xlsx'])
    st.caption('ファイル2〜4は未指定でも生成できます。未指定ファイルに対応するシートは、ファイル1の内容をそのまま維持します。')
    go=st.form_submit_button('Excelを生成',type='primary',use_container_width=True)

if go:
    if not f1:
        st.error('ファイル1（テンプレート）は必須です。')
    else:
        try:
            bar=st.progress(0,text='処理を開始します…')
            def update_progress(pct,msg):
                bar.progress(pct,text=msg)
            out,months,mn,rn,miss,unmatched,timings=generate(f1,f2,f3,f4,update_progress)
            st.balloons()
            parts=[]
            if f2 and months: parts.append(f'コスト {months[0]}〜{months[-1]}')
            else: parts.append('コストデータ：維持')
            if f3: parts.append(f'媒体コード {mn:,}行')
            else: parts.append('媒体コードマスタ：維持')
            if f4: parts.append(f'後方数値 {rn:,}行')
            else: parts.append('後方数値データ：維持')
            st.success('生成完了：' + ' / '.join(parts))
            st.caption('処理時間：' + ' / '.join(f'{k} {v:.1f}秒' for k,v in timings.items()))
            if miss:
                st.warning(f'媒体コードマスタ：既存テンプレートに同一媒体コードがなく、汎用式を使用した行が {miss:,} 行あります。D:G/I列をご確認ください。')
            st.download_button('生成Excelをダウンロード',data=out,file_name='後方数値分析ツール_生成.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
        except Exception as e:
            st.exception(e)
