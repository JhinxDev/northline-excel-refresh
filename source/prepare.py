"""Rebuild synthetic fixtures and workbook shell in a disposable project copy."""
from pathlib import Path
from collections import defaultdict, Counter
from datetime import date, datetime
import argparse, csv, hashlib, json, shutil
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
ORIGINAL = ROOT / 'fixtures/original-sales.csv'
FIELDS = ['order_id','date','channel','sku','product','category','quantity','unit_price_cents','unit_cost_cents']

def write_csv(path, rows, fields=FIELDS):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', newline='', encoding='utf-8') as f:
        w=csv.DictWriter(f, fieldnames=fields); w.writeheader(); w.writerows(rows)

def fixtures():
    rows=list(csv.DictReader(ORIGINAL.open(encoding='utf-8')))
    bundled=ROOT/'fixtures/original-sales.csv'
    if ORIGINAL.resolve()!=bundled.resolve(): shutil.copyfile(ORIGINAL,bundled)
    by_month=defaultdict(list)
    for row in rows: by_month[row['date'][:7]].append(row)
    for month, records in by_month.items(): write_csv(ROOT/f'fixtures/monthly/{month}.csv', records)
    master={r['sku']:{k:r[k] for k in ['sku','product','category']} for r in rows}
    write_csv(ROOT/'fixtures/products.csv', sorted(master.values(), key=lambda r:r['sku']), ['sku','product','category'])
    october=[]
    for n, sku in enumerate(sorted(master)):
        exemplar=next(r for r in rows if r['sku']==sku)
        october.append({**exemplar,'order_id':f'OCT{n+1:04d}','date':f'2026-10-{n+1:02d}','quantity':str(n%4+1),'channel':['Online','Retail','Wholesale'][n%3]})
    def bad(oid, **values): return {**october[0], 'order_id':oid, **values}
    october.extend([
        bad('DUP-OCT'), bad('DUP-OCT'),
        bad('UNKNOWN-SKU',sku='ZZ99'), bad('BAD-QTY',quantity='two'),
        bad('BAD-DATE',date='2026-02-30'), bad('BAD-CHANNEL',channel='Partner'),
        bad('BAD-PRICE',unit_price_cents='-50'), bad('NAME-MISMATCH',product='Wrong product'),
        bad('',quantity='3'), bad('DECIMAL-QTY',quantity='1.5'), bad('   '),
    ])
    write_csv(ROOT/'fixtures/incoming/2026-10.csv',october)
    # Independent Python oracle, not workbook formulas or M output.
    def expected(data):
        counts=Counter(r['order_id'] for r in data); valid=[]; exceptions=[]; months=defaultdict(lambda:dict(orders=0,revenue_cents=0,cost_cents=0))
        def integer(s,positive=False): return bool(s) and s.isascii() and s.isdigit() and int(s)<10**9 and (not positive or int(s)>0)
        raw=0; rejected=0; unvalued=0
        for r in data:
            errors=[]
            if not r['order_id'].strip(): errors.append('missing_order_id')
            elif counts[r['order_id']]>1: errors.append('duplicate_order_id')
            try:
                parsed=date.fromisoformat(r['date'])
                if parsed.isoformat()!=r['date'] or not 2020<=parsed.year<=2100: raise ValueError()
            except ValueError: errors.append('invalid_date')
            if r['channel'] not in ['Online','Retail','Wholesale']: errors.append('invalid_channel')
            product=master.get(r['sku'])
            if product is None: errors.append('unmatched_sku')
            elif (product['product'],product['category'])!=(r['product'],r['category']): errors.append('product_mismatch')
            for key, flag in [('quantity',True),('unit_price_cents',False),('unit_cost_cents',False)]:
                if not integer(r[key],flag): errors.append('invalid_'+key)
            valued=integer(r['quantity'],True) and integer(r['unit_price_cents']) and integer(r['unit_cost_cents'])
            amount=int(r['quantity'])*int(r['unit_price_cents']) if valued else 0
            raw+=amount; unvalued+=not valued
            if errors:
                exceptions.append(dict(order_id=r['order_id'],issues='; '.join(errors))); rejected+=amount
            else:
                valid.append(r); m=months[r['date'][:7]];m['orders']+=1;m['revenue_cents']+=amount;m['cost_cents']+=int(r['quantity'])*int(r['unit_cost_cents'])
        return dict(input_rows=len(data),accepted_rows=len(valid),exception_rows=len(exceptions),unvalued_rows=unvalued,raw_valued_revenue_cents=raw,excluded_valued_revenue_cents=rejected,revenue_cents=sum(m['revenue_cents'] for m in months.values()),cost_cents=sum(m['cost_cents'] for m in months.values()),months=dict(months),exceptions=exceptions)
    (ROOT/'evidence/expected.json').write_text(json.dumps({'baseline':expected(rows),'october':expected(rows+october)},indent=2))

def shell():
    book=xlsxwriter.Workbook(ROOT/'northline-refresh.xlsx')
    book.set_properties({'title':'Northline monthly refresh v1.1.0','author':'JhinxDev','comments':'Synthetic portfolio exercise. AI-assisted implementation. Native Power Query, no VBA.'})
    title=book.add_format({'font_name':'Aptos Display','font_size':26,'font_color':'#183A3D','bold':True})
    note=book.add_format({'font_name':'Aptos','font_size':11,'font_color':'#456164','text_wrap':True,'valign':'top'})
    banner=book.add_format({'font_name':'Aptos','font_size':11,'font_color':'#FFFFFF','bg_color':'#176C61','bold':True})
    for name,heading,sub in [
        ('Overview','A repeatable monthly close.','NORTHLINE SUPPLY / SYNTHETIC PORTFOLIO EXERCISE'),
        ('Monthly','Sales by month','ACCEPTED ORDERS ONLY / CAD'),
        ('Exceptions','Rows that need attention','QUARANTINED / EXCLUDED FROM SALES TOTALS'),
        ('Accepted','Validated sales','ONE ROW PER ORDER / SYNTHETIC DATA'),
        ('Reconciliation','Every row accounted for','INTEGER CENTS / NO SILENT ROW REMOVAL')]:
        s=book.add_worksheet(name);s.hide_gridlines(2);s.set_zoom(90);s.set_column('A:A',3);s.set_column('B:L',16);s.set_row(1,36)
        s.merge_range('B2:L2',heading,title);s.merge_range('B4:L4',sub,banner);s.freeze_panes(6,2)
        s.set_landscape();s.set_paper(9);s.fit_to_pages(1,1);s.set_margins(.3,.3,.4,.4)
    s=book.get_worksheet_by_name('Overview');s.activate();s.set_column('B:L',13)
    guidance=[
        '01  Add the next monthly CSV to the designated monthly folder. Keep the exact header and one product per order.',
        '02  In Excel, Data > Refresh All. Wait until Queries & Connections shows completion. A failed refresh can leave older rows visible.',
        '03  Read Reconciliation and Exceptions before using Monthly. All copies of a repeated order ID are excluded until corrected.',
        '04  Correct the source CSV, then refresh again. Never repair loaded result tables: the next refresh replaces them.',
        'First setup: Data > Queries & Connections > DataRoot > Edit > Advanced Editor. Set the local fixtures folder path in quotes.',
        'The workbook needs local folder access and Excel Power Query. It has no VBA, cloud accounts, scheduled updates or live website connection.',
        'Revenue less cost of goods is gross profit. Shipping, overhead, tax and returns are outside this exercise. All records are synthetic.',
    ]
    for i,line in enumerate(guidance): s.merge_range(6+i*3,1,7+i*3,11,line,note)
    s.print_area('B2:L28')
    book.close()

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--sales-source',type=Path)
    args=parser.parse_args()
    if args.sales_source: ORIGINAL=args.sales_source
    fixtures();shell()
