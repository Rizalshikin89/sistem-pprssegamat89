from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from copy import copy
import shutil, subprocess, zipfile, re, os

BASE = Path(__file__).resolve().parent
TEMPLATES = BASE / 'templates'
OUTPUT = BASE / 'output'
OUTPUT.mkdir(exist_ok=True)
app = Flask(__name__)
app.secret_key = os.environ.get('PPRS_SECRET', 'local-only-change-me')

DOCS = [
 ('lamp-r','Lampiran R – Perakuan Serahan Siap Kerja','20. LAMP.R  - PERAKUAN SERAHAN SIAP KERJA.docx','docx'),
 ('lamp-h','Lampiran H – Borang Pengesahan Pendapatan','9. LAMP. H - BORANG PENGESAHAN PENDAPATAN (Update 22 Februari 2024).docx','docx'),
 ('lamp-b','Lampiran B – Kebenaran Bertulis Guna Tanah','LAMP. B - BORANG KEBENARAN BERTULIS GUNA TANAH.docx','docx'),
 ('lamp-f','Lampiran F – Borang Permohonan PPRT','23. LAMP. F - BORANG PERMOHONAN PPRT.xlsx','xlsx'),
 ('doa','Surat Ikatan Serah Hak / Pelantikan Sub Kontraktor','(D.O.A) SURAT PERJANJIAN PELANTIKAN SEBAGAI SUB KONTRAKTOR(1).docx','docx'),
 ('ibs','Surat Setuju Terima (IBS)','SURAT SETUJU TERIMA (IBS)(1).docx','docx'),
 ('sst','Surat Setuju Terima Tawaran PPRS','SURAT SETUJU-TERIMA TAWARAN PPRS.docx','docx'),
 ('penilaian','Borang Penilaian Kontraktor','10. 08 GMBPP PPRT - Borang Penilaian Kontraktor.xlsx','xlsx'),
 ('eot','Lampiran 13 – Borang Permohonan EOT','37. 16 - BORANG PERMOHONAN PELANJUTAN TEMPOH MASA (EOT).docx','docx'),
 ('lamp-n','Lampiran N – Profil Peserta','15. LAMP. N - PROFIL PESERTA PSERTA SEBELUM SEMASA SELEPAS Update Januari 2024.docx','docx'),
 ('q1','Lampiran Q1 – Perakuan Siap Kerja Sub Kontraktor','LAMP. Q1 - PERAKUAN SIAP KERJA SUB KONTRAKTOR - Copy.doc','doc'),
]
FIELDS = [
 ('nama','Nama penerima'),('ic','No. kad pengenalan'),('alamat','Alamat rumah / alamat penerima'),('telefon','No. telefon'),('umur','Umur'),('jantina','Jantina'),('bangsa','Bangsa'),('kaum','Kaum'),('status','Taraf perkahwinan'),('usia_perkahwinan','Usia perkahwinan (tahun)'),('pekerjaan','Pekerjaan'),('pendapatan','Pendapatan bulanan (RM)'),('nama_pasangan','Nama pasangan'),('ic_pasangan','No. K/P pasangan'),('pekerjaan_pasangan','Pekerjaan pasangan'),('pendapatan_pasangan','Pendapatan pasangan (RM)'),('bil_anak','Bilangan anak'),('tanggungan','Bilangan tanggungan'),('harta','Maklumat pemilikan harta KIR & pasangan'),('kampung','Kampung'),('jkkk','JKKK'),('negeri','Negeri'),('daerah','Daerah'),('parlimen','Parlimen'),('dun','DUN'),('jenis_projek','Jenis projek (Bina Baharu / Baik Pulih)'),('tahun','Tahun'),('tarikh_mula','Tarikh mula'),('tarikh_siap','Tarikh siap'),('tarikh_serahan','Tarikh serahan siap kerja'),('tarikh_serahan_penerima','Tarikh serahan kepada penerima'),('no_po','No. Pesanan Belian / PO'),('kos_projek','Jumlah kos projek (RM)'),('nama_kontraktor','Nama kontraktor / syarikat'),('alamat_kontraktor','Alamat kontraktor'),('no_rujukan','No. rujukan surat tawaran'),('tarikh_surat','Tarikh surat tawaran'),('nama_pemilik_tanah','Nama pemilik tanah'),('ic_pemilik_tanah','No. K/P pemilik tanah'),('lot','No. lot'),('tempat','Tempat'),('mukim','Mukim'),('tarikh_dokumen','Tarikh dokumen (Lampiran B)'),('hari','Hari'),('bulan','Bulan'),('tempoh_tanah','Tempoh kebenaran menduduki (tahun)'),('nama_subkontraktor','Nama sub-kontraktor'),('alamat_subkontraktor','Alamat sub-kontraktor'),('jumlah_subkontraktor','Jumlah bayaran sub-kontraktor (RM)'),('no_sijil_giatmara','No. sijil GIATMARA'),('no_kod_program','No. kod program'),('gps_n','Koordinat GPS N'),('gps_s','Koordinat GPS S'),('bil_ahli_isi_rumah','Bilangan ahli isi rumah'),('jenis_bantuan','Jenis bantuan'),('kesediaan_tapak','Kesediaan tapak'),('kriteria_pemohon','Kriteria pemohon'),('status_ekasih','Status e-Kasih'),('tarikh_permohonan','Tarikh permohonan'),('sebab_eot','Punca kelewatan'),('tarikh_akhir_eot','Tarikh akhir dipohon (EOT)'),
]

def val(data,k): return (data.get(k) or '').strip()
def blacken(run):
    run.font.color.rgb = __import__('docx').shared.RGBColor(0,0,0)

def replace_in_paragraph(p, old, new):
    """Replace text across runs without rebuilding the paragraph; inserted text inherits first placeholder run formatting then turns black."""
    if not old: return False
    runs = p.runs
    full = ''.join(r.text for r in runs)
    pos = full.find(old)
    if pos < 0: return False
    end = pos + len(old)
    offsets=[]; n=0
    for r in runs:
        offsets.append((n,n+len(r.text))); n += len(r.text)
    first = next((i for i,(a,b) in enumerate(offsets) if b>pos), None)
    last = next((i for i,(a,b) in enumerate(offsets) if b>=end), None)
    if first is None: return False
    a,b=offsets[first]; r=runs[first]
    left=r.text[:max(0,pos-a)]
    if first==last:
        r.text=left+new+r.text[end-a:]
        blacken(r)
    else:
        last_run=runs[last]; la,lb=offsets[last]
        right=last_run.text[max(0,end-la):]
        r.text=left+new; blacken(r)
        for j in range(first+1,last): runs[j].text=''
        last_run.text=right
    return True

def replace_first_ellipsis(p, new):
    # Use the first contiguous placeholder sequence in the paragraph.
    m=re.search(r'[.…_]{3,}', ''.join(r.text for r in p.runs))
    if not m: return False
    old=m.group(0)
    return replace_in_paragraph(p, old, new)

def iter_all_paragraphs(doc):
    for p in doc.paragraphs: yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs: yield p

def set_cell_preserve(cell, text):
    # Keep cell, paragraph and run properties. Only change text in first existing run.
    p=cell.paragraphs[0]
    if p.runs:
        p.runs[0].text=str(text); blacken(p.runs[0])
        for r in p.runs[1:]: r.text=''
    else:
        r=p.add_run(str(text)); blacken(r)

def fill_docx(key, path, d):
    doc=Document(path)
    if key=='lamp-r':
        t=doc.tables[0]
        keys=['nama','ic','alamat','tarikh_serahan','tarikh_serahan_penerima']
        for row,k in zip(t.rows,keys):
            if val(d,k): set_cell_preserve(row.cells[1],val(d,k))
    elif key=='lamp-h':
        ps=doc.paragraphs
        if val(d,'nama'): replace_in_paragraph(ps[11],'……………………………………………………………',val(d,'nama'))
        if val(d,'ic'): replace_in_paragraph(ps[11],'…………………………',val(d,'ic'))
        if val(d,'pendapatan'): replace_in_paragraph(ps[19],'………………',val(d,'pendapatan'))
        # User specified no signature / external verification fields; those remain untouched.
    elif key=='lamp-b':
        # Replace only the explicitly identified blanks in the opening permission paragraph.
        if val(d,'hari'): replace_first_ellipsis(doc.paragraphs[7],val(d,'hari'))
        if val(d,'bulan'): replace_first_ellipsis(doc.paragraphs[7],val(d,'bulan'))
        if val(d,'tahun'): replace_first_ellipsis(doc.paragraphs[9],val(d,'tahun'))
        if val(d,'nama_pemilik_tanah'): replace_in_paragraph(doc.paragraphs[9],'………......................………………………………….......……….',val(d,'nama_pemilik_tanah'))
        if val(d,'ic_pemilik_tanah'): replace_first_ellipsis(doc.paragraphs[12],val(d,'ic_pemilik_tanah'))
        if val(d,'nama'): replace_first_ellipsis(doc.paragraphs[14],val(d,'nama'))
        if val(d,'ic'): replace_first_ellipsis(doc.paragraphs[17],val(d,'ic'))
        # Lot/place/mukim/daerah/negeri in the same paragraph, in order.
        for k in ['lot','tempat','mukim','daerah','negeri']:
            if val(d,k): replace_first_ellipsis(doc.paragraphs[17],val(d,k))
        if val(d,'tempoh_tanah'): replace_first_ellipsis(doc.paragraphs[24],val(d,'tempoh_tanah'))
        # Signature and witness areas deliberately not populated.
    elif key=='lamp-n':
        mapping={5:'tahun',6:'negeri',9:'nama',10:'ic',11:'alamat',12:'telefon',13:'gps_n',14:'pekerjaan',15:'status',16:'nama_kontraktor',17:'tarikh_mula'}
        for idx,k in mapping.items():
            if val(d,k): replace_first_ellipsis(doc.paragraphs[idx],val(d,k))
        if val(d,'daerah'): replace_in_paragraph(doc.paragraphs[6],'…………………',val(d,'daerah'))
        if val(d,'umur'): replace_in_paragraph(doc.paragraphs[10],'……………..',val(d,'umur'))
        if val(d,'gps_s'): replace_first_ellipsis(doc.paragraphs[13],val(d,'gps_s'))
        if val(d,'pendapatan'): replace_first_ellipsis(doc.paragraphs[14],val(d,'pendapatan'))
        if val(d,'bil_ahli_isi_rumah'): replace_first_ellipsis(doc.paragraphs[15],val(d,'bil_ahli_isi_rumah'))
        if val(d,'tarikh_siap'): replace_first_ellipsis(doc.paragraphs[17],val(d,'tarikh_siap'))
    elif key=='eot':
        t=doc.tables[0]
        pairs={1:('jenis_projek',),2:('no_po',),3:('kos_projek',),4:('tarikh_mula',),5:('tarikh_akhir_eot',),6:('nama',),7:('alamat',)}
        for ri,ks in pairs.items():
            k=ks[0]
            if val(d,k): set_cell_preserve(t.rows[ri].cells[1],val(d,k))
        # Delay reason checkbox / comments are not altered automatically.
    elif key=='sst':
        # First original copy is the active form; second copy retained unchanged.
        ps=doc.paragraphs
        repl=[(15,'…………….....', 'no_rujukan'),(15,'…………..', 'tarikh_surat'),(15,'…………………………………………………………','nama'),(16,'…………………………………………………………………………………………','alamat_kontraktor'),(19,'……………………………..','no_po'),(20,'………………….','tarikh_permohonan')]
        for pi,needle,k in repl:
            if val(d,k): replace_in_paragraph(ps[pi],needle,val(d,k))
    elif key=='ibs':
        ps=doc.paragraphs
        if val(d,'nama'): replace_first_ellipsis(ps[3],val(d,'nama'))
        if val(d,'ic'): replace_first_ellipsis(ps[3],val(d,'ic'))
        if val(d,'nama_kontraktor'): replace_first_ellipsis(ps[4],val(d,'nama_kontraktor'))
        if val(d,'jenis_projek'): replace_first_ellipsis(ps[4],val(d,'jenis_projek'))
        if val(d,'no_po'): replace_first_ellipsis(ps[13],val(d,'no_po'))
    elif key=='doa':
        # Do not alter the sample names, addresses or amounts until project-specific values and the exact editable fields are confirmed.
        pending_note = 'Surat Ikatan Serah Hak: pemetaan nama/alamat/nilai kontrak belum disahkan; salinan templat dikekalkan.'
        pass
    doc.save(path)

def fill_xlsx(key,path,d):
    wb=load_workbook(path)
    ws=wb[wb.sheetnames[0]]
    if key=='lamp-f':
        cells={'M18':'nama','M20':'ic','M22':'umur','M26':'usia_perkahwinan','M30':'bangsa','M32':'kaum','M34':'pekerjaan','M36':'pendapatan','M38':'nama_pasangan','M40':'ic_pasangan','M42':'pekerjaan_pasangan','M44':'pendapatan_pasangan','M46':'bil_anak','M48':'tanggungan','M106':'alamat','M111':'kampung','M113':'jkkk','M115':'negeri','M117':'daerah','M119':'parlimen','M121':'dun','M135':'tarikh_permohonan'}
        for cell,k in cells.items():
            if val(d,k): ws[cell]=val(d,k)
        # Choice fields: mark the matching existing option; do not alter labels/layout.
        choices={'status':{'Bujang':'N24','Berkahwin':'Q24','Janda':'U24','Duda':'X24'},'jantina':{'Lelaki':'N28','Perempuan':'Q28'},'jenis_bantuan':{'Bina Baharu':'N125','Baik Pulih':'R125'}}
        for k,opts in choices.items():
            v=val(d,k)
            if v in opts:
                c=ws[opts[v]]; c.value='✓ '+str(c.value)
        if val(d,'nama') or val(d,'ic'):
            # Declaration sentence is in one original cell; replace only underline placeholders.
            c=ws['C143']; s=str(c.value)
            if val(d,'nama'): s=s.replace('Saya __________________________________________','Saya '+val(d,'nama'))
            if val(d,'ic'): s=s.replace('No. K/P ____________________________________','No. K/P '+val(d,'ic'))
            c.value=s
    elif key=='penilaian':
        for cell,k in {'H7':'nama_kontraktor','H8':'no_sijil_giatmara','J12':'jenis_projek','H14':'no_kod_program','H16':'nama','H18':'daerah'}.items():
            if val(d,k): ws[cell]=val(d,k)
    wb.save(path)

def convert_pdf(src):
    outdir=src.parent
    try:
        subprocess.run(['libreoffice','--headless','--convert-to','pdf','--outdir',str(outdir),str(src)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=90)
        p=outdir/(src.stem+'.pdf')
        return p if p.exists() else None
    except Exception:
        return None

@app.route('/')
def index(): return render_template('index.html',docs=DOCS,fields=FIELDS)

@app.route('/generate',methods=['POST'])
def generate():
    data={k:request.form.get(k,'').strip() for k,_ in FIELDS}
    keys=request.form.getlist('docs')
    if not keys: flash('Pilih sekurang-kurangnya satu dokumen.'); return redirect(url_for('index'))
    person=re.sub(r'[^A-Za-z0-9_-]+','_',data.get('nama') or 'Penerima_PPRS')[:45]
    job=OUTPUT/person
    if job.exists(): shutil.rmtree(job)
    job.mkdir(parents=True)
    made=[]; pending=[]
    for key in keys:
        item=next((x for x in DOCS if x[0]==key),None)
        if not item: continue
        _,label,filename,kind=item; src=TEMPLATES/filename
        if not src.exists(): pending.append(label+' (templat tiada)'); continue
        dest=job/filename
        shutil.copy2(src,dest)
        try:
            if kind=='docx': fill_docx(key,dest,data)
            elif kind=='xlsx': fill_xlsx(key,dest,data)
            else: pending.append(label+' (format .doc lama – disalin tanpa perubahan; belum ada isian automatik)')
            if key=='doa': pending.append(label+' (ruangan nama/alamat/nilai perlu disahkan sebelum pemetaan automatik)')
            if key=='lamp-f': pending.append(label+' (pilihan harta dan beberapa pilihan bertanda perlu pemetaan/semakan lanjut)')
            if key=='penilaian': pending.append(label+' (ruangan markah penilaian belum dipetakan; hanya maklumat asas diisi)')
            if key=='eot' and val(data,'sebab_eot'): pending.append(label+' (punca kelewatan belum ditanda automatik)')
            pdf=convert_pdf(dest)
            made.append(dest)
            if pdf: made.append(pdf)
            else: pending.append(label+' (PDF belum dapat dijana; fail Microsoft disediakan)')
        except Exception as e:
            pending.append(label+' (salinan asal disediakan; isian gagal: '+str(e)+')')
    bundle=OUTPUT/(person+'_Dokumen_PPRS.zip')
    with zipfile.ZipFile(bundle,'w',zipfile.ZIP_DEFLATED) as z:
        for p in made:
            if p.exists(): z.write(p,p.name)
        z.writestr('STATUS_SEMAKAN.txt','Dokumen ini ialah DRAF UNTUK SEMAKAN. Semak setiap fail berbanding templat asal sebelum pengesahan.\n\nBelum lengkap / tindakan diperlukan:\n'+'\n'.join('- '+x for x in pending))
    flash('Fail draf dijana. Semak STATUS_SEMAKAN.txt dalam ZIP. Dokumen belum disahkan untuk kegunaan rasmi.')
    return send_file(bundle,as_attachment=True,download_name=bundle.name)

if __name__=='__main__':
    app.run(host='127.0.0.1',port=5000,debug=False)
