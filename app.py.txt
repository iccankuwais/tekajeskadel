import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import io
import random
import re
import subprocess
import tempfile
import os

# ==========================================
# 1. DATABASE KURIKULUM MERDEKA TKJ
# ==========================================
DATA_KURIKULUM_TKJ = {
    "Fase E (Kelas X - Dasar-dasar TJKT)": {
        "elemen": {
            "E1. Proses Bisnis di Bidang TJKT": {
                "jp": 48,
                "cp": "Memahami proses bisnis pada bidang teknik jaringan komputer dan telekomunikasi, meliputi perencanaan, analisis kebutuhan pelanggan, strategi implementasi, pelayanan pelanggan, dan penerapan K3LH.",
                "materi": ["Wawasan industri dan Technopreneurship", "Analisis kebutuhan pelanggan", "Manajemen proyek IT", "K3LH"],
                "deep_learning": {
                    "mindful": "Siswa merefleksikan peran etika profesi sebelum menyentuh perangkat keras.",
                    "meaningful": "Studi kasus langsung merancang proposal penawaran jaringan UMKM.",
                    "joyful": "Simulasi role-play konsultan IT dan klien."
                },
                "asesmen_soal": [
                    {"tanya": "1. Jelaskan 3 komponen utama SLA!", "jawab": "Layanan yang dijanjikan, waktu respons, dan kompensasi jika gagal."},
                    {"tanya": "2. Langkah analisis kebutuhan Wi-Fi kafe?", "jawab": "Survei lokasi, estimasi pengguna, dan pemilihan ISP."},
                    {"tanya": "3. Sebutkan penerapan 5R di ruang server!", "jawab": "Ringkas, Rapi, Resik, Rawat, Rajin di area perangkat."}
                ],
                "praktik_indikator": "Menyusun dokumen survei kebutuhan klien."
            },
            "E5. Dasar-Dasar Teknik Jaringan Komputer": {
                "jp": 108,
                "cp": "Memahami prinsip dasar sistem IPV4/IPV6, TCP/IP, Networking Service, dan Media Transmisi.",
                "materi": ["OSI Layer", "IPv4 & Subnetting", "Crimping UTP & FO", "Sistem Seluler & WLAN"],
                "deep_learning": {
                    "mindful": "Ketelitian langkah saat splicing Fiber Optic.",
                    "meaningful": "Merancang IP Addressing gedung bertingkat.",
                    "joyful": "Estafet crimping kabel tercepat."
                },
                "asesmen_soal": [
                    {"tanya": "1. Hitung Network ID IP 192.168.10.65/26!", "jawab": "Network ID: 192.168.10.64"},
                    {"tanya": "2. Urutan warna UTP standar 568B?", "jawab": "PO, O, PH, B, PB, H, PC, C"},
                    {"tanya": "3. Fungsi Layer 3 (Network) pada OSI?", "jawab": "Mengatur routing dan pengalamatan logis (IP)."}
                ],
                "praktik_indikator": "Melakukan crimping UTP lulus uji LAN Tester."
            }
        }
    },
    "Fase F (Kelas XI & XII - Konsentrasi TJKT)": {
        "elemen": {
            "F2. Teknologi Jaringan Kabel dan Nirkabel": {
                "jp": 216,
                "cp": "Menginstalasi dan mengkonfigurasi jaringan nirkabel, VLAN, serta routing statis dan dinamis.",
                "materi": ["Fiber Optic (Splicing/OTDR)", "VLAN Switch", "Routing OSPF", "Standar WLAN"],
                "deep_learning": {
                    "mindful": "Menerapkan K3LH ketat saat memotong core FO.",
                    "meaningful": "Membangun arsitektur RT/RW Net lengkap.",
                    "joyful": "Kompetisi estafet VLAN tercepat."
                },
                "asesmen_soal": [
                    {"tanya": "1. Perbedaan routing statis dan OSPF?", "jawab": "OSPF otomatis mencari jalur alternatif saat link mati, statis harus diubah manual."},
                    {"tanya": "2. Kelebihan frekuensi 5.8 GHz?", "jawab": "Minim interferensi, bandwidth lebih besar, namun jarak tembus halangan kurang."},
                    {"tanya": "3. Fungsi VLAN Trunking?", "jawab": "Membawa banyak ID VLAN berbeda melewati satu link fisik antar switch."}
                ],
                "praktik_indikator": "Berhasil melakukan splicing FO dan konfigurasi OSPF."
            },
            "F3. Sistem Keamanan Jaringan": {
                "jp": 144,
                "cp": "Mengkonfigurasi firewall, mengevaluasi sistem keamanan, dan manajemen bandwidth.",
                "materi": ["Firewall ACL & NAT", "Arsitektur DMZ", "Manajemen Bandwidth", "Konfigurasi VPN"],
                "deep_learning": {
                    "mindful": "Menganalisis traffic jaringan sebelum menerapkan blokir.",
                    "meaningful": "Mengamankan server ujian dengan DMZ.",
                    "joyful": "Simulasi Role-Play Red Team vs Blue Team."
                },
                "asesmen_soal": [
                    {"tanya": "1. Cara kerja stateful inspection?", "jawab": "Memantau status sesi koneksi aktif, tidak hanya header paket."},
                    {"tanya": "2. Logika Simple Queue 2 Mbps?", "jawab": "Target IP/Subnet, Max Limit=2M."},
                    {"tanya": "3. Alasan server publik ditaruh di DMZ?", "jawab": "Agar peretas terisolasi dan tidak bisa masuk ke LAN internal."}
                ],
                "praktik_indikator": "Mengimplementasikan Firewall Rules & Simple Queue pada MikroTik."
            }
        }
    }
}

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================
def set_cell_background(cell, fill_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def convert_docx_to_pdf(docx_bytes):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "doc.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        try:
            subprocess.run([
                "libreoffice", "--headless", "--convert-to", "pdf", 
                "doc.docx", "--outdir", tmpdir
            ], cwd=tmpdir, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            pdf_path = os.path.join(tmpdir, "doc.pdf")
            with open(pdf_path, "rb") as f:
                return f.read()
        except:
            return None

def create_header(doc, logo_kiri, logo_kanan):
    header = doc.sections[0].header
    header.paragraphs[0].text = ""
    htable = header.add_table(rows=1, cols=3, width=Inches(6.5))
    htable.columns[0].width = Inches(1.0)
    htable.columns[1].width = Inches(4.5)
    htable.columns[2].width = Inches(1.0)
    
    # Logo Kiri
    run_kiri = htable.cell(0, 0).paragraphs[0].add_run()
    if logo_kiri:
        try: run_kiri.add_picture(logo_kiri, width=Inches(0.9))
        except: pass
    htable.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Teks Kop
    p_text = htable.cell(0, 1).paragraphs[0]
    p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_text.add_run("PEMERINTAH PROVINSI SULAWESI SELATAN\nDINAS PENDIDIKAN\n").bold = True
    p_text.runs[0].font.size = Pt(12)
    p_text.add_run("UPT SMK NEGERI 8 JENEPONTO\n").bold = True
    p_text.runs[1].font.size = Pt(16)
    p_text.add_run("Program Keahlian: Teknik Jaringan Komputer dan Telekomunikasi\n").font.size = Pt(11)
    p_text.add_run("Alamat: Jl. Poros Jeneponto - Bantaeng, Sulawesi Selatan\n").font.size = Pt(10)
    
    # Logo Kanan
    run_kanan = htable.cell(0, 2).paragraphs[0].add_run()
    if logo_kanan:
        try: run_kanan.add_picture(logo_kanan, width=Inches(0.9))
        except: pass
    htable.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    header.add_paragraph().add_run("_" * 75).bold = True
    doc.add_paragraph("")

# ==========================================
# 3. GENERATORS
# ==========================================
def generate_docx(fase, elemen_nama, elemen_data, logo_kiri=None, logo_kanan=None):
    doc = Document()
    create_header(doc, logo_kiri, logo_kanan)
    
    # MODUL AJAR
    doc.add_heading(f"Modul Ajar TJKT - {fase}", level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(elemen_nama, level=1)
    
    doc.add_heading('Capaian Pembelajaran (CP)', level=2)
    doc.add_paragraph(elemen_data['cp'])
    
    doc.add_heading('Materi Pembelajaran', level=2)
    for materi in elemen_data['materi']: doc.add_paragraph(materi, style='List Bullet')
        
    doc.add_heading('Pendekatan Deep Learning', level=2)
    doc.add_paragraph(f"Mindful: {elemen_data['deep_learning']['mindful']}", style='List Bullet')
    doc.add_paragraph(f"Meaningful: {elemen_data['deep_learning']['meaningful']}", style='List Bullet')
    doc.add_paragraph(f"Joyful: {elemen_data['deep_learning']['joyful']}", style='List Bullet')
    
    doc.add_heading('Indikator Praktik', level=2)
    doc.add_paragraph(elemen_data['praktik_indikator'])
    
    # KKTP
    doc.add_heading('KKTP (Skala Interval)', level=2)
    kktp_table = doc.add_table(rows=1, cols=3)
    kktp_table.style = 'Table Grid'
    kktp_hdr = kktp_table.rows[0].cells
    for i, title in enumerate(['Interval Nilai', 'Kriteria Ketuntasan', 'Tindak Lanjut']):
        kktp_hdr[i].text = title
        kktp_hdr[i].paragraphs[0].runs[0].bold = True
        kktp_hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(kktp_hdr[i], 'D9D9D9')
        
    kktp_data = [("0% - 40%", "Belum Tuntas", "Remedial seluruh materi"), 
                 ("41% - 65%", "Belum Tuntas", "Remedial materi tertentu"), 
                 ("66% - 85%", "Tuntas", "Lanjut materi berikutnya"), 
                 ("86% - 100%", "Tuntas (Sangat Baik)", "Pengayaan / Tutor Sebaya")]
    for item in kktp_data:
        row = kktp_table.add_row().cells
        row[0].text, row[1].text, row[2].text = item[0], item[1], item[2]
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n*Ketuntasan minimal berada pada interval 66% - 85%.").runs[0].italic = True

    # JOB SHEET
    doc.add_page_break()
    doc.add_heading('LEMBAR KERJA (JOB SHEET) PRAKTIK', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Materi: {elemen_nama}\n")
    doc.add_paragraph("Nama Siswa\t: ..............................................................")
    doc.add_paragraph("Kelas\t\t: ..............................................................")
    doc.add_heading('A. Tujuan Praktik', level=2)
    doc.add_paragraph(f"Siswa mampu: {elemen_data['praktik_indikator']}")
    doc.add_heading('B. Hasil Pengujian & Kesimpulan', level=2)
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    t.rows[0].cells[0].text = "Tuliskan hasil analisis di sini...\n\n\n"
    
    # RUBRIK
    doc.add_heading('C. Rubrik Penilaian Praktik', level=2)
    rubrik = doc.add_table(rows=1, cols=5)
    rubrik.style = 'Table Grid'
    rhdr = rubrik.rows[0].cells
    for i, title in enumerate(['Aspek', 'Sangat Baik (4)', 'Baik (3)', 'Cukup (2)', 'Kurang (1)']):
        rhdr[i].text = title
        rhdr[i].paragraphs[0].runs[0].bold = True
        rhdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(rhdr[i], 'D9D9D9')
    
    data_r = [("Prosedur Kerja", "Sistematis & mandiri", "Sistematis", "Kurang", "Gagal"),
              ("Hasil Praktik", "Berhasil sempurna", "Berhasil minor", "Sebagian", "Gagal total"),
              ("K3LH & 5R", "Sangat taat & rapi", "Taat K3", "Kurang", "Mengabaikan")]
    for item in data_r:
        row = rubrik.add_row().cells
        for j in range(5): row[j].text = item[j]
        
    doc.add_paragraph("\nNilai Total: .......  |  Paraf Guru: ..........................").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def generate_soal_docx(list_soal, logo_kiri=None, logo_kanan=None):
    doc = Document()
    create_header(doc, logo_kiri, logo_kanan)
    
    # HALAMAN SOAL
    doc.add_heading("LEMBAR SOAL EVALUASI ASESMEN", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Mata Pelajaran\t: TJKT\nNama Siswa\t: ..............................................\n")
    for i, data in enumerate(list_soal, 1):
        doc.add_paragraph(f"{i}. {data['teks']}")
        doc.add_paragraph(f"   (Sumber: {data['elemen']})").runs[0].italic = True
        doc.add_paragraph("\n")
        
    # HALAMAN KUNCI
    doc.add_page_break()
    doc.add_heading("KUNCI JAWABAN & PENSKORAN", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("")
    for i, data in enumerate(list_soal, 1):
        doc.add_paragraph(f"Soal {i}: {data['teks']}").runs[0].italic = True
        p_jawab = doc.add_paragraph("Jawaban: ")
        p_jawab.add_run(data['kunci']).bold = True
        doc.add_paragraph("Skor Maksimal: 20  |  Skor Diperoleh: .......\n")
        
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 4. STREAMLIT UI
# ==========================================
st.set_page_config(page_title="App Kurikulum TJKT", layout="wide")
st.title("?? Sistem Manajemen Kurikulum TJKT")

st.sidebar.markdown("### ?? Pengaturan Kop Surat")
logo_kiri = st.sidebar.file_uploader("Upload Logo Kiri (Pemprov):", type=["png", "jpg"])
logo_kanan = st.sidebar.file_uploader("Upload Logo Kanan (Sekolah):", type=["png", "jpg"])

tab1, tab2 = st.tabs(["?? Modul Ajar & Job Sheet", "?? Generator Soal Acak"])

with tab1:
    col1, col2 = st.columns(2)
    fase_pilihan = col1.selectbox("Pilih Fase:", list(DATA_KURIKULUM_TKJ.keys()))
    elemen_pilihan = col2.selectbox("Pilih Elemen:", list(DATA_KURIKULUM_TKJ[fase_pilihan]["elemen"].keys()))
    data_elemen = DATA_KURIKULUM_TKJ[fase_pilihan]["elemen"][elemen_pilihan]
    
    st.info(f"**Capaian Pembelajaran:** {data_elemen['cp']}")
    st.success(f"**Praktik:** {data_elemen['praktik_indikator']}")
    
    if st.button("Generate Modul Ajar (DOCX & PDF)"):
        st.session_state['modul_doc'] = generate_docx(fase_pilihan, elemen_pilihan, data_elemen, logo_kiri, logo_kanan)
        st.session_state['modul_pdf'] = convert_docx_to_pdf(st.session_state['modul_doc'])
        st.success("Dokumen siap diunduh!")
        
    if 'modul_doc' in st.session_state:
        c_dl1, c_dl2 = st.columns(2)
        c_dl1.download_button("?? Download Word (DOCX)", st.session_state['modul_doc'], "Modul.docx")
        if st.session_state['modul_pdf']:
            c_dl2.download_button("?? Download PDF (Siswa)", st.session_state['modul_pdf'], "Modul.pdf")

with tab2:
    bank_soal = []
    for el, content in DATA_KURIKULUM_TKJ["Fase F (Kelas XI & XII - Konsentrasi TJKT)"]["elemen"].items():
        for soal in content["asesmen_soal"]:
            bank_soal.append({"elemen": el, "teks": re.sub(r'^\d+\.\s*', '', soal["tanya"]), "kunci": soal["jawab"]})
            
    jml = st.slider("Jumlah Soal:", 1, len(bank_soal), 3)
    if st.button("?? Acak Soal Ujian"):
        st.session_state['soal_acak'] = random.sample(bank_soal, jml)
        st.session_state['soal_doc'] = generate_soal_docx(st.session_state['soal_acak'], logo_kiri, logo_kanan)
        st.session_state['soal_pdf'] = convert_docx_to_pdf(st.session_state['soal_doc'])
        
    if 'soal_acak' in st.session_state:
        for idx, item in enumerate(st.session_state['soal_acak'], 1):
            st.write(f"**{idx}.** {item['teks']}")
        
        c_dl1, c_dl2 = st.columns(2)
        c_dl1.download_button("?? Download Lembar Soal (DOCX)", st.session_state['soal_doc'], "Soal.docx")
        if st.session_state['soal_pdf']:
            c_dl2.download_button("?? Download Lembar Soal (PDF)", st.session_state['soal_pdf'], "Soal.pdf")