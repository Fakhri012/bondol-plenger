from io import BytesIO
from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, Table, TableStyle, Image
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib import colors
from docx import Document
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
    static_url_path="/static",
)
app.config["SECRET_KEY"] = "change-me"

LETTER_HEADER_TEMPLATE = "Jakarta, {tanggal}"
LETTER_LIST_ITEMS = [
    "Surat Lamaran Kerja",
    "Daftar Riwayat Hidup (CV)",
    "Photocopy Ijazah dan Transkrip Nilai",
    "Photocopy Kartu Tanda Penduduk",
    "Pas Photo Berwarna",
    "Dan Berkas Pendukung Lainnya",
]


def load_word_template():
    path = os.path.join(os.path.dirname(__file__), "letter_template.docx")
    if not os.path.exists(path):
        return None

    document = Document(path)
    paragraphs = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            paragraphs.append(text)

    return "\n\n".join(paragraphs)


def build_story_from_word_text(
    letter_text,
    header_style,
    subject_style,
    recipient_style,
    body_style,
    body_no_indent,
    signature_style,
):
    story = []
    paragraphs = [p.strip() for p in letter_text.split("\n\n") if p.strip()]

    for index, paragraph in enumerate(paragraphs):
        if index == 0 and paragraph.lower().startswith("Jakarta,"):
            story.append(Paragraph(paragraph, header_style))
            continue

        if paragraph.startswith("Perihal"):
            story.append(Paragraph(paragraph, subject_style))
            continue

        if paragraph.startswith("Kepada Yth.") or paragraph.startswith("Bapak/Ibu") or paragraph.startswith("di-"):
            for line in paragraph.splitlines():
                if line.strip():
                    story.append(Paragraph(line.strip(), recipient_style))
            continue

        if paragraph.startswith("Hormat saya"):
            story.append(Spacer(1, 3 * cm))
            story.append(Paragraph(paragraph, signature_style))
            continue

        story.append(Paragraph(paragraph.replace("\n", "<br/>"), body_style))

    return story


def format_date_indonesia(value: str) -> str:
    months = {
        "01": "Januari",
        "02": "Februari",
        "03": "Maret",
        "04": "April",
        "05": "Mei",
        "06": "Juni",
        "07": "Juli",
        "08": "Agustus",
        "09": "September",
        "10": "Oktober",
        "11": "November",
        "12": "Desember",
    }
    parts = value.split("-")
    if len(parts) == 3:
        year, month, day = parts
        return f"{int(day)} {months.get(month, month)} {year}"
    return value


def sanitize_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_-]+", "_", value.strip())
    return value or "surat"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", errors=None, form=None)


@app.route("/generate", methods=["POST"])
def generate_pdf():
    nama_perusahaan = request.form.get("nama_perusahaan", "").strip()
    lokasi_perusahaan = request.form.get("lokasi_perusahaan", "").strip()
    job_posisi = request.form.get("job_posisi", "").strip()
    tanggal = request.form.get("tanggal", "").strip()

    errors = []
    if not nama_perusahaan:
        errors.append("Nama Perusahaan harus diisi.")
    if not lokasi_perusahaan:
        errors.append("Lokasi Perusahaan harus diisi.")
    if not job_posisi:
        errors.append("Posisi yang Dilamar harus diisi.")
    if not tanggal:
        errors.append("Tanggal Surat harus diisi.")

    if errors:
        form = {
            "nama_perusahaan": nama_perusahaan,
            "lokasi_perusahaan": lokasi_perusahaan,
            "job_posisi": job_posisi,
            "tanggal": tanggal,
        }
        return render_template("index.html", errors=errors, form=form)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=3 * cm,
        rightMargin=2.5 * cm,
        topMargin=2 * cm,
        bottomMargin=2.5 * cm,
    )

    header_style = ParagraphStyle(
        "Header",
        fontName="Times-Roman",
        fontSize=12,
        leading=18,
        alignment=TA_RIGHT,
        spaceAfter=12,
    )

    subject_style = ParagraphStyle(
        "Subject",
        fontName="Times-Roman",
        fontSize=12,
        leading=18,
        alignment=TA_LEFT,
        spaceAfter=18,
    )

    recipient_style = ParagraphStyle(
        "Recipient",
        fontName="Times-Roman",
        fontSize=12,
        leading=12,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0,
    )

    body_style = ParagraphStyle(
        "Body",
        fontName="Times-Roman",
        fontSize=12,
        leading=12,
        alignment=TA_JUSTIFY,
        firstLineIndent=0,
        spaceAfter=6,
    )

    label_style = ParagraphStyle(
        "Label",
        fontName="Times-Roman",
        fontSize=12,
        leading=12,
        alignment=TA_LEFT,
        firstLineIndent=0,
        leftIndent=0,
        spaceAfter=3,
    )

    value_style = ParagraphStyle(
        "Value",
        fontName="Times-Roman",
        fontSize=12,
        leading=12,
        alignment=TA_LEFT,
        firstLineIndent=0,
        leftIndent=0,
        spaceAfter=3,
    )

    body_no_indent = ParagraphStyle(
        "BodyNoIndent",
        fontName="Times-Roman",
        fontSize=12,
        leading=12,
        alignment=TA_JUSTIFY,
        firstLineIndent=0,
        spaceAfter=6,
    )

    signature_style = ParagraphStyle(
        "Signature",
        fontName="Times-Roman",
        fontSize=12,
        leading=18,
        alignment=TA_RIGHT,
        spaceAfter=0,
    )

    word_template = load_word_template()
    formatted_date = format_date_indonesia(tanggal)

    if word_template:
        letter_text = word_template.format(
            tanggal=formatted_date,
            nama_perusahaan=nama_perusahaan,
            lokasi_perusahaan=lokasi_perusahaan,
            job_posisi=job_posisi,
        )
        story = build_story_from_word_text(
            letter_text,
            header_style,
            subject_style,
            recipient_style,
            body_style,
            body_no_indent,
            signature_style,
        )
    else:
        story = []
        story.append(Paragraph(LETTER_HEADER_TEMPLATE.format(tanggal=formatted_date), header_style))
        story.append(Paragraph("Perihal : Lamaran Pekerjaan", subject_style))

        story.append(Paragraph("Kepada Yth.", recipient_style))
        story.append(Paragraph("Bapak/Ibu Pimpinan", recipient_style))
        story.append(Paragraph(f"<b>{nama_perusahaan}</b>", recipient_style))
        story.append(Paragraph(f"di- {lokasi_perusahaan}", recipient_style))
        story.append(Spacer(1, 12))

        story.append(Paragraph("<b>Dengan hormat,</b>", body_no_indent))
        story.append(Spacer(1, 12))
        story.append(
            Paragraph(
                f"Sehubungan dengan informasi tentang dibukanya lowongan pekerjaan untuk posisi {job_posisi}, yang saya baca melalui media sosial, maka saya bermaksud mengajukan diri guna mengisi posisi tersebut.",
                body_style,
            )
        )
        story.append(Paragraph("Yang bertandatangan di bawah ini:", body_style))
        detail_data = [
            [Paragraph("Nama", label_style), Paragraph(":", label_style), Paragraph("Fatma Zahratun Nisa", value_style)],
            [Paragraph("Tempat & Tanggal Lahir", label_style), Paragraph(":", label_style), Paragraph("Jakarta, 4 Februari, 2007", value_style)],
            [Paragraph("Pendidikan Terakhir", label_style), Paragraph(":", label_style), Paragraph("SMKN 19 Jakarta Jurusan Akutansi", value_style)],
            [Paragraph("Alamat", label_style), Paragraph(":", label_style), Paragraph("Pedurenan Mesjid Rt/Rw 010/007 Karet Kuningan, Setiabudi", value_style)],
            [Paragraph("Nomor Handphone", label_style), Paragraph(":", label_style), Paragraph("0857-7483-1517", value_style)],
            [Paragraph("Email", label_style), Paragraph(":", label_style), Paragraph("fatmanissa5@gmail.com", value_style)],
            [Paragraph("Status", label_style), Paragraph(":", label_style), Paragraph("Belum Menikah", value_style)],
        ]
        detail_table = Table(
            detail_data,
            colWidths=[6 * cm, 0.5 * cm, 9.5 * cm],
            hAlign="LEFT",
        )
        detail_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(detail_table)
        story.append(Spacer(1, 6))
        story.append(
            Paragraph(
                f"Dengan ini mengajukan permohonan lamaran kerja untuk menempati posisi {job_posisi} dengan bekal kemampuan yang saya miliki. Sebagai bahan pertimbangan, bersama ini saya lampirkan:",
                body_style,
            )
        )
        list_data = [
            [
                Paragraph(f"{index + 1}.", label_style),
                Paragraph(item, body_style),
            ]
            for index, item in enumerate(LETTER_LIST_ITEMS)
        ]
        list_table = Table(
            list_data,
            colWidths=[0.6 * cm, None],
            hAlign="LEFT",
        )
        list_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (0, -1), 0),
                    ("RIGHTPADDING", (1, 0), (1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(list_table)
        story.append(Spacer(1, 12))
        story.append(
            Paragraph(
                f"Demikian surat lamaran ini saya buat dengan sebenarnya. Besar harapan saya untuk dapat diterima bekerja menjadi bagian dari {job_posisi} di {nama_perusahaan}. Atas perhatian dan kesempatan yang diberikan, saya ucapkan terima kasih.",
                body_style,
            )
        )
        story.append(Spacer(1, 0.5 * cm))
        story.append(Paragraph("Hormat saya,", signature_style))
        image_path = os.path.join(os.path.dirname(__file__), "ttd-bondol.png")
        if os.path.exists(image_path):
            signature_image = Image(image_path, width=3 * cm, height=1.2 * cm)
            signature_image.hAlign = "RIGHT"
            story.append(signature_image)
            story.append(Spacer(1, 0.4 * cm))
        story.append(Paragraph("Fatma Zahratun Nisa", signature_style))

    doc.build(story)
    buffer.seek(0)

    safe_company = sanitize_filename(nama_perusahaan)
    safe_position = sanitize_filename(job_posisi)
    filename = f"Lamaran_{safe_company}_{safe_position}.pdf"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf",
    )


if __name__ == "__main__":
    app.run(debug=True)
