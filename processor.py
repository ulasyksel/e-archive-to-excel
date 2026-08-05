from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
import re
import zipfile

import fitz
import pandas as pd


MAIN_COLUMNS = [
    "MÜŞTERİ ADI",
    "FATURA NO",
    "VERGİ HARİÇ TUTAR",
    "KDV",
    "ÖDENECEK TUTAR",
    "ÖDEME ŞEKLİ",
    "MAĞAZA",
    "PARA BİRİMİ",
]

CONTROL_COLUMNS = [
    "MÜŞTERİ ADI",
    "FATURA NO",
    "PARA BİRİMİ",
    "FATURA VERGİ HARİÇ",
    "SATIR VERGİ HARİÇ TOPLAMI",
    "FATURA KDV",
    "SATIR KDV TOPLAMI",
    "FATURA ÖDENECEK",
    "ÖDEME SATIRLARI TOPLAMI",
    "FARK",
    "DURUM",
    "NOT",
]


def money(value):
    return float(
        Decimal(str(value)).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    )


def tr_number(value):
    if value is None:
        return None

    value = re.sub(r"[^\d,.\-]", "", value.strip())

    if not value:
        return None

    if "," in value:
        value = value.replace(".", "").replace(",", ".")
    elif value.count(".") == 1 and len(value.split(".")[1]) == 3:
        value = value.replace(".", "")

    try:
        return float(value)
    except ValueError:
        return None


def find_amount(text, label):
    match = re.search(
        rf"{label}\s*:?\s*([\-]?[0-9][0-9\.,]*)\s*(TL|EUR|USD)?",
        text,
        re.IGNORECASE,
    )

    if not match:
        return None, ""

    return tr_number(match.group(1)), (match.group(2) or "")


def normalize_payment(raw):
    upper = raw.upper()

    if any(x in upper for x in ("HEDİYE", "HEDIYE", "KREDİ ÇEK", "KREDI ÇEK")):
        return "KREDİ ÇEKLERİ"
    if "NAKİT" in upper or "NAKIT" in upper:
        return "NAKİT"
    if "QNB" in upper or "FİNANSBANK" in upper or "FINANSBANK" in upper:
        return "QNB FİNANSBANK POS"
    if "GARANTİ" in upper or "GARANTI" in upper:
        return "GARANTİ POS"
    if any(x in upper for x in (
        "İŞ BANK",
        "IŞ BANK",
        "IS BANK",
        "İSBANK",
        "ISBANK",
        "T.IŞ",
        "T.İŞ",
    )):
        return "İŞ BANKASI POS"
    if "AKBANK" in upper:
        return "AKBANK POS"
    if "HALK" in upper:
        return "HALKBANK POS"
    if "YAPI" in upper:
        return "YAPI KREDİ POS"

    return "DİĞER/BELİRSİZ"


def normalize_store(raw_store, invoice_no):
    by_prefix = {
        "DBA": "Missoni Mandarin",
        "DMA": "Missoni Yalıkavak",
        "DJA": "Jacquemus",
        "DJH": "Jacquemus",
        "DKA": "Antalya Kemer",
        "DAP": "Mandarin Printemps",
        "DXA": "MAX Royale",
        "DAR": "Missoni Printemps",
    }

    prefix = invoice_no[:3].upper()

    if prefix in by_prefix:
        return by_prefix[prefix]

    upper = raw_store.upper()

    if "JACQUEMUS" in upper:
        return "Jacquemus"
    if "MO BODRUM" in upper:
        return "Missoni Mandarin"
    if "MAXX ROYAL KEMER" in upper:
        return "Antalya Kemer"
    if "MAXX ROYAL" in upper:
        return "MAX Royale"
    if "MARINA PR" in upper:
        return "Mandarin Printemps"
    if "MO PR" in upper:
        return "Missoni Printemps"
    if "MISSONI" in upper:
        return "Missoni Yalıkavak"

    return raw_store.strip()


def extract_customer(text):
    parts = re.split(
        r"SAYIN",
        text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )

    if len(parts) < 2:
        return ""

    section = parts[1]

    title = re.search(
        r"Unvanı\s*:\s*([^\n\r]+)",
        section,
        re.IGNORECASE,
    )

    if title:
        customer = re.split(
            r"\s+(?:Tarih|Fatura No|Özelleştirme No|Senaryo)\s*:",
            title.group(1),
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0].strip()

        if "GÜMRÜK VE TİCARET BAKANLIĞI" not in customer.upper():
            return customer

    ignored = (
        "TARİH:",
        "FATURA NO:",
        "ÖZELLEŞTİRME",
        "SENARYO:",
        "FATURA TİPİ:",
        "OLUŞMA ZAMANI:",
        "E-POSTA:",
        "VERGİ DAİRESİ:",
        "TCKN:",
        "VKN:",
        "MUSTERINO:",
        "MÜŞTERİNO:",
        "ETTN:",
        "SOKAK",
        "BİNA",
        "İLÇE:",
        "İL:",
        "POSTA KODU:",
        "ÜLKE:",
    )

    for line in section.splitlines():
        line = re.sub(r"\s+", " ", line).strip()

        if line and not line.upper().startswith(ignored):
            return line

    return ""


def pdf_text(pdf_bytes):
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        return "\n".join(
            page.get_text("text", sort=True)
            for page in document
        )


def collect_pdfs(uploaded_files):
    results = []

    for uploaded in uploaded_files:
        name = uploaded.name
        data = uploaded.getvalue()

        if name.lower().endswith(".pdf"):
            results.append((name, data))

        elif name.lower().endswith(".zip"):
            with zipfile.ZipFile(BytesIO(data)) as archive:
                for member in archive.namelist():
                    if (
                        member.lower().endswith(".pdf")
                        and not member.endswith("/")
                    ):
                        results.append(
                            (
                                member.rsplit("/", 1)[-1],
                                archive.read(member),
                            )
                        )

    return results


def count_pdfs(uploaded_files):
    return len(collect_pdfs(uploaded_files))


def parse_payments(text, payable):
    lines = re.findall(
        r"\*\s*Ödeme\s*:\s*([^\n\r]+)",
        text,
        re.IGNORECASE,
    )

    if not lines:
        return [
            {
                "method": "DİĞER/BELİRSİZ",
                "amount": payable,
            }
        ], "Faturada ödeme satırı bulunamadı."

    payments = []

    for line in lines:
        amounts = re.findall(
            r"([0-9][0-9\.,]*)\s*(?:TL|EUR|USD)",
            line,
            re.IGNORECASE,
        )

        payments.append(
            {
                "method": normalize_payment(line),
                "amount": tr_number(amounts[-1]) if amounts else None,
            }
        )

    known_total = sum(payment["amount"] or 0 for payment in payments)
    missing = [
        payment
        for payment in payments
        if payment["amount"] is None
    ]

    if missing:
        remaining = money(payable - known_total)
        share = money(remaining / len(missing))

        for payment in missing[:-1]:
            payment["amount"] = share
            remaining = money(remaining - share)

        missing[-1]["amount"] = remaining

    return payments, ""


def process_files(uploaded_files, progress_callback=None):
    pdf_files = collect_pdfs(uploaded_files)

    main_rows = []
    control_rows = []
    errors = []

    for index, (source_name, pdf_bytes) in enumerate(
        pdf_files,
        start=1,
    ):
        try:
            text = pdf_text(pdf_bytes)

            invoice_match = re.search(
                r"Fatura No\s*:\s*([A-Z0-9]+)",
                text,
                re.IGNORECASE,
            )
            invoice_no = (
                invoice_match.group(1)
                if invoice_match
                else ""
            )

            tax_exclusive, currency_1 = find_amount(
                text,
                r"Vergi Hariç Tutar",
            )
            payable, currency_2 = find_amount(
                text,
                r"Ödenecek Tutar",
            )

            if tax_exclusive is None:
                raise ValueError(
                    "Vergi Hariç Tutar bulunamadı."
                )

            if payable is None:
                raise ValueError(
                    "Ödenecek Tutar bulunamadı."
                )

            tax_exclusive = money(tax_exclusive)
            payable = money(payable)
            vat = money(payable - tax_exclusive)
            currency = currency_2 or currency_1 or "TL"
            customer = extract_customer(text)

            store_match = re.search(
                r"\*\s*Mağaza Adı\s*:\s*([^\n\r]+)",
                text,
                re.IGNORECASE,
            )
            raw_store = (
                store_match.group(1).strip()
                if store_match
                else ""
            )
            store = normalize_store(
                raw_store,
                invoice_no,
            )

            payments, payment_note = parse_payments(
                text,
                payable,
            )
            payment_total = money(
                sum(payment["amount"] for payment in payments)
            )

            allocated_tax = []
            allocated_vat = []

            for payment_index, payment in enumerate(payments):
                payment_amount = money(payment["amount"])

                if payment_index == len(payments) - 1:
                    row_tax = money(
                        tax_exclusive - sum(allocated_tax)
                    )
                    row_vat = money(
                        vat - sum(allocated_vat)
                    )
                else:
                    ratio = (
                        payment_amount / payment_total
                        if payment_total
                        else 0
                    )
                    row_tax = money(
                        tax_exclusive * ratio
                    )
                    row_vat = money(
                        vat * ratio
                    )

                allocated_tax.append(row_tax)
                allocated_vat.append(row_vat)

                main_rows.append(
                    [
                        customer,
                        invoice_no,
                        row_tax,
                        row_vat,
                        payment_amount,
                        payment["method"],
                        store,
                        currency,
                    ]
                )

            row_tax_total = money(sum(allocated_tax))
            row_vat_total = money(sum(allocated_vat))

            difference = money(
                max(
                    abs(tax_exclusive - row_tax_total),
                    abs(vat - row_vat_total),
                    abs(payable - payment_total),
                )
            )

            notes = []

            if payment_note:
                notes.append(payment_note)
            if not customer:
                notes.append("Müşteri adı okunamadı.")
            if not invoice_no:
                notes.append("Fatura numarası okunamadı.")
            if not store:
                notes.append("Mağaza adı okunamadı.")
            if abs(payable - payment_total) > 0.01:
                notes.append(
                    "Ödeme toplamı faturayla uyuşmuyor."
                )

            status = (
                "KONTROL"
                if notes or difference > 0.01
                else "OK"
            )

            control_rows.append(
                [
                    customer,
                    invoice_no,
                    currency,
                    tax_exclusive,
                    row_tax_total,
                    vat,
                    row_vat_total,
                    payable,
                    payment_total,
                    difference,
                    status,
                    " ".join(notes) or None,
                ]
            )

        except Exception as error:
            errors.append(
                {
                    "DOSYA": source_name,
                    "HATA": str(error),
                }
            )

        if progress_callback:
            progress_callback(
                index,
                len(pdf_files),
            )

    return (
        pd.DataFrame(
            main_rows,
            columns=MAIN_COLUMNS,
        ),
        pd.DataFrame(
            control_rows,
            columns=CONTROL_COLUMNS,
        ),
        pd.DataFrame(errors),
    )
