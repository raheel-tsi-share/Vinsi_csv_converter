import csv
import io
import re

import openpyxl
import streamlit as st

# ── Core logic (no file I/O — returns CSV bytes) ──────────────────────────────

_EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def normalize_email(raw: str):
    email = re.sub(r'\s+', '', raw)
    return email if _EMAIL_RE.match(email) else None


def normalize_phone(raw: str):
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 10:
        return f"+1{digits}"
    if len(digits) == 11 and digits.startswith("1"):
        return f"+{digits}"
    return None


def get_fieldnames(file_bytes: bytes, filename: str) -> list:
    if filename.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb.active
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        wb.close()
        return [str(c) if c is not None else "" for c in first_row]
    text = file_bytes.decode("utf-8-sig")
    return csv.DictReader(io.StringIO(text)).fieldnames or []


def iter_rows(file_bytes: bytes, filename: str, fieldnames: list):
    if filename.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            yield {fieldnames[i]: (str(v) if v is not None else "") for i, v in enumerate(row) if i < len(fieldnames)}
        wb.close()
    else:
        text = file_bytes.decode("utf-8-sig")
        yield from csv.DictReader(io.StringIO(text))


def process_batch(file_bytes, filename, phone_col, name_format,
                  name_col, first_name_col, last_name_col, extra_cols, dedup):
    seen_phones, seen_names = set(), set()
    rows, skipped, failed = [], 0, []
    out_fieldnames = ["name", "phoneNumber"] + extra_cols
    fieldnames = get_fieldnames(file_bytes, filename)

    for row in iter_rows(file_bytes, filename, fieldnames):
        if name_format == "first_last":
            name = f"{row.get(first_name_col,'').strip()} {row.get(last_name_col,'').strip()}".strip()
        else:
            name = row.get(name_col, "").strip()

        phone = normalize_phone(row.get(phone_col, "").strip())
        if phone is None:
            failed.append(name or "(no name)")
            skipped += 1
            continue
        if dedup == "phones" and phone in seen_phones:
            skipped += 1
            continue
        if dedup == "names" and name in seen_names:
            skipped += 1
            continue

        seen_phones.add(phone)
        seen_names.add(name)
        out = {"name": name, "phoneNumber": phone}
        for col in extra_cols:
            out[col] = row.get(col, "").strip()
        rows.append(out)

    if not rows:
        return 0, skipped, failed, None

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=out_fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    return len(rows), skipped, failed, buf.getvalue().encode()


def process_email(file_bytes, filename, email_col, name_format,
                  first_name_col, last_name_col, full_name_col, dedup):
    seen_emails = set()
    rows, skipped, failed = [], 0, []

    if name_format == "split" and full_name_col:
        out_fieldnames = ["email", "firstName", "lastName"]
    else:
        out_fieldnames = ["email"]
        if first_name_col:
            out_fieldnames.append("firstName")
        if last_name_col:
            out_fieldnames.append("lastName")

    fieldnames = get_fieldnames(file_bytes, filename)
    for row in iter_rows(file_bytes, filename, fieldnames):
        email = normalize_email(row.get(email_col, ""))
        if email is None:
            failed.append(row.get(email_col, "").strip() or "(empty)")
            skipped += 1
            continue
        if dedup == "emails" and email in seen_emails:
            skipped += 1
            continue
        if name_format == "split" and full_name_col:
            full_name = row.get(full_name_col, "").strip()
            parts = full_name.split(None, 1)
            if len(parts) < 2:
                failed.append(full_name or "(no name)")
                skipped += 1
                continue
            first, last = parts[0], parts[1]

        seen_emails.add(email)
        out = {"email": email}
        if name_format == "split" and full_name_col:
            out["firstName"] = first
            out["lastName"] = last
        else:
            if first_name_col:
                out["firstName"] = row.get(first_name_col, "").strip()
            if last_name_col:
                out["lastName"] = row.get(last_name_col, "").strip()
        rows.append(out)

    if not rows:
        return 0, skipped, failed, None

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=out_fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    return len(rows), skipped, failed, buf.getvalue().encode()


def process_contact(file_bytes, filename, email_col, name_format,
                    full_name_col, first_name_col, last_name_col, extra_cols, dedup):
    seen_emails = set()
    rows, skipped, failed = [], 0, []

    out_fieldnames = ["email"]
    if name_format == "split" and full_name_col:
        out_fieldnames += ["firstName", "lastName"]
    else:
        if first_name_col:
            out_fieldnames.append("firstName")
        if last_name_col:
            out_fieldnames.append("lastName")
    out_fieldnames += extra_cols

    fieldnames = get_fieldnames(file_bytes, filename)
    for row in iter_rows(file_bytes, filename, fieldnames):
        email = normalize_email(row.get(email_col, ""))
        if email is None:
            failed.append(row.get(email_col, "").strip() or "(empty)")
            skipped += 1
            continue
        if dedup == "emails" and email in seen_emails:
            skipped += 1
            continue
        if name_format == "split" and full_name_col:
            full_name = row.get(full_name_col, "").strip()
            parts = full_name.split(None, 1)
            if len(parts) < 2:
                failed.append(full_name or "(no name)")
                skipped += 1
                continue
            first, last = parts[0], parts[1]

        seen_emails.add(email)
        out = {"email": email}
        if name_format == "split" and full_name_col:
            out["firstName"] = first
            out["lastName"] = last
        else:
            if first_name_col:
                out["firstName"] = row.get(first_name_col, "").strip()
            if last_name_col:
                out["lastName"] = row.get(last_name_col, "").strip()
        for col in extra_cols:
            out[col] = row.get(col, "").strip()
        rows.append(out)

    if not rows:
        return 0, skipped, failed, None

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=out_fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    return len(rows), skipped, failed, buf.getvalue().encode()


# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="VINSI CSV Formatter", page_icon="📋", layout="centered")

st.title("VINSI CSV Formatter")

mode = st.radio(
    "What would you like to create?",
    ["Batch Phone Import", "Email Campaign", "Batch Contact Import"],
    horizontal=True,
)

uploaded = st.file_uploader("Upload your LineLeader export", type=["xlsx", "csv"])

if uploaded:
    file_bytes = uploaded.read()
    filename = uploaded.name
    try:
        columns = get_fieldnames(file_bytes, filename)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()

    if not columns:
        st.error("No columns found in the uploaded file.")
        st.stop()

    st.divider()

    # ── Batch Phone Import ────────────────────────────────────────────────────
    if mode == "Batch Phone Import":
        col1, col2 = st.columns(2)
        with col1:
            name_format = st.radio("Name format", ["Full Name Column", "Separate First / Last Name Columns"])
            phone_col = st.selectbox("Phone number column", columns)
        with col2:
            if name_format == "Full Name Column":
                name_col = st.selectbox("Name column", columns)
                first_name_col = last_name_col = ""
            else:
                name_col = ""
                first_name_col = st.selectbox("First name column", columns)
                last_name_col = st.selectbox("Last name column", columns)

        extra_cols = st.multiselect("Extra columns to include (optional)", columns)
        dedup = st.selectbox("Remove Duplicates By", ["None", "Phone numbers", "Names"])
        dedup_key = {"None": "none", "Phone numbers": "phones", "Names": "names"}[dedup]
        nf_key = "first_last" if name_format == "Separate First / Last Name Columns" else "full"

        if st.button("Convert", type="primary"):
            count, skipped, failed, csv_bytes = process_batch(
                file_bytes, filename, phone_col, nf_key,
                name_col, first_name_col, last_name_col, extra_cols, dedup_key,
            )
            if csv_bytes is None:
                st.error("No valid phone numbers found. Nothing was converted.")
            else:
                st.success(f"✓ {count} rows converted, {skipped} skipped.")
                if failed:
                    with st.expander(f"Rows with invalid phone numbers ({len(failed)})"):
                        st.write(failed)
                st.download_button("Download CSV", csv_bytes, "vinsi_batch_calls.csv", "text/csv")

    # ── Email Campaign ────────────────────────────────────────────────────────
    elif mode == "Email Campaign":
        col1, col2 = st.columns(2)
        with col1:
            name_format = st.radio("Name format", ["Separate First / Last Name Columns", "Full Name Column (will be split)", "No name"])
            email_col = st.selectbox("Email column", columns)
        with col2:
            if name_format == "Separate First / Last Name Columns":
                first_name_col = st.selectbox("First name column", columns)
                last_name_col = st.selectbox("Last name column", columns)
                full_name_col = ""
            elif name_format == "Full Name Column (will be split)":
                full_name_col = st.selectbox("Full name column (will be split)", columns)
                first_name_col = last_name_col = ""
            else:
                first_name_col = last_name_col = full_name_col = ""

        dedup = st.selectbox("Remove Duplicates By", ["None", "Email addresses"])
        dedup_key = "emails" if dedup == "Email addresses" else "none"
        nf_key = {"Full Name Column (will be split)": "split", "Separate First / Last Name Columns": "separate", "No name": "none"}[name_format]

        if st.button("Convert", type="primary"):
            count, skipped, failed, csv_bytes = process_email(
                file_bytes, filename, email_col, nf_key,
                first_name_col, last_name_col, full_name_col, dedup_key,
            )
            if csv_bytes is None:
                st.error("No valid email addresses found. Nothing was converted.")
            else:
                st.success(f"✓ {count} rows converted, {skipped} skipped.")
                if failed:
                    with st.expander(f"Invalid/skipped emails ({len(failed)})"):
                        st.write(failed)
                st.download_button("Download CSV", csv_bytes, "email_campaign.csv", "text/csv")

    # ── Batch Contact Import ──────────────────────────────────────────────────
    else:
        col1, col2 = st.columns(2)
        with col1:
            name_format = st.radio("Name format", ["Separate First / Last Name Columns", "Full Name Column (will be split)", "No name"])
            email_col = st.selectbox("Email column", columns)
        with col2:
            if name_format == "Separate First / Last Name Columns":
                first_name_col = st.selectbox("First name column", columns)
                last_name_col = st.selectbox("Last name column", columns)
                full_name_col = ""
            elif name_format == "Full Name Column (will be split)":
                full_name_col = st.selectbox("Full name column (will be split)", columns)
                first_name_col = last_name_col = ""
            else:
                first_name_col = last_name_col = full_name_col = ""

        extra_cols = st.multiselect("Extra columns to include (optional)", columns)
        dedup = st.selectbox("Remove Duplicates By", ["None", "Email addresses"])
        dedup_key = "emails" if dedup == "Email addresses" else "none"
        nf_key = {"Full Name Column (will be split)": "split", "Separate First / Last Name Columns": "separate", "No name": "none"}[name_format]

        if st.button("Convert", type="primary"):
            count, skipped, failed, csv_bytes = process_contact(
                file_bytes, filename, email_col, nf_key,
                full_name_col, first_name_col, last_name_col, extra_cols, dedup_key,
            )
            if csv_bytes is None:
                st.error("No valid contacts found. Nothing was converted.")
            else:
                st.success(f"✓ {count} rows converted, {skipped} skipped.")
                if failed:
                    with st.expander(f"Invalid/skipped entries ({len(failed)})"):
                        st.write(failed)
                st.download_button("Download CSV", csv_bytes, "vinsi_contact_import.csv", "text/csv")
