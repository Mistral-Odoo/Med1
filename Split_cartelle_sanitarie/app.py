import streamlit as st
import re, io, os, copy, zipfile, unicodedata, subprocess, tempfile
from pathlib import Path
from openpyxl import load_workbook
from docx import Document
from docx.oxml.ns import qn

st.set_page_config(page_title="Cartelle Sanitarie", page_icon="📁", layout="centered")

st.title("📁 Generazione Cartelle Sanitarie")
st.markdown(
    "Carica i file TOTALI (`.docx`) e le anagrafiche (`.xlsx`). "
    "L'applicazione genera automaticamente un PDF per ogni lavoratore, "
    "organizzato in cartelle nominate **COGNOME_NOME_CF**."
)
st.info(
    "🔒 I file vengono elaborati interamente in memoria e non vengono "
    "salvati né trasmessi a terzi. Appena chiudi la pagina non rimane nulla.",
    icon="🔒",
)

def sanitizza(s):
    s = str(s).upper().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Z0-9]+", "_", s)
    return s.strip("_")

def split_blocchi(doc):
    body = doc.element.body
    blocchi, blocco = [], []
    for child in list(body):
        if child.tag == qn('w:sectPr'):
            blocco.append(child)
            if blocco:
                blocchi.append(blocco)
            break
        blocco.append(child)
        if child.tag == qn('w:p'):
            ppr = child.find(qn('w:pPr'))
            if ppr is not None and ppr.find(qn('w:sectPr')) is not None:
                blocchi.append(blocco)
                blocco = []
    return blocchi

def ha_nif(blocco):
    for el in blocco:
        if re.search(r'NIF\)[^:]*[:–\-]\s*\d{5,}', ''.join(el.itertext())):
            return True
    return False

def estrai_docx_bytes(blocco, docx_bytes_originali):
    doc = Document(io.BytesIO(docx_bytes_originali))
    body = doc.element.body
    for child in list(body):
        body.remove(child)
    for el in blocco:
        body.append(copy.deepcopy(el))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def docx_bytes_to_pdf_bytes(docx_bytes, tmpdir):
    tmp = Path(tmpdir)
    docx_path = tmp / "input.docx"
    docx_path.write_bytes(docx_bytes)
    lo_home = tmp / "lo_home"
    lo_home.mkdir(exist_ok=True)
    env = {
        **os.environ,
        "HOME": str(lo_home),
        "UserInstallation": f"file://{lo_home}/lo_profile",
    }
    res = subprocess.run(
        ["libreoffice", "--headless", "--norestore", "--nofirststartwizard",
         "--convert-to", "pdf", "--outdir", str(tmp), str(docx_path)],
        capture_output=True, text=True, env=env,
    )
    pdf_path = tmp / "input.pdf"
    if not pdf_path.exists():
        raise RuntimeError(
            f"Conversione PDF fallita.\nSTDOUT: {res.stdout[-300:]}\nSTDERR: {res.stderr[-300:]}"
        )
    data = pdf_path.read_bytes()
    docx_path.unlink(missing_ok=True)
    pdf_path.unlink(missing_ok=True)
    return data

def analizza_docx(fileobj):
    doc = Document(fileobj)
    if not doc.tables:
        return None
    tipo = "idoneita"
    for p in doc.paragraphs[:5]:
        if "CARTELLA SANITARIA" in p.text.upper():
            tipo = "cartella"
            break
    n_lavoratori = 0
    azienda = None
    for table in doc.tables:
        cell = table.rows[0].cells[0].text
        if re.search(r'NIF\)[^:]*[:–\-]\s*\d{5,}', cell):
            n_lavoratori += 1
            if azienda is None:
                m = re.search(r'Azienda[^:]*:\s*\*{0,2}([^\n*]+?)\*{0,2}\s*(?:Sede|Mansione|$)', cell)
                if m:
                    azienda = m.group(1).strip()
    return {"tipo": tipo, "azienda": azienda, "n_lavoratori": n_lavoratori}

def analizza_xlsx(fileobj):
    wb = load_workbook(fileobj, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    lavoratori, azienda = [], None
    for row in rows[1:]:
        if row[0] is None:
            continue
        cognome, nome, _, _, _, _, cf, az = row[:8]
        if azienda is None:
            azienda = str(az).strip()
        lavoratori.append(f"{sanitizza(cognome)}_{sanitizza(nome)}_{int(cf)}")
    return {"azienda": azienda, "lavoratori": lavoratori}

# ── UI ───────────────────────────────────────────────────────────────────────

uploaded = st.file_uploader(
    "Trascina qui i file oppure clicca per selezionarli",
    type=["docx", "xlsx"],
    accept_multiple_files=True,
)

if uploaded:
    docx_files = [f for f in uploaded if f.name.lower().endswith(".docx")]
    xlsx_files  = [f for f in uploaded if f.name.lower().endswith(".xlsx")]
    st.markdown(f"**File caricati:** {len(docx_files)} docx · {len(xlsx_files)} xlsx")

    with st.expander("📋 Dettaglio file riconosciuti", expanded=False):
        for f in docx_files:
            info = analizza_docx(io.BytesIO(f.read()))
            f.seek(0)
            if info and info["n_lavoratori"] > 1:
                st.success(f"✅ **{f.name}** — {info['tipo']} · {info['n_lavoratori']} lavoratori · {info['azienda']}")
            elif info:
                st.warning(f"⏭️ **{f.name}** — file singolo (BASE), verrà ignorato")
            else:
                st.error(f"❌ **{f.name}** — non riconoscibile")
        for f in xlsx_files:
            info = analizza_xlsx(io.BytesIO(f.read()))
            f.seek(0)
            st.info(f"📊 **{f.name}** — {info['azienda']} · {len(info['lavoratori'])} lavoratori")

    st.divider()

    if st.button("🚀 Genera cartelle", type="primary", use_container_width=True):
        progress = st.progress(0, text="Avvio elaborazione...")
        log = st.empty()
        try:
            # Analisi
            docx_info = []
            for f in docx_files:
                raw = f.read(); f.seek(0)
                info = analizza_docx(io.BytesIO(raw))
                if info and info["n_lavoratori"] > 1:
                    info["raw"] = raw
                    docx_info.append(info)

            xlsx_map = {}
            for f in xlsx_files:
                info = analizza_xlsx(io.BytesIO(f.read())); f.seek(0)
                if info["azienda"]:
                    xlsx_map[sanitizza(info["azienda"])] = info

            # Abbinamento
            jobs, errori = [], []
            for d in docx_info:
                key = sanitizza(d["azienda"] or "")
                if key not in xlsx_map:
                    errori.append(f"Nessuna anagrafica xlsx per: **{d['azienda']}**")
                    continue
                xl = xlsx_map[key]
                doc = Document(io.BytesIO(d["raw"]))
                blocchi_reali = [b for b in split_blocchi(doc) if ha_nif(b)]
                n = min(len(xl["lavoratori"]), len(blocchi_reali))
                jobs.append({
                    "label": f"{d['tipo'].upper()} – {d['azienda']}",
                    "tipo": d["tipo"],
                    "raw": d["raw"],
                    "blocchi": blocchi_reali[:n],
                    "lavoratori": xl["lavoratori"][:n],
                })

            if errori:
                for e in errori: st.error(e)
                st.stop()

            # Elaborazione
            n_jobs = len(jobs)
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                with tempfile.TemporaryDirectory() as tmpdir:
                    for ji, job in enumerate(jobs):
                        n_lav = len(job["lavoratori"])
                        for i, (nome_cartella, blocco) in enumerate(zip(job["lavoratori"], job["blocchi"])):
                            log.markdown(f"⏳ **{job['label']}** — {nome_cartella} ({i+1}/{n_lav})")
                            progress.progress(
                                (ji * n_lav + i) / (n_jobs * n_lav),
                                text=f"{job['tipo']} {i+1}/{n_lav}"
                            )
                            docx_b = estrai_docx_bytes(blocco, job["raw"])
                            pdf_b  = docx_bytes_to_pdf_bytes(docx_b, tmpdir)
                            zf.writestr(f"{nome_cartella}/{job['tipo']}_{nome_cartella}.pdf", pdf_b)

            progress.progress(1.0, text="✅ Completato!")
            log.empty()

            n_cartelle = len({
                zi.filename.split("/")[0]
                for zi in zipfile.ZipFile(io.BytesIO(zip_buf.getvalue())).infolist()
                if "/" in zi.filename
            })
            st.success(f"✅ **Elaborazione completata!** {n_cartelle} cartelle generate.")
            zip_buf.seek(0)
            st.download_button(
                label="⬇️  Scarica cartelle_sanitarie.zip",
                data=zip_buf,
                file_name="cartelle_sanitarie.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            progress.empty()
            log.empty()
            st.error(f"❌ Errore: {e}")

else:
    st.markdown(
        "<br><p style='text-align:center;color:gray;'>"
        "Nessun file caricato — trascina i file qui sopra per iniziare."
        "</p>",
        unsafe_allow_html=True,
    )
