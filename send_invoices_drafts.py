"""
send_invoices_drafts_v11.py
Améliorations appliquées vs v8 :
  1. Tkinter : root.destroy() systématique après chaque dialogue
  2. Instance Outlook unique partagée (initialisée une seule fois dans main)
  3. header Excel flexible : détection automatique de la ligne d'en-tête
  4. TEST_MODE activable via case à cocher dans la fenêtre de sélection Outlook
  5. Fallback store_choice=None (plus de chaîne magique "___DEFAULT___")
  6. Nettoyage des PDFs verrouillés en cas d'échec de création du brouillon

v10 :
  7. Support des courriers de subsides (PDF nommés "FG {contrat}.pdf")
     - Dossier subsides sélectionnable séparément via pop-up de confirmation
     - Verrouillage avec le même mot de passe que les factures
     - Copie verrouillée renommée "Subside {contrat}_locked.pdf"
     - Ajout en 2e pièce jointe si trouvé, sinon mail envoyé normalement sans
     - Rapport CSV enrichi : colonne "Subside" (OUI / NON)

v11 — Optimisations performance (200+ contrats) :
  8. Pré-indexation des PDFs factures ET subsides en dictionnaires au démarrage
     → supprime les glob/itération répétés à chaque contrat
  9. Verrouillage PDF parallèle via ThreadPoolExecutor
     → factures et subsides verrouillés simultanément (N workers)
     → la création des brouillons Outlook reste séquentielle (COM non thread-safe)
  10. Skip des PDFs déjà verrouillés (reprise après erreur sans re-traitement)

v12 — Optimisations & fiabilité :
  11. Dossier Brouillons + SMTP résolus UNE seule fois (plus de parcours des
      Stores Outlook à chaque mail → ~200 appels COM économisés)
  12. Lecture Excel unique (détection d'en-tête sur les données déjà chargées)
  13. Cellules vides : fillna("") → plus de "nan" dans les mails / contrats
  14. Recherche facture partielle stricte : le contrat 123 ne matche plus 1234
  15. Skip PDF verrouillé seulement s'il est plus récent que la source
      (évite d'envoyer la facture du mois précédent)
  16. Barre de progression sur le verrouillage parallèle
"""

from pathlib import Path
import pandas as pd
import pikepdf
import win32com.client as win32
import csv
import logging
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import re
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed


# --------------------------- CONFIGURATION ---------------------------
CONTRACT_COL = "N°_de_contrat"

POSSIBLE_EMAIL_COLS = [
    "EMail1_0", "EMail2_0",
    "EMail1_4", "EMail2_4",
    "EMail1_V", "EMail2_V",
    "Mail facturation", "Mail_facturation",
    "email_facturation", "AdresseEmail", "Mail"
]

RESPONSABLE_COL = "EMailResponsable"

PDF_PASSWORD = "Fondation&Facture.26"

TEST_ROWS = 5

# v11 : nombre de workers parallèles pour le verrouillage PDF
# 4 est un bon compromis sur la plupart des machines ; augmenter si SSD rapide
PDF_LOCK_WORKERS = 4

mois = datetime.now().strftime("%m/%Y")
SUBJECT_TEMPLATE = f"Facture {mois} – Contrat {{contract}}"

BODY_TEMPLATE = """
Mail automatique – Ne pas répondre à ce mail

Bonjour,

Nous vous remettons ci-joint notre facture mensuelle.

Vous pouvez contacter directement votre gestionnaire ({responsable}) pour obtenir le mot de passe ou contacter Mme Widmer au 021 310 12 30.

Avec nos meilleures salutations.

**************************************************************************************************************************

Automatische Mail - Auf diese Mail nicht antworten

Guten Tag!

Anbei erhalten Sie unsere monatliche Rechnung.

Sie können sich direkt an Ihren Sachbearbeiter ({responsable}) wenden, um das Passwort zu erhalten, oder Frau Widmer unter 021 310 12 30 kontaktieren.

Freundliche Grüsse

**************************************************************************************************************************

Automatic mail - Do not reply to this mail

Hello,

Please find enclosed our monthly invoice.

You can contact your manager ({responsable}) directly to get the password or contact Mrs Widmer on 021 310 12 30.

Best regards,
"""

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


# ---------------------------------------------------------
# FIX 1 — Sélection Excel / Dossier avec destroy() propre
# ---------------------------------------------------------

def ask_excel_file() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    try:
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
    finally:
        root.destroy()
    return Path(path) if path else None


def ask_pdf_folder() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    try:
        folder = filedialog.askdirectory(
            title="Sélectionner le dossier contenant les PDF de factures"
        )
    finally:
        root.destroy()
    return Path(folder) if folder else None


def ask_subsidy_folder() -> Path | None:
    root_confirm = tk.Tk()
    root_confirm.withdraw()
    try:
        confirm = messagebox.askyesno(
            title="Subsides",
            message="Y a-t-il des courriers de subsides à joindre ce mois-ci ?"
        )
    finally:
        root_confirm.destroy()

    if not confirm:
        logging.info("Pas de subsides ce mois-ci.")
        return None

    root_folder = tk.Tk()
    root_folder.withdraw()
    try:
        folder = filedialog.askdirectory(
            title="Sélectionner le dossier contenant les PDFs de subsides (FG ...)"
        )
    finally:
        root_folder.destroy()

    return Path(folder) if folder else None


# ---------------------------------------------------------
# FIX 2 — Instance Outlook unique
# ---------------------------------------------------------

def get_outlook_instance():
    return win32.Dispatch("Outlook.Application")


# ---------------------------------------------------------
# OUTLOOK : choix de la boîte (Store) + TEST_MODE UI
# ---------------------------------------------------------

def list_outlook_stores(outlook) -> list[str]:
    namespace = outlook.GetNamespace("MAPI")
    names = []
    try:
        for store in namespace.Stores:
            try:
                dn = store.DisplayName
                if dn:
                    names.append(dn)
            except Exception:
                continue
    except Exception:
        pass
    return names


def ask_outlook_store_and_options(outlook) -> tuple[str | None, bool]:
    stores = list_outlook_stores(outlook)
    result = {"store": None, "test_mode": False}

    root = tk.Tk()
    root.title("Paramètres d'envoi – Factures")
    root.geometry("540x220")
    root.resizable(False, False)

    tk.Label(root, text="Sélectionnez la boîte Outlook\n(dossier Brouillons cible) :", pady=8).pack()

    combo_values = stores if stores else ["(aucune boîte détectée)"]
    combo = ttk.Combobox(root, values=combo_values, state="readonly", width=70)
    combo.current(0)
    combo.pack(pady=4)

    test_var = tk.BooleanVar(value=False)
    tk.Checkbutton(
        root,
        text=f"Mode Test (traiter seulement les {TEST_ROWS} premières lignes)",
        variable=test_var
    ).pack(pady=6)

    def validate():
        result["store"] = combo.get() if stores else None
        result["test_mode"] = test_var.get()
        root.destroy()

    def cancel():
        result["store"] = None
        result["test_mode"] = False
        root.destroy()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="Valider", width=16, command=validate).pack(side="left", padx=8)
    tk.Button(btn_frame, text="Annuler", width=16, command=cancel).pack(side="left", padx=8)

    root.mainloop()
    return result["store"], result["test_mode"]


def get_store_and_drafts_folder(outlook, store_display_name: str):
    namespace = outlook.GetNamespace("MAPI")
    store_found = None
    for s in namespace.Stores:
        try:
            if s.DisplayName.lower() == store_display_name.lower():
                store_found = s
                break
        except Exception:
            continue

    if store_found is None:
        return None, None, ""

    try:
        drafts = store_found.GetDefaultFolder(16)
    except Exception:
        drafts = None

    smtp = ""
    try:
        smtp = store_found.GetRootFolder().PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        ) or ""
    except Exception:
        smtp = ""

    return store_found, drafts, smtp


# ---------------------------------------------------------
# FIX 3 — Détection automatique de la ligne d'en-tête Excel
# ---------------------------------------------------------

def load_excel(excel_path: Path, target_col: str, max_scan: int = 20) -> tuple[pd.DataFrame, int]:
    """
    v12 — Lit le fichier Excel UNE seule fois, détecte la ligne d'en-tête
    (celle qui contient target_col) et renvoie (DataFrame, index_ligne_entete).
    Les cellules vides sont remplacées par "" (au lieu de NaN → "nan").
    """
    raw = pd.read_excel(excel_path, engine="openpyxl", dtype=str, header=None).fillna("")

    header_row = 0
    for i in range(min(max_scan, len(raw))):
        if (raw.iloc[i].str.strip() == target_col).any():
            header_row = i
            logging.info(f"En-tête détecté à la ligne {i} (0-based).")
            break
    else:
        logging.warning(f"'{target_col}' introuvable dans les {max_scan} premières lignes. header=0.")

    # Noms de colonnes nettoyés + dédoublonnés (comme pandas : "Col", "Col.1", ...)
    columns, seen = [], {}
    for c in raw.iloc[header_row]:
        name = str(c).strip()
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        columns.append(name)

    df = raw.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = columns
    return df, header_row


# ---------------------------------------------------------
# Extraction email
# ---------------------------------------------------------

def extract_facturation_email(text: str) -> str:
    if not isinstance(text, str):
        return ""
    if "facturation" not in text.lower():
        return ""
    emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    return emails[-1] if emails else ""


def get_email_from_row(row, email_cols: list[str]) -> str:
    for col in email_cols:
        raw = str(row.get(col, "")).strip()
        email = extract_facturation_email(raw)
        if email:
            return email
    return ""


# ---------------------------------------------------------
# PDF — Verrouillage
# ---------------------------------------------------------

def lock_pdf(input_path: Path, output_path: Path, password: str) -> tuple[bool, str | None]:
    try:
        with pikepdf.open(input_path) as pdf:
            pdf.save(
                output_path,
                encryption=pikepdf.Encryption(owner=password, user=password, R=4)
            )
        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------
# v11 — OPT 8 : Pré-indexation des PDFs en dictionnaires
# ---------------------------------------------------------

def normalize_name(text: str) -> str:
    """Minuscules + espaces multiples réduits à un seul."""
    return " ".join(text.split()).lower()


def index_pdfs(folder: Path) -> dict[str, Path]:
    """
    Construit un dictionnaire {nom_normalise: path} pour tous les PDFs du dossier.
    Appelé une seule fois au démarrage, évite les glob répétés à chaque contrat.
    """
    index = {}
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() == ".pdf":
            index[normalize_name(p.stem)] = p
    logging.info(f"{len(index)} PDFs indexés dans : {folder}")
    return index


def find_invoice_pdf(contract: str, pdf_index: dict[str, Path]) -> Path | None:
    """
    Cherche la facture d'un contrat dans l'index pré-construit.
    Stratégie : correspondance exacte, puis recherche partielle.
    v12 : la recherche partielle exige que le contrat ne soit pas collé à
    d'autres chiffres → "123" ne matche plus "facture 1234" ni "0123".
    """
    key = normalize_name(contract)

    exact = pdf_index.get(key)
    if exact:
        return exact

    pattern = re.compile(rf"(?<!\d){re.escape(key)}(?!\d)")
    for name, path in pdf_index.items():
        if pattern.search(name):
            return path

    return None


def find_subsidy_pdf(contract: str, subsidy_index: dict[str, Path]) -> Path | None:
    """
    Cherche le subside d'un contrat dans l'index pré-construit.
    Convention : "FG {contract}" (insensible à la casse).
    """
    return subsidy_index.get(normalize_name(f"FG {contract}"))


# ---------------------------------------------------------
# v11 — OPT 9 : Verrouillage PDF parallèle
# ---------------------------------------------------------

def lock_all_pdfs_parallel(
    lock_tasks: list[tuple[str, Path, Path, Path | None, Path | None]],
    password: str,
    workers: int = PDF_LOCK_WORKERS
) -> dict[str, dict]:
    """
    Verrouille tous les PDFs (factures + subsides) en parallèle.

    lock_tasks : liste de tuples (contract, pdf_input, locked_facture, subsidy_input_or_None, locked_subsidy_or_None)
    Retourne un dict {contract: {"facture": (ok, err, path), "subside": (ok, err, path) | None}}

    OPT 10 : skip si le fichier verrouillé existe déjà.
    """
    results = {}

    def _lock_one(contract, src, dst):
        # OPT 10 (v12) : skip seulement si la copie verrouillée est plus récente
        # que la source — sinon c'est probablement celle du mois précédent.
        if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime:
            logging.debug(f"[{contract}] PDF déjà verrouillé, skip : {dst.name}")
            return True, None, dst
        ok, err = lock_pdf(src, dst, password)
        return ok, err, dst

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {}  # future -> (contract, kind)
        for (contract, pdf_input, locked_facture, subsidy_input, locked_subsidy) in lock_tasks:
            results[contract] = {"facture": None, "subside": None}
            futures[executor.submit(_lock_one, contract, pdf_input, locked_facture)] = (contract, "facture")
            if subsidy_input and locked_subsidy:
                futures[executor.submit(_lock_one, contract, subsidy_input, locked_subsidy)] = (contract, "subside")

        # as_completed : on récupère les résultats dès qu'ils sont prêts → progression réelle
        for future in tqdm(as_completed(futures), total=len(futures), desc="Verrouillage PDF"):
            contract, kind = futures[future]
            try:
                results[contract][kind] = future.result()
            except Exception as e:
                results[contract][kind] = (False, str(e), None)

    return results


# ---------------------------------------------------------
# Création brouillon Outlook (séquentiel — COM non thread-safe)
# ---------------------------------------------------------

def create_outlook_draft_in_store(
    outlook,
    to_addr: str,
    subject: str,
    body: str,
    attachment_paths: list[Path],
    drafts_folder,
    smtp: str = "",
    locked_paths: list[Path] | None = None
) -> tuple[bool, str | None]:
    """
    v12 : reçoit directement le dossier Brouillons (résolu une seule fois dans
    main) au lieu de re-parcourir tous les Stores Outlook à chaque mail.
    drafts_folder=None → boîte par défaut.
    """

    def _attach_all(mail):
        for p in attachment_paths:
            mail.Attachments.Add(str(p))

    def _cleanup_locked():
        for p in (locked_paths or []):
            if p and p.exists():
                try:
                    p.unlink()
                    logging.info(f"PDF verrouillé supprimé (échec brouillon) : {p.name}")
                except Exception as cleanup_err:
                    logging.warning(f"Impossible de supprimer {p.name} : {cleanup_err}")

    if drafts_folder is not None:
        try:
            mail = drafts_folder.Items.Add("IPM.Note")
            if smtp:
                try:
                    mail.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x5D01001F", smtp
                    )
                except Exception:
                    pass
                try:
                    mail.SentOnBehalfOfName = smtp
                except Exception:
                    pass

            mail.To = to_addr
            mail.Subject = subject
            mail.Body = body
            _attach_all(mail)
            mail.Save()
            return True, None

        except Exception as e:
            logging.warning(f"Échec dossier Brouillons du Store : {e}. Tentative boîte par défaut.")

    try:
        mail = outlook.CreateItem(0)
        mail.To = to_addr
        mail.Subject = subject
        mail.Body = body
        _attach_all(mail)
        mail.Save()
        fallback_msg = "Fallback boîte par défaut" if drafts_folder is not None else "Boîte par défaut"
        return True, fallback_msg

    except Exception as e2:
        _cleanup_locked()
        return False, str(e2)


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------

def main():
    start_time = datetime.now()
    logging.info("=== Démarrage ===")

    outlook = get_outlook_instance()

    store_choice, test_mode = ask_outlook_store_and_options(outlook)
    if not store_choice:
        logging.warning("Aucune boîte sélectionnée. Les brouillons iront dans la boîte par défaut.")
        store_choice = None

    # v12 OPT 11 — Résolution unique du dossier Brouillons + adresse SMTP
    drafts_folder, smtp = None, ""
    if store_choice:
        _, drafts_folder, smtp = get_store_and_drafts_folder(outlook, store_choice)
        if drafts_folder is None:
            logging.warning(f"Brouillons introuvables pour '{store_choice}'. Boîte par défaut utilisée.")
        else:
            logging.info(f"Brouillons cible : {store_choice} ({smtp or 'SMTP inconnu'})")

    if test_mode:
        logging.info(f"Mode Test activé : seulement {TEST_ROWS} lignes seront traitées.")

    excel_file = ask_excel_file()
    if not excel_file:
        logging.error("Aucun fichier Excel sélectionné.")
        sys.exit(1)

    pdf_folder = ask_pdf_folder()
    if not pdf_folder:
        logging.error("Aucun dossier PDF sélectionné.")
        sys.exit(1)

    subsidy_folder = ask_subsidy_folder()
    if subsidy_folder:
        logging.info(f"Dossier subsides : {subsidy_folder}")
    else:
        logging.info("Aucun dossier de subsides — mails sans subside.")

    locked_folder = pdf_folder / "Factures_Locked"
    locked_folder.mkdir(parents=True, exist_ok=True)

    # FIX 3 + v12 OPT 12/13 : lecture unique, en-tête auto, cellules vides = ""
    df, header_row = load_excel(excel_file, CONTRACT_COL)

    if CONTRACT_COL not in df.columns:
        logging.error(f"Colonne '{CONTRACT_COL}' non trouvée (en-tête détecté ligne {header_row}).")
        sys.exit(1)

    email_cols = [c for c in POSSIBLE_EMAIL_COLS if c in df.columns]
    if not email_cols:
        logging.error("Aucune colonne email trouvée.")
        sys.exit(1)

    if RESPONSABLE_COL not in df.columns:
        logging.warning("Colonne EMailResponsable absente.")

    if test_mode:
        df = df.head(TEST_ROWS)

    # v11 OPT 8 — Pré-indexation des PDFs (une seule lecture du dossier)
    logging.info("Indexation des PDFs...")
    pdf_index = index_pdfs(pdf_folder)
    subsidy_index = index_pdfs(subsidy_folder) if subsidy_folder else {}

    # --- Première passe : préparer les tâches de verrouillage ---
    # On collecte les infos de chaque contrat valide avant de lancer le parallèle
    logging.info("Préparation des contrats...")
    report_rows = []
    valid_rows = []   # contrats à traiter
    success = 0

    # to_dict("records") : liste de dicts, bien plus rapide que df.iterrows()
    for row in df.to_dict("records"):
        contract = row.get(CONTRACT_COL, "").strip()
        email_facturation = get_email_from_row(row, email_cols)
        responsable_email = row.get(RESPONSABLE_COL, "").strip()

        if not contract:
            report_rows.append([contract, "", "SKIPPED", "Contrat manquant", ""])
            continue
        if not email_facturation:
            report_rows.append([contract, "", "SKIPPED", "Aucun email (facturation)", ""])
            continue

        pdf_input = find_invoice_pdf(contract, pdf_index)
        if not pdf_input:
            report_rows.append([contract, email_facturation, "FAILED", "PDF introuvable", ""])
            continue

        locked_facture = locked_folder / f"{pdf_input.stem}_locked.pdf"

        subsidy_input = find_subsidy_pdf(contract, subsidy_index) if subsidy_index else None
        locked_subsidy = (locked_folder / f"Subside {contract}_locked.pdf") if subsidy_input else None

        valid_rows.append({
            "contract": contract,
            "email": email_facturation,
            "responsable": responsable_email,
            "pdf_input": pdf_input,
            "locked_facture": locked_facture,
            "subsidy_input": subsidy_input,
            "locked_subsidy": locked_subsidy,
        })

    # v11 OPT 9 — Verrouillage parallèle de tous les PDFs d'un coup
    logging.info(f"Verrouillage parallèle de {len(valid_rows)} contrats ({PDF_LOCK_WORKERS} workers)...")
    lock_tasks = [
        (r["contract"], r["pdf_input"], r["locked_facture"], r["subsidy_input"], r["locked_subsidy"])
        for r in valid_rows
    ]
    lock_results = lock_all_pdfs_parallel(lock_tasks, PDF_PASSWORD, PDF_LOCK_WORKERS)

    # --- Deuxième passe : création des brouillons Outlook (séquentiel) ---
    logging.info("Création des brouillons Outlook...")
    for r in tqdm(valid_rows, desc="Brouillons Outlook"):
        contract      = r["contract"]
        email         = r["email"]
        responsable   = r["responsable"] or "votre gestionnaire"
        res           = lock_results.get(contract, {})

        # Vérification verrouillage facture
        fac_ok, fac_err, locked_facture = res.get("facture") or (False, "Résultat manquant", None)
        if not fac_ok:
            report_rows.append([contract, email, "FAILED_LOCK", fac_err, ""])
            continue

        # Vérification verrouillage subside (non bloquant)
        subside_statut = "NON"
        locked_subsidy = None
        sub_result = res.get("subside")
        if sub_result:
            sub_ok, sub_err, sub_path = sub_result
            if sub_ok:
                locked_subsidy = sub_path
                subside_statut = "OUI"
            else:
                logging.warning(f"Échec verrouillage subside {contract} : {sub_err}")

        # Construction du mail
        body = BODY_TEMPLATE.format(responsable=responsable)
        subject = SUBJECT_TEMPLATE.format(contract=contract)
        attachment_paths = [locked_facture]
        if locked_subsidy:
            attachment_paths.append(locked_subsidy)

        ok_mail, info = create_outlook_draft_in_store(
            outlook=outlook,
            to_addr=email,
            subject=subject,
            body=body,
            attachment_paths=attachment_paths,
            drafts_folder=drafts_folder,
            smtp=smtp,
            locked_paths=list(attachment_paths)
        )

        if not ok_mail:
            report_rows.append([contract, email, "FAILED_MAIL", info, subside_statut])
            continue

        success += 1
        report_rows.append([
            contract, email, "OK",
            f"{responsable} | {info or 'Store OK'}",
            subside_statut
        ])

    # Rapport CSV
    report_path = excel_file.parent / f"rapport_envoi_factures_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(report_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Contrat", "Email Facturation", "Statut", "Info", "Subside"])
        writer.writerows(report_rows)

    elapsed = datetime.now() - start_time
    minutes, seconds = divmod(int(elapsed.total_seconds()), 60)
    duration_str = f"{minutes}m {seconds}s" if minutes else f"{seconds}s"

    logging.info(f"=== Terminé — {success}/{len(df)} contrats traités ===")
    logging.info(f"Durée totale : {duration_str}")
    logging.info(f"Rapport : {report_path}")


if __name__ == "__main__":
    main()
