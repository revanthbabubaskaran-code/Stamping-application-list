"""Create a date-sorted client equipment list from a text-based PDF.

Run with:  python client_list_generator.py
Dependencies: pdfplumber, openpyxl (install with: pip install -r requirements.txt)
"""

from __future__ import annotations

import re
import sys
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime
from math import ceil
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


HEADERS = [
    "Date of Stamping", "Address", "Machine Details", "SI",
]

DATE_FORMATS = (
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y",
    "%Y-%m-%d", "%Y/%m/%d", "%d %B %Y", "%d %b %Y",
)


@dataclass
class ClientRecord:
    stamping_date: date | None = None
    address: str = ""
    machine_details: list[str] | None = None
    sis: list[str] | None = None
    source_pdf: str = ""
    file_status: str = ""

    def __post_init__(self) -> None:
        self.machine_details = self.machine_details or []
        self.sis = self.sis or []

    def values(self) -> list[str | date]:
        return [
            self.stamping_date or "", self.address,
            "\n".join(self.machine_details), "\n".join(self.sis),
        ]


def parse_date(value: str) -> date | None:
    value = value.strip().rstrip(".,;:")
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    # Accept a date embedded in surrounding text.
    match = re.search(r"\b(?:\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})\b", value)
    if match:
        return parse_date(match.group(0))
    return None


def clean_value(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip(" \t:;|-")


def normalize_si(text: str) -> str:
    """Normalize inspector designations and locations to a short SI label."""
    candidates: list[str] = []
    deputy_pattern = re.compile(
        r"Deputy\s+Inspector\s+of\s+Labou?r\s*[-–]\s*(\d+|[IVXLCDM]+)"
        r"\s*,\s*([A-Za-z][A-Za-z .'-]*)",
        re.I,
    )
    circle_pattern = re.compile(
        r"(?:Stamping\s+Inspector\s*/\s*)?Inspector\s+of\s+Legal\s+Metrology\s*,\s*"
        r"(\d+)(?:st|nd|rd|th)\s+Circle\s*,\s*([A-Za-z][A-Za-z .'-]*)",
        re.I,
    )
    numbered_pattern = re.compile(
        r"Stamping\s+Inspector\s*-\s*(\d+)\s*,?\s*([A-Za-z][A-Za-z .'-]*)?",
        re.I,
    )
    location_pattern = re.compile(
        r"Stamping\s+Inspector\s*,?\s+([A-Za-z][A-Za-z.'-]*)",
        re.I,
    )

    for line in text.splitlines():
        match = deputy_pattern.search(line)
        if match:
            inspector_level = match.group(1).upper()
            if inspector_level.isdigit():
                level = int(inspector_level)
            else:
                roman_values = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
                level = 0
                previous = 0
                for numeral in reversed(inspector_level):
                    value = roman_values[numeral]
                    level += -value if value < previous else value
                    previous = max(previous, value)
            location = match.group(2).strip().rstrip(" ,.-").title()
            candidates.append(f"DIL {level}, {location}" if location else f"DIL {level}")
            continue
        match = circle_pattern.search(line)
        if match:
            location = match.group(2).strip().rstrip(" ,.-").title()
            candidates.append(f"SI {int(match.group(1))}, {location}" if location else f"SI {int(match.group(1))}")
            continue
        match = numbered_pattern.search(line)
        if match:
            location = (match.group(2) or "").strip().rstrip(" ,.-").title()
            candidates.append(f"SI {int(match.group(1))}, {location}" if location else f"SI {int(match.group(1))}")
            continue
        match = location_pattern.search(line)
        if match:
            location = match.group(1).strip().rstrip(" ,.-").title()
            candidates.append(f"SI, {location}" if location else "SI")
    return candidates[-1] if candidates else ""


def _words_in_row(words: list[dict], *, x_min: float, x_max: float,
                  top_min: float, top_max: float) -> str:
    selected = [w for w in words if x_min <= w["x0"] < x_max and top_min <= w["top"] <= top_max]
    selected.sort(key=lambda w: (w["top"], w["x0"]))
    return clean_value(" ".join(w["text"] for w in selected))


def _extract_machine(page) -> dict[str, str] | None:
    """Read the equipment row by its printed table columns and labels."""
    words = page.extract_words()
    text = page.extract_text() or ""

    # The Items column is the leftmost table column. Join its wrapped lines.
    item_words = [w for w in words if 40 <= w["x0"] < 96 and 395 <= w["top"] <= 480]
    row_lines: list[list] = []
    for word in sorted(item_words, key=lambda w: (w["top"], w["x0"])):
        if not row_lines or abs(word["top"] - row_lines[-1][0]) > 2:
            row_lines.append([word["top"], word["text"]])
        else:
            row_lines[-1][1] += " " + word["text"]
    item_text = " ".join(line[1] for line in row_lines)
    item_text = item_text.replace("\ufffd", "–")
    item_text = re.sub(r"\bNon-\s+Automatic\b", "Non-Automatic", item_text, flags=re.I)
    item_text = re.sub(r"\s+([,)])", r"\1", item_text)
    item_text = re.sub(r"\(\s+", "(", item_text)
    item_text = re.sub(r"\s+", " ", item_text).strip()

    # Some certificates list physical weights rather than a numbered machine.
    # Those rows have no make, Max capacity, or machine number; use their item
    # description and denomination so the client PDF is still represented.
    item_top = min((word["top"] for word in item_words), default=405)
    if re.search(r"\bweights?\b", item_text, re.I) and not re.search(
        r"\bNon-Automatic\s+Weighing\s+Instruments\b", item_text, re.I
    ):
        denomination = _words_in_row(
            words, x_min=120, x_max=172,
            top_min=item_top - 2, top_max=item_top + 3,
        ).rstrip(".")
        if item_text and denomination:
            if re.search(
                r"\bIron\s+hexagonal\s+knob\s+type\s+weights\s+and\s+parallelepiped\s+weights\b",
                item_text, re.I,
            ):
                return {"machine_details": f"{denomination}, Test Weight"}
            return {"machine_details": f"{item_text}, {denomination}"}

    if not re.search(r"\bMachine\s+No\.?\s*:", text, re.I):
        return None

    type_match = re.search(
        r"(Non-Automatic\s+Weighing\s+Instruments\s*[–-]\s*"
        r"(?:Electronic|Mechanical)(?:\s*\([^)]*\))?\s+Class\s+III\s+and\s+IIII)",
        item_text, re.I,
    )
    item_description = type_match.group(1) if type_match else item_text

    make_label = next((w for w in words if w["text"].lower() == "make:" and 290 <= w["x0"] <= 330), None)
    max_label = next((w for w in words if w["text"].lower().startswith("max") and 290 <= w["x0"] <= 330), None)
    machine_label = next((w for w in words if w["text"].lower() == "machine" and 370 <= w["x0"] <= 410), None)
    make_row_top = make_label["top"] if make_label else 405

    make = _words_in_row(words, x_min=make_label["x1"] if make_label else 320,
                         x_max=350, top_min=(make_label["top"] - 1) if make_label else 400,
                         top_max=(make_label["top"] + 3) if make_label else 410)
    if not make and make_label:
        # Some certificates place the manufacturer on the line below "Make:".
        make = _words_in_row(words, x_min=295, x_max=350,
                             top_min=make_label["top"] + 4,
                             top_max=make_label["top"] + 14)
    capacity = _words_in_row(words, x_min=(max_label["x1"] if max_label else 295), x_max=350,
                             top_min=(max_label["top"] - 2) if max_label else 400,
                             top_max=(max_label["top"] + 3) if max_label else 410)
    if not capacity and max_label:
        # On some forms Max and its numeric value are printed on adjacent lines.
        capacity = _words_in_row(words, x_min=295, x_max=350,
                                 top_min=max_label["top"] + 4,
                                 top_max=max_label["top"] + 14)
    machine_no = _words_in_row(words, x_min=377, x_max=428,
                               top_min=(machine_label["top"] + 3) if machine_label else 405,
                               top_max=(machine_label["top"] + 12) if machine_label else 420)
    machine_no = clean_value(re.sub(r"^No\.?\s*:\s*", "", machine_no, flags=re.I))
    machine_no = clean_value(re.sub(r"^:\s*", "", machine_no))
    machine_class = _words_in_row(words, x_min=270, x_max=300,
                                  top_min=make_row_top - 2, top_max=make_row_top + 2)
    class_match = re.search(r"\b(IIII|III|II)\b", machine_class, re.I)
    class_value = class_match.group(1).upper() if class_match else ""
    if re.search(r"\bMechanical\b", item_description, re.I):
        machine_type = f"Mechanical, Class {class_value}" if class_value else ""
    elif re.search(r"\bElectronic\b", item_description, re.I):
        machine_type = f"Electronic Class {class_value}" if class_value else ""
    else:
        machine_type = ""

    # Keep values that can be matched to an equipment row.
    if not all((capacity, make, machine_no, machine_type)):
        return None
    machine_details = f"{make}, {capacity}, {machine_type}, M/C: {machine_no}"
    return {"machine_details": machine_details}


def extract_certificate_page(page) -> tuple[dict | None, str, str, str]:
    text = page.extract_text() or ""
    date_match = re.search(r"\bDate\s*:\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text, re.I)
    stamping_date = parse_date(date_match.group(1)) if date_match else None
    address_match = re.search(r"belonging\s+to\s+(M/S\s+.+?)\s+Items\b", text, re.I | re.S)
    address = clean_value(address_match.group(1)) if address_match else ""
    si = normalize_si(text)
    machine = _extract_machine(page)
    return machine, address, si, stamping_date.isoformat() if stamping_date else ""


def read_pdf(pdf_path: Path) -> tuple[list[ClientRecord], int]:
    record: ClientRecord | None = None
    blank_pages = 0
    unrecognized_pages = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            machine, address, si, iso_date = extract_certificate_page(page)
            if machine is None:
                footer_only = re.sub(r"Powered by TCPDF(?:\s*\([^)]*\))?", "", text, flags=re.I)
                footer_only = re.sub(r"\b\d+\s*/\s*\d+\b", "", footer_only)
                if not footer_only.strip():
                    blank_pages += 1
                else:
                    unrecognized_pages += 1
                continue
            current_date = date.fromisoformat(iso_date) if iso_date else None
            if record is None:
                record = ClientRecord(stamping_date=current_date, address=address)
            elif current_date and not record.stamping_date:
                record.stamping_date = current_date
            if not record.address:
                record.address = address
            if si and si not in record.sis:
                record.sis.append(si)
            record.machine_details.append(machine["machine_details"])
    if record is None:
        reason = "FAILED: no machine rows could be extracted. The PDF may be scanned or use an unsupported layout."
        if unrecognized_pages:
            reason += f" {unrecognized_pages} nonblank page(s) had text but no recognized machine row."
        record = ClientRecord(source_pdf=str(pdf_path), file_status=reason)
    else:
        record.source_pdf = str(pdf_path)
        machine_count = len(record.machine_details)
        if unrecognized_pages:
            record.file_status = (
                f"PARTIAL: added 1 row with {machine_count} item entr{'y' if machine_count == 1 else 'ies'}; "
                f"{unrecognized_pages} nonblank page(s) had text but no recognized machine row."
            )
        else:
            record.file_status = (
                f"SUCCESS: added 1 row with {machine_count} item entr{'y' if machine_count == 1 else 'ies'}."
            )
            if blank_pages:
                record.file_status += f" Ignored {blank_pages} blank page(s)."
    return [record], blank_pages + unrecognized_pages


def write_excel(records: list[ClientRecord], output_path: Path) -> None:
    # The alphabetical tie-break is applied to Address, the requested
    # client-identifying field. Failed files have no client row to export.
    export_records = sorted((row for row in records if not row.file_status.startswith("FAILED")), key=lambda row: (
        row.stamping_date is None,
        row.stamping_date or date.max,
        row.address.casefold(),
        row.source_pdf.casefold(),
    ))
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Client List"
    sheet.append(HEADERS)
    for record in export_records:
        sheet.append(record.values())
        sheet.cell(sheet.max_row, 1).number_format = "dd/mm/yyyy"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    widths = [19, 48, 82, 38]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        wrapped_line_counts = []
        for index, cell in enumerate(row, start=1):
            column_width = widths[index - 1]
            wrapped_line_counts.append(sum(
                max(1, ceil(len(line) / max(column_width - 2, 10)))
                for line in str(cell.value or "").splitlines()
            ))
        sheet.row_dimensions[row[0].row].height = max(30, 18 * max(wrapped_line_counts))
    workbook.save(output_path)


class GeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Stamping List Generator")
        self.geometry("900x680")
        self.minsize(700, 560)
        self.pdf_paths: list[Path] = []
        self.status = tk.StringVar(value="Add one or more searchable PDF files to begin.")

        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.main_frame = ttk.Frame(self, padding=24, style="App.TFrame")
        self.main_frame.pack(fill="both", expand=True)
        ttk.Label(self.main_frame, text="Stamping List Generator", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            self.main_frame,
            text="Select stamping certificate PDFs. The Excel list will contain one row per PDF, combining machine entries from its certificate pages.",
            style="Muted.TLabel", wraplength=740,
        ).pack(anchor="w", pady=(5, 18))

        ttk.Label(self.main_frame, text="Selected PDF files", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        list_area = ttk.Frame(self.main_frame, style="App.TFrame")
        list_area.pack(fill="both", expand=True)
        self.file_list = tk.Listbox(
            list_area, selectmode=tk.EXTENDED, activestyle="dotbox",
            font=("Segoe UI", 10), relief="solid", borderwidth=1,
        )
        self.file_list.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_area, orient="vertical", command=self.file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_list.configure(yscrollcommand=scrollbar.set)

        ttk.Label(self.main_frame, text="Failed / partial files", style="Section.TLabel").pack(anchor="w", pady=(12, 7))
        results_area = ttk.Frame(self.main_frame, style="App.TFrame")
        results_area.pack(fill="both", expand=True)
        self.results_text = tk.Text(
            results_area, height=7, wrap="word", font=("Segoe UI", 9),
            relief="solid", borderwidth=1, state="disabled",
        )
        self.results_text.pack(side="left", fill="both", expand=True)
        results_scrollbar = ttk.Scrollbar(results_area, orient="vertical", command=self.results_text.yview)
        results_scrollbar.pack(side="right", fill="y")
        self.results_text.configure(yscrollcommand=results_scrollbar.set)

        actions = ttk.Frame(self.main_frame, style="App.TFrame")
        actions.pack(fill="x", pady=(12, 10))
        ttk.Button(actions, text="Add PDF file(s)…", command=self.browse).pack(side="left")
        ttk.Button(actions, text="Remove selected", command=self.remove_selected).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Clear all", command=self.clear_all).pack(side="left", padx=(8, 0))
        self.create_button = ttk.Button(actions, text="Create list", command=self.generate, style="Accent.TButton")
        self.create_button.pack(side="right")

        ttk.Label(self.main_frame, textvariable=self.status, style="Status.TLabel", wraplength=740).pack(anchor="w", pady=(5, 0))
        self.apply_system_theme()
        self.after(1500, self.check_theme)

    @staticmethod
    def system_prefers_dark() -> bool:
        """Read Windows' current app theme preference; use light elsewhere."""
        if sys.platform == "win32":
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
                    return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
            except (OSError, ImportError):
                pass
        return False

    def apply_system_theme(self) -> None:
        dark = self.system_prefers_dark()
        self.current_theme = "dark" if dark else "light"
        if dark:
            colors = {"background": "#202124", "surface": "#292A2D", "foreground": "#E8EAED",
                      "muted": "#B8BCC2", "field": "#17181A", "select": "#3B5B82", "button": "#34363A"}
        else:
            colors = {"background": "#F4F6F8", "surface": "#FFFFFF", "foreground": "#202124",
                      "muted": "#5F6368", "field": "#FFFFFF", "select": "#CFE2F7", "button": "#E8EDF2"}
        self.configure(background=colors["background"])
        self.style.configure("App.TFrame", background=colors["background"])
        self.style.configure("Title.TLabel", background=colors["background"], foreground=colors["foreground"], font=("Segoe UI", 20, "bold"))
        self.style.configure("Section.TLabel", background=colors["background"], foreground=colors["foreground"], font=("Segoe UI", 11, "bold"))
        self.style.configure("Muted.TLabel", background=colors["background"], foreground=colors["muted"], font=("Segoe UI", 10))
        self.style.configure("Status.TLabel", background=colors["background"], foreground=colors["muted"], font=("Segoe UI", 9))
        self.style.configure("TButton", background=colors["button"], foreground=colors["foreground"], padding=(12, 8))
        self.style.map("TButton", background=[("active", colors["select"])])
        self.style.configure("Accent.TButton", background="#2463A6", foreground="#FFFFFF", padding=(18, 9), font=("Segoe UI", 10, "bold"))
        self.style.map("Accent.TButton", background=[("active", "#174A83"), ("disabled", colors["button"])])
        self.style.configure("Vertical.TScrollbar", background=colors["button"], troughcolor=colors["background"], arrowcolor=colors["foreground"])
        self.file_list.configure(
            background=colors["field"], foreground=colors["foreground"],
            selectbackground=colors["select"], selectforeground=colors["foreground"],
            highlightbackground=colors["background"], highlightcolor=colors["select"],
        )
        self.results_text.configure(
            background=colors["field"], foreground=colors["foreground"],
            insertbackground=colors["foreground"], selectbackground=colors["select"],
            selectforeground=colors["foreground"], highlightbackground=colors["background"],
        )
        self.results_text.tag_configure("success", foreground="#81C784" if dark else "#237A3B")
        self.results_text.tag_configure("partial", foreground="#FFCC80" if dark else "#A45B00")
        self.results_text.tag_configure("failed", foreground="#EF9A9A" if dark else "#B3261E")
        self.results_text.tag_configure("filename", foreground=colors["muted"], font=("Segoe UI", 9, "bold"))

    def clear_results(self) -> None:
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", tk.END)
        self.results_text.configure(state="disabled")

    def display_results(self, records: list[ClientRecord]) -> None:
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", tk.END)
        attention_records = [record for record in records if not record.file_status.startswith("SUCCESS")]
        if not attention_records:
            self.results_text.insert(tk.END, "No failed or partial files.")
        for record in attention_records:
            state = record.file_status.partition(":")[0].lower()
            self.results_text.insert(tk.END, f"{record.source_pdf}\n", "filename")
            self.results_text.insert(tk.END, f"{record.file_status}\n\n", state)
        self.results_text.configure(state="disabled")
        self.results_text.see("1.0")

    def check_theme(self) -> None:
        if ("dark" if self.system_prefers_dark() else "light") != self.current_theme:
            self.apply_system_theme()
        self.after(1500, self.check_theme)

    def browse(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Add client PDF files", filetypes=[("PDF files", "*.pdf")],
        )
        existing = {str(path).casefold() for path in self.pdf_paths}
        added = False
        for filename in selected:
            path = Path(filename)
            key = str(path).casefold()
            if key not in existing:
                self.pdf_paths.append(path)
                self.file_list.insert(tk.END, str(path))
                existing.add(key)
                added = True
        if added:
            self.clear_results()
        self.status.set(f"{len(self.pdf_paths)} PDF file(s) selected.")

    def remove_selected(self) -> None:
        indexes = list(self.file_list.curselection())
        for index in reversed(indexes):
            self.file_list.delete(index)
            del self.pdf_paths[index]
        if indexes:
            self.clear_results()
        self.status.set(f"{len(self.pdf_paths)} PDF file(s) selected.")

    def clear_all(self) -> None:
        self.pdf_paths.clear()
        self.file_list.delete(0, tk.END)
        self.clear_results()
        self.status.set("All selected files cleared.")

    def generate(self) -> None:
        if not self.pdf_paths:
            messagebox.showerror("PDF required", "Add at least one PDF file first.", parent=self)
            return
        first_available = next((path for path in self.pdf_paths if path.is_file()), None)
        initial_dir = str(first_available.parent) if first_available else str(Path.cwd())
        initial_name = f"{self.pdf_paths[0].stem}_client_list.xlsx" if len(self.pdf_paths) == 1 else "combined_client_list.xlsx"
        output = filedialog.asksaveasfilename(
            title="Save Excel list", initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        try:
            records: list[ClientRecord] = []
            for source in self.pdf_paths:
                if not source.is_file():
                    records.append(ClientRecord(
                        source_pdf=str(source),
                        file_status="FAILED: the selected file could not be found. It may have been moved or deleted.",
                    ))
                    continue
                try:
                    file_records, _ = read_pdf(source)
                    records.extend(file_records)
                except Exception as exc:
                    reason = re.sub(r"\s+", " ", str(exc)).strip() or type(exc).__name__
                    records.append(ClientRecord(
                        source_pdf=str(source),
                        file_status=f"FAILED: could not read this PDF ({reason[:500]}).",
                    ))
            try:
                write_excel(records, Path(output))
            except Exception as exc:
                reason = re.sub(r"\s+", " ", str(exc)).strip() or type(exc).__name__
                for record in records:
                    record.file_status = f"FAILED: the Excel workbook could not be saved ({reason[:500]})."
                self.display_results(records)
                raise
            self.display_results(records)
            attention_count = sum(not record.file_status.startswith("SUCCESS") for record in records)
            if attention_count:
                note = f"Created the Excel file at {output}. {attention_count} failed or partial file(s) are listed in the app."
            else:
                note = f"Created the Excel file at {output}. No failed or partial files."
            self.status.set(note)
            messagebox.showinfo("Excel created", note, parent=self)
        except Exception as exc:
            messagebox.showerror("Could not generate list", str(exc), parent=self)


if __name__ == "__main__":
    GeneratorApp().mainloop()
