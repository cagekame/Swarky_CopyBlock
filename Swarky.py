#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations
import sys, re, time, logging, json, os, shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple

# Solo Windows
if sys.platform != "win32":
    raise RuntimeError("Questo programma è supportato solo su Windows.")

# ---- CONFIG DATACLASS ----------------------------------------------------------------

@dataclass(frozen=True)
class Config:
    DIR_HPLOTTER: Path
    ARCHIVIO_DISEGNI: Path
    ERROR_DIR: Path
    PARI_REV_DIR: Path
    PLM_DIR: Path
    ARCHIVIO_STORICO: Path
    DIR_ISS: Path
    DIR_FIV_LOADING: Path
    DIR_HENGELO: Path
    DIR_PLM_ERROR: Path
    DIR_TABELLARI: Path
    LOG_DIR: Optional[Path] = None
    LOG_LEVEL: int = logging.INFO
    ACCEPT_PDF: bool = True
    LOG_PHASES: bool = True

    @staticmethod
    def from_json(d: Dict[str, Any]) -> "Config":
        p = d.get("paths", {})
        def P(key: str, default: Optional[str]=None) -> Path:
            val = p.get(key, default)
            if val is None:
                raise KeyError(f"Config mancante: paths.{key}")
            return Path(val)
        log_dir = p.get("log_dir")
        return Config(
            DIR_HPLOTTER=P("hplotter"),
            ARCHIVIO_DISEGNI=P("archivio"),
            ERROR_DIR=P("error_dir"),
            PARI_REV_DIR=P("pari_rev"),
            PLM_DIR=P("plm"),
            ARCHIVIO_STORICO=P("storico"),
            DIR_ISS=P("iss"),
            DIR_FIV_LOADING=P("fiv"),
            DIR_HENGELO=P("heng"),
            DIR_PLM_ERROR=P("error_plm"),
            DIR_TABELLARI=P("tab"),
            LOG_DIR=Path(log_dir) if log_dir else None,
            LOG_LEVEL=logging.INFO,
            ACCEPT_PDF=bool(d.get("ACCEPT_PDF", True)),
            LOG_PHASES=bool(d.get("LOG_PHASES", True)),
        )

# ---- REGEX ---------------------------------------------------------------------------

BASE_NAME = re.compile(r"D(\w)(\w)(\d{6})R(\d{2})S(\d{2})(\w)\.(tif|pdf)$", re.IGNORECASE)
ISS_BASENAME = re.compile(r"G(\d{4})([A-Za-z0-9]{4})([A-Za-z0-9]{6})ISSR(\d{2})S(\d{2})\.pdf$", re.IGNORECASE)

# ---- PREFISSO DOCNO: LISTA NOMI SENZA ENUM COMPLETA -------------------------

import ctypes
import ctypes.wintypes as wt

INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
FILE_ATTRIBUTE_DIRECTORY = 0x10
FIND_FIRST_EX_LARGE_FETCH = 2
FindExInfoBasic = 1
FindExSearchNameMatch = 0
ERROR_FILE_NOT_FOUND = 2
ERROR_PATH_NOT_FOUND = 3

class WIN32_FIND_DATAW(ctypes.Structure):
    _fields_ = [
        ("dwFileAttributes", wt.DWORD),
        ("ftCreationTime", wt.FILETIME),
        ("ftLastAccessTime", wt.FILETIME),
        ("ftLastWriteTime", wt.FILETIME),
        ("nFileSizeHigh", wt.DWORD),
        ("nFileSizeLow", wt.DWORD),
        ("dwReserved0", wt.DWORD),
        ("dwReserved1", wt.DWORD),
        ("cFileName", ctypes.c_wchar * 260),
        ("cAlternateFileName", ctypes.c_wchar * 14),
    ]

_k32 = ctypes.WinDLL("kernel32", use_last_error=True)
_FindFirstFileW = _k32.FindFirstFileW
_FindFirstFileW.argtypes = [wt.LPCWSTR, ctypes.POINTER(WIN32_FIND_DATAW)]
_FindFirstFileW.restype = wt.HANDLE
_FindNextFileW = _k32.FindNextFileW
_FindNextFileW.argtypes = [wt.HANDLE, ctypes.POINTER(WIN32_FIND_DATAW)]
_FindNextFileW.restype = wt.BOOL
_FindClose = _k32.FindClose
_FindClose.argtypes = [wt.HANDLE]
_FindClose.restype = wt.BOOL

try:
    _FindFirstFileExW = _k32.FindFirstFileExW
    _FindFirstFileExW.argtypes = [
        wt.LPCWSTR,
        ctypes.c_int,
        ctypes.POINTER(WIN32_FIND_DATAW),
        ctypes.c_int,
        ctypes.c_void_p,
        wt.DWORD,
    ]
    _FindFirstFileExW.restype = wt.HANDLE
except AttributeError:
    _FindFirstFileExW = None

def _win_find_names(dirp: Path, pattern: str) -> tuple[str, ...]:
    query = str(dirp / pattern)
    data = WIN32_FIND_DATAW()
    h = _FindFirstFileW(query, ctypes.byref(data))
    if h == INVALID_HANDLE_VALUE:
        return tuple()
    names: list[str] = []
    try:
        while True:
            nm = data.cFileName
            if nm not in (".", "..") and not (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY):
                names.append(nm)
            if not _FindNextFileW(h, ctypes.byref(data)):
                break
    finally:
        _FindClose(h)
    return tuple(names)

def _win_find_names_ex(dirp: Path, pattern: str) -> tuple[str, ...]:
    if _FindFirstFileExW is None:
        return _win_find_names(dirp, pattern)
    query = str(dirp / pattern)
    data = WIN32_FIND_DATAW()
    h = _FindFirstFileExW(
        query,
        FindExInfoBasic,
        ctypes.byref(data),
        FindExSearchNameMatch,
        None,
        FIND_FIRST_EX_LARGE_FETCH,
    )
    if h == INVALID_HANDLE_VALUE:
        err = ctypes.get_last_error()
        if err in (ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND):
            return tuple()
        return _win_find_names(dirp, pattern)
    names: list[str] = []
    try:
        while True:
            nm = data.cFileName
            if nm not in (".", "..") and not (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY):
                names.append(nm)
            if not _FindNextFileW(h, ctypes.byref(data)):
                break
    finally:
        _FindClose(h)
    return tuple(names)

# ---- COPIA A BLOCCHI ----------------------------------

# Buffer di default (MiB) – modifiare il valore 8 per cambiare il buffer
_DEF_BUF_MIB = int(os.environ.get("SWARKY_COPY_BUF_MIB", "8"))
_DEF_BUF = max(1, _DEF_BUF_MIB) * 1024 * 1024

def _copy_file_blocks(src: Path, dst: Path, *, overwrite: bool = True, bufsize: int = _DEF_BUF) -> None:
    """
    Copia "user-mode" a blocchi, opzionalmente senza overwrite.
    - Se overwrite=False il file di destinazione viene creato con O_CREAT|O_EXCL (fail se esiste).
    - Preserva mtime/atime (copystat) quando possibile.
    """
    if not overwrite:
        # Crea il file solo se NON esiste
        flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
        # su Windows aggiungi flag binari
        if hasattr(os, "O_BINARY"):
            flags |= os.O_BINARY
        fd = os.open(str(dst), flags)
        try:
            with open(src, "rb", buffering=0) as sf, os.fdopen(fd, "wb", buffering=0) as df:
                shutil.copyfileobj(sf, df, length=bufsize)
        except Exception:
            try:
                os.close(fd)
            except Exception:
                pass
            # se qualcosa va storto, prova a rimuovere il parziale
            try:
                os.unlink(dst)
            except Exception:
                pass
            raise
    else:
        with open(src, "rb", buffering=0) as sf, open(dst, "wb", buffering=0) as df:
            shutil.copyfileobj(sf, df, length=bufsize)

    # preserva metadati base (best effort)
    try:
        shutil.copystat(src, dst)
    except Exception:
        pass

def _fast_copy_or_link(src: Path, dst: Path, *, overwrite: bool = True) -> None:
    """
    1) tenta hardlink (istantaneo se stesso volume/share)
    2) altrimenti copia a blocchi
    """
    try:
        os.link(src, dst)
        return
    except OSError:
        pass
    _copy_file_blocks(src, dst, overwrite=overwrite)

# ---- UTILS PREFISSO ---------------------------------------------------------

def _docno_from_match(m: re.Match) -> str:
    return f"D{m.group(1)}{m.group(2)}{m.group(3)}"

def _parse_prefixed(names: tuple[str, ...]) -> list[tuple[str, str, str, str]]:
    """-> [(rev, name, metric, sheet)]  (rev='02', metric in {M,I,D,N}, sheet='01')"""
    out: list[tuple[str, str, str, str]] = []
    for nm in names:
        mm = BASE_NAME.fullmatch(nm)
        if mm:
            out.append((mm.group(4), nm, mm.group(6).upper(), mm.group(5)))
    return out

def _list_same_doc_prefisso(dirp: Path, m: re.Match) -> list[tuple[str, str, str, str]]:
    """Riduce i round-trip SMB enumerando docno* una sola volta e filtrando in RAM, senza ordinare."""
    docno = _docno_from_match(m)
    names_all = _win_find_names_ex(dirp, f"{docno}*")
    if not names_all:
        return []
    names = tuple(nm for nm in names_all if nm.lower().endswith((".tif", ".pdf")))
    return _parse_prefixed(names)

# ---- LOGGING -------------------------------------------------------------------------

_FILE_LOG_BUF: list[str] = []  # buffer per log-file batch

def month_tag() -> str:
    return datetime.now().strftime("%b.%Y")

def setup_logging(cfg: Config):
    log_dir = cfg.LOG_DIR or cfg.DIR_HPLOTTER
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"Swarky_{month_tag()}.log"

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)

    class _PhaseFilter(logging.Filter):
        def __init__(self, enable_phases: bool):
            super().__init__()
            self.enable_phases = enable_phases
        def filter(self, record: logging.LogRecord) -> bool:
            ui = getattr(record, "ui", None)
            if not self.enable_phases and ui:
                return False
            return True

    fh.addFilter(_PhaseFilter(cfg.LOG_PHASES))

    root = logging.getLogger()
    root.setLevel(cfg.LOG_LEVEL)

    # mantieni altri handler (es. GUI), sostituisci solo il FileHandler
    new_handlers = [h for h in root.handlers if not isinstance(h, logging.FileHandler)]
    new_handlers.append(fh)
    root.handlers = new_handlers

    logging.debug("Log file: %s", log_file)

def _append_filelog_line(line: str) -> None:
    _FILE_LOG_BUF.append(line)

def _flush_file_log(cfg: Config) -> None:
    if not _FILE_LOG_BUF:
        return
    log_path = (cfg.LOG_DIR or cfg.DIR_HPLOTTER) / f"Swarky_{month_tag()}.log"
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as f:
            f.write("\n".join(_FILE_LOG_BUF) + "\n")
    finally:
        _FILE_LOG_BUF.clear()

# ---- FS UTILS ------------------------------------------------------------------------

def _is_same_file(src: Path, dst: Path, *, mtime_slack_ns: int = 2_000_000_000) -> bool:
    try:
        s1 = os.stat(src)
        s2 = os.stat(dst)
    except OSError:
        return False
    return s1.st_size == s2.st_size and abs(s1.st_mtime_ns - s2.st_mtime_ns) <= mtime_slack_ns

def copy_to(src: Path, dst_dir: Path):
    dst_dir.mkdir(parents=True, exist_ok=True)
    _fast_copy_or_link(src, dst_dir / src.name, overwrite=True)

def move_to(src: Path, dst_dir: Path):
    """
    Sposta con rename atomico se possibile; se fallisce (volumi diversi),
    copia a blocchi (overwrite=True) + delete della sorgente.
    """
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    try:
        # rename/replace: su Windows os.rename fallisce se esiste; qui ci aspettiamo nome nuovo
        os.replace(src, dst)   # permettiamo overwrite qui (parità rev già gestita a monte)
        return
    except OSError:
        # cross-volume/share o lock: copia e cancella
        _fast_copy_or_link(src, dst, overwrite=True)
        try:
            src.unlink(missing_ok=True)
        except Exception:
            pass

def move_to_storico_safe(src: Path, dst_dir: Path) -> tuple[bool, int]:
    """
    Sposta src in dst_dir in modalità 'safe' (NON sovrascrivere mai).
    Ritorna (copied_and_deleted, rc_simbolico) dove rc=1 ok, 0 skip perché già presente, 8 errore.
    Nessuna enumerazione preventiva: usiamo errori di sistema.
    """
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    # 1) stesso volume: prova rename che fallisce se il target esiste?
    try:
        # os.rename su Windows fallisce se il target esiste (non sovrascrive)
        os.rename(src, dst)
        return (True, 1)
    except FileExistsError:
        return (False, 0)
    except OSError:
        # 2) cross-volume: copia "safe" con O_EXCL
        try:
            _copy_file_blocks(src, dst, overwrite=False)
            # se la copia è riuscita, cancella sorgente
            try:
                src.unlink()
            except Exception:
                pass
            return (True, 1)
        except FileExistsError:
            return (False, 0)
        except Exception:
            return (False, 8)

def write_lines(p: Path, lines: List[str]):
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

# ---- MAPPATURE, VALIDAZIONI E LOG WRITERS --------------------------------------------

LOCATION_MAP = {
    ("M", "*"): ("costruttivi", "Costruttivi", "m", "DETAIL", "Italian"),
    ("K", "*"): ("bozzetti", "Bozzetti", "k", "Customer Drawings", "English"),
    ("F", "*"): ("fornitori", "Fornitori", "f", "Vendor Supplied Data", "English"),
    ("T", "*"): ("tenute_meccaniche", "T_meccaniche", "t", "Customer Drawings", "English"),
    ("E", "*"): ("sezioni", "Sezioni", "s", "Customer Drawings", "English"),
    ("S", "*"): ("sezioni", "Sezioni", "s", "Customer Drawings", "English"),
    ("N", "*"): ("marcianise", "Marcianise", "n", "DETAIL", "Italian"),
    ("P", "*"): ("preventivi", "Preventivi", "p", "Customer Drawings", "English"),
    ("*", "4"): ("pID_ELETTRICI", "Pid_Elettrici", "m", "Customer Drawings", "Italian"),
    ("*", "5"): ("piping", "Piping", "m", "Customer Drawings", "Italian"),
}
DEFAULT_LOCATION = ("unknown", "Unknown", "m", "Customer Drawings", "English")

def map_location(m: re.Match, cfg: Config) -> dict:
    first = m.group(3)[0]
    l2 = m.group(2).upper()
    loc = (
        LOCATION_MAP.get((l2, first))
        or LOCATION_MAP.get((l2, "*"))
        or LOCATION_MAP.get(("*", first))
        or DEFAULT_LOCATION
    )
    folder, log_name, subloc, doctype, lang = loc
    arch_tif_loc = m.group(1).upper() + subloc
    dir_tif_loc = cfg.ARCHIVIO_DISEGNI / folder / arch_tif_loc
    return dict(folder=folder, log_name=log_name, subloc=subloc, doctype=doctype, lang=lang,
                arch_tif_loc=arch_tif_loc, dir_tif_loc=dir_tif_loc)

def size_from_letter(ch: str) -> str:
    return dict(A="A4",B="A3",C="A2",D="A1",E="A0").get(ch.upper(),"A4")

def uom_from_letter(ch: str) -> str:
    return dict(N="(Not applicable)",M="Metric",I="Inch",D="Dual").get(ch.upper(),"Metric")

# ---- ORIENTAMENTO TIFF: parser header-only -------------------------------

def _tiff_read_size_vfast(path: Path) -> Optional[Tuple[int,int]]:
    import struct
    try:
        with open(path, 'rb') as f:
            hdr = f.read(8)
            if len(hdr) < 8:
                return None
            endian = hdr[:2]
            if endian == b'II':
                u16 = lambda b: struct.unpack('<H', b)[0]
                u32 = lambda b: struct.unpack('<I', b)[0]
            elif endian == b'MM':
                u16 = lambda b: struct.unpack('>H', b)[0]
                u32 = lambda b: struct.unpack('>I', b)[0]
            else:
                return None
            if u16(hdr[2:4]) != 42:
                return None
            ifd_off = u32(hdr[4:8])
            f.seek(ifd_off)
            nbytes = f.read(2)
            if len(nbytes) < 2:
                return None
            n = u16(nbytes)
            TAG_W, TAG_H = 256, 257
            TYPE_SIZES = {1:1,2:1,3:2,4:4,5:8,7:1,9:4,10:8}
            w = h = None
            for _ in range(n):
                ent = f.read(12)
                if len(ent) < 12:
                    break
                tag = u16(ent[0:2]); typ = u16(ent[2:4]); cnt = u32(ent[4:8]); val = ent[8:12]
                unit = TYPE_SIZES.get(typ)
                if not unit:
                    continue
                datasz = unit * cnt
                if datasz <= 4:
                    if typ == 3: v = u16(val[0:2])
                    elif typ == 4: v = u32(val)
                    else: continue
                else:
                    off = u32(val); cur = f.tell()
                    f.seek(off); raw = f.read(unit); f.seek(cur)
                    if typ == 3: v = u16(raw)
                    elif typ == 4: v = u32(raw)
                    else: continue
                if tag == TAG_W: w = v
                elif tag == TAG_H: h = v
                if w is not None and h is not None:
                    return (w, h)
    except Exception:
        return None
    return None

def check_orientation_ok(tif_path: Path) -> bool:
    if tif_path.suffix.lower() == ".pdf":
        return True
    wh = _tiff_read_size_vfast(tif_path)
    if wh is None:
        return True
    w, h = wh
    return w > h

# ---- LOG WRAPPERS (GUI + buffer file) --------------------------------------

def _now_ddmonYYYY() -> str:
    return datetime.now().strftime("%d.%b.%Y")
def _now_HHMMSS() -> str:
    return datetime.now().strftime("%H:%M:%S")

def log_swarky(cfg: Config, file_name: str, loc: str, process: str,
               archive_dwg: str = "", dest: str = ""):
    line = f"{_now_ddmonYYYY()} # {_now_HHMMSS()} # {file_name}\t# {loc}\t# {process}\t# {archive_dwg}"
    _append_filelog_line(line)  # TXT batch
    logging.info("processed %s", file_name,
                 extra={"ui": ("processed", file_name, process, archive_dwg, dest)})

def log_error(cfg: Config, file_name: str, err: str, archive_dwg: str = ""):
    line = f"{_now_ddmonYYYY()} # {_now_HHMMSS()} # {file_name}\t# ERRORE\t# {err}\t# {archive_dwg}"
    _append_filelog_line(line)  # TXT batch
    logging.error("anomaly %s", file_name,
                  extra={"ui": ("anomaly", file_name, err)})

# ---- UI PHASES → eventi per la GUI ------------------------------------------

class _UIPhase:
    def __init__(self, label: str):
        self.label = label
        I = 0.0
        self.t0 = 0.0

    def __enter__(self):
        logging.info(self.label, extra={"ui": ("phase", self.label)})
        self.t0 = time.perf_counter()
        return self

    def __exit__(self, exc_type, exc, tb):
        elapsed_ms = int((time.perf_counter() - self.t0) * 1000)
        logging.info(f"{self.label} finita in {elapsed_ms} ms",
                     extra={"ui": ("phase_done", self.label, elapsed_ms)})
        return False

def ui_phase(label: str) -> _UIPhase:
    return _UIPhase(label)

# ---- EDI WRITER --------------------------------------------------------------

def _edi_body(
    *,
    document_no: str,
    rev: str,
    sheet: str,
    description: str,
    actual_size: str,
    uom: str,
    doctype: str,
    lang: str,
    file_name: str,
    file_type: str,
    now: Optional[str] = None
) -> List[str]:
    now = now or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    header = [
        "[Database]",
        "ServerName=ORMDB33",
        "ProjectName=FPD Engineering",
        "[DatabaseFields]",
        f"DocumentNo={document_no}",
        f"DocumentRev={rev}",
        f"SheetNumber={sheet}",
        f"Description={description}",
        f"ActualSize={actual_size}",
        "PumpModel=(UNKNOWN)",
        "OEM=Flowserve",
        "PumpSize=",
        "OrderNumber=",
        "SerialNumber=",
        f"Document_Type={doctype}",
        "DrawingClass=COMMERCIAL",
        "DesignCenter=Desio, Italy",
        "OEMSite=Desio, Italy",
        "OEMDrawingNumber=",
        f"UOM={uom}",
        f"DWGLanguage={lang}",
        "CurrentRevision=Y",
        "EnteredBy=10150286",
        "Notes=",
        "NonEnglishDesc=",
        "SupersededBy=",
        "NumberOfStages=",
        "[DrawingInfo]",
        f"DocumentNo={document_no}",
        f"SheetNumber={sheet}",
        ("Document_Type=Detail" if doctype == "DETAIL" else "Document_Type=Customer Drawings"),
        f"DocumentRev={rev}",
        f"FileName={file_name}",
        f"FileType={file_type}",
        f"Currentdate={now}",
    ]
    return header

def write_edi(
    cfg: Config,
    file_name: str,
    out_dir: Path,
    *,
    m: Optional[re.Match] = None,
    iss_match: Optional[re.Match] = None,
    loc: Optional[dict] = None
) -> None:
    edi = out_dir / (Path(file_name).stem + ".DESEDI")
    if edi.exists():
        return
    if iss_match is not None:
        g1 = iss_match.group(1); g2 = iss_match.group(2); g3 = iss_match.group(3)
        rev = iss_match.group(4); sheet = iss_match.group(5)
        docno = f"G{g1}{g2}{g3}"
        body = _edi_body(
            document_no=docno, rev=rev, sheet=sheet,
            description=" Impeller Specification Sheet",
            actual_size="A4", uom="Metric", doctype="DETAIL", lang="English",
            file_name=file_name, file_type="Pdf",
        )
        write_lines(out_dir / (Path(file_name).stem + ".DESEDI"), body)
        return
    if m is None or loc is None:
        raise ValueError("write_edi: per STANDARD/FIV servono 'm' (BASE_NAME) e 'loc' (map_location)")
    document_no = f"D{m.group(1)}{m.group(2)}{m.group(3)}"
    rev = m.group(4); sheet = m.group(5)
    file_type = "Pdf" if Path(file_name).suffix.lower() == ".pdf" else "Tiff"
    body = _edi_body(
        document_no=document_no, rev=rev, sheet=sheet, description="",
        actual_size=size_from_letter(m.group(1)), uom=uom_from_letter(m.group(6)),
        doctype=loc["doctype"], lang=loc["lang"],
        file_name=file_name, file_type=file_type,
    )
    write_lines(edi, body)

# ---- STORICO: routing ---------------------------------------------------------------

def _storico_dest_dir_for_name(cfg: Config, nm: str) -> Path:
    mm = BASE_NAME.fullmatch(nm)
    if not mm:
        return cfg.ARCHIVIO_STORICO / "unknown"
    return cfg.ARCHIVIO_STORICO / f"D{mm.group(1).upper()}"

# ---- PIPELINE PRINCIPALE -------------------------------------------------------------

def _iter_candidates(dirp: Path, accept_pdf: bool):
    exts = {".tif"}
    if accept_pdf:
        exts.add(".pdf")
    with os.scandir(dirp) as it:
        for de in it:
            if de.is_file():
                suf = os.path.splitext(de.name)[1].lower()
                if suf in exts:
                    yield Path(de.path)

def _process_candidate(p: Path, cfg: Config) -> bool:
    try:
        # --- normalizzazione estensione on-the-fly ---
        suf = p.suffix
        if suf == ".TIF":
            q = p.with_suffix(".tif")
            try:
                p.rename(q); p = q
            except Exception:
                pass
        elif suf.lower() == ".tiff":
            q = p.with_suffix(".tif")
            try:
                p.rename(q); p = q
            except Exception:
                pass

        name = p.name

        # ---- ORIENTAMENTO: subito in testa ----
        with ui_phase(f"{name} • orientamento"):
            if not check_orientation_ok(p):
                log_error(cfg, name, "Immagine Girata")
                move_to(p, cfg.ERROR_DIR)
                return True

        # ---- Regex + validazioni ----
        with ui_phase(f"{name} • regex+validate"):
            m = BASE_NAME.fullmatch(name)
            if not m:
                log_error(cfg, name, "Nome File Errato"); move_to(p, cfg.ERROR_DIR); return True
            if m.group(1).upper() not in "ABCDE":
                log_error(cfg, name, "Formato Errato"); move_to(p, cfg.ERROR_DIR); return True
            if m.group(2).upper() not in "MKFTESNP":
                log_error(cfg, name, "Location Errata"); move_to(p, cfg.ERROR_DIR); return True
            if m.group(6).upper() not in "MIDN":
                log_error(cfg, name, "Metrica Errata"); move_to(p, cfg.ERROR_DIR); return True

        new_rev    = m.group(4)
        new_sheet  = m.group(5)
        new_metric = m.group(6).upper()
        MI = {"M","I"}; DN = {"D","N"}
        new_group = "MI" if new_metric in MI else "DN"
        new_rev_i = int(new_rev)

        # ---- Mappatura destinazione archivio ----
        with ui_phase(f"{name} • map_location"):
            loc = map_location(m, cfg)
            dir_tif_loc = loc["dir_tif_loc"]
            tiflog      = loc["log_name"]

        # ---- Elenco file con stesso DOCNO ----
        with ui_phase(f"{name} • list_same_doc_prefisso"):
            same_doc = _list_same_doc_prefisso(dir_tif_loc, m)

        with ui_phase(f"{name} • derive_same_sheet"):
            same_sheet = [(r, nm, met, sh) for (r, nm, met, sh) in same_doc if sh == new_sheet]

        # ---- Pari revisione (verifica via lista) ----
        with ui_phase(f"{name} • check_same_filename"):
            if any((nm == name and r == new_rev) for (r, nm, met, sh) in same_sheet):
                log_error(cfg, name, "Pari Revisione")
                move_to(p, cfg.PARI_REV_DIR)
                return True

        # ---- Partizionamento e max rev ----
        same_sheet_mi = [(int(r), nm, met) for (r, nm, met, sh) in same_sheet if met in MI]
        same_sheet_dn = [(int(r), nm, met) for (r, nm, met, sh) in same_sheet if met in DN]
        same_sheet_same_metric = [(int(r), nm, met) for (r, nm, met, sh) in same_sheet if met == new_metric]

        def _max_rev(entries: List[Tuple[int,str,str]]) -> Optional[int]:
            return max((rv for (rv, _, _) in entries), default=None)

        max_mi = _max_rev(same_sheet_mi)
        max_dn = _max_rev(same_sheet_dn)
        own_max = _max_rev(same_sheet_same_metric)

        # ---- Revisioni precedenti rispetto all'altro gruppo ----
        other_entries = same_sheet_dn if new_group == "MI" else same_sheet_mi
        other_max = max_dn if new_group == "MI" else max_mi
        if other_max is not None and new_rev_i < other_max:
            ref = next((nm for (rv, nm, _met) in other_entries if rv == other_max), "")
            log_error(cfg, name, "Revisione Precendente", ref)
            move_to(p, cfg.ERROR_DIR)
            return True

        # ---- Revisioni precedenti rispetto stessa metrica ----
        if own_max is not None and new_rev_i < own_max:
            ref = next((nm for (rv, nm, _met) in same_sheet_same_metric if rv == own_max), "")
            log_error(cfg, name, "Revisione Precendente", ref)
            move_to(p, cfg.ERROR_DIR)
            return True

        # ---- Conflitti pari rev tra gruppi/metrica ----
        same_rev_mi = [(rv, nm, met) for (rv, nm, met) in same_sheet_mi if rv == new_rev_i]
        same_rev_dn = [(rv, nm, met) for (rv, nm, met) in same_sheet_dn if rv == new_rev_i]

        if new_group == "MI":
            if same_rev_dn:
                ref = same_rev_dn[0][1]
                log_error(cfg, name, "Conflitto Metrica (DN a pari revisione)", ref)
                move_to(p, cfg.ERROR_DIR)
                return True
            other_mi = next((nm for (_rv, nm, met) in same_rev_mi if met != new_metric), None)
            if other_mi:
                log_swarky(cfg, name, tiflog, "Metrica Diversa", other_mi)
        else:
            if same_rev_mi:
                ref = same_rev_mi[0][1]
                log_error(cfg, name, "Conflitto Metrica (MI a pari revisione)", ref)
                move_to(p, cfg.ERROR_DIR)
                return True
            other_dn = next((nm for (_rv, nm, met) in same_rev_dn if met != new_metric), None)
            if other_dn:
                log_error(cfg, name, "Conflitto Metrica (D/N a pari revisione)", other_dn)
                move_to(p, cfg.ERROR_DIR)
                return True

        # ---- ACCETTAZIONE del NUOVO ----
        with ui_phase(f"{name} • move_to_archivio"):
            move_to(p, dir_tif_loc)
            new_path = dir_tif_loc / name

        # ---- STORICIZZAZIONI (dopo l'accettazione) ----
        to_storico_same: list[tuple[Path, Path, str]] = []
        to_storico_other: list[tuple[Path, Path, str]] = []
        if own_max is None or new_rev_i > own_max:
            for rv, nm, _met in same_sheet_same_metric:
                if rv < new_rev_i:
                    to_storico_same.append((dir_tif_loc / nm, _storico_dest_dir_for_name(cfg, nm), nm))
        if other_max is not None and new_rev_i > other_max:
            for rv, nm, _met in other_entries:
                if rv < new_rev_i:
                    to_storico_other.append((dir_tif_loc / nm, _storico_dest_dir_for_name(cfg, nm), nm))

        if to_storico_same:
            with ui_phase(f"{name} • move_old_revs_same_metric"):
                for old_path, dest_dir, nm in to_storico_same:
                    try:
                        copied, rc = move_to_storico_safe(old_path, dest_dir)
                        if rc >= 8:
                            logging.exception("Storico (same metric) errore: %s → %s", old_path, dest_dir)
                        elif copied:
                            log_swarky(cfg, name, tiflog, "Rev superata", nm, "Storico")
                        else:
                            log_error(cfg, nm, "Presente in Storico")
                            try:
                                move_to(old_path, cfg.ERROR_DIR)
                            except FileNotFoundError:
                                pass
                    except Exception as e:
                        logging.exception("Storico (same metric): %s → %s: %s", old_path, dest_dir, e)

        if to_storico_other:
            with ui_phase(f"{name} • move_old_revs_other_group"):
                for old_path, dest_dir, nm in to_storico_other:
                    try:
                        copied, rc = move_to_storico_safe(old_path, dest_dir)
                        if rc >= 8:
                            logging.exception("Storico (other grp) errore: %s → %s", old_path, dest_dir)
                        elif copied:
                            log_swarky(cfg, name, tiflog, "Rev superata", nm, "Storico")
                        else:
                            log_error(cfg, nm, "Presente in Storico")
                            try:
                                move_to(old_path, cfg.ERROR_DIR)
                            except FileNotFoundError:
                                pass
                    except Exception as e:
                        logging.exception("Storico (other grp): %s → %s: %s", old_path, dest_dir, e)

        # ---- PLM + EDI ----
        with ui_phase(f"{name} • link/copy_to_PLM"):
            try:
                _fast_copy_or_link(new_path, cfg.PLM_DIR / name, overwrite=True)
            except Exception as e:
                logging.exception("PLM copy/link fallita per %s: %s", new_path, e)

        with ui_phase(f"{name} • write_EDI"):
            try:
                write_edi(cfg, name, cfg.PLM_DIR, m=m, loc=loc)
            except Exception as e:
                logging.exception("Impossibile creare DESEDI per %s: %s", name, e)

        log_swarky(cfg, name, tiflog, "Archiviato", "", dest=tiflog)
        return True

    except Exception:
        logging.exception("Errore inatteso per %s", p)
        return False

# ---- ISS / FIV ----------------------------------------------------------------------

def iss_loading(cfg: Config) -> bool:
    did = False
    try:
        candidates = [p for p in cfg.DIR_ISS.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    except Exception as e:
        logging.exception("ISS: impossibile leggere la cartella %s: %s", cfg.DIR_ISS, e)
        return False

    for p in candidates:
        m = ISS_BASENAME.fullmatch(p.name)
        if not m:
            log_error(cfg, p.name, "Nome ISS Errato")
            continue
        try:
            with ui_phase(f"{p.name} • ISS move_to_PLM"):
                move_to(p, cfg.PLM_DIR)
            with ui_phase(f"{p.name} • ISS write_EDI"):
                write_edi(cfg, file_name=p.name, out_dir=cfg.PLM_DIR, iss_match=m)
            log_swarky(cfg, p.name, "ISS", "ISS", "", "")
            did = True
        except Exception as e:
            logging.exception("Impossibile processare ISS %s: %s", p.name, e)
        try:
            now = datetime.now()
            stem = p.stem
            log = cfg.DIR_ISS / "SwarkyISS.log"
            write_lines(log, [f"{now.strftime('%d.%b.%Y')} # {now.strftime('%H:%M:%S')} # {stem}"])
        except Exception:
            logging.exception("ISS: impossibile aggiornare SwarkyISS.log")

    return did

def fiv_loading(cfg: Config) -> bool:
    did = False
    try:
        files = [p for p in cfg.DIR_FIV_LOADING.iterdir() if p.is_file()]
    except Exception as e:
        logging.exception("FIV: lettura cartella fallita: %s", e)
        return False

    for p in files:
        ext = p.suffix.lower()
        if ext not in (".tif", ".tiff") and not (cfg.ACCEPT_PDF and ext == ".pdf"):
            continue
        m = BASE_NAME.fullmatch(p.name)
        if not m:
            log_error(cfg, p.name, "Nome FIV Errato")
            continue
        try:
            with ui_phase(f"{p.name} • FIV map_location"):
                loc = map_location(m, cfg)
            with ui_phase(f"{p.name} • FIV write_EDI"):
                write_edi(cfg, m=m, file_name=p.name, loc=loc, out_dir=cfg.PLM_DIR)
            with ui_phase(f"{p.name} • FIV move_to_PLM"):
                move_to(p, cfg.PLM_DIR)
            log_swarky(cfg, p.name, "FIV", "FIV loading", "", "")
            did = True
        except Exception as e:
            logging.exception("Impossibile processare FIV %s: %s", p.name, e)

    return did

# ---- STATS --------------------------------------------------------------------------

_LAST_STATS_TS: float = 0.0

def _count_files_quick(d: Path, exts: tuple[str, ...]) -> int:
    try:
        with os.scandir(d) as it:
            return sum(1 for de in it if de.is_file() and os.path.splitext(de.name)[1].lower() in exts)
    except (OSError, FileNotFoundError):
        return 0

def _stats_interval_sec() -> int:
    val = os.environ.get("SWARKY_STATS_EVERY", "300")
    try:
        n = int(val)
    except Exception:
        n = 300
    return n if n >= 10 else 10

def _should_emit_stats() -> bool:
    global _LAST_STATS_TS
    now = time.monotonic()
    if now - _LAST_STATS_TS >= _stats_interval_sec():
        _LAST_STATS_TS = now
        return True
    return False

def count_tif_files(cfg: Config) -> dict:
    return {
        "Same Rev Dwg": _count_files_quick(cfg.PARI_REV_DIR, (".tif", ".pdf")),
        "Check Dwg": _count_files_quick(cfg.ERROR_DIR, (".tif", ".pdf")),
        "Heng Dwg": _count_files_quick(cfg.DIR_HENGELO, (".tif", ".pdf")),
        "Tab Dwg": _count_files_quick(cfg.DIR_TABELLARI, (".tif", ".pdf")),
        "Plm error Dwg": _count_files_quick(cfg.DIR_PLM_ERROR, (".tif", ".pdf")),
    }

# ---- LOOP ----------------------------------------------------------------------------

def run_once(cfg: Config) -> bool:
    start_all = time.time()

    with ui_phase("Scan candidati (hplotter)"):
        candidates: List[Path] = list(_iter_candidates(cfg.DIR_HPLOTTER, cfg.ACCEPT_PDF))

    did_something = False
    for p in candidates:
        try:
            did_something |= _process_candidate(p, cfg)
        except Exception:
            logging.exception("Errore nel processing")

    did_arch = did_something
    did_iss  = iss_loading(cfg)
    did_fiv  = fiv_loading(cfg)

    elapsed_all = time.time() - start_all
    minutes = int(elapsed_all // 60)
    seconds = int(elapsed_all % 60)
    _append_filelog_line(f"ProcessTime # {minutes:02d}:{seconds:02d}")

    _flush_file_log(cfg)

    if logging.getLogger().isEnabledFor(logging.DEBUG) and _should_emit_stats():
        logging.debug("Counts: %s", count_tif_files(cfg))

    return did_arch or did_iss or did_fiv

def watch_loop(cfg: Config, interval: int):
    logging.info("Watch ogni %ds...", interval)
    while True:
        run_once(cfg); time.sleep(interval)

# ---- CLI -----------------------------------------------------------------------------

def parse_args(argv: List[str]):
    import argparse
    ap = argparse.ArgumentParser(description="Swarky - batch archiviazione/EDI")
    ap.add_argument("--watch", type=int, default=0, help="Loop di polling in secondi, 0=una sola passata")
    return ap.parse_args(argv)

def load_config(path: Path) -> Config:
    if not path.exists():
        raise FileNotFoundError(f"Config non trovato: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    return Config.from_json(data)

def main(argv: List[str]):
    args = parse_args(argv)
    cfg = load_config(Path("config.json"))
    setup_logging(cfg)

    if args.watch > 0:
        watch_loop(cfg, args.watch)
    else:
        run_once(cfg)

if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except KeyboardInterrupt:
        print("Interrotto")
