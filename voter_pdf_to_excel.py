#!/usr/bin/env python3
"""
Gujarati voter-roll PDFs -> one Excel file, using the Gemini API.

Setup (once):
    pip install google-genai pymupdf openpyxl
    set GEMINI_API_KEY=your_key        (Windows)   /   export GEMINI_API_KEY=your_key   (Mac/Linux)
    Get a free key at https://aistudio.google.com/apikey

Run (one Excel per PDF, same name as the PDF, saved in the output folder):
    python voter_pdf_to_excel.py pdfs excel_output
Options:
    --model gemini-3.5-flash-lite  model name (python voter_pdf_to_excel.py --list-models shows yours)
    --fallback-model gemini-3.5-flash   stronger model, used only to re-read pages with bad letters
    --rpm 10                       max requests per minute (stay under your free-tier limit)
    --ward-map ward_map.csv        optional CSV with two columns: bhag_no,ward_no
    --skip-existing                skip PDFs whose Excel already exists
    --build-only                   do not call Gemini, rebuild Excel files from cached results

It is RESUMABLE: every page result is saved in <output>_cache/. If the daily free quota
runs out the script stops cleanly; run the same command again the next day and it
continues where it stopped.
"""
import argparse, csv, glob, hashlib, json, os, re, sys, time

SKIP_FIRST_PAGES = 2      # page 1 = summary, page 2 = maps/photos (no voters)

PROMPT = """This is one page of a Gujarati electoral roll (મતદારયાદી). The page has a header
line and a grid of voter cards (usually 3 columns x 10 rows).

Return JSON only. Copy every value EXACTLY as printed (keep Gujarati script, do not translate,
do not transliterate, do not correct spellings). Read cards left-to-right, top-to-bottom.

SUPPLEMENTARY pages: some pages (header like "ઘટક - 1 : વધારા યાદી - 1 ...") list added voters. They have NO "વિભાગનો નંબર અને નામ" line (return vibhag_line "")
but every card has TWO small boxes at the top-left: the first is the serial number, the second is that voter's vibhag number -> put it in "card_vibhag_no". On normal pages return card_vibhag_no "".
These pages ARE voter pages - read all their cards exactly like normal pages.

Header (top of the page) has TWO lines. Do NOT confuse them:
- Line 1 starts with "વિધાનસભા મત વિભાગનો નંબર અને નામ" (assembly constituency, e.g. 69-રાજકોટ પશ્ચિમ, and "ભાગ નંબર"). IGNORE this line.
- Line 2 starts with only "વિભાગનો નંબર અને નામ :" (the neighbourhood / society section). Put the complete text after its colon
  into "vibhag_line" exactly as printed, e.g. "9-સ્વપ્રલોક રેસીડેન્સી, રામાપીર ચોકડી પાસે, રૈયા-34, રાજકોટ". If this line is missing, return "".
If the page has no voter cards (maps, summary, blank), return an empty "voters" list.

For each voter card:
- serial: the number in the (first) small box at the top-left of the card.
- card_vibhag_no: only on supplementary pages, see above; otherwise "".
- epic: the ID number at the top-right of the card (like WQT3033412 - 3 capital letters + 7 digits; it can also be a different format, copy exactly). Be careful: W/V, 0/O, 1/I, 5/S, 8/B.
- name: text after "મતદારનું નામ :" (may wrap over 2 lines - join them with a space).
- relation_type: which label the second line has: "પિતાનું નામ" -> "પિતા", "પતિનું નામ" -> "પતિ", "માતાનું નામ" -> "માતા", anything else (e.g. "અન્યનું નામ") -> "અન્ય".
- relative_name: text after that relation label (join wrapped lines with a space).
- house_no: text after "ઘર નંબર :" exactly as printed (words, letters and slashes kept; write every digit as an ASCII digit 0-9, e.g. ૧૯૨ -> 192; empty string if blank).
- age: integer after "ઉંમર :" (convert Gujarati digits to normal digits).
- gender: "પુરુષ", "સ્ત્રી" or "અન્ય" as printed after "જાતિ :".
If a card is deleted / struck through or the text is unreadable, still include it and put what you can read.
STRICT RULES:
- Use ONLY Gujarati letters (plus ASCII digits/letters/punctuation). NEVER output Devanagari, Bengali, Kannada, Telugu, Malayalam, Thai or any other script;
  Gujarati conjuncts like "જ્ઞ", "ત્ર", "ક્ષ", "દ્ય" must be written with Gujarati characters (example: યાજ્ઞિક, પંડ્યા, રાવલ, ચૌહાણ, કલ્યાણભાઈ).
- If the age/gender line is NOT visible on a card (text overflowed), return age 0 and gender "" - never guess.
- Never invent a card: only include cards that are really printed on the page.
"""

SCHEMA = {
    "type": "object",
    "properties": {
        "vibhag_line": {"type": "string"},
        "voters": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "serial": {"type": "integer"},
                    "card_vibhag_no": {"type": "string"},
                    "epic": {"type": "string"},
                    "name": {"type": "string"},
                    "relation_type": {"type": "string"},
                    "relative_name": {"type": "string"},
                    "house_no": {"type": "string"},
                    "age": {"type": "integer"},
                    "gender": {"type": "string"},
                },
                "required": ["serial", "card_vibhag_no", "epic", "name", "relation_type", "relative_name", "house_no", "age", "gender"],
            },
        },
    },
    "required": ["vibhag_line", "voters"],
}

COLUMNS = [
    "ક્રમ નંબર (Serial No)", "ઓળખપત્ર નંબર (EPIC No)", "મતદારનું નામ (Full Name)",
    "સંબંધનો પ્રકાર (Relation Type)", "સંબંધીનું નામ (Relative Name)", "પિતાનું નામ (Father Name)",
    "પતિનું નામ (Husband Name)", "માતાનું નામ (Mother Name)", "અન્ય સંબંધી (Other Relation)",
    "ઘર / મકાન નંબર (House No)", "ઉંમર (Age)", "જાતિ (Gender)", "વિભાગ નંબર (Vibhag No)",
    "વિભાગનું નામ (Vibhag Name)", "વોર્ડ નંબર (Ward No)", "ભાગ નંબર (Bhag No)",
    "Source File", "Page", "Check",
]


def file_meta(path):
    """Bhag no from the file name (...-GUJ-11-WI.pdf)."""
    m = re.search(r"-(\d+)-WI\.pdf$", os.path.basename(path), re.I)
    return m.group(1) if m else ""


def render_pages(pdf_path, dpi=150):
    import fitz  # pymupdf
    doc = fitz.open(pdf_path)
    for i in range(doc.page_count):
        yield i + 1, doc[i].get_pixmap(dpi=dpi).tobytes("jpeg", jpg_quality=85)


def render_one(doc, n, dpi=150):
    return doc[n - 1].get_pixmap(dpi=dpi).tobytes("jpeg", jpg_quality=85)


def cache_path(cache_dir, pdf_path, page_no):
    h = hashlib.md5(os.path.abspath(pdf_path).encode()).hexdigest()[:8]
    return os.path.join(cache_dir, f"{os.path.basename(pdf_path)[:40]}_{h}_p{page_no:03d}.json")


class QuotaExhausted(Exception):
    pass


def parse_json_text(text):
    """Pull the JSON object out of a model reply (handles ```json fences and extra words)."""
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    i, j = text.find("{"), text.rfind("}")
    d = json.loads(text[i:j + 1])
    d.setdefault("vibhag_line", "")
    d.setdefault("voters", [])
    for v in d["voters"]:
        for k, default in (("card_vibhag_no", ""), ("epic", ""), ("name", ""), ("relation_type", ""), ("relative_name", ""),
                           ("house_no", ""), ("gender", ""), ("age", 0), ("serial", 0)):
            v.setdefault(k, default)
        for k in ("age", "serial"):
            try: v[k] = int(normalize_digits(str(v[k])))
            except ValueError: v[k] = 0
    return d


def call_ollama(model, images, temperature=0.0):
    """Local model through Ollama (http://localhost:11434). model like 'ollama:gemma3:12b'."""
    import base64, urllib.request
    body = {"model": model.split(":", 1)[1], "stream": False, "format": SCHEMA,
            "options": {"temperature": temperature, "num_ctx": 8192},
            "messages": [{"role": "user", "content": PROMPT,
                          "images": [base64.b64encode(b).decode() for b in images]}]}
    req = urllib.request.Request(os.environ.get("OLLAMA_URL", "http://localhost:11434") + "/api/chat",
                                 data=json.dumps(body).encode(), headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=3600) as r:
            text = json.loads(r.read().decode())["message"]["content"]
    except OSError as e:
        sys.exit(f"Cannot reach Ollama at localhost:11434 ({e}). Install/start Ollama and run: ollama pull {model.split(':', 1)[1]}")
    return parse_json_text(text)


def call_gemini(client, types, model, images, temperature=0.0, max_retries=6):
    if model.startswith("ollama:"):
        return call_ollama(model, images, temperature)
    parts = [types.Part.from_bytes(data=b, mime_type="image/jpeg") for b in images]
    gemma = model.lower().startswith("gemma")
    if gemma:   # Gemma models: no JSON-schema mode, so ask for JSON in the prompt and parse the text
        cfg = types.GenerateContentConfig(temperature=temperature)
        prompt = PROMPT + "\nReturn ONLY one JSON object (no markdown, no explanation) of this shape:\n" + json.dumps(SCHEMA["properties"], ensure_ascii=False)
    else:
        cfg = types.GenerateContentConfig(
            response_mime_type="application/json", response_schema=SCHEMA, temperature=temperature)
        prompt = PROMPT
    delay = 15
    for attempt in range(max_retries):
        try:
            r = client.models.generate_content(model=model, contents=parts + [prompt], config=cfg)
            return parse_json_text(r.text) if gemma else json.loads(r.text)
        except Exception as e:  # noqa
            msg = str(e)
            low = msg.lower()
            if "403" in msg or "permission_denied" in low or "api key not valid" in low or "api_key_invalid" in low:
                sys.exit("\nGemini refused this key/account (permission problem) - retrying will not help:\n" + msg +
                         "\n\nFix: use a key from a project/account where the Gemini API free tier is enabled and available in your region, or enable billing.")
            if "429" in msg or "resource_exhausted" in low or "quota" in low:
                if "per day" in low or "perday" in low.replace(" ", "") or "daily" in low:
                    raise QuotaExhausted(msg)
                print(f"   rate limited, waiting {delay}s ...", flush=True)
                time.sleep(delay); delay = min(delay * 2, 120); continue
            if "503" in msg or "500" in msg or "unavailable" in low or "deadline" in low or isinstance(e, json.JSONDecodeError):
                print(f"   temporary error ({msg[:80]}), retrying in {delay}s ...", flush=True)
                time.sleep(delay); delay = min(delay * 2, 120); continue
            raise
    raise RuntimeError("too many retries")


CACHE_VERSION = 2
GUJ_DIGITS = str.maketrans("૦૧૨૩૪૫૬૭૮૯", "0123456789")


def bad_script_chars(text):
    """Characters from Indic/other scripts that are NOT Gujarati (Devanagari, Kannada, Telugu, Gurmukhi, Thai ...)."""
    bad = set()
    for ch in text or "":
        o = ord(ch)
        if o in (0x0964, 0x0965):          # danda signs are harmless
            continue
        if (0x0900 <= o <= 0x0DFF and not 0x0A80 <= o <= 0x0AFF) or 0x0E00 <= o <= 0x0EFF:
            bad.add(ch)
    return bad


def normalize_digits(text):
    """Any script's digits (૧૯૨, १९२ ...) -> ASCII digits with the same numeric value."""
    import unicodedata
    return "".join(str(unicodedata.digit(ch)) if unicodedata.category(ch) == "Nd" else ch for ch in (text or ""))


def voter_bad(v):
    return (bad_script_chars(v.get("name", "")) or bad_script_chars(v.get("relative_name", ""))
            or bad_script_chars(normalize_digits(v.get("house_no", ""))))   # foreign DIGITS are fixed later, not an error


def page_bad(res):
    return bool(bad_script_chars(res.get("vibhag_line", ""))) or any(voter_bad(v) for v in res.get("voters", []))


def page_has_bad_script(res):
    return page_bad(res)


def cache_ok(path):
    if not os.path.exists(path): return False
    try:
        return json.load(open(path, encoding="utf-8")).get("cache_version") == CACHE_VERSION
    except Exception:
        return False


def merge_better(res, alt):
    """Take from `alt` only what is better than `res`: clean versions of bad voters, missing voters, clean vibhag line."""
    by = {v["serial"]: v for v in alt.get("voters", [])}
    out = []
    for v in res.get("voters", []):
        c = by.pop(v["serial"], None)
        out.append(c if (voter_bad(v) and c is not None and not voter_bad(c)) else v)
    out += [v for v in by.values() if not voter_bad(v)]          # cards the first read missed
    out.sort(key=lambda v: v["serial"])
    line = res.get("vibhag_line", "")
    if (not line or bad_script_chars(line)) and alt.get("vibhag_line") and not bad_script_chars(alt["vibhag_line"]):
        line = alt["vibhag_line"]
    return {"vibhag_line": line, "voters": out}


def read_page(client, types, args, img, is_last=False):
    """Read one page; re-read (different sampling, then fallback model) if empty or if non-Gujarati letters appear."""
    res = call_gemini(client, types, args.model, [img])
    if not res.get("voters") and not is_last:
        print("   no voters found, re-reading once ...", flush=True)
        res = call_gemini(client, types, args.model, [img], temperature=0.4)
    for mdl, temp in [(args.model, 0.4), (args.model, 0.8), (args.fallback_model or args.model, 0.3)][:args.rereads]:
        if not page_bad(res):
            break
        print("   non-Gujarati letters found, re-reading page ...", flush=True)
        res = merge_better(res, call_gemini(client, types, mdl, [img], temperature=temp))
    return res


def get_client():
    from google import genai
    from google.genai import types
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key:
        return None, types       # fine for local (ollama:) models; Gemini calls will fail without a key
    return genai.Client(api_key=key), types


_last_call = [0.0]


def extract_pdf(client, types, args, pdf):
    """Read all pages of one PDF into the cache. Returns True when every page is cached.
    Raises QuotaExhausted when the daily free quota is over."""
    import fitz
    doc = fitz.open(pdf)
    total = doc.page_count
    todo = [n for n in range(SKIP_FIRST_PAGES + 1, total + 1) if not cache_ok(cache_path(args.cache, pdf, n))]
    print(f"{os.path.basename(pdf)}: {total} pages, {len(todo)} to read", flush=True)
    min_gap = 60.0 / args.rpm
    for i, n in enumerate(todo, 1):
        img = render_one(doc, n, args.dpi)
        wait = min_gap - (time.time() - _last_call[0])
        if wait > 0: time.sleep(wait)
        _last_call[0] = time.time()
        try:
            res = read_page(client, types, args, img, is_last=(n == total))
        except QuotaExhausted:
            raise
        except Exception as e:  # noqa
            print(f"   FAILED page {n}: {str(e)[:120]} (will retry on next run)", flush=True)
            continue
        res["cache_version"] = CACHE_VERSION
        json.dump(res, open(cache_path(args.cache, pdf, n), "w", encoding="utf-8"), ensure_ascii=False)
        print(f"   page {n}/{total} ok ({len(res.get('voters', []))} voters)  [{i}/{len(todo)}]", flush=True)
    return all(cache_ok(cache_path(args.cache, pdf, n)) for n in range(SKIP_FIRST_PAGES + 1, total + 1))


def parse_vibhag(line):
    """'9-સ્વપ્રલોક રેસીડેન્સી, ...' -> ('9', 'સ્વપ્રલોક રેસીડેન્સી, ...'); Gujarati digits converted."""
    line = normalize_digits((line or "").strip())
    m = re.match(r"^\s*([0-9૦-૯]+)\s*[-–—]\s*(.*)$", line, re.S)
    if not m: return "", line
    return m.group(1).translate(GUJ_DIGITS), m.group(2).strip()


def load_ward_map(path):
    if not path: return {}
    with open(path, encoding="utf-8-sig") as f:
        return {r[0].strip(): r[1].strip() for r in csv.reader(f) if len(r) >= 2 and r[0].strip().isdigit()}


def build_excel(args, pdfs, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    ward = load_ward_map(args.ward_map)
    wb = Workbook(); ws = wb.active; ws.title = "Voters"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    epic_re = re.compile(r"^([A-Z]{3}\d{7}|GJ/\d{2}/\d{3}/\d{6})$")
    total = flagged = 0
    for pdf in pdfs:
        bhag = file_meta(pdf)
        pages = []
        for n in range(SKIP_FIRST_PAGES + 1, 400):
            p = cache_path(args.cache, pdf, n)
            if not os.path.exists(p):
                if n > SKIP_FIRST_PAGES + 1 and not any(os.path.exists(cache_path(args.cache, pdf, n + j)) for j in range(1, 3)):
                    break
                continue
            pages.append((n, json.load(open(p, encoding="utf-8"))))
        # vibhag number -> name, learned from the normal pages (supplementary cards only carry the number)
        names = {}
        for n, d in pages:
            no, nm = parse_vibhag(d.get("vibhag_line", ""))
            if no and nm and not bad_script_chars(nm): names.setdefault(no, nm)
        rows = []
        for n, d in pages:
            pno, pname = parse_vibhag(d.get("vibhag_line", ""))
            for v in d.get("voters", []):
                cno = normalize_digits(v.get("card_vibhag_no", "")).strip()
                rows.append((n, cno, pno, pname, v))
        empty_pages = [n for n, d in pages if not d.get("voters")]
        last_no = last_name = ""
        prev_serial = 0
        seen_epic = set(); serials = set()
        for n, cno, vno, vname, v in rows:
            if cno:                                   # supplementary card: number from the card itself
                vno, vname = cno, names.get(cno, "")
            elif vno:
                last_no, last_name = vno, vname
            else:
                vno, vname = last_no, last_name
            issues = []
            if not vno: issues.append("vibhag missing")
            elif not vname or bad_script_chars(vname): issues.append("vibhag name")
            rt, rel = v.get("relation_type", ""), v.get("relative_name", "")
            house = normalize_digits(v.get("house_no", ""))
            if v["serial"] != prev_serial + 1: issues.append("serial jump")
            prev_serial = v["serial"]; serials.add(v["serial"])
            if not epic_re.match(v.get("epic", "")): issues.append("EPIC format")
            if v["epic"] in seen_epic: issues.append("duplicate EPIC")
            seen_epic.add(v["epic"])
            if not v.get("name"): issues.append("name empty")
            if not rel: issues.append("relative empty")
            if rt not in ("પિતા", "પતિ", "માતા", "અન્ય"): issues.append("relation type")
            if not (18 <= v.get("age", 0) <= 120): issues.append("age/gender not visible or wrong")
            if voter_bad(v): issues.append("non-Gujarati letters")
            if v.get("gender") not in ("પુરુષ", "સ્ત્રી", "અન્ય"): issues.append("gender")
            ws.append([
                v["serial"], v["epic"], v["name"], rt, rel,
                rel if rt == "પિતા" else "", rel if rt == "પતિ" else "",
                rel if rt == "માતા" else "", rel if rt == "અન્ય" else "",
                house, v["age"], v["gender"], vno, vname,
                ward.get(bhag, ""), bhag, os.path.basename(pdf), n, ", ".join(issues)])
            total += 1; flagged += bool(issues)
        if serials:
            missing = sorted(set(range(1, max(serials) + 1)) - serials)
            if missing:
                print(f"!! {os.path.basename(pdf)}: MISSING serials: {missing[:40]}{' ...' if len(missing) > 40 else ''}")
        if len(empty_pages) > 1 or (empty_pages and empty_pages[0] != pages[-1][0]):
            print(f"!! {os.path.basename(pdf)}: pages with no voters read: {empty_pages} (last page is normally the summary)")
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    for col, w in zip("ABCDEFGHIJKLMNOPQRS", [9, 14, 28, 12, 28, 24, 24, 24, 24, 16, 7, 9, 10, 40, 9, 9, 30, 6, 22]):
        ws.column_dimensions[col].width = w
    wb.save(out_path)
    print(f"   Excel saved: {out_path}  rows={total}  rows needing a look (Check column)={flagged}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", default="", help="PDF file or folder of PDFs")
    ap.add_argument("output", nargs="?", default="", help="folder for the Excel files (one per PDF, same name as the PDF)")
    ap.add_argument("--model", default="gemini-3.5-flash-lite")
    ap.add_argument("--rpm", type=float, default=10, help="requests per minute limit")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--fallback-model", default="", help="stronger model used only to re-read pages with bad letters, e.g. gemini-3.5-flash")
    ap.add_argument("--rereads", type=int, default=2, help="max re-reads of a page that has non-Gujarati letters (each costs 1 request)")
    ap.add_argument("--ward-map", default="")
    ap.add_argument("--skip-existing", action="store_true", help="skip PDFs whose Excel already exists")
    ap.add_argument("--cache", default="")
    ap.add_argument("--build-only", action="store_true")
    ap.add_argument("--list-models", action="store_true", help="show models your key can use, then exit")
    a = ap.parse_args()
    if a.list_models:
        from google import genai
        c = genai.Client(api_key=os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"))
        for mo in c.models.list():
            if "generateContent" in (getattr(mo, "supported_actions", None) or []):
                print(mo.name.replace("models/", ""))
        sys.exit(0)
    if not a.input or not a.output:
        ap.error("input and output are required")
    pdfs = sorted(glob.glob(os.path.join(a.input, "**", "*.pdf"), recursive=True)) if os.path.isdir(a.input) else [a.input]
    if not pdfs:
        sys.exit("No PDF files found in " + a.input)
    if os.path.isdir(a.input):
        out_dir = a.output if not a.output.lower().endswith(".xlsx") else (os.path.dirname(a.output) or ".")
    else:
        out_dir = (os.path.dirname(a.output) or ".") if a.output.lower().endswith(".xlsx") else a.output
    os.makedirs(out_dir, exist_ok=True)
    a.cache = a.cache or os.path.join(out_dir, "_cache")
    os.makedirs(a.cache, exist_ok=True)
    client = types = None
    if not a.build_only:
        client, types = get_client()
        if client is None and not a.model.startswith("ollama:"):
            sys.exit("Set GEMINI_API_KEY first (free key: https://aistudio.google.com/apikey)")
    done, incomplete = [], []
    stopped = False
    for i, pdf in enumerate(pdfs, 1):
        print(f"\n[{i}/{len(pdfs)}]", flush=True)
        out_path = os.path.join(out_dir, os.path.splitext(os.path.basename(pdf))[0] + ".xlsx")
        if a.skip_existing and os.path.exists(out_path):
            print(f"{os.path.basename(pdf)}: Excel already exists, skipped"); done.append(out_path); continue
        complete = True
        if not a.build_only:
            try:
                complete = extract_pdf(client, types, a, pdf)
            except QuotaExhausted as e:
                print("\nDaily free quota finished. Run the same command again later - it resumes from here.\n", str(e)[:200])
                incomplete.append(pdf); stopped = True; break
        if complete:
            build_excel(a, [pdf], out_path); done.append(out_path)
        else:
            print(f"   {os.path.basename(pdf)}: some pages failed - Excel not written. Run again to retry them.")
            incomplete.append(pdf)
    if stopped:
        incomplete += pdfs[i:]
    print(f"\nFINISHED: {len(done)} Excel file(s) in '{out_dir}'. Not finished: {len(incomplete)}")
    for p in incomplete: print("   -", os.path.basename(p))