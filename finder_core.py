"""Statement Finder: the engine. No window in this file, so it can be tested on its own.

What it does
    1. walks a folder and every folder under it and lists the files that could be statements
    2. reads each file (PDF, Word, text / bank report / HTML) in its own worker process, several at the same time
    3. decides: is it a bank statement, and which bank
    4. copies the statements into <destination>/<Bank name>/ (never moves, never changes the original)
    5. writes a CSV report of everything it found

How it decides (no AI model, no internet)
    * title words ("statement of account"), column words ("narration", "withdrawal", "balance"),
      account words ("IFSC", "opening balance") and, most of all, rows that look like transactions
      (a date and an amount on the same line)
    * when a statement has no header or column titles, it checks for transaction rows alongside common
      banking terms ("UPI", "NEFT", "RTGS", "IMPS", "BY TRANSFER", "Cr"/"Dr")
    * the bank comes from the IFSC code or the bank's name in the heading, or (when the heading is blank)
      from page footers, bank website addresses, bank-specific phrases, or the file / folder name
    * scanned PDFs, password-protected PDFs, Excel/CSV spreadsheets, invoices, ledgers, mutual fund reports
      and the Statement Analyser's own audit reports are left out

A worker that hangs or crashes on one bad file is stopped and replaced; the others carry on.
"""
import csv
import hashlib
import html
import multiprocessing as mp
import os
import queue
import re
import shutil
import time

VERSION = '1.1'

FILE_TYPES = {
    'pdf': ('.pdf',),
    'word': ('.docx',),
    'text': ('.txt', '.rpt', '.rptnfs', '.nfs', '.tsv', '.htm', '.html'),
}
FILE_TYPE_NAMES = {'pdf': 'PDF', 'word': 'Word', 'text': 'Text / Report (.rpt, .nfs, .txt)'}

DEFAULTS = {
    'kinds': ('pdf', 'word', 'text'),
    'include_not_sure': True,        # also copy the "likely" ones, into a "(not sure)" folder
    'skip_system': True,             # Windows, Program Files, AppData, .git, caches ...
    'dry_run': False,                # list what would be copied, copy nothing
    'workers': 0,                    # 0 = one per processor thread
    'timeout': 60,                   # seconds allowed for one file
    'max_mb': 300,                   # bigger files are not opened
    'pages': 4,                      # PDF pages read (the first ones are enough to tell)
}

NOT_SURE_FOLDER = '(not sure)'
UNKNOWN_BANK = '_Bank not identified'
CARD_FOLDER = '_Credit card statements'

SYSTEM_FOLDERS = {
    '$recycle.bin', 'system volume information', 'windows', 'program files', 'program files (x86)', 'programdata',
    'appdata', 'node_modules', '.git', '__pycache__', 'site-packages', '$windows.~bt', 'recovery', 'msocache',
    '.venv', 'venv', '.cache', '.gradle', '.m2', '.npm',
}


class Locked(Exception):
    """The file needs a password."""


# ---------------------------------------------------------------------------------------------------------- banks
# (folder name, name pattern, IFSC prefixes). The pattern is searched in the heading of the statement first.
BANKS = [
    ('State Bank of India', r'state\s+bank\s+of\s+india|\bsbi\b(?!\s*(?:life|card|mutual|general|cap))', ('SBIN',)),
    ('HDFC Bank', r'\bhdfc\s*bank\b', ('HDFC',)),
    ('ICICI Bank', r'\bicici\s*bank\b', ('ICIC',)),
    ('Axis Bank', r'\baxis\s*bank\b|statement\s+of\s+axis\s+account', ('UTIB',)),
    ('Kotak Mahindra Bank', r'kotak\s*mahindra|\bkotak\b', ('KKBK',)),
    ('Punjab National Bank', r'punjab\s+national\s+bank|\bpnb\b', ('PUNB',)),
    ('Bank of Baroda', r'bank\s+of\s+baroda', ('BARB',)),
    ('Canara Bank', r'canara\s*bank', ('CNRB',)),
    ('Union Bank of India', r'union\s+bank\s+of\s+india', ('UBIN',)),
    ('Bank of India', r'bank\s+of\s+india', ('BKID',)),
    ('Indian Bank', r'\bindian\s+bank\b', ('IDIB',)),
    ('Central Bank of India', r'central\s+bank\s+of\s+india', ('CBIN',)),
    ('IDBI Bank', r'\bidbi\s*bank\b', ('IBKL',)),
    ('IndusInd Bank', r'\bindusind\b', ('INDB',)),
    ('Yes Bank', r'\byes\s+bank\b', ('YESB',)),
    ('Federal Bank', r'\bfederal\s+bank\b', ('FDRL',)),
    ('South Indian Bank', r'south\s+indian\s+bank', ('SIBL',)),
    ('Karnataka Bank', r'karnataka\s+bank', ('KARB',)),
    ('Karur Vysya Bank', r'karur\s+vysya', ('KVBL',)),
    ('City Union Bank', r'city\s+union\s+bank', ('CIUB',)),
    ('Tamilnad Mercantile Bank', r'tamilnad\s+mercantile', ('TMBL',)),
    ('Dhanlaxmi Bank', r'dhanlaxmi', ('DLXB',)),
    ('J&K Grameen Bank', r'(?:jammu\s*(?:&|and)\s*kashmir|j\s*&\s*k)\s+grameen\s+bank|ellaquai\s+dehati', ()),
    ('J&K Bank', r'jammu\s*(?:&|and)\s*kashmir\s+bank|\bj\s*&\s*k\s+bank\b|\bjkb\b|\bjk\s+bank\b', ('JAKA',)),
    ('IDFC First Bank', r'idfc\s*first|\bidfc\s+bank', ('IDFB',)),
    ('UCO Bank', r'\buco\s+bank\b', ('UCBA',)),
    ('Indian Overseas Bank', r'indian\s+overseas\s+bank', ('IOBA',)),
    ('Bandhan Bank', r'bandhan\s+bank', ('BDBL',)),
    ('RBL Bank', r'\brbl\s+bank\b', ('RATN',)),
    ('AU Small Finance Bank', r'\bau\s+small\s+finance', ('AUBL',)),
    ('Paytm Payments Bank', r'paytm\s+payments\s+bank', ('PYTM',)),
    ('Airtel Payments Bank', r'airtel\s+payments\s+bank', ('AIRP',)),
    ('India Post Payments Bank', r'india\s+post\s+payments\s+bank', ('IPOS',)),
    ('Punjab & Sind Bank', r'punjab\s*(?:&|and)\s*sind\s+bank', ('PSIB',)),
    ('Bank of Maharashtra', r'bank\s+of\s+maharashtra', ('MAHB',)),
    ('Equitas Small Finance Bank', r'equitas\s+small', ('ESFB',)),
    ('Ujjivan Small Finance Bank', r'ujjivan\s+small', ('UJVN',)),
    ('Jana Small Finance Bank', r'\bjana\s+small\s+finance', ('JSFB',)),
    ('ESAF Small Finance Bank', r'\besaf\s+small', ('ESMF',)),
    ('Fino Payments Bank', r'fino\s+payments', ('FINO',)),
    ('DBS Bank', r'\bdbs\s+bank\b', ('DBSS',)),
    ('Standard Chartered Bank', r'standard\s+chartered', ('SCBL',)),
    ('HSBC', r'\bhsbc\b', ('HSBC',)),
    ('Citibank', r'\bciti\s*bank\b', ('CITI',)),
    ('Deutsche Bank', r'deutsche\s+bank', ('DEUT',)),
    ('Mintifi', r'\bmintifi\b', ()),
    ('Mahindra Finance', r'\bmahindra\s+(?:&\s+mahindra\s+financial\s+services|finance)\b', ()),
    ('Tata Motors Finance', r'\btata\s+motors\s+finance\b', ()),
    ('Urban Co-operative Bank', r'\bu\.?\s*c\.?\s*b\.?\s*ltd\b', ()),
]
_BANK_RX = [(name, re.compile(pat, re.I)) for name, pat, _ifsc in BANKS]
_IFSC_BANK = {p: name for name, _pat, prefixes in BANKS for p in prefixes}

# Website addresses and unmistakable bank-specific phrases found in page footers or report dumps
_BANK_FINGERPRINTS = [
    ('State Bank of India', re.compile(
        r'(?<![a-z])(?:onlinesbi\.com|sbi\.co\.in)|\bby\s+transfer-inb\b|\bto\s+transfer-upi\b|\bmod\s+balance\b|'
        r'Account\s+Number\s*:\s*0{6}\d{11}[\s\S]{0,800}\bC\.Y\.YTD\s+Interest\b', re.I)),
    ('J&K Grameen Bank', re.compile(r'(?<![a-z])jkgrameen\.com|\bellaquai\s+dehati\b', re.I)),
    ('J&K Bank', re.compile(
        r'(?<![a-z])(?:jkbank\.com|jkb\.com)|\bjkb\s+mconnect\b|'
        r'\((?:within|with\s+in)\s+j\s*&\s*k\s+bank\)|'
        r'Account:\s*\d{16}\|[^|\n]+\|(?:SBA|CAA|CCA|ODA|LAA|TDA)\|\d{4}\b|'
        r'Transactions\s+List\s*-\s*(?:CD|CC|SB|OD|LA)100-', re.I)),
    ('HDFC Bank', re.compile(r'(?<![a-z])hdfcbank\.com', re.I)),
    ('ICICI Bank', re.compile(
        r'(?<![a-z])icicibank\.com|\((?:within|with\s+in)\s+icici\s+bank\)|'
        r'DETAILED\s+STATEMENT[\s\S]{0,600}Account\s+Number\s+\d{12}\(INR\)[\s\S]{0,400}Cheque\s+number\s+from\s+NA\s+to\s+NA|'
        r'DETAILED\s+STATEMENT[\s\S]{0,300}Transactions\s+List\s*-\s*-|'
        r'Consolidated\s+Transaction\s+Summary\s+for\s+the\s+Account\s+Number:\s*\d{12}\b|'
        r'Loan\s+Account\s+Statement\s+for\s+LA[A-Z]{3}\d{11}\b', re.I)),
    ('Axis Bank', re.compile(r'(?<![a-z])axisbank\.com', re.I)),
    ('Kotak Mahindra Bank', re.compile(r'(?<![a-z])kotak\.com', re.I)),
    ('Punjab National Bank', re.compile(r'(?<![a-z])(?:pnbindia\.in|netpnb\.com)', re.I)),
    ('Bank of Baroda', re.compile(r'(?<![a-z])(?:bankofbaroda\.in|bobibanking\.com)', re.I)),
    ('Canara Bank', re.compile(r'(?<![a-z])canarabank\.com', re.I)),
    ('Union Bank of India', re.compile(r'(?<![a-z])unionbankofindia\.co\.in', re.I)),
    ('Central Bank of India', re.compile(r'(?<![a-z])centralbankofindia\.co\.in', re.I)),
    ('Bank of India', re.compile(r'(?<![a-z])bankofindia\.co\.in', re.I)),
    ('Indian Bank', re.compile(r'(?<![a-z])indianbank\.in', re.I)),
    ('IDBI Bank', re.compile(r'(?<![a-z])idbibank\.in', re.I)),
    ('IndusInd Bank', re.compile(r'(?<![a-z])indusind\.com', re.I)),
    ('Yes Bank', re.compile(r'(?<![a-z])yesbank\.in', re.I)),
    ('Federal Bank', re.compile(r'(?<![a-z])federalbank\.co\.in', re.I)),
    ('IDFC First Bank', re.compile(r'(?<![a-z])idfcfirstbank\.com', re.I)),
    ('UCO Bank', re.compile(r'(?<![a-z])ucobank\.com', re.I)),
    ('Indian Overseas Bank', re.compile(r'(?<![a-z])iob\.in', re.I)),
    ('Bandhan Bank', re.compile(r'(?<![a-z])bandhanbank\.com', re.I)),
    ('RBL Bank', re.compile(r'(?<![a-z])rblbank\.com', re.I)),
    ('AU Small Finance Bank', re.compile(r'(?<![a-z])aubank\.in', re.I)),
    ('Mahindra Finance', re.compile(r'(?<![a-z])mahindrafinance\.com', re.I)),
]

# Short bank names often used in file names or folder names (checked only when the document itself has no bank name)
_PATH_BANK_RX = [
    ('J&K Grameen Bank', re.compile(r'(?<![a-z])(?:jk\s*grameen|j\s*&\s*k\s*grameen|grameen\s*bank|edbank)(?![a-z])', re.I)),
    ('J&K Bank', re.compile(r'(?<![a-z])(?:jkb|jk\s*bank|j\s*&\s*k\s*bank|jammu\s*(?:&|and)\s*kashmir\s*bank)(?![a-z])', re.I)),
    ('State Bank of India', re.compile(r'(?<![a-z])(?:sbi|state\s+bank\s+of\s+india)(?![a-z])', re.I)),
    ('HDFC Bank', re.compile(r'(?<![a-z])hdfc(?![a-z])', re.I)),
    ('ICICI Bank', re.compile(r'(?<![a-z])icici(?![a-z])', re.I)),
    ('Axis Bank', re.compile(r'(?<![a-z])axis(?:\s*bank)?(?![a-z])', re.I)),
    ('Kotak Mahindra Bank', re.compile(r'(?<![a-z])kotak(?![a-z])', re.I)),
    ('Punjab National Bank', re.compile(r'(?<![a-z])(?:pnb|punjab\s+national)(?![a-z])', re.I)),
    ('Bank of Baroda', re.compile(r'(?<![a-z])(?:bob|bank\s+of\s+baroda|baroda)(?![a-z])', re.I)),
    ('Canara Bank', re.compile(r'(?<![a-z])canara(?![a-z])', re.I)),
    ('Union Bank of India', re.compile(r'(?<![a-z])(?:ubin|union\s*bank)(?![a-z])', re.I)),
    ('Bank of India', re.compile(r'(?<![a-z])(?:boi|bank\s+of\s+india)(?![a-z])', re.I)),
    ('Indian Bank', re.compile(r'(?<![a-z])indian\s*bank(?![a-z])', re.I)),
    ('Central Bank of India', re.compile(r'(?<![a-z])(?:cbi|central\s*bank)(?![a-z])', re.I)),
    ('IDBI Bank', re.compile(r'(?<![a-z])idbi(?![a-z])', re.I)),
    ('IndusInd Bank', re.compile(r'(?<![a-z])indusind(?![a-z])', re.I)),
    ('Yes Bank', re.compile(r'(?<![a-z])yes\s*bank(?![a-z])', re.I)),
    ('Federal Bank', re.compile(r'(?<![a-z])federal\s*bank(?![a-z])', re.I)),
    ('IDFC First Bank', re.compile(r'(?<![a-z])idfc(?![a-z])', re.I)),
    ('UCO Bank', re.compile(r'(?<![a-z])uco(?:\s*bank)?(?![a-z])', re.I)),
    ('Indian Overseas Bank', re.compile(r'(?<![a-z])(?:iob|indian\s+overseas)(?![a-z])', re.I)),
    ('Bandhan Bank', re.compile(r'(?<![a-z])bandhan(?![a-z])', re.I)),
    ('RBL Bank', re.compile(r'(?<![a-z])rbl(?:\s*bank)?(?![a-z])', re.I)),
    ('AU Small Finance Bank', re.compile(r'(?<![a-z])au\s*(?:bank|small\s*finance)(?![a-z])', re.I)),
]

# banks that are named in a general way: "<name> Co-operative Bank", "<name> Gramin Bank" ...
_GENERIC_BANK_RX = [
    (re.compile(r"([A-Z][A-Za-z&.'\- ]{2,40}?)\s+(?i:co-?\s*op(?:erative)?)\.?\s+(?i:bank)\b"), 'Co-operative Bank'),
    (re.compile(r"([A-Z][A-Za-z&.'\- ]{2,30}?)\s+(?i:gramin|grameen)\s+(?i:bank)\b"), 'Gramin Bank'),
    (re.compile(r"([A-Z][A-Za-z&.'\- ]{2,30}?)\s+(?i:small\s+finance\s+bank)\b"), 'Small Finance Bank'),
    (re.compile(r"([A-Z][A-Za-z&.'\- ]{2,30}?)\s+(?i:payments\s+bank)\b"), 'Payments Bank'),
]
_RBI_RX = re.compile(r'reserve\s+bank\s+of\s+india', re.I)          # mentioned in many documents; never the bank of the account
_IFSC_RX = re.compile(r'\b([A-Z]{4})0([A-Z0-9]{6})\b')
# Sub-member co-op / RRB sponsor IFSCs, UPI virtual handles, and RBI tax payment codes share a sponsor prefix but are not the sponsor bank's own branches
_NON_OWN_IFSC_RX = re.compile(r'^(?:HDFC0C|UTIB0S|SBIN0RR|RBIS0|[A-Z]{4}0[A-Z]*UPI$)')
_CLIPPED_JKB_IFSC_RX = re.compile(r'\bIFSC\s*(?:CODE)?\s*[:\-]\s*(JAKA0(?!GRAM|GRMN|EDB)[A-Z]{5})\b', re.I)
_ADDRESS_LINE_RX = re.compile(
    r'^\s*(?:(?:customer|permanent|communication|mailing|correspondence|residen(?:tial|ce))\s+)?address\s*[:\-]', re.I)
_ADDRESS_PREP_RX = re.compile(
    r'\b(?:opposite|opp\.?|near|nr\.?|behind|beside|adjacent\s+to|in\s+front\s+of|next\s+to)\b', re.I)
_TRAILING_PREP_RX = re.compile(
    r'\b(?:opposite|opp\.?|near|nr\.?|behind|beside|adjacent\s+to|in\s+front\s+of|next\s+to)\s*[-:,]?\s*(?:the\s+)?$', re.I)
_FOOTER_CUE_RX = re.compile(
    r'registered\s+office|corporate\s+office|head\s+office|(?:computer|system|electronically)[- ]generated|'
    r'does\s+not\s+require\s+(?:any\s+)?signature|toll[- ]free|customer\s+care|helpline|legends\s+used\s+in|'
    r'\bpage\s+\d+(?:\s+of\s+\d+)?|\bgenerated\s+(?:on|at|from|by)\b|\bifsc\s*(?:code)?\s*[:\-]|'
    r'\bmicr\s*(?:code)?\s*[:\-]|statement\s+summary|end\s+of\s+statement|https?://|www\.', re.I)
_NARRATION_FRAG_RX = re.compile(
    r'^\s*/|\b(?:upi|neft|rtgs|imps|bil|mmt|mtfr|inb|ach|ecs|nach|pos|atm|phonepe|gpay|paytm|bhim|paymen|'
    r'to\s+transfer|by\s+transfer)\b|@|\d+-[A-Z]{4}0[A-Z0-9]{6}\b|\b[A-Z]{4}0[A-Z0-9]{6}/|[/\-][A-Z]{4}0[A-Z0-9]{6}\b', re.I)


def _is_address_mention(text, start, matched=''):
    """True when a bank name appears inside a customer address or after a landmark preposition like 'Opposite ...'."""
    if matched and _ADDRESS_PREP_RX.search(matched):
        return True
    line_start = text.rfind('\n', 0, start) + 1
    prefix = text[line_start:start]
    if _ADDRESS_LINE_RX.match(prefix):
        return True
    if _TRAILING_PREP_RX.search(prefix[-50:]):
        return True
    if not prefix.strip() and line_start > 1:
        prev_start = text.rfind('\n', 0, line_start - 1) + 1
        if _TRAILING_PREP_RX.search(text[prev_start:line_start - 1].rstrip()):
            return True
    return False


def _drop_contained_spans(spans):
    """Drops bank matches nested inside a longer match of another bank (e.g. 'Bank of India' inside 'Central Bank of India' or RBI)."""
    kept = [s for s in spans if not any(
        o[2] != s[2] and o[0] <= s[0] and s[1] <= o[1] and (o[1] - o[0]) > (s[1] - s[0])
        for o in spans
    )]
    return sorted(s for s in kept if s[2] is not None)


_HONORIFICS = {'mr', 'mr.', 'mrs', 'mrs.', 'ms', 'ms.', 'dr', 'dr.', 'shri', 'smt', 'sh', 'sh.', 'm/s', 's/o', 'd/o', 'w/o', 'c/o'}
_NAME_COL_RX = re.compile(r'\b(?:customer|account|holder)\s+name\s*:', re.I)


def _generic_bank(text, skip_address=True):
    """-> (folder name, start, end) for a co-operative / gramin / small-finance bank that is not in the list, or None."""
    best = None
    for rx, tail in _GENERIC_BANK_RX:
        for m in rx.finditer(text):
            if skip_address and _is_address_mention(text, m.start(), m.group(1)):
                continue
            words = [w for w in m.group(1).replace('&', ' & ').split()
                     if w.lower() not in ('the', 'of', 'from', 'a', 'an', 'statement', 'account')]
            words = [w for w in words if w[:1].isalpha()]
            # A two-column header ("Customer Name : Mr. SAMPLE HOLDER   Sample Central Co-op Bank") can leak the person's name
            h_idx = max((i for i, w in enumerate(words[:-1]) if w.lower() in _HONORIFICS), default=-1)
            line_prefix = text[text.rfind('\n', 0, m.start()) + 1:m.start()]
            if h_idx >= 0 or (skip_address and _NAME_COL_RX.search(line_prefix)):
                if h_idx >= 0:
                    words = words[h_idx + 1:]
                first_title = next((i for i, w in enumerate(words) if not w.isupper()), -1)
                if first_title > 0:
                    words = words[first_title:]
            words = words[-3:]
            if not words:
                continue
            name = ' '.join(w.capitalize() if w.isupper() or w.islower() else w for w in words) + ' ' + tail
            if best is None or m.start() < best[1]:
                best = (name, m.start(), m.end())
            break
    return best


def _bank_from_path(path):
    """Checks the file name first, then up to 2 parent folder names, for a known bank name."""
    if not path:
        return '', ''
    parts = [os.path.splitext(os.path.basename(path))[0]]
    parent = os.path.dirname(path)
    for _ in range(2):
        base = os.path.basename(parent)
        if not base or base.lower() in ('source', 'statements', 'bank', 'banks', 'files', 'documents', 'desktop', 'downloads'):
            break
        parts.append(base)
        parent = os.path.dirname(parent)
    for i, part in enumerate(parts):
        where = 'file name' if i == 0 else 'folder name'
        cleaned = re.sub(r'[_-]+', ' ', part)
        spans = []
        for candidate in (part, cleaned):
            for name, rx in _BANK_RX + _PATH_BANK_RX + [(None, _RBI_RX)]:
                for m in rx.finditer(candidate):
                    spans.append((m.start(), m.end(), name))
        kept = _drop_contained_spans(spans)
        if kept:
            return kept[0][2], f'from the {where}'
        generic = _generic_bank(part, skip_address=False) or _generic_bank(cleaned, skip_address=False)
        if generic:
            return generic[0], f'from the {where}'
    return '', ''


def _non_tx_text(text):
    """Header and footer/summary lines from `text`, excluding dated rows and wrapped transaction narrations."""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    dated = _dated_flags(lines)
    out, in_table, in_footer, since_dated = [], False, False, 0
    for i, ln in enumerate(lines):
        if dated[i]:
            in_table, in_footer, since_dated = True, False, 0
            continue
        is_cue = bool(_FOOTER_CUE_RX.search(ln))
        if is_cue:
            in_footer = True
        if not in_table or is_cue or (in_footer and not _NARRATION_FRAG_RX.search(ln)):
            out.append(ln)
        else:
            since_dated += 1
            if since_dated > 5 and not _NARRATION_FRAG_RX.search(ln):
                out.append(ln)
    return '\n'.join(out)


def find_bank(head, whole='', path=''):
    """The bank a statement belongs to. `head` is the heading (before the first transaction), `whole` more of the text,
    and `path` the file's path on disk (used as a last resort when the document text itself does not name the bank).
    The bank that is named first wins; an IFSC code that agrees with a name settles it.
    -> (name or '', how)"""
    def ifsc_in(text):
        hits, upper, seen = [], text.upper(), set()          # (position, bank, code)
        for m in _IFSC_RX.finditer(upper):
            code = m.group(1) + '0' + m.group(2)
            seen.add(m.start())
            if code.startswith(('JAKA0GRAM', 'JAKA0GRMN', 'JAKA0EDB')):
                hits.append((m.start(), 'J&K Grameen Bank', code))
            elif _NON_OWN_IFSC_RX.match(code):
                continue
            elif m.group(1) in _IFSC_BANK:
                hits.append((m.start(), _IFSC_BANK[m.group(1)], code))
        for m in _CLIPPED_JKB_IFSC_RX.finditer(upper):
            if m.start(1) not in seen:
                hits.append((m.start(1), 'J&K Bank', m.group(1)))
        return sorted(hits)

    def by_name(text):
        spans = []
        for name, rx in _BANK_RX + [(None, _RBI_RX)]:
            for m in rx.finditer(text):
                spans.append((m.start(), m.end(), name))
        spans = [s for s in _drop_contained_spans(spans) if not _is_address_mention(text, s[0])]
        generic = _generic_bank(text)
        if generic and not any(s[0] < generic[2] and generic[1] < s[1] for s in spans):    # an exact match over the same words wins
            spans.append((generic[1], generic[2], generic[0]))
        return sorted(spans)

    ifsc_hits = ifsc_in(head)
    names = by_name(head)
    if ifsc_hits and names:
        for _pos, bank, code in ifsc_hits:
            if any(n[2] == bank for n in names):
                return bank, f'IFSC {code}'
        # they disagree (a counterparty's IFSC in a narration, say): whichever comes first in the statement is the account's bank
        if names[0][0] < ifsc_hits[0][0]:
            return names[0][2], f'"{names[0][2]}" in the heading'
        return ifsc_hits[0][1], f'IFSC {ifsc_hits[0][2]}'
    if ifsc_hits:
        return ifsc_hits[0][1], f'IFSC {ifsc_hits[0][2]}'
    if names:
        return names[0][2], f'"{names[0][2]}" in the heading'

    if whole:
        # 1. Website URLs or bank-specific phrases in footers/text (never counterparty names)
        for bank, rx in _BANK_FINGERPRINTS:
            if rx.search(whole):
                return bank, 'bank website / phrase on the page'
        # 2. Non-transaction lines (page footers, summary lines) that are not wrapped transaction narrations
        non_tx = _non_tx_text(whole)
        footer_ifsc = ifsc_in(non_tx)
        if footer_ifsc:
            return footer_ifsc[0][1], f'IFSC {footer_ifsc[0][2]} on the page'
        footer_names = by_name(non_tx)
        if footer_names:
            return footer_names[0][2], 'named in the page footer'

    if path:
        return _bank_from_path(path)
    return '', ''


# ---------------------------------------------------------------------------------------------------------- the text test
_MONTH = r'(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*'
DATE_RX = re.compile(
    r'\b\d{1,2}[-–—/.]\d{1,2}[-–—/.]\d{2,4}\b|\b\d{1,2}[-–—/ .]' + _MONTH + r'[-–—/ .,]*\d{2,4}\b|'
    r'\b\d{4}[-–—]\d{2}[-–—]\d{2}\b|\b(?:0?[1-9]|[12]\d|3[01]),(?:0?[1-9]|[12]\d|3[01]),(?:19|20)\d{2}\b', re.I)
_WRAP_DATE_RX = re.compile(r'\b\d{1,2}[-–—/ ]' + _MONTH + r'[-–—/](?=\s|$)', re.I)
_WRAP_YEAR_RX = re.compile(r'^\s*(?:19|20)\d{2}\b')
AMOUNT_RX = re.compile(r'(?<![\d.,])(?:\d{1,3}(?:,\d{2,3})+|\d+)\.\d{2}(?!\d|\.\d)')
ENDS_WITH_AMOUNT = re.compile(r'\d\.\d{2}[\s"\',;|]*(?:cr|dr|c|d)?\.?[\s"\',;|]*$', re.I)       # a transaction line ends with an amount (or its balance)
CR_DR_END_RX = re.compile(r'\d\.\d{2}[\s"\',;|]*(?:cr|dr)\.?[\s"\',;|]*$', re.I)

TITLES = ('statement of account', 'account statement', 'bank statement', 'statement for the period', 'account summary',
          'transaction history', 'transaction statement', 'e-passbook', 'epassbook', 'passbook', 'mini statement',
          'statement of transactions', 'account activity', 'detailed statement', 'statement of acc')
COLUMNS = ('narration', 'particulars', 'description', 'withdrawal', 'deposit', 'debit', 'credit', 'balance', 'chq', 'cheque',
           'ref no', 'reference', 'value date', 'txn date', 'transaction date', 'tran date', 'post date', 'dr/cr', 'amount')
MONEY_COLUMNS = ('debit', 'credit', 'withdrawal', 'deposit', 'dr/cr')
ACCOUNT_WORDS = ('account no', 'account number', 'a/c no', 'a/c number', 'acct no', 'ifsc', 'branch', 'customer id', 'customer name',
                 'cif', 'opening balance', 'closing balance', 'brought forward', 'statement period', 'from date', 'micr',
                 'nomination', 'account type', 'account open', 'sol id', 'tran id', 'sanction limit', 'drawing power')
BANKING_WORDS_RX = re.compile(
    r'\b(?:upi|neft|rtgs|imps|atm|chq|cheque|by\s+transfer|to\s+transfer|cash\s+dep|clearing|clg|ecs|nach|ach|pos|emi|sweep|inb)\b', re.I)

# things that look like transaction lists but are not bank statements; searched in the heading only
NOT_STATEMENT = [
    ('made by the Statement Analyser', re.compile(
        r'audit\s+report|executive\s+(?:audit|financial)|bank\s+statement\s+(?:overview|summary)\s+report|contra\s*/\s*self\s+transfers|'
        r'statement\s+analyser|receipts\s+from\s+contracts|net\s+operating\s+deposits|monthly\s+movement', re.I), 20),
    ('an invoice or bill', re.compile(r'tax\s+invoice|proforma\s+invoice|invoice\s+no|bill\s+of\s+supply|purchase\s+order|e-?way\s+bill|delivery\s+challan|quotation', re.I), 12),
    ('an accounting report', re.compile(
        r'day\s+book|ledger\s+account|trial\s+balance|profit\s*(?:and|&)\s*loss|balance\s+sheet|sales\s+register|purchase\s+register|'
        r'\bgstr[\s-]*\d+[a-z]?\b|\bdrc-?0?3\b|\bitc\s+claimed\b|taxable\s+value|cash\s+book|notes?\s+to\s+accounts|'
        r'\bvch\.?\s*(?:type|no)\b|\bvou-?no\b|\bsgledger\d*\b|'
        r'\bgroup\s*:\s*sundry\s+(?:debtors|creditors|purchase|sales)\b|'
        r'\b(?:purchase|sales?|receipt|payment)\s+(?:[a-z]+\s+){0,2}register\b|'
        r'\bcash\s*/\s*bank\s+book\b|\bL\s+E\s+D\s+G\s+E\s+R\b|'
        r'\bdr\s+amt\.?\s*\([^)]*\)\s*\|\s*particulars\s+cr\s+amt\b', re.I), 25),
    ('a GST portal ledger', re.compile(
        r'electronic\s+(?:credit|cash)\s+ledger|electronic\s+liability\s+register|'
        r'integrated\s+central\s+state\s+tax\s+cess', re.I), 25),
    ('a mutual fund / demat report', re.compile(r'consolidated\s+account\s+statement|mutual\s+fund|folio\s+no|\bnsdl\b|\bcdsl\b|\bdemat\b', re.I), 15),
    ('a salary slip or tax form', re.compile(r'salary\s+slip|pay\s*slip|form\s+16|form\s+26as|income\s+tax\s+return', re.I), 25),
    ('a loan repayment schedule', re.compile(
        r'repayment\s+schedule(?![\s\S]{0,600}\b(?:withdrawal|deposit|debit\s*/\s*credit)\b)|'
        r'instl\.?\s+no\s+due\s+date\s+instl\.?\s+amt[\s\S]{0,120}\bo/s\s+principal\b', re.I), 25),
    ('a loan or legal document', re.compile(
        r'hypothecation|guarantee\s+deed|sanction\s+letter|covering\s+letter|affidavit|\bd\.?p\.?\s+note|letter\s+of\s+continuity|power\s+of\s+attorney|'
        r'restructuring\s+proposal|name\s+and\s+address\s+of\s+the\s+borrower', re.I), 20),
]
# unmistakable non-bank documents even when dated rows begin near the top of the first page (note: never match bare "GSTIN", which HDFC prints in footers)
DOC_NOT_STATEMENT = [
    ('a Form 26AS / TDS tax statement', re.compile(
        r'annual\s+tax\s+statement|tan\s+of\s+deductor|name\s+of\s+assessee|details\s+of\s+tax\s+deducted\s+at\s+source|'
        r'tinpan\.proteantech\.in|tin-?nsdl\.com|tdscpc\.gov\.in', re.I), 25),
    ('a broker / equity statement', re.compile(
        r'equity\s+transaction\s+statement|sebi\s+regn|unique\s+client\s+code|trading\s+code\s+no|'
        r'member\s+(?:of\s+)?national\s+stock\s+exchange|member\s+bombay\s+stock\s+exchange', re.I), 25),
    ('a GST / bill schedule', re.compile(
        r'bills\s+verified\s+against\s+original\s+receipts|details\s+of\s+payments\s+received\s+in\s+bank\s+account|'
        r'\bgstr[\s-]*7\s+details\b|taxable\s+value[\s\S]{0,120}\b(?:cgst|sgst|igst)\b|'
        r'\b(?:cgst|sgst|igst)\b[\s\S]{0,120}taxable\s+value|\bsale\s+bill\s+no\.?|\bpurchase\s+bill\s+no\.?', re.I), 25),
    ('a telecom / recharge usage bill', re.compile(
        r'account\s+usage\s+and\s+recharge\s+statement|'
        r'\bjio\s+number\s*:[\s\S]{0,500}(?:plan\s+recharges|usage\s+charges|itemized\s+usage)|'
        r'itemized\s+usage[\s\S]{0,200}(?:billed|chargeable)\s+usage', re.I), 25),
]
CARD_RX = re.compile(
    r'credit\s+card\s+statement|total\s+amount\s+due|minimum\s+(?:amount\s+)?due|payment\s+due\s+date|'
    r'card\s+account\s+no|credit\s+card\s+no|available\s+credit\s+limit|credit\s+limit|cash\s+limit|'
    r'card\s+no|card\s+number|reward\s+points', re.I)
_STRONG_CARD_TERMS = {
    'credit card statement', 'total amount due', 'minimum due', 'minimum amount due',
    'payment due date', 'card account no', 'credit card no', 'available credit limit',
}
OWN_REPORT_NAME = re.compile(r'_audit_(?:report|summary)', re.I)            # how the Statement Analyser names what it exports


_PURE_WORD = re.compile(r'(?<![\w])[A-Za-z]{3,}(?![\w])')


def _starts_with_date(line):
    """A transaction line begins with its date (maybe after a serial number or an ID such as TXN8845). A title such as
    "BALANCE SHEET AS ON 31-03-2025" or "STATEMENT FOR THE PERIOD 01-04-2025" has words before its date."""
    m = DATE_RX.search(line)
    return bool(m) and not _PURE_WORD.search(line[:m.start()])


def _dated_flags(lines):
    """Flags each non-empty line in `lines` that begins a transaction date, including 'DD-Mon-' whose year wraps to the next line."""
    flags = [_starts_with_date(ln) for ln in lines]
    for i, ln in enumerate(lines):
        if flags[i]:
            continue
        m = _WRAP_DATE_RX.search(ln)
        if not m or _PURE_WORD.search(ln[:m.start()]):
            continue
        for j in (i + 1, i + 2):
            if j < len(lines) and not flags[j] and _WRAP_YEAR_RX.match(lines[j]):
                flags[i] = True
                break
    return flags


def _is_row(line, amounts):
    """A transaction line ends with an amount (usually the balance) or carries two. Prose that mentions a date and an
    amount ("dated 12-05-2025 for Rs 15.00 lacs") does not."""
    return bool(amounts) and (len(amounts) >= 2 or bool(ENDS_WITH_AMOUNT.search(line)))


def analyse_text(text, filename='', path=''):
    """Is this text a bank statement, and which bank? Pure function, no file access.
    -> {'verdict': 'statement' | 'likely' | 'not', 'kind': 'bank' | 'card', 'bank', 'bank_how', 'score', 'rows', 'reasons'}"""
    low = text.lower()
    lines = [ln for ln in text.splitlines() if ln.strip()]
    offsets, pos = [], 0
    for ln in lines:
        offsets.append(pos)
        pos += len(ln) + 1
    dated = _dated_flags(lines)
    rows = rows2 = cr_dr_rows = 0
    first_row_at = None
    for i, ln in enumerate(lines):
        if not dated[i]:
            continue
        amounts = AMOUNT_RX.findall(ln)
        if _is_row(ln, amounts):
            n = len(amounts)
            if CR_DR_END_RX.search(ln):
                cr_dr_rows += 1
        else:
            # many banks print the date and the narration on one line and the amounts on the next
            n, j = 0, i + 1
            while j < len(lines) and j <= i + 3 and not dated[j]:
                a = AMOUNT_RX.findall(lines[j])
                if a and len(lines[j]) <= 90 and _is_row(lines[j], a):
                    n += len(a)
                    if CR_DR_END_RX.search(lines[j]):
                        cr_dr_rows += 1
                j += 1
            n += len(amounts) if n else 0
        if n:
            rows += 1
            if first_row_at is None:
                first_row_at = offsets[i]
            if n >= 2:
                rows2 += 1
    head_end = 3000 if first_row_at is None else min(first_row_at, 3000)
    head = text[:head_end]

    reasons, score = [], 0.0
    title = next((t for t in TITLES if t in low[:8000]), '')
    if title:
        score += 3
        reasons.append(f'title "{title}"')
    cols = sorted({w for w in COLUMNS if w in low})
    score += min(len(cols), 6)
    if cols:
        reasons.append('columns: ' + ', '.join(cols[:6]))
    acct = sorted({w for w in ACCOUNT_WORDS if (re.search(rf'\b{re.escape(w)}\b', low) if len(w) <= 4 else w in low)})
    score += min(len(acct), 5)

    tx_words = sorted({' '.join(m.group(0).lower().split()) for m in BANKING_WORDS_RX.finditer(low)})
    has_banking_signals = len(tx_words) >= 1 or cr_dr_rows >= 3
    if tx_words:
        score += min(len(tx_words), 4)
        reasons.append('banking terms: ' + ', '.join(tx_words[:5]))
    if cr_dr_rows >= 3:
        score += 2

    whole = text if len(text) <= 16000 else text[:12000] + '\n' + text[-4000:]
    bank, how = find_bank(head, whole, path or filename)
    if bank:
        score += 2
        reasons.append(f'bank: {bank} ({how})')
    if how.startswith('IFSC'):
        score += 2
    score += min(rows, 20) * 0.5 + min(rows2, 20) * 0.25
    if rows:
        reasons.append(f'{rows} rows with a date and an amount')

    notes = []
    for label, rx, penalty, scope in (
        [(l, r, p, head) for l, r, p in NOT_STATEMENT] +
        [(l, r, p, text[:8000]) for l, r, p in DOC_NOT_STATEMENT]
    ):
        if rx.search(scope) and label not in notes:
            score -= penalty
            notes.append(label)
    if OWN_REPORT_NAME.search(filename) and not any('Statement Analyser' in n for n in notes):
        score -= 20
        notes.append('made by the Statement Analyser')
    if notes:
        reasons.append('looks like ' + ', '.join(notes))

    raw_card = {' '.join(m.group(0).lower().split()) for m in CARD_RX.finditer(low[:6000])}
    canon_card = {'card number' if h == 'card no' else h for h in raw_card}
    kind = 'card' if (raw_card & _STRONG_CARD_TERMS) or len(canon_card) >= 2 else 'bank'
    if kind == 'card':
        reasons.append('credit card statement')

    # Dated, amount-bearing rows alone prove little (invoice lists, beneficiary lists, reports have them too). A statement also
    # shows a sign of being one: a bank's name, a title, balance/debit/credit columns, or a strong table of banking transactions.
    statement_columns = ('balance' in cols and any(w in cols for w in MONEY_COLUMNS)) or \
        ('debit' in cols and 'credit' in cols) or ('withdrawal' in cols and 'deposit' in cols)
    strong_rows = rows >= 8 and rows2 >= 0.5 * rows
    if strong_rows and has_banking_signals and not notes and not title and not statement_columns:
        score += 6                                                     # compensates for a missing header when transaction rows are unmistakably banking
    sure_gate = bool(bank or title or statement_columns or (strong_rows and has_banking_signals))
    likely_gate = bool(sure_gate or len(acct) >= 2 or 'balance' in cols or (rows >= 3 and has_banking_signals))

    if any('Statement Analyser' in n for n in notes):
        verdict = 'not'                                                # the Statement Analyser's own reports are never statements
    elif sure_gate and score >= 15 and rows >= 5 and (len(cols) >= 2 or strong_rows or (rows >= 6 and has_banking_signals)):
        verdict = 'statement'
    elif likely_gate and ((score >= 9 and rows >= 2) or (title and len(acct) >= 2 and rows >= 1 and score >= 8) or
                          (score >= 12 and rows >= 1)):
        verdict = 'likely'
    else:
        verdict = 'not'
    return {'verdict': verdict, 'kind': kind, 'bank': bank, 'bank_how': how, 'score': round(score, 1), 'rows': rows,
            'reasons': reasons}


# ---------------------------------------------------------------------------------------------------------- reading files
def _lp(path):
    """Very long Windows paths need the \\\\?\\ prefix."""
    if os.name == 'nt':
        p = os.path.abspath(path)
        if len(p) >= 240 and not p.startswith('\\\\?\\'):
            return '\\\\?\\UNC\\' + p[2:] if p.startswith('\\\\') else '\\\\?\\' + p
    return path


def _rows_from_words(words, tol=3.0):
    """Words with positions -> text lines, one per visual row. A table cell on its own line would hide the row."""
    items = sorted(((w[1] + w[3]) / 2.0, w[0], w[4]) for w in words if str(w[4]).strip())
    rows, cur, cur_y = [], [], None
    for yc, x, t in items:
        if cur_y is None or abs(yc - cur_y) <= tol:
            cur.append((x, t))
            cur_y = yc if cur_y is None else (cur_y + yc) / 2.0
        else:
            rows.append(cur)
            cur, cur_y = [(x, t)], yc
    if cur:
        rows.append(cur)
    return [' '.join(t for _x, t in sorted(r)) for r in rows]


def _pymupdf():
    """The PyMuPDF module, or None. 'pymupdf' is its current name; 'fitz' is the old one (newer releases warn about it)."""
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        pass
    try:
        import fitz
        return fitz
    except ImportError:
        return None


def _read_pdf(path, pages):
    info = {'pages': 0}
    fitz = _pymupdf()
    if fitz is not None:
        doc = fitz.open(_lp(path))
        try:
            if doc.needs_pass and not doc.authenticate(''):
                raise Locked()
            info['pages'] = doc.page_count
            out = []
            for i in range(min(pages, doc.page_count)):
                out.extend(_rows_from_words(doc[i].get_text('words')))
            return '\n'.join(out), info
        finally:
            doc.close()
    try:
        import pdfplumber
        try:
            with pdfplumber.open(_lp(path)) as pdf:
                info['pages'] = len(pdf.pages)
                return '\n'.join((pg.extract_text() or '') for pg in pdf.pages[:pages]), info
        except Exception as exc:
            if exc.__class__.__name__ in ('PDFPasswordIncorrect', 'PdfminerException') or 'password' in str(exc).lower():
                try:
                    from pypdf import PdfReader
                    rd = PdfReader(_lp(path))
                    if rd.is_encrypted and not rd.decrypt(''):
                        raise Locked()
                except Locked:
                    raise
                except Exception:
                    pass
            raise
    except ImportError:
        pass
    from pypdf import PdfReader
    rd = PdfReader(_lp(path))
    if rd.is_encrypted and not rd.decrypt(''):
        raise Locked()
    info['pages'] = len(rd.pages)
    return '\n'.join((pg.extract_text() or '') for pg in rd.pages[:pages]), info


def _read_docx(path):
    import docx
    d = docx.Document(_lp(path))
    out = [p.text for p in d.paragraphs[:400] if p.text.strip()]
    n = 0
    for table in d.tables[:6]:
        for row in table.rows:
            cells = []
            for c in row.cells:
                t = ' '.join(c.text.split())
                if t and (not cells or cells[-1] != t):               # a merged cell repeats itself
                    cells.append(t)
            if cells:
                out.append(' '.join(cells))
                n += 1
                if n >= 400:
                    break
        if n >= 400:
            break
    return '\n'.join(out), {'pages': len(d.tables)}


def _read_text(path):
    with open(_lp(path), 'rb') as f:
        raw = f.read(400_000)
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
        text = raw.decode('utf-16', errors='replace')
    else:
        try:
            text = raw.decode('utf-8-sig')
        except UnicodeDecodeError as exc:
            # a 400KB read slice can split a multi-byte UTF-8 character at the end
            if exc.reason == 'unexpected end of data' and exc.start >= len(raw) - 3 and len(raw) == 400_000:
                text = raw[:exc.start].decode('utf-8-sig', errors='replace')
            else:
                text = raw.decode('cp1252', errors='replace')
    if re.search(r'<\s*(table|tr|html)\b', text[:5000], re.I):
        text = re.sub(r'(?is)<(script|style).*?</\1>', ' ', text)
        text = re.sub(r'(?i)</t[dh]\s*>', ' ', text)
        text = re.sub(r'(?i)</tr\s*>|<br\s*/?>|</p\s*>|</div\s*>', '\n', text)
        text = html.unescape(re.sub(r'<[^>]+>', ' ', text))
        text = '\n'.join(' '.join(ln.split()) for ln in text.splitlines())
    return text, {'pages': 1}


def read_text(path, pages=4):
    """-> (text, info). Raises Locked for a password-protected file."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.pdf':
        return _read_pdf(path, pages)
    if ext == '.docx':
        return _read_docx(path)
    return _read_text(path)


def available_readers():
    """Which libraries are installed: {'PDF': name or None, 'Word': name or None}."""
    def has(mod):
        try:
            __import__(mod)
            return True
        except Exception:
            return False
    pdf = 'PyMuPDF' if _pymupdf() is not None else next((n for m, n in (('pdfplumber', 'pdfplumber'), ('pypdf', 'pypdf')) if has(m)), None)
    word = 'python-docx' if has('docx') else None
    return {'PDF': pdf, 'Word': word}


# ---------------------------------------------------------------------------------------------------------- one file
def _sha256(path):
    h = hashlib.sha256()
    with open(_lp(path), 'rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block)
    return h.hexdigest()


def _test_hook(name):
    """For the tests only: STATEMENT_FINDER_TEST_HOOK='hang:<file name>' or 'crash:<file name>'."""
    hook = os.environ.get('STATEMENT_FINDER_TEST_HOOK', '')
    if ':' in hook:
        what, target = hook.split(':', 1)
        if target == name:
            if what == 'hang':
                time.sleep(3600)
            elif what == 'crash':
                time.sleep(0.05)
                os._exit(3)


def classify_file(path, opts):
    """Reads one file and says what it is. Runs in a worker process.
    verdict: statement / likely / not / locked / scanned / unreadable / too_big"""
    name = os.path.basename(path)
    ext = os.path.splitext(name)[1].lower()
    out = {'path': path, 'name': name, 'ext': ext, 'verdict': 'not', 'kind': 'bank', 'bank': '', 'bank_how': '', 'score': 0.0,
           'rows': 0, 'reasons': [], 'sha': '', 'pages': 0, 'size': 0, 'mtime': 0.0, 'error': ''}
    try:
        st = os.stat(_lp(path))
        out['size'], out['mtime'] = st.st_size, st.st_mtime
    except OSError as exc:
        out.update(verdict='unreadable', error=f'could not open ({exc.__class__.__name__}; a Google Drive file not downloaded yet?)')
        return out
    _test_hook(name)
    if st.st_size == 0:
        out.update(verdict='unreadable', error='empty file')
        return out
    if st.st_size > int(opts.get('max_mb', 300)) * 1048576:
        out.update(verdict='too_big', error=f'bigger than {opts.get("max_mb", 300)} MB, not opened')
        return out
    try:
        text, info = read_text(path, int(opts.get('pages', 4)))
    except Locked:
        out.update(verdict='locked', reasons=['password-protected'])
        return out
    except Exception as exc:
        if exc.__class__.__name__ == 'PackageNotFoundError' or 'File is not a zip file' in str(exc):
            msg = 'not a valid Word file (an old .doc renamed, or a damaged file)'
        else:
            msg = f'{exc.__class__.__name__}: {exc}'
        out.update(verdict='unreadable', error=' '.join(msg.split())[:200])
        return out
    out['pages'] = info.get('pages', 0)
    if ext == '.pdf' and len(text.strip()) < 40:
        out.update(verdict='scanned', reasons=['no text in the PDF (a scan or a picture)'])
        return out
    res = analyse_text(text, name, path)
    out.update(res)
    if res['verdict'] in ('statement', 'likely'):
        try:
            out['sha'] = _sha256(path)
        except OSError as exc:
            out['error'] = f'could not read for the copy ({exc.__class__.__name__})'
    return out


def _worker_main(task_q, result_q, opts):
    """A worker process: takes file paths from the queue until it is told to stop."""
    os.environ.setdefault('PYTHONDONTWRITEBYTECODE', '1')
    try:
        _pymupdf().TOOLS.mupdf_display_errors(False)
    except Exception:
        pass
    while True:
        item = task_q.get()
        if item is None:
            return
        idx, path = item
        result_q.put(('start', os.getpid(), idx))
        try:
            res = classify_file(path, opts)
        except Exception as exc:                                      # never let one file stop a worker
            res = {'path': path, 'name': os.path.basename(path), 'ext': os.path.splitext(path)[1].lower(), 'verdict': 'unreadable',
                   'kind': 'bank', 'bank': '', 'bank_how': '', 'score': 0.0, 'rows': 0, 'reasons': [], 'sha': '', 'pages': 0,
                   'size': 0, 'mtime': 0.0, 'error': f'{exc.__class__.__name__}: {exc}'[:200]}
        result_q.put(('done', os.getpid(), idx, res))


# ---------------------------------------------------------------------------------------------------------- the folder walk
def _inside(path, folder):
    a, b = os.path.normcase(os.path.abspath(path)), os.path.normcase(os.path.abspath(folder))
    return a == b or a.startswith(b.rstrip('\\/') + os.sep)


def walk_files(source, exts, skip_system, exclude, stop, problems):
    """Yields (path, size, mtime) of every file under `source` with one of `exts`. Never follows links."""
    exts = tuple(exts)
    stack = [source]
    while stack:
        if stop.is_set():
            return
        folder = stack.pop()
        try:
            with os.scandir(_lp(folder)) as it:
                entries = list(it)
        except OSError as exc:
            problems.append(f'{folder}: {exc.__class__.__name__}')
            continue
        for e in entries:
            try:
                name = e.name
                path = os.path.join(folder, name)
                if e.is_dir(follow_symlinks=False):
                    if skip_system and name.lower() in SYSTEM_FOLDERS:
                        continue
                    if any(_inside(path, x) for x in exclude):
                        continue
                    stack.append(path)
                elif e.is_file(follow_symlinks=False):
                    if name.startswith('~$') or not name.lower().endswith(exts):
                        continue
                    st = e.stat(follow_symlinks=False)
                    yield path, st.st_size, st.st_mtime
            except OSError as exc:
                problems.append(f'{os.path.join(folder, e.name)}: {exc.__class__.__name__}')


# ---------------------------------------------------------------------------------------------------------- the pool
def classify_all(files, opts, on_result, on_progress, stop):
    """Classifies every file in worker processes, several at the same time. A worker that runs past the time limit or
    dies is replaced, and its file is reported as not read; the rest carry on."""
    ctx = mp.get_context('spawn')
    n = max(1, min(int(opts.get('workers') or (os.cpu_count() or 4)), 64, len(files)))
    timeout = max(5, int(opts.get('timeout', 60)))
    task_q, result_q = ctx.Queue(), ctx.Queue()
    procs, current = {}, {}

    def spawn():
        p = ctx.Process(target=_worker_main, args=(task_q, result_q, opts), daemon=True)
        p.start()
        procs[p.pid] = p

    def lost(idx, why):
        path = files[idx][0]
        return {'path': path, 'name': os.path.basename(path), 'ext': os.path.splitext(path)[1].lower(), 'verdict': 'unreadable',
                'kind': 'bank', 'bank': '', 'bank_how': '', 'score': 0.0, 'rows': 0, 'reasons': [], 'sha': '', 'pages': 0,
                'size': files[idx][1], 'mtime': files[idx][2], 'error': why}

    for _ in range(n):
        spawn()
    next_i = done = next_emit = 0
    handed, finished = set(), set()
    pending = {}
    window = n * 3
    quiet_since = None

    def flush_ready():
        nonlocal next_emit
        while next_emit in pending:
            on_result(pending.pop(next_emit))
            next_emit += 1

    try:
        while done < len(files) and not stop.is_set():
            while len(handed) - len(finished) < window and next_i < len(files):
                task_q.put((next_i, files[next_i][0]))
                handed.add(next_i)
                next_i += 1
            got_any = False
            while True:
                try:
                    msg = result_q.get(timeout=0.0 if got_any else 0.2)
                except queue.Empty:
                    break
                got_any = True
                quiet_since = None
                if msg[0] == 'start':
                    current[msg[1]] = (msg[2], time.time())
                else:
                    _kind, pid, idx, res = msg
                    current.pop(pid, None)
                    if idx not in finished:
                        finished.add(idx)
                        done += 1
                        pending[idx] = res
                        flush_ready()
                        on_progress(done, len(files), res.get('name', ''))
            now = time.time()
            for pid, p in list(procs.items()):
                cur = current.get(pid)
                why = ''
                if cur and now - cur[1] > timeout:
                    why = f'stopped after {timeout} seconds (the file may be damaged)'
                elif not p.is_alive():
                    why = 'the reader crashed on this file' if cur else ''
                if cur and why:
                    p.terminate()
                    p.join(2)
                    del procs[pid]
                    current.pop(pid, None)
                    if cur[0] not in finished:
                        finished.add(cur[0])
                        done += 1
                        pending[cur[0]] = lost(cur[0], why)
                        flush_ready()
                        on_progress(done, len(files), os.path.basename(files[cur[0]][0]))
                    spawn()
                elif not p.is_alive():                                 # died between files: replace it
                    del procs[pid]
                    spawn()
            # a file handed to a worker that died before it reported anything would never finish: give up on it after a while
            if next_i >= len(files) and not current and not got_any:
                quiet_since = quiet_since or now
                if now - quiet_since > 6:
                    for idx in sorted(handed - finished):
                        finished.add(idx)
                        done += 1
                        pending[idx] = lost(idx, 'the reader stopped before it could read this file')
                    flush_ready()
                    break
        for idx in sorted(pending):
            on_result(pending[idx])
        pending.clear()
    finally:
        if stop.is_set():
            for p in procs.values():
                p.terminate()
        else:
            for _ in procs:
                task_q.put(None)
        for p in procs.values():
            p.join(3)
            if p.is_alive():
                p.terminate()
        for q in (task_q, result_q):
            q.cancel_join_thread()
            q.close()


# ---------------------------------------------------------------------------------------------------------- copying
def safe_name(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')
    return name[:80] or '_'


def safe_filename(name):
    """Sanitizes a file name and keeps its extension intact when truncating to 80 characters."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')
    if len(cleaned) <= 80:
        return cleaned or '_'
    stem, ext = os.path.splitext(cleaned)
    if not ext or len(ext) > 15:
        return cleaned[:80]
    keep = max(1, 80 - len(ext))
    return (stem[:keep].rstrip(' .') or '_') + ext


def target_folder(res):
    """The folder (relative to the destination) a statement belongs in."""
    v = res['verdict']
    base = safe_name(res['bank']) if res['bank'] else UNKNOWN_BANK
    if res.get('kind') == 'card':
        base = os.path.join(CARD_FOLDER, base if res['bank'] else UNKNOWN_BANK)
    if v == 'likely':
        base = os.path.join(base, NOT_SURE_FOLDER)
    return base


class Copier:
    """Puts statements in their bank folder. Identical files (same bytes) are copied once."""

    def __init__(self, dest, dry_run):
        self.dest, self.dry = dest, dry_run
        self.by_hash = {}
        self.reserved = {}

    def place(self, res):
        """-> (action, target path). action: copied / would copy / duplicate / already there / failed: <why>"""
        src = res['path']
        folder = os.path.join(self.dest, target_folder(res))
        target = os.path.join(folder, safe_filename(res['name']))
        base, ext = os.path.splitext(target)
        try:
            sha = res['sha'] or _sha256(src)
            if sha and sha in self.by_hash:
                return 'duplicate', self.by_hash[sha]
            k = 1
            while True:
                cand = target if k == 1 else f'{base} ({k}){ext}'
                key = os.path.normcase(_lp(cand))
                existing_sha = self.reserved.get(key)
                if existing_sha is None and os.path.exists(_lp(cand)):
                    existing_sha = _sha256(cand)
                    self.reserved[key] = existing_sha
                if existing_sha is None:
                    target = cand
                    self.reserved[key] = sha
                    break
                if existing_sha == sha:
                    self.by_hash[sha] = cand
                    return 'already there', cand
                k += 1
            self.by_hash[sha] = target
            if self.dry:
                return 'would copy', target
            os.makedirs(_lp(folder), exist_ok=True)
            shutil.copy2(_lp(src), _lp(target))
            return 'copied', target
        except OSError as exc:
            return f'failed: {exc.__class__.__name__}', ''


# ---------------------------------------------------------------------------------------------------------- the run
REPORT_COLUMNS = ['Result', 'Bank', 'Certainty', 'Type', 'Score', 'Rows with date and amount', 'File', 'Original location',
                  'Copied to', 'Why', 'Size (KB)', 'File date']


def run(source, dest, opts, emit, stop):
    """Looks through `source`, copies the statements into `dest`, writes a report. `emit(event)` gets progress events:
        ('phase', text)  ('walk', files so far)  ('total', n)  ('progress', done, total, name)  ('found', record)  ('done', summary)
    `stop` is a threading.Event; set it to stop early. Returns the summary dict."""
    o = dict(DEFAULTS)
    o.update(opts or {})
    source, dest = os.path.abspath(source), os.path.abspath(dest)
    t0 = time.time()
    s = {'source': source, 'dest': dest, 'dry_run': bool(o['dry_run']), 'files': 0, 'read': 0, 'statements': 0, 'not_sure': 0,
         'cards': 0, 'locked': 0, 'scanned': 0, 'not_statements': 0, 'unreadable': 0, 'copied': 0, 'duplicates': 0,
         'by_bank': {}, 'seconds': 0.0, 'stopped': False, 'report': '', 'folder_problems': []}
    exts = tuple(e for k in o['kinds'] for e in FILE_TYPES.get(k, ()))
    exclude = [dest] if _inside(dest, source) else []

    emit(('phase', 'Looking for files ...'))
    files, problems = [], []
    for item in walk_files(source, exts, o['skip_system'], exclude, stop, problems):
        files.append(item)
        if len(files) % 400 == 0:
            emit(('walk', len(files)))
    emit(('walk', len(files)))
    s['files'] = len(files)
    s['folder_problems'] = problems[:200]
    files.sort(key=lambda f: -f[1])                                   # the biggest first, so the run does not end waiting for one
    emit(('total', len(files)))

    records = []
    copier = Copier(dest, o['dry_run'])

    def on_result(res):
        s['read'] += 1
        v = res['verdict']
        wanted = (v == 'statement' or (v == 'likely' and o['include_not_sure']))
        if v == 'statement':
            s['statements'] += 1
        elif v == 'likely':
            s['not_sure'] += 1
        elif v == 'locked':
            s['locked'] += 1
        elif v == 'scanned':
            s['scanned'] += 1
        elif v in ('unreadable', 'too_big'):
            s['unreadable'] += 1
        else:
            s['not_statements'] += 1
        if v in ('statement', 'likely') and res.get('kind') == 'card':
            s['cards'] += 1
        if v in ('statement', 'likely'):
            b = s['by_bank'].setdefault(res['bank'] or 'Bank not identified', {'sure': 0, 'not_sure': 0})
            b['sure' if v == 'statement' else 'not_sure'] += 1
        res['action'], res['target'] = '', ''
        if wanted:
            res['action'], res['target'] = copier.place(res)
            if res['action'] in ('copied', 'would copy'):
                s['copied'] += 1
            elif res['action'] == 'duplicate':
                s['duplicates'] += 1
        if wanted or v in ('unreadable', 'too_big'):
            records.append(res)
            emit(('found', res))

    if files and not stop.is_set():
        emit(('phase', 'Reading the files ...'))
        classify_all(files, o, on_result, lambda d, t, n: emit(('progress', d, t, n)), stop)
    s['stopped'] = stop.is_set()
    s['seconds'] = round(time.time() - t0, 1)

    try:
        os.makedirs(_lp(dest), exist_ok=True)
        s['report'] = os.path.join(dest, 'Statement Finder report ' + time.strftime('%Y-%m-%d_%H%M%S') + '.csv')
        write_report(s['report'], records, source)
    except OSError as exc:
        s['report'] = ''
        s['folder_problems'].append(f'report not written: {exc}')
    emit(('done', s))
    return s


def write_report(path, records, source):
    pretty = {'statement': 'Sure', 'likely': 'Not sure', 'unreadable': '-', 'too_big': '-'}
    with open(_lp(path), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(REPORT_COLUMNS)
        for r in sorted(records, key=lambda r: (r['bank'] or '~', r['name'].lower())):
            v = r['verdict']
            if v in ('statement', 'likely'):
                result = {'copied': 'Copied', 'would copy': 'Would copy (scan only)', 'duplicate': 'Same file as one already found',
                          'already there': 'Already in the destination'}.get(r['action'], r['action'] or 'Not copied')
                bank = r['bank'] or 'Bank not identified'
            else:
                result = {'unreadable': 'Could not be read', 'too_big': 'Too big, not opened'}.get(v, v)
                bank = ''
            why = r['error'] or '; '.join(r['reasons'])
            w.writerow([result, bank, pretty.get(v, '-'), 'Credit card' if r.get('kind') == 'card' and v in ('statement', 'likely') else
                        ('Bank account' if v in ('statement', 'likely') else ''), r['score'] or '', r['rows'] or '', r['name'],
                        os.path.dirname(os.path.relpath(r['path'], source)) or '.', r['target'], why, round(r['size'] / 1024),
                        time.strftime('%d-%m-%Y', time.localtime(r['mtime'])) if r['mtime'] else ''])
