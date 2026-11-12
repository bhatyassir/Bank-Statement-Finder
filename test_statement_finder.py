"""Tests for Statement Finder. Only made-up statements are used here (no real client data).

    python test_statement_finder.py
"""
import csv
import hashlib
import io
import multiprocessing as mp
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import finder_core as core

# ------------------------------------------------------------------------------------------------ made-up statements
HEADINGS = {
    'hdfc': 'HDFC BANK LIMITED\nStatement of account\nAccount No : 50100012345678   IFSC : HDFC0001234\nBranch : MAIN ROAD',
    'sbi': 'State Bank of India\nAccount Statement\nAccount Number : 123456789012   Branch : MAIN BRANCH',
    'jkb': 'JAMMU AND KASHMIR BANK LTD\nMAIN MARKET\nIFSC Code : JAKA0SAMPLE , MICR Code : 190000123\nSTATEMENT OF ACCOUNT FOR THE PERIOD 01-04-2025 to 31-03-2026',
    'nobank': 'Statement of account\nAccount Number : 123456789\nBranch : ROAD',
}
COLUMNS_LINE = 'Date Narration Chq./Ref.No. Value Dt Withdrawal Amt. Deposit Amt. Closing Balance'


def rows(n=12, narration='UPI-SHOP-4455'):
    out, bal = [], 100000.0
    for i in range(n):
        amt = 100.0 + i * 7.25
        bal -= amt
        out.append(f'{(i % 27) + 1:02d}/04/2025 {narration} 00001234{i:02d} {(i % 27) + 1:02d}/04/2025 {amt:,.2f} {bal:,.2f}')
    return out


def statement(kind='hdfc', n=12, cols=True, narration='UPI-SHOP-4455'):
    return '\n'.join([HEADINGS[kind]] + ([COLUMNS_LINE] if cols else []) + rows(n, narration))


INVOICE = '\n'.join(['TAX INVOICE', 'Invoice No : 4455  Date : 12-05-2025', 'GSTIN of supplier : 01ABCDE1234F1Z5', 'Item Qty Rate Amount'] +
                    [f'Spare part {i} 2 {100 + i}.00 {200 + 2 * i}.00' for i in range(8)] + ['Total 2,345.00', 'Due date 30-05-2025 1,000.00'])
DAYBOOK = '\n'.join(['Day Book', 'Sample Traders', 'Date Particulars Vch Type Debit Credit'] +
                    [f'{(i % 27) + 1:02d}-Apr-2025 Cash Receipt {i} {500 + i}.00 {500 + i}.00' for i in range(15)])
CARD = '\n'.join(['HDFC BANK Credit Card Statement', 'Card No : 4567 XXXX XXXX 1234', 'Total Amount Due 12,345.00', 'Minimum Amount Due 617.00',
                  'Payment Due Date 05/05/2025', 'Credit Limit 100000.00', COLUMNS_LINE] + rows(10))


def md5(path):
    with open(path, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()


def tree_state(folder):
    out = {}
    for d, _dirs, files in os.walk(folder):
        for f in files:
            out[os.path.relpath(os.path.join(d, f), folder)] = md5(os.path.join(d, f))
    return out


def write(path, text, encoding='utf-8'):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding=encoding) as f:
        f.write(text)
    return path


def _minimal_pdf_bytes(lines):
    """Builds a minimal valid single-page PDF with text lines when PyMuPDF is not installed."""
    def esc(s):
        return s.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')
    ops = ['BT', '/F1 8 Tf']
    y = 750
    for ln in lines:
        ops.append(f'1 0 0 1 40 {y} Tm ({esc(ln)}) Tj')
        y -= 12
    ops.append('ET')
    stream = '\n'.join(ops).encode('latin-1', errors='replace')
    objs = [
        b'1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
        b'2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
        b'3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n',
        f'4 0 obj\n<< /Length {len(stream)} >>\nstream\n'.encode('ascii') + stream + b'\nendstream\nendobj\n',
        b'5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n',
    ]
    buf = io.BytesIO()
    buf.write(b'%PDF-1.4\n')
    offsets = []
    for o in objs:
        offsets.append(buf.tell())
        buf.write(o)
    xref_pos = buf.tell()
    buf.write(f'xref\n0 {len(objs) + 1}\n0000000000 65535 f \n'.encode('ascii'))
    for off in offsets:
        buf.write(f'{off:010d} 00000 n \n'.encode('ascii'))
    buf.write(f'trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n'.encode('ascii'))
    return buf.getvalue()


def make_pdf(path, lines, password=None, blank=False):
    """A PDF whose words are placed at columns, so each statement line is one visual row."""
    fitz = core._pymupdf()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if fitz is not None:
        doc = fitz.open()
        page = doc.new_page()
        if blank:
            page.draw_rect(fitz.Rect(50, 50, 300, 200), color=(0, 0, 0))     # a picture only: no text layer
        y = 50
        for ln in lines:
            x = 40
            for word in ln.split(' '):
                page.insert_text((x, y), word, fontsize=8)
                x += 6.2 * len(word) + 5
            y += 12
            if y > 800:
                page = doc.new_page()
                y = 50
        if password:
            doc.save(path, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=password, owner_pw=password + 'x')
        else:
            doc.save(path)
        doc.close()
        return path
    # Fallback when PyMuPDF is not installed in the test environment
    raw = _minimal_pdf_bytes([] if blank else lines)
    if password:
        from pypdf import PdfReader, PdfWriter
        rd = PdfReader(io.BytesIO(raw))
        wr = PdfWriter()
        for pg in rd.pages:
            wr.add_page(pg)
        wr.encrypt(password)
        with open(path, 'wb') as f:
            wr.write(f)
    else:
        with open(path, 'wb') as f:
            f.write(raw)
    return path


# ------------------------------------------------------------------------------------------------ the text test
class Analyse(unittest.TestCase):
    def test_a_statement_with_its_bank(self):
        r = core.analyse_text(statement('hdfc'))
        self.assertEqual((r['verdict'], r['bank'], r['kind']), ('statement', 'HDFC Bank', 'bank'))
        self.assertTrue(r['bank_how'].startswith('IFSC HDFC0001234'))
        self.assertEqual(r['rows'], 12)

    def test_bank_by_name_and_not_confused_with_bank_of_india(self):
        r = core.analyse_text(statement('sbi'))
        self.assertEqual((r['verdict'], r['bank']), ('statement', 'State Bank of India'))
        self.assertEqual(core.find_bank('Bank of India\nAccount Statement')[0], 'Bank of India')
        self.assertEqual(core.find_bank('Reserve Bank of India notice')[0], '')

    def test_the_bank_is_the_one_in_the_heading_not_in_the_narrations(self):
        text = statement('jkb', narration='NEFT/HDFC0001234/ICICI BANK LTD/STATE BANK OF INDIA/SBIN0000123')
        r = core.analyse_text(text)
        self.assertEqual((r['verdict'], r['bank']), ('statement', 'J&K Bank'))

    def test_grameen_is_not_the_j_and_k_bank(self):
        self.assertEqual(core.find_bank('IFSC : JAKA0GRAMEN\nBranch')[0], 'J&K Grameen Bank')
        self.assertEqual(core.find_bank('Jammu and Kashmir Grameen Bank\nStatement')[0], 'J&K Grameen Bank')

    def test_a_cooperative_bank_that_is_not_in_the_list(self):
        name, _how = core.find_bank('THE SAMPLE URBAN CO-OPERATIVE BANK LTD\nStatement of account')
        self.assertTrue(name.endswith('Co-operative Bank'), name)
        self.assertIn('Sample', name.title())

    def test_a_statement_from_an_unknown_bank_is_still_a_statement(self):
        r = core.analyse_text(statement('nobank'))
        self.assertEqual((r['verdict'], r['bank']), ('statement', ''))

    def test_headerless_statement_with_banking_terms_is_recognised(self):
        # No header, no title, no column names, no bank name — just transaction rows with UPI / NEFT and Cr balances
        raw_rows = [f'{(i % 27) + 1:02d}-04-2025 UPI/DR/503122/SHOP   {150 + i}.00   {95000 - i * 10}.00Cr' for i in range(10)]
        r = core.analyse_text('\n'.join(raw_rows))
        self.assertEqual(r['verdict'], 'statement')
        self.assertEqual(r['rows'], 10)

    def test_bank_found_from_footer_website_or_file_path_when_header_is_blank(self):
        raw_rows = '\n'.join(rows(10, narration='BY TRANSFER-INB/UPI'))
        r_sbi = core.analyse_text(raw_rows)
        self.assertEqual((r_sbi['verdict'], r_sbi['bank']), ('statement', 'State Bank of India'))

        with_footer = '\n'.join(rows(10)) + '\nGenerated from https://www.jkbank.com on 12-05-2025'
        r_jkb = core.analyse_text(with_footer)
        self.assertEqual((r_jkb['verdict'], r_jkb['bank']), ('statement', 'J&K Bank'))

        # Bank from file name or parent folder name when text has no bank clues at all
        plain_tx = '\n'.join(rows(10, narration='UPI-LOCAL-SHOP'))
        r_file = core.analyse_text(plain_tx, 'JKB_April_2025.rpt', '/data/JKB_April_2025.rpt')
        self.assertEqual((r_file['verdict'], r_file['bank']), ('statement', 'J&K Bank'))
        r_folder = core.analyse_text(plain_tx, 'acct_01.nfs', '/data/clients/HDFC Bank/acct_01.nfs')
        self.assertEqual((r_folder['verdict'], r_folder['bank']), ('statement', 'HDFC Bank'))

    def test_things_that_are_not_statements(self):
        for name, text in (('invoice', INVOICE), ('day book', DAYBOOK), ('nothing', 'Hello\nThis is a letter about a meeting on 12-05-2025.')):
            self.assertEqual(core.analyse_text(text)['verdict'], 'not', name)

    def test_a_credit_card_statement_is_told_apart(self):
        r = core.analyse_text(CARD)
        self.assertEqual((r['verdict'], r['kind'], r['bank']), ('statement', 'card', 'HDFC Bank'))

    def test_the_statement_analysers_own_reports_are_not_statements(self):
        report = 'Audit Report\nDashboard\nMonthly Movement\n' + '\n'.join(rows(20))
        self.assertEqual(core.analyse_text(report)['verdict'], 'not')
        self.assertEqual(core.analyse_text(statement('hdfc'), 'Client_Audit_Report_2026-01-01_1200.pdf')['verdict'], 'not')
        self.assertEqual(core.analyse_text(statement('hdfc'), 'Client_Audit_Summary.pdf')['verdict'], 'not')

    def test_rows_over_two_lines_are_counted(self):
        # date and narration on one line, the amounts on the next: how some banks print
        text = HEADINGS['jkb'] + '\n' + '\n'.join(f'{(i % 27) + 1:02d}-04-2025 UPI/HDFC/5031/DR/SHOP\n{100 + i}.00 {99000 - i}.50Cr' for i in range(10))
        r = core.analyse_text(text)
        self.assertEqual((r['rows'], r['verdict']), (10, 'statement'))

    def test_a_date_is_not_an_amount(self):
        self.assertEqual(core.AMOUNT_RX.findall('01.04.2025 12.05.2025 5,000.00 1,23,456.78 0.50'), ['5,000.00', '1,23,456.78', '0.50'])

    def test_a_short_statement_is_not_sure_and_a_scrap_is_not_a_statement(self):
        short = '\n'.join(['Statement of account', 'Account No 12345678', 'IFSC SBIN0001234', 'Branch MAIN'] + rows(3))
        self.assertEqual(core.analyse_text(short)['verdict'], 'likely')
        self.assertEqual(core.analyse_text('\n'.join(rows(1)))['verdict'], 'not')

    def test_tax_26as_broker_equity_and_gst_schedules_are_rejected(self):
        form26as = '\n'.join([
            'Annual Tax Statement',
            'Permanent Account Number (PAN): ABCDE1234F   Financial Year: 2024-25   Assessment Year: 2025-26',
            'Name of Assessee: TEST USER   Address: SRINAGAR JAMMU & KASHMIR',
            'TAN of Deductor: DELA12345B   Details of Tax Deducted at Source',
            'Date of Deposit   Amount Paid   Tax Deducted   TDS Deposited',
        ] + [f'{(i % 27) + 1:02d}-Apr-2025   50,000.00   500.00   500.00' for i in range(25)] + [
            'Downloaded from https://tinpan.proteantech.in',
        ])
        self.assertEqual(core.analyse_text(form26as)['verdict'], 'not')

        equity = '\n'.join([
            'ICICI Securities Limited',
            'Equity Transaction Statement',
            'MEMBER OF NATIONAL STOCK EXCHANGE OF INDIA LTD   SEBI REGN. NO.: INZ000000000',
            'UNIQUE CLIENT CODE: 123456   TRADING CODE NO: 998877',
            'Trade Date   Scrip Name   Buy/Sell   Rate   Net Amount',
            '02-04-2025 FEDERAL BANK LTD   BUY   180.00   18,000.00',
        ] + [f'{(i % 27) + 3:02d}-04-2025 YES BANK LIMITED BUY 20.00 2,000.00' for i in range(10)])
        self.assertEqual(core.analyse_text(equity)['verdict'], 'not')

        gst_itc = '\n'.join([
            'Details Attached for Differences in ITC Claimed in GSTR-3B',
            'DRC-03 Payment Particulars   Inv Date   Taxable Value   IGST   CGST',
        ] + [f'{(i % 27) + 1:02d}-05-2025   10,000.00   1,800.00' for i in range(12)])
        self.assertEqual(core.analyse_text(gst_itc)['verdict'], 'not')

        contractor = '\n'.join([
            'BILLS VERIFIED AGAINST ORIGINAL RECEIPTS',
            'Details of Payments Received in Bank Account',
        ] + [f'{(i % 27) + 1:02d}-06-2025 BILL NO {i + 1}   1,00,000.00   95,000.00' for i in range(12)] + [
            'GSTR 7 DETAILS',
            'Taxable Value   CGST   SGST',
        ])
        self.assertEqual(core.analyse_text(contractor)['verdict'], 'not')

        hdfc_gstin = statement('hdfc') + '\nHDFC Bank GSTIN number details are available at www.hdfcbank.com'
        r_hdfc = core.analyse_text(hdfc_gstin)
        self.assertEqual((r_hdfc['verdict'], r_hdfc['bank'], r_hdfc['kind']), ('statement', 'HDFC Bank', 'bank'))

    def test_axis_footer_card_wording_vs_real_credit_card(self):
        axis_bank_stmt = '\n'.join([
            'Statement of Axis Account No :900000000000001 for the period 01-04-2025 to 31-03-2026',
            'Customer Name: TEST USER',
            'Tran Date   Particulars   Chq No   Debit   Credit   Balance',
        ] + rows(10, narration='UPI/P2A/123456/SHOP') + [
            'Never share your card\nnumber, CVV or PIN. Do not disclose your card number to anyone.',
        ])
        r_axis = core.analyse_text(axis_bank_stmt)
        self.assertEqual((r_axis['verdict'], r_axis['bank'], r_axis['kind']), ('statement', 'Axis Bank', 'bank'))

        icici_card = '\n'.join([
            'ICICI BANK',
            'Statement of account',
            'Card Account No : 4315XXXXXXXX1234   Card Number : 4315XXXXXXXX5678',
            'Date   Transaction Details   Amount   Balance',
        ] + rows(10, narration='AMAZON RETAIL'))
        r_card = core.analyse_text(icici_card)
        self.assertEqual((r_card['verdict'], r_card['bank'], r_card['kind']), ('statement', 'ICICI Bank', 'card'))

    def test_endash_and_wrapped_year_dates_and_wrapped_narrations(self):
        endash_stmt = '\n'.join([
            'ICICI Bank   Statement of Account',
            'Tran Date   Particulars   Withdrawal   Deposit   Balance',
        ] + [f'{(i % 27) + 1:02d}\u201304\u20132025  UPI/SHOP/{i}   100.00   {9000 - i * 100}.00' for i in range(8)])
        r_endash = core.analyse_text(endash_stmt)
        self.assertEqual((r_endash['verdict'], r_endash['rows']), ('statement', 8))

        wrapped_stmt = '\n'.join([
            'ICICI Bank   Detailed Statement',
            'Sl No   Tran ID   Date   Particulars   Debit   Balance',
        ] + [f'{i + 1}   TXN000{i}   {(i % 27) + 1:02d}-Sep-\n2026   UPI/SHOP   250.00   {12000 - i * 250}.00' for i in range(8)])
        r_wrapped = core.analyse_text(wrapped_stmt)
        self.assertEqual((r_wrapped['verdict'], r_wrapped['rows']), ('statement', 8))

        wrapped_narration = '\n'.join([
            'Statement of Account',
            'Date   Particulars   Debit   Credit   Balance',
            '02-04-2025   NEFT/PAYMENT/123   500.00   9,500.00',
            '50200000000001-HDFC0000001',
            '03-04-2025   UPI/P2A/456/PhonePe/   200.00   9,300.00',
            'Yes Bank',
            '04-04-2025   IMPS/789/   100.00   9,200.00',
            'BARB0000001/TEST',
            '05-04-2025   /SBI Credit   100.00   9,100.00',
        ])
        self.assertEqual(core.find_bank('Statement of Account', wrapped_narration)[0], '')

    def test_address_bank_ignored_and_coop_and_path_priority(self):
        addr_text = '\n'.join([
            'Statement of Account',
            'Customer Name : TEST',
            'Address : MAIN ROAD OPPOSITE SAMPLE CENTRAL COOPERATIVE BANK',
            'Date   Particulars   Debit   Credit   Balance',
        ] + rows(8))
        self.assertEqual(core.find_bank(addr_text)[0], '')

        coop_text = 'THE SAMPLE CENTRAL CO-OP BANK LTD\nStatement of Account'
        self.assertEqual(core.find_bank(coop_text)[0], 'Sample Central Co-operative Bank')

        two_col_coop = 'Customer Name : Mr. SAMPLE HOLDER   Sample Central Co-op Bank\nStatement of Account'
        self.assertEqual(core.find_bank(two_col_coop)[0], 'Sample Central Co-operative Bank')

        sbi_addr = 'Address : 10 SAMPLE STREET\nNEAR STATE BANK OF INDIA\nIFS Code : SBIN0000123\nAccount Statement'
        self.assertEqual(core.find_bank(sbi_addr)[0], 'State Bank of India')

        loan_fee = '\n'.join([
            'STATEMENT OF ACCOUNTS\nMintifi ID: TEST001\nDate   Narration   Debit   Credit   Balance',
            '2025-03-04 IGST 18% 3240.00 50000.00',
        ] + [f'2025-04-{i + 1:02d} Schedule EMI 5000.00 {45000 - i * 1000}.00' for i in range(8)])
        r_loan = core.analyse_text(loan_fee)
        self.assertEqual((r_loan['verdict'], r_loan['bank']), ('statement', 'Mintifi'))

        self.assertEqual(core.find_bank('', '', '/tmp/Central_Bank_of_India_May2025.pdf')[0], 'Central Bank of India')
        self.assertEqual(core.find_bank('', '', '/tmp/Reserve_Bank_of_India_Circular.pdf')[0], '')

        non_acct = '\n'.join([
            'Statement of account',
            'Microsoft specification sheet   Amount',
            '01-04-2025  Spec item A  100.00  200.00',
            '02-04-2025  Spec item B  100.00  300.00',
        ])
        self.assertEqual(core.analyse_text(non_acct)['verdict'], 'not')

    def test_comma_dates_ifsc_leaks_and_bank_fingerprints(self):
        # Comma-separated MM,DD,YYYY rows must still count as dated rows so wrapped counterparty IFSCs are skipped
        comma_stmt = '\n'.join([
            'DETAILED STATEMENT',
            'Search',
            'Account Number 000012345678(INR)',
            'Cheque number from NA to NA',
            'Transaction Date   Value Date   Cheque Number   Transaction Remarks   Withdrawal Amt   Deposit Amt   Balance',
        ] + [
            f'08,{i + 10:02d},2023  08,{i + 10:02d},2023  -  NEFT/TEST/valid/HDFC0001234  {100 + i}.00  0.00  {9000 - i * 100}.00'
            for i in range(8)
        ])
        r_icici = core.analyse_text(comma_stmt)
        self.assertEqual((r_icici['verdict'], r_icici['bank'], r_icici['rows']), ('statement', 'ICICI Bank', 8))

        # Sponsor IFSC HDFC0CTEST1 belongs to U.C.B.LTD (Urban Co-operative Bank), not HDFC Bank; UPI/sub-member IFSCs are skipped
        coop_text = 'BRANCH : U.C.B.LTD, BRANCH MAIN\nBranch IFSC Code : HDFC0CTEST1\nStatement of Account'
        self.assertEqual(core.find_bank(coop_text)[0], 'Urban Co-operative Bank')
        self.assertEqual(core.find_bank('IFSC : YESB0PTMUPI\nStatement of Account')[0], '')
        self.assertEqual(core.find_bank('IFSC : UTIB0STEST1\nStatement of Account')[0], '')

        # Clipped 10-char J&K Bank IFSC and Finacle account line / legend
        self.assertEqual(core.find_bank('Account Statement\nIFSC: JAKA0SMPLB\nMICR: 190000456')[0], 'J&K Bank')
        self.assertEqual(core.find_bank('DETAILED ACCOUNT STATEMENT', 'Account: 0005010100001234|TRADER|CAA|0005\nTransactions List - CD100-0005')[0], 'J&K Bank')
        jkb_legend = 'Statement of Account\n' + '\n'.join(rows(8)) + '\nLegends Used in Account Statement\n1. INFT - Internal Fund Transfer(Within J&K Bank)'
        self.assertEqual(core.find_bank('Statement of Account', jkb_legend)[0], 'J&K Bank')

        # ICICI variants (corporate, consolidated, loan, and legend) and SBI Term Loan
        self.assertEqual(core.find_bank('DETAILED STATEMENT', 'DETAILED STATEMENT\nTransactions List - -TEST(INR) - 000012345678')[0], 'ICICI Bank')
        self.assertEqual(core.find_bank('Summary', 'Consolidated Transaction Summary for the Account Number: 000012345678')[0], 'ICICI Bank')
        self.assertEqual(core.find_bank('Loan Statement', 'Loan Account Statement for LAAXX00012345678')[0], 'ICICI Bank')
        self.assertEqual(core.find_bank('Statement of Account', 'Account Number : 00000012345678901\nDescription : EB-MSME-TL\nC.Y.YTD Interest : 12000.00')[0], 'State Bank of India')

        # Mahindra Finance SOA with bounced counterparty Axis check stays Mahindra Finance
        mf_soa = '\n'.join([
            'STATEMENT OF ACCOUNT',
            'Contract No : 1234567',
            'RECEIPT DETAILS',
            'Date   Receipt No   Amount   Balance',
        ] + rows(8) + ['Bounced Instrument Details\nAXIS BANK   CHQ 12345\nwww.mahindrafinance.com'])
        r_mf = core.analyse_text(mf_soa)
        self.assertEqual((r_mf['verdict'], r_mf['bank']), ('statement', 'Mahindra Finance'))

    def test_new_non_statements_rejected_without_blocking_bank_ledgers(self):
        repay_sched = '\n'.join([
            'HDFC BANK LTD',
            'Repayment Schedule',
            'Amount Financed : 500000.00   Total Instl : 60',
            'Instl No Due Date Instl Amt Prin Comp Int Comp O/s Principal',
        ] + [f'{i + 1}  05-{i + 1:02d}-2025  10500.00  8000.00  2500.00  {490000 - i * 8000}.00' for i in range(8)])
        self.assertEqual(core.analyse_text(repay_sched)['verdict'], 'not')

        tally_vch = '\n'.join([
            'Firm Traders',
            'Ledger Account',
            'Date   Particulars   Vch Type   Vch No.   Debit   Credit',
        ] + [f'{i + 1:02d}-04-2025  Sales Account  Sales  10{i}  5000.00  5000.00' for i in range(8)])
        self.assertEqual(core.analyse_text(tally_vch)['verdict'], 'not')

        sale_bill_ledger = '\n'.join([
            'M/S ABC Enterprises',
            'Purchase Ledger 01/04/2024 to 31/03/2025',
            'Date   Sale Bill No.   Party   Amount   Balance',
        ] + [f'{i + 1:02d}-04-2024  SB-{100 + i}  Buyer  2500.00  25000.00' for i in range(8)])
        self.assertEqual(core.analyse_text(sale_bill_ledger)['verdict'], 'not')

        gst_ecl = '\n'.join([
            'Goods and Services Tax',
            'Electronic Credit ledger',
            'Integrated Central State Tax CESS Balance',
        ] + [f'{i + 1:02d}/04/2025  ITC  1000.00  500.00  500.00  0.00  2000.00' for i in range(8)])
        self.assertEqual(core.analyse_text(gst_ecl)['verdict'], 'not')

        proposal = '\n'.join([
            'RESTRUCTURING PROPOSAL',
            'Name and Address of the Borrower : M/S ABC',
            'Bank Balance Sheet Statement of Account',
        ] + rows(8))
        self.assertEqual(core.analyse_text(proposal)['verdict'], 'not')

        jio_bill = '\n'.join([
            'Statement of Account',
            'Jio Number : 7000000000  Account Number : 1234567890',
            'Plan Recharges   Usage Charges   Total Amount',
            'Set ACH mandate via MyJio',
        ] + [f'{i + 1:02d}/04/2025  VOICE CALL  0.00  199.00  199.00' for i in range(8)])
        self.assertEqual(core.analyse_text(jio_bill)['verdict'], 'not')

        # Real bank "Account Ledger Inquiry" (PNB) must still pass as a bank statement
        pnb_ledger = '\n'.join([
            'PUNJAB NATIONAL BANK',
            'Account Ledger Inquiry',
            'Account No : 0123008700001234   IFSC : PUNB0012300',
            'Date   Particulars   Debit   Credit   Balance',
        ] + rows(10))
        r_pnb = core.analyse_text(pnb_ledger)
        self.assertEqual((r_pnb['verdict'], r_pnb['bank']), ('statement', 'Punjab National Bank'))


# ------------------------------------------------------------------------------------------------ reading files
class Readers(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp(prefix='sf_read_')

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def classify(self, path):
        return core.classify_file(path, dict(core.DEFAULTS))

    def test_text_rpt_nfs_and_utf16(self):
        for name, enc in (('a.txt', 'utf-8'), ('b.rpt', 'utf-8'), ('c.rptnfs', 'cp1252'), ('d.nfs', 'utf-8'), ('e.txt', 'utf-16')):
            r = self.classify(write(os.path.join(self.d, name), statement('hdfc'), enc))
            self.assertEqual((r['verdict'], r['bank']), ('statement', 'HDFC Bank'), name)
            self.assertEqual(len(r['sha']), 64)

    def test_utf8_split_at_read_limit_stays_utf8(self):
        prefix = 'HDFC BANK\nStatement of account\n' + 'a' * (400_000 - len('HDFC BANK\nStatement of account\n') - 1)
        p = os.path.join(self.d, 'big_utf8.txt')
        with open(p, 'wb') as f:
            f.write(prefix.encode('utf-8') + '\u20b9'.encode('utf-8'))   # 3-byte rupee sign split at byte 400,000
        text, _meta = core._read_text(p)
        self.assertNotIn('\xe2', text[-10:])
        self.assertIn('HDFC BANK', text)

    def test_a_pdf_is_read_row_by_row(self):
        r = self.classify(make_pdf(os.path.join(self.d, 'x.pdf'), statement('hdfc').splitlines()))
        self.assertEqual((r['verdict'], r['bank'], r['rows']), ('statement', 'HDFC Bank', 12))

    def test_word_with_a_table(self):
        import docx
        d = docx.Document()
        d.add_paragraph('JAMMU AND KASHMIR BANK LTD')
        d.add_paragraph('IFSC Code : JAKA0SAMPLE  STATEMENT OF ACCOUNT')
        t = d.add_table(rows=1, cols=6)
        for c, h in zip(t.rows[0].cells, ('Date', 'Particulars', 'Chq', 'Withdrawals', 'Deposits', 'Balance')):
            c.text = h
        for i in range(10):
            cells = t.add_row().cells
            for c, v in zip(cells, (f'{i + 1:02d}-04-2025', 'UPI/SHOP', '', f'{100 + i}.00', '', f'{9000 - i}.50Cr')):
                c.text = v
        p = os.path.join(self.d, 's.docx')
        d.save(p)
        r = self.classify(p)
        self.assertEqual((r['verdict'], r['bank']), ('statement', 'J&K Bank'))

    def test_html_statement(self):
        html_doc = '<html><body><table><tr><td>ICICI Bank Limited</td></tr><tr><td>Statement of Account</td><td>IFSC ICIC0001234</td></tr>' + \
                   '<tr><th>Date</th><th>Narration</th><th>Withdrawal</th><th>Deposit</th><th>Balance</th></tr>' + \
                   ''.join(f'<tr><td>{i + 1:02d}/04/2025</td><td>UPI</td><td>{100 + i}.00</td><td></td><td>{9000 - i}.00</td></tr>' for i in range(10)) + \
                   '</table></body></html>'
        r = self.classify(write(os.path.join(self.d, 'h.html'), html_doc))
        self.assertEqual((r['verdict'], r['bank']), ('statement', 'ICICI Bank'))

    def test_locked_and_scanned_pdfs(self):
        locked = self.classify(make_pdf(os.path.join(self.d, 'l.pdf'), statement('hdfc').splitlines(), password='secret'))
        self.assertEqual(locked['verdict'], 'locked')
        scanned = self.classify(make_pdf(os.path.join(self.d, 's.pdf'), [], blank=True))
        self.assertEqual(scanned['verdict'], 'scanned')

    def test_unreadable_empty_and_missing(self):
        self.assertEqual(self.classify(write(os.path.join(self.d, 'e.txt'), ''))['verdict'], 'unreadable')
        bad = write(os.path.join(self.d, 'bad.pdf'), 'this is not a pdf')
        self.assertEqual(self.classify(bad)['verdict'], 'unreadable')
        self.assertEqual(self.classify(os.path.join(self.d, 'gone.pdf'))['verdict'], 'unreadable')

    def test_too_big_is_not_opened(self):
        p = write(os.path.join(self.d, 'big.txt'), statement('hdfc'))
        r = core.classify_file(p, dict(core.DEFAULTS, max_mb=0))
        self.assertEqual(r['verdict'], 'too_big')


# ------------------------------------------------------------------------------------------------ folders and copying
class Walk(unittest.TestCase):
    def test_what_is_looked_at_and_excel_csv_are_skipped(self):
        d = tempfile.mkdtemp(prefix='sf_walk_')
        try:
            for rel in ('a/one.pdf', 'a/b/two.XLSX', 'a/b/old.xls', 'a/b/data.csv', 'a/b/c/three.txt', 'a/b/c/four.rpt',
                        'a/b/c/five.nfs', 'a/b/c/six.rptnfs', 'node_modules/x/skip.pdf', 'Windows/skip.pdf', 'a/~$lock.docx',
                        'a/notes.md', 'a/photo.jpg', 'out/inside.pdf', '.git/objects/skip.txt'):
                write(os.path.join(d, rel), 'x')
            exts = core.FILE_TYPES['pdf'] + core.FILE_TYPES['word'] + core.FILE_TYPES['text']
            names = lambda **kw: sorted(os.path.relpath(p, d).replace('\\', '/') for p, _s, _m in core.walk_files(
                d, exts, kw.get('skip', True), kw.get('exclude', []), threading.Event(), []))
            self.assertEqual(names(exclude=[os.path.join(d, 'out')]),
                             ['a/b/c/five.nfs', 'a/b/c/four.rpt', 'a/b/c/six.rptnfs', 'a/b/c/three.txt', 'a/one.pdf'])
            problems = []
            list(core.walk_files(os.path.join(d, 'missing'), exts, True, [], threading.Event(), problems))
            self.assertEqual(len(problems), 1)
        finally:
            shutil.rmtree(d, ignore_errors=True)


class Copying(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp(prefix='sf_copy_')

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def rec(self, path, **kw):
        with open(path, 'rb') as f:
            sha = hashlib.sha256(f.read()).hexdigest()
        base = {'path': path, 'name': os.path.basename(path), 'verdict': 'statement', 'kind': 'bank', 'bank': 'HDFC Bank',
                'sha': sha}
        base.update(kw)
        return base

    def test_folders(self):
        t = core.target_folder
        self.assertEqual(t({'verdict': 'statement', 'kind': 'bank', 'bank': 'HDFC Bank'}), 'HDFC Bank')
        self.assertEqual(t({'verdict': 'likely', 'kind': 'bank', 'bank': 'HDFC Bank'}), os.path.join('HDFC Bank', '(not sure)'))
        self.assertEqual(t({'verdict': 'statement', 'kind': 'bank', 'bank': ''}), '_Bank not identified')
        self.assertEqual(t({'verdict': 'statement', 'kind': 'card', 'bank': 'HDFC Bank'}), os.path.join('_Credit card statements', 'HDFC Bank'))
        self.assertEqual(core.safe_name('A/B:C?.pdf '), 'A_B_C_.pdf')
        long_pdf = core.safe_filename('A' * 95 + '.pdf')
        self.assertTrue(len(long_pdf) <= 80 and long_pdf.endswith('.pdf'), long_pdf)

    def test_copy_keeps_the_original_and_its_date(self):
        src = write(os.path.join(self.d, 'in', 'a.txt'), 'one')
        os.utime(src, (1_600_000_000, 1_600_000_000))
        cp = core.Copier(os.path.join(self.d, 'out'), False)
        action, target = cp.place(self.rec(src))
        self.assertEqual(action, 'copied')
        self.assertEqual(target, os.path.join(self.d, 'out', 'HDFC Bank', 'a.txt'))
        with open(src, encoding='utf-8') as f:
            self.assertEqual(f.read(), 'one')
        self.assertEqual(int(os.path.getmtime(target)), 1_600_000_000)

    def test_same_name_different_content_gets_a_number_and_same_content_is_copied_once(self):
        a = write(os.path.join(self.d, 'in1', 's.txt'), 'AAA')
        b = write(os.path.join(self.d, 'in2', 's.txt'), 'BBB')
        c = write(os.path.join(self.d, 'in3', 'other name.txt'), 'AAA')
        cp = core.Copier(os.path.join(self.d, 'out'), False)
        self.assertEqual(cp.place(self.rec(a))[0], 'copied')
        action, target = cp.place(self.rec(b))
        self.assertEqual((action, os.path.basename(target)), ('copied', 's (2).txt'))
        self.assertEqual(cp.place(self.rec(c))[0], 'duplicate')
        again = core.Copier(os.path.join(self.d, 'out'), False)                      # a second run finds them already there
        self.assertEqual(again.place(self.rec(a))[0], 'already there')
        self.assertEqual(again.place(self.rec(b))[0], 'already there')

    def test_dry_run_collision_numbers_targets(self):
        a = write(os.path.join(self.d, 'd1', 's.txt'), 'AAA')
        b = write(os.path.join(self.d, 'd2', 's.txt'), 'BBB')
        cp = core.Copier(os.path.join(self.d, 'dry_out'), True)
        act1, t1 = cp.place(self.rec(a))
        act2, t2 = cp.place(self.rec(b))
        self.assertEqual((act1, os.path.basename(t1)), ('would copy', 's.txt'))
        self.assertEqual((act2, os.path.basename(t2)), ('would copy', 's (2).txt'))

    def test_scan_only_copies_nothing(self):
        src = write(os.path.join(self.d, 'in', 'a.txt'), 'one')
        cp = core.Copier(os.path.join(self.d, 'out'), True)
        self.assertEqual(cp.place(self.rec(src))[0], 'would copy')
        self.assertFalse(os.path.exists(os.path.join(self.d, 'out')))


# ------------------------------------------------------------------------------------------------ a whole run
def collect(events):
    return lambda ev: events.append(ev)


class WholeRun(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d = tempfile.mkdtemp(prefix='sf_run_')
        s = os.path.join(cls.d, 'source')
        write(os.path.join(s, 'a', 'hdfc1.rpt'), statement('hdfc'))
        write(os.path.join(s, 'a', 'b', 'copy of hdfc1.rpt'), statement('hdfc'))                     # the same bytes
        write(os.path.join(s, 'c', 'sbi.nfs'), statement('sbi'))
        write(os.path.join(s, 'c', 'unknown.txt'), statement('nobank'))
        write(os.path.join(s, 'c', 'invoice.txt'), INVOICE)
        write(os.path.join(s, 'c', 'card.txt'), CARD)
        write(os.path.join(s, 'c', 'maybe.txt'), '\n'.join(['Statement of account', 'Account No 12345678', 'IFSC SBIN0001234', 'Branch MAIN'] + rows(3)))
        write(os.path.join(s, 'c', 'ignored.csv'), statement('hdfc'))                                # CSV and Excel must be ignored completely
        write(os.path.join(s, 'c', 'ignored.xlsx'), statement('hdfc'))
        make_pdf(os.path.join(s, 'd', 'locked statement.pdf'), statement('hdfc').splitlines(), password='secret')
        make_pdf(os.path.join(s, 'd', 'scan statement.pdf'), [], blank=True)
        make_pdf(os.path.join(s, 'e', 'hdfc2.pdf'), statement('hdfc', n=15, narration='ATM-CASH').splitlines())
        cls.source = s
        cls.before = tree_state(s)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.d, ignore_errors=True)

    def go(self, dest, **opts):
        events = []
        summary = core.run(self.source, dest, dict({'workers': 3, 'timeout': 30}, **opts), collect(events), threading.Event())
        return summary, events

    def test_full_run_skips_excel_csv_and_locked_or_scanned_pdfs(self):
        dest = os.path.join(self.d, 'out_full')
        s, events = self.go(dest)
        got = tree_state(dest)
        report = [k for k in got if k.startswith('Statement Finder report')]
        self.assertEqual(len(report), 1)
        files = sorted(k.replace('\\', '/') for k in got if k not in report)
        self.assertEqual(files, sorted([
            'HDFC Bank/hdfc1.rpt', 'HDFC Bank/hdfc2.pdf', 'State Bank of India/sbi.nfs', '_Bank not identified/unknown.txt',
            '_Credit card statements/HDFC Bank/card.txt', 'State Bank of India/(not sure)/maybe.txt']))
        self.assertEqual((s['files'], s['duplicates'], s['stopped']), (10, 1, False))
        self.assertEqual((s['statements'], s['not_sure'], s['locked'], s['scanned'], s['not_statements']), (6, 1, 1, 1, 1))
        self.assertEqual(self.before, tree_state(self.source), 'the originals must not change')
        kinds = [e[0] for e in events]
        self.assertEqual(kinds[0], 'phase')
        self.assertEqual(kinds[-1], 'done')
        self.assertIn(('total', 10), events)
        self.assertEqual(max(e[1] for e in events if e[0] == 'progress'), 10)

    def test_the_report(self):
        dest = os.path.join(self.d, 'out_report')
        s, _events = self.go(dest)
        with open(s['report'], encoding='utf-8-sig', newline='') as f:
            rows_ = list(csv.DictReader(f))
        self.assertEqual(list(rows_[0].keys()), core.REPORT_COLUMNS)
        by = {r['File']: r for r in rows_}
        self.assertEqual(by['hdfc1.rpt']['Result'], 'Copied')
        self.assertEqual(by['hdfc1.rpt']['Bank'], 'HDFC Bank')
        self.assertEqual(by['hdfc1.rpt']['Certainty'], 'Sure')
        self.assertEqual(by['maybe.txt']['Certainty'], 'Not sure')
        self.assertEqual(by['copy of hdfc1.rpt']['Result'], 'Same file as one already found')
        self.assertEqual(by['card.txt']['Type'], 'Credit card')
        self.assertNotIn('locked statement.pdf', by)
        self.assertNotIn('scan statement.pdf', by)
        self.assertNotIn('invoice.txt', by)
        self.assertTrue(by['hdfc1.rpt']['Why'].startswith('title'))

    def test_scan_only(self):
        dest = os.path.join(self.d, 'out_dry')
        s, _events = self.go(dest, dry_run=True)
        self.assertEqual([f for f in tree_state(dest) if not f.startswith('Statement Finder report')], [])
        self.assertEqual(s['copied'], 6)
        with open(s['report'], encoding='utf-8-sig') as f:
            self.assertIn('Would copy (scan only)', f.read())

    def test_not_sure_can_be_left_out(self):
        dest = os.path.join(self.d, 'out_strict')
        self.go(dest, include_not_sure=False)
        files = sorted(k.replace('\\', '/') for k in tree_state(dest) if not k.startswith('Statement Finder report'))
        self.assertEqual(files, sorted(['HDFC Bank/hdfc1.rpt', 'HDFC Bank/hdfc2.pdf', 'State Bank of India/sbi.nfs',
                                        '_Bank not identified/unknown.txt', '_Credit card statements/HDFC Bank/card.txt']))

    def test_file_types(self):
        dest = os.path.join(self.d, 'out_pdf')
        s, _events = self.go(dest, kinds=('pdf',))
        self.assertEqual(s['files'], 3)
        self.assertEqual(sorted(k.replace('\\', '/') for k in tree_state(dest) if not k.startswith('Statement Finder')),
                         ['HDFC Bank/hdfc2.pdf'])

    def test_the_destination_inside_the_source_is_not_searched_again(self):
        src = os.path.join(self.d, 'src2')
        write(os.path.join(src, 'x', 'hdfc.txt'), statement('hdfc'))
        dest = os.path.join(src, 'found')
        first = core.run(src, dest, {'workers': 2}, lambda e: None, threading.Event())
        second = core.run(src, dest, {'workers': 2}, lambda e: None, threading.Event())
        self.assertEqual((first['files'], first['copied']), (1, 1))
        self.assertEqual((second['files'], second['copied'], second['duplicates']), (1, 0, 0))        # the copy was not found again
        self.assertEqual(sorted(k for k in tree_state(dest) if not k.startswith('Statement Finder')), [os.path.join('HDFC Bank', 'hdfc.txt')])

    def test_an_empty_folder(self):
        empty = os.path.join(self.d, 'empty')
        os.makedirs(empty)
        s = core.run(empty, os.path.join(self.d, 'out_empty'), {}, lambda e: None, threading.Event())
        self.assertEqual((s['files'], s['read'], s['copied']), (0, 0, 0))
        self.assertTrue(os.path.isfile(s['report']))


class Robust(unittest.TestCase):
    """A file that hangs or crashes its reader must not stop the run."""

    def test_hang_and_crash(self):
        d = tempfile.mkdtemp(prefix='sf_robust_')
        try:
            s = os.path.join(d, 'src')
            for i in range(6):
                write(os.path.join(s, f'ok{i}.txt'), statement('hdfc', n=10 + i))
            write(os.path.join(s, 'hang.txt'), statement('hdfc'))
            write(os.path.join(s, 'crash.txt'), statement('hdfc'))
            events = []
            for hook in ('hang:hang.txt', 'crash:crash.txt'):
                os.environ['STATEMENT_FINDER_TEST_HOOK'] = hook
                events.clear()
                t0 = time.time()
                summary = core.run(s, os.path.join(d, 'out_' + hook[:4]), {'workers': 2, 'timeout': 5}, collect(events), threading.Event())
                took = time.time() - t0
                self.assertLess(took, 40, hook)
                self.assertEqual((summary['read'], summary['statements'], summary['unreadable']), (8, 7, 1), hook)     # 6 ok files + the one the hook leaves alone
                bad = [e[1] for e in events if e[0] == 'found' and e[1]['verdict'] == 'unreadable']
                self.assertEqual(len(bad), 1)
                self.assertEqual(bad[0]['name'], hook.split(':')[1])
                self.assertIn('stopped after 5 seconds' if 'hang' in hook else 'crashed', bad[0]['error'])
                self.assertEqual(summary['read'], 8, 'every file is accounted for')
        finally:
            os.environ.pop('STATEMENT_FINDER_TEST_HOOK', None)
            shutil.rmtree(d, ignore_errors=True)

    def test_stop_ends_the_run_quickly_even_with_a_stuck_reader(self):
        d = tempfile.mkdtemp(prefix='sf_stop_')
        try:
            s = os.path.join(d, 'src')
            write(os.path.join(s, 'hang.txt'), statement('hdfc'))
            for i in range(3):
                write(os.path.join(s, f'ok{i}.txt'), statement('hdfc'))
            os.environ['STATEMENT_FINDER_TEST_HOOK'] = 'hang:hang.txt'
            stop = threading.Event()
            out = {}
            t = threading.Thread(target=lambda: out.update(core.run(s, os.path.join(d, 'out'), {'workers': 2, 'timeout': 600}, lambda e: None, stop)))
            t0 = time.time()
            t.start()
            time.sleep(3)
            stop.set()
            t.join(20)
            self.assertFalse(t.is_alive())
            self.assertLess(time.time() - t0, 15)
            self.assertTrue(out['stopped'])
            self.assertTrue(os.path.isfile(out['report']), 'a report is written even when stopped')
        finally:
            os.environ.pop('STATEMENT_FINDER_TEST_HOOK', None)
            shutil.rmtree(d, ignore_errors=True)

    def test_stopped_before_it_starts(self):
        d = tempfile.mkdtemp(prefix='sf_stop0_')
        try:
            write(os.path.join(d, 'src', 'a.txt'), statement('hdfc'))
            stop = threading.Event()
            stop.set()
            s = core.run(os.path.join(d, 'src'), os.path.join(d, 'out'), {}, lambda e: None, stop)
            self.assertEqual((s['stopped'], s['copied']), (True, 0))
        finally:
            shutil.rmtree(d, ignore_errors=True)


# ------------------------------------------------------------------------------------------------ the window
class Window(unittest.TestCase):
    def setUp(self):
        try:
            import tkinter as tk
            self.root = tk.Tk()
        except Exception as exc:                                       # no display
            self.skipTest(f'no window available: {exc}')
        self.root.withdraw()
        import statement_finder as sf
        self.sf = sf
        self.d = tempfile.mkdtemp(prefix='sf_gui_')
        self.patches = [mock.patch.object(sf, 'SETTINGS_DIR', os.path.join(self.d, 'cfg')),
                        mock.patch.object(sf, 'SETTINGS_FILE', os.path.join(self.d, 'cfg', 's.json')),
                        mock.patch.object(sf, 'CRASH_FILE', os.path.join(self.d, 'cfg', 'crash.log'))]
        for p in self.patches:
            p.start()
        self.app = sf.App(self.root)

    def tearDown(self):
        for p in self.patches:
            p.stop()
        try:
            self.root.destroy()
        except Exception:
            pass
        shutil.rmtree(self.d, ignore_errors=True)

    def pump(self, until, seconds=90):
        t0 = time.time()
        while not until() and time.time() - t0 < seconds:
            self.root.update()
            time.sleep(0.02)
        for _ in range(20):
            self.root.update()
        return until()

    def test_a_search_from_the_window(self):
        src = os.path.join(self.d, 'src')
        write(os.path.join(src, 'a', 'h.rpt'), statement('hdfc'))
        write(os.path.join(src, 'b', 's.nfs'), statement('sbi'))
        write(os.path.join(src, 'b', 'inv.txt'), INVOICE)
        self.app.source.set(src)
        self.app.dest.set(os.path.join(self.d, 'dest'))
        self.app._start()
        self.assertEqual(str(self.app.stop_btn['state']), 'normal')
        self.assertEqual(str(self.app.start_btn['state']), 'disabled')
        self.assertTrue(self.pump(lambda: self.app.result is not None), 'the search did not finish')
        self.assertEqual((self.app.sure, self.app.not_sure, self.app.copied), (2, 0, 2))
        rows_ = {self.app.tree.item(i, 'values')[0]: self.app.tree.item(i, 'values') for i in self.app.tree.get_children()}
        self.assertEqual(set(rows_), {'HDFC Bank', 'State Bank of India', 'All banks'})
        self.assertEqual(str(self.app.start_btn['state']), 'normal')
        self.assertEqual((str(self.app.stop_btn['state']), str(self.app.open_btn['state']), str(self.app.report_btn['state'])), ('disabled', 'normal', 'normal'))
        self.assertIn('2 look like bank statements', self.app.status.get())
        self.assertTrue(os.path.isfile(os.path.join(self.d, 'dest', 'HDFC Bank', 'h.rpt')))
        self.assertTrue(os.path.isfile(self.sf.SETTINGS_FILE), 'the folders are remembered')

    def test_it_will_not_start_with_bad_input(self):
        with mock.patch.object(self.sf.messagebox, 'showwarning') as warn:
            self.app._start()                                                          # nothing chosen
            self.app.source.set(self.d)
            self.app._start()                                                          # no destination
            self.app.dest.set(self.d)
            self.app._start()                                                          # the same folder
            self.app.dest.set(os.path.join(self.d, 'x'))
            for v in self.app.kind_vars.values():
                v.set(False)
            self.app._start()                                                          # no file type
        self.assertEqual(warn.call_count, 4)
        self.assertFalse(self.app.running)
        self.assertFalse(os.path.exists(os.path.join(self.d, 'x')))

    def test_cleared_spinbox_falls_back_cleanly(self):
        src = os.path.join(self.d, 'src')
        write(os.path.join(src, 'h.txt'), statement('hdfc'))
        self.app.source.set(src)
        self.app.dest.set(os.path.join(self.d, 'dest'))
        self.app.workers.set('')                                       # simulate clearing Spinbox
        self.app.timeout.set('')
        self.app._start()
        self.assertTrue(self.pump(lambda: self.app.result is not None), 'the search did not finish')
        self.assertEqual(self.app.sure, 1)
        self.assertGreaterEqual(self.app.workers.get(), 1)
        self.assertGreaterEqual(self.app.timeout.get(), 10)

    def test_cli_run_mode(self):
        src = os.path.join(self.d, 'cli_src')
        dest = os.path.join(self.d, 'cli_dest')
        write(os.path.join(src, 'h.txt'), statement('hdfc'))
        with mock.patch.object(sys, 'argv', ['sf', '--run', src, dest]):
            self.sf.main()
        self.assertTrue(os.path.isfile(os.path.join(dest, 'HDFC Bank', 'h.txt')))
        with mock.patch.object(sys, 'argv', ['sf', '--run', src, src]):
            with self.assertRaises(SystemExit) as ctx:
                self.sf.main()
            self.assertEqual(ctx.exception.code, 2)


if __name__ == '__main__':
    mp.freeze_support()
    unittest.main(verbosity=1)
