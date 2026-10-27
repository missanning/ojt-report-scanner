import io
import unittest
import zipfile

from app import parse_report, years_for_record, make_workbook, repair_saved_record


class ReportTests(unittest.TestCase):
    def test_repairs_saved_unreviewed_address_without_losing_other_fields(self):
        old = {'id': 'saved-id', 'company': 'UPHSL - Bifian', 'address': 'ITS Department', 'student': 'Dufigo, Ezra', 'section': 'J4A', 'verified': False, 'raw_text': 'TRAINING NARRATIVE REPORT\nUPHSL - Bifian\nITS Department\nNational Highway, Sto. Nifio, Bifian, Laguna\nOJT Duration:\nJuly 17, 2025 - Nov 12, 2025'}
        fixed = repair_saved_record(old)
        self.assertEqual(fixed['company'], 'UPHSL - Biñan')
        self.assertEqual(fixed['address'], 'National Highway, Sto. Niño, Biñan, Laguna')
        self.assertEqual(fixed['student'], 'Duñgo, Ezra')
        self.assertEqual(fixed['section'], 'J4A')
        self.assertEqual(fixed['id'], 'saved-id')
        self.assertEqual(old['address'], 'ITS Department')
        self.assertEqual(repair_saved_record(dict(old, verified=True)), dict(old, verified=True))
        self.assertEqual(repair_saved_record({'id': 'no-ocr'}), {'id': 'no-ocr'})

    def test_address_skips_department_and_keeps_street_lines(self):
        for department in ('ITS Department', 'INFORMATION TECHNOLOGY DEPARTMENT'):
            with self.subTest(department=department):
                text = f'TRAINING NARRATIVE REPORT\nUPHSL - BIFIAN\n{department}\n83JP+F52, National Highway, Sto. Niöo,\nBifian, Laguna\nOJT Duration:\nJuly 17, 2025 - Nov 12, 2025\nSubmitted by:\nANTONIO, NICO JOHN Y.\nBSIT GD\nJ4A'
                row = parse_report(text, 'address.jpg')
                self.assertEqual(row['company'], 'UPHSL - BIÑAN')
                self.assertEqual(row['address'], '83JP+F52, National Highway, Sto. Niño, Biñan, Laguna')
                self.assertEqual(row['student'], 'ANTONIO, NICO JOHN Y.')

    def test_department_without_street_address_is_flagged(self):
        row = parse_report('TRAINING NARRATIVE REPORT\nUPHSL\nITS Department\nOJT Duration:\nJuly 17, 2025 - Nov 12, 2025\nSubmitted by:\nJane Doe\nBSIT - J4A', 'missing.jpg')
        self.assertEqual(row['address'], '')
        self.assertIn('address', row['needs_review'])

    def test_extracts_inline_section_after_program_or_major(self):
        for printed, program, section in (
            ('BSCS - J4S', 'BSCS', 'J4S'),
            ('BSIT-J4A', 'BSIT', 'J4A'),
            ('BSIT-GD - J4T', 'BSIT', 'J4T'),
            ('BSIT — Major in Game Development J4S', 'BSIT', 'J4S'),
            ('BSCS / j3a', 'BSCS', 'J3A'),
        ):
            with self.subTest(printed=printed):
                row = parse_report(f'Submitted by:\nJane Doe\n{printed}\nSubmitted to:\nProfessor', 'inline.jpg')
                self.assertEqual(row['program'], program)
                self.assertEqual(row['section'], section)
                self.assertNotIn('section', row['needs_review'])

    def test_repairs_confirmed_enye_ocr_errors_without_changing_real_fi(self):
        text = 'TRAINING NARRATIVE REPORT\nUPHSL - Bifian\nSto. Nifio, Bifian, Laguna\nSubmitted by:\nDufigo, Ezra Eleazar F.\nBSIT - J4T'
        row = parse_report(text, 'enye.jpg')
        self.assertEqual(row['company'], 'UPHSL - Biñan')
        self.assertEqual(row['address'], 'Sto. Niño, Biñan, Laguna')
        self.assertEqual(row['student'], 'Duñgo, Ezra Eleazar F.')
        self.assertEqual(row['raw_text'], text)
        for name in ('DUFIGO, EZRA', 'Duigo, Ezra', 'Duñgo, Ezra', 'Peña, Niño', 'PEÑA, NIÑO', 'Fidel Delfin', 'Office Filinvest'):
            with self.subTest(name=name):
                row = parse_report(f'Submitted by:\n{name}\nBSCS - J4S', 'name.jpg')
                expected = {'DUFIGO, EZRA': 'DUÑGO, EZRA', 'Duigo, Ezra': 'Duñgo, Ezra'}.get(name, name)
                self.assertEqual(row['student'], expected)

    def test_normalizes_combining_enye_and_preserves_it_in_excel(self):
        row = parse_report('TRAINING NARRATIVE REPORT\nAcme\nBin\u0303an\nSubmitted by:\nPen\u0303a, NIN\u0303O\nBSCS - J4S', 'unicode.jpg')
        self.assertEqual(row['student'], 'Peña, NIÑO')
        self.assertEqual(row['address'], 'Biñan')
        with zipfile.ZipFile(io.BytesIO(make_workbook([row]))) as workbook:
            sheet = workbook.read('xl/worksheets/sheet1.xml').decode('utf-8')
            self.assertIn('Peña, NIÑO', sheet)
            self.assertIn('Biñan', sheet)
            self.assertIn('J4S', sheet)

    def test_program_majors_are_classified_with_section_preserved(self):
        for printed, expected in (
            ('BSCS-DA', 'BSCS'), ('BSIT-GD', 'BSIT'),
            ('bscs – Software Engineering', 'BSCS'),
            ('B S I T - GD', 'BSIT'), ('B.S.C.S. - DA', 'BSCS'),
            ('Bachelor of Science in Computer Science major in Data Analytics', 'BSCS'),
            ('Bachelor of Science in Information Technology - Game Development', 'BSIT'),
        ):
            with self.subTest(printed=printed):
                row = parse_report(f'Submitted by:\nJane Doe\n{printed}\n4th/J4T\nSubmitted to:\nProfessor', 'major.jpg')
                self.assertEqual(row['program'], expected)
                self.assertEqual(row['section'], '4th/J4T')

    def test_program_detection_does_not_match_inside_other_words(self):
        row = parse_report('Submitted by:\nJane Doe\nBSITEST\nJ4T', 'unknown.jpg')
        self.assertEqual(row['program'], '')

    def test_extracts_cover_fields_and_both_years(self):
        text = '''University of Perpetual Help System Laguna
TRAINING NARRATIVE REPORT
<Motorcentral>
<2nd flr. Morales Building, Brgy. San Antonio, Binan Laguna>
OJT Duration:
<July 28 2025 - Dec 1 2026>
Submitted by:
<Ventura Vixcen Erean Antonio>
<BSIT>
<4th/J4T>
Submitted to:
Oliver M. Junio'''
        row = parse_report(text, 'ojt2.jpg')
        self.assertEqual(row['company'], 'Motorcentral')
        self.assertEqual(row['address'], '2nd flr. Morales Building, Brgy. San Antonio, Binan Laguna')
        self.assertEqual(row['student'], 'Ventura Vixcen Erean Antonio')
        self.assertEqual(row['program'], 'BSIT')
        self.assertEqual(row['section'], '4th/J4T')
        self.assertEqual(years_for_record(row), [2025, 2026])

    def test_missing_fields_are_flagged_without_inventing_values(self):
        row = parse_report('TRAINING NARRATIVE REPORT\nAcme\nSubmitted by:\nJane Doe\nBSCS', 'a.jpg')
        self.assertEqual(row['address'], '')
        self.assertEqual(row['section'], '')
        self.assertIn('address', row['needs_review'])
        self.assertIn('section', row['needs_review'])

    def test_workbook_repeats_spanning_record_for_each_filter_year(self):
        row = parse_report('TRAINING NARRATIVE REPORT\nAcme\nLaguna\nOJT Duration:\nDec 1, 2025 - Jan 2, 2026\nSubmitted by:\nJane Doe\nBSCS\nJ3A', 'a.jpg')
        data = make_workbook([row])
        with zipfile.ZipFile(io.BytesIO(data)) as workbook:
            sheet = workbook.read('xl/worksheets/sheet1.xml').decode()
            self.assertIn('autoFilter ref=', sheet)
            self.assertIn('>2025<', sheet)
            self.assertIn('>2026<', sheet)
        filtered = make_workbook([row], 2026)
        with zipfile.ZipFile(io.BytesIO(filtered)) as workbook:
            sheet = workbook.read('xl/worksheets/sheet1.xml').decode()
            self.assertIn('>2026<', sheet)
            self.assertNotIn('>2025<', sheet)


if __name__ == '__main__':
    unittest.main()
