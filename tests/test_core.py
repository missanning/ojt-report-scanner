import io
import unittest
import zipfile

from app import parse_report, years_for_record, make_workbook


class ReportTests(unittest.TestCase):
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
