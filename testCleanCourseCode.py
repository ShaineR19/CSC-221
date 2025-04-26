import unittest
from m_pro_4_groupa_functions import clean_course_code

class TestCleanCourseCode(unittest.TestCase):
    def test_clean_course_code(self):
        # Test with a typical course code
        result = clean_course_code('CSI-120-0003')
        expected = "csi1200003_FTE.xlsx"
        self.assertEqual(result, expected)
        
        # You can add more test cases if needed
        result2 = clean_course_code('MAT-171-0001')
        expected2 = "mat1710001_FTE.xlsx"
        self.assertEqual(result2, expected2)

        # Test with a typical course code
        result = clean_course_code('WBL-111-5001')
        expected = "wbl1115001_FTE.xlsx"
        self.assertEqual(result, expected)

if __name__ == '__main__':
    unittest.main()