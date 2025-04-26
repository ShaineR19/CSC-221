import unittest
from m_pro_4_groupa_functions import clean_instructor_name  # Replace 'module_name' with the actual module name

class TestCleanInstructorName(unittest.TestCase):
    
    def test_comma_separated_names(self):
        """Test names in 'Last, First' format."""
        self.assertEqual(clean_instructor_name("Seidi, H"), "seidih_FTE.xlsx")
        self.assertEqual(clean_instructor_name("Seidi, H."), "seidih_FTE.xlsx")
        self.assertEqual(clean_instructor_name("Smith, John"), "smithj_FTE.xlsx")
        self.assertEqual(clean_instructor_name("  O'Connor,  J. "), "o'connorj_FTE.xlsx")
    
    def test_space_separated_names(self):
        """Test names in 'First Last' format."""
        self.assertEqual(clean_instructor_name("H Seidi"), "seidih_FTE.xlsx")
        self.assertEqual(clean_instructor_name("H. Seidi"), "seidih_FTE.xlsx")
        self.assertEqual(clean_instructor_name("John Smith"), "smithj_FTE.xlsx")
        self.assertEqual(clean_instructor_name("J. O'Connor  "), "o'connorj_FTE.xlsx")
    
    def test_multiple_word_names(self):
        """Test names with multiple words."""
        self.assertEqual(clean_instructor_name("J Van Der Beek"), "beekj_FTE.xlsx")
        self.assertEqual(clean_instructor_name("Van Der Beek, J"), "van der beekj_FTE.xlsx")
        


if __name__ == '__main__':
    unittest.main()