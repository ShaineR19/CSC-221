import unittest
from m_pro_4_groupa_functions import clean_name_for_search  # Replace 'module_name' with the actual module name

class TestCleanNameForSearch(unittest.TestCase):
    
    def test_remove_periods(self):
        """Test removal of periods."""
        self.assertEqual(clean_name_for_search("J.R. Smith"), "jr smith")
        self.assertEqual(clean_name_for_search("Ph.D."), "phd")
        self.assertEqual(clean_name_for_search("St. James"), "st james")
        
    def test_strip_spaces(self):
        """Test stripping extra spaces."""
        self.assertEqual(clean_name_for_search("  John Smith  "), "john smith")
        self.assertEqual(clean_name_for_search("John  Smith"), "john  smith")  # Note: internal spaces are preserved
        self.assertEqual(clean_name_for_search("\tJohn Smith\n"), "john smith")
        
    def test_lowercase_conversion(self):
        """Test conversion to lowercase."""
        self.assertEqual(clean_name_for_search("JOHN SMITH"), "john smith")
        self.assertEqual(clean_name_for_search("John SMITH"), "john smith")
        self.assertEqual(clean_name_for_search("john smith"), "john smith")
        
    def test_combined_operations(self):
        """Test all operations together."""
        self.assertEqual(clean_name_for_search("  J.R. SMITH  "), "jr smith")
        self.assertEqual(clean_name_for_search("Dr. Jane Doe, Ph.D."), "dr jane doe, phd")
        self.assertEqual(clean_name_for_search("O'CONNOR, J.T."), "o'connor, jt")
        
    def test_special_characters(self):
        """Test names with special characters."""
        self.assertEqual(clean_name_for_search("José Martínez"), "josé martínez")
        self.assertEqual(clean_name_for_search("Jean-Claude"), "jean-claude")
        self.assertEqual(clean_name_for_search("O'Hara"), "o'hara")
        
    def test_edge_cases(self):
        """Test edge cases."""
        self.assertEqual(clean_name_for_search(""), "")  # Empty string
        self.assertEqual(clean_name_for_search("."), "")  # Just a period
        self.assertEqual(clean_name_for_search("   "), "")  # Just spaces
        self.assertEqual(clean_name_for_search("..."), "")  # Multiple periods

if __name__ == '__main__':
    unittest.main()