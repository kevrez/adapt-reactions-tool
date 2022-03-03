import unittest
import xlrd

from adapt_reactions_parser import AdaptReactionsParser


INVALID_FILETYPE_PATH = r'"C:\Users\kreznicek\Desktop\KR-Programs\Adapt-Reactions\test-reports\Invalid_Filetype.pdf"'
RUN_MISSING_REACTIONS_PATH = r'"C:\Users\kreznicek\Desktop\KR-Programs\Adapt-Reactions\test-reports\Invalid_Report.xls"'
SKIPPED_RUN_PATH = r'"C:\Users\kreznicek\Desktop\KR-Programs\Adapt-Reactions\test-reports\Skipped_Live_Loads.xls"'
VALID_RUN_PATH = r'"C:\Users\kreznicek\Desktop\KR-Programs\Adapt-Reactions\test-reports\Valid_Reactions.xls"'
VALID_RUN_REACTIONS = ([8.93, 29.36, 6.17],
                       [0.5, 8.47, 6.81],
                       [5.14, 37.0, 16.77])


class TestAdaptReactionsParser(unittest.TestCase):

    def test_reactions_parsing(self):
        self.assertEqual(AdaptReactionsParser._reactions(
            VALID_RUN_PATH), VALID_RUN_REACTIONS)

    def test_error_with_skipped_run(self):
        with self.assertRaises(ValueError):
            AdaptReactionsParser._reactions(SKIPPED_RUN_PATH)

    def test_error_on_invalid_path(self):
        with self.assertRaises(FileNotFoundError):
            AdaptReactionsParser._reactions(INVALID_FILETYPE_PATH)

    def test_error_on_invalid_spreadsheet(self):
        with self.assertRaises(xlrd.biffh.XLRDError):
            AdaptReactionsParser._reactions(RUN_MISSING_REACTIONS_PATH)


if __name__ == '__main__':
    unittest.main()
