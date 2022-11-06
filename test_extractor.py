import unittest
from extractor import extract, patterns


class ExtractionTests(unittest.TestCase):
    def setUp(self):
        self.file = open(r"sample.txt", "r")
        self.test_data_a = self.file.read()

    def tearDown(self):
        self.file.close()

    def test_extracts_last_first_name(self):
        extracted_data = extract(
            self.test_data_a,
            last_first_name=patterns["last_first_name"],
        )
        print(extracted_data)
        self.assertTrue(bool(extracted_data))


if __name__ == "__main__":
    unittest.main()

