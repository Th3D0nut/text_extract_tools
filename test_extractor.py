import unittest
from extractor import extract, patterns


names = [
    ('van der Brug', 'Kim'), ('Es Ouasilla', 'Fatima'), ('Kakelstaart', 'Dom'),
    ('Taartén', 'Ølaf'), ('es Doorda', 'Tequilla rocha Dos'),
    ("ñan De Ñie Stàarten Stàat", "Ðenk")
]


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
        )["last_first_name"]
        self.assertEqual(extracted_data, names)


if __name__ == "__main__":
    unittest.main()

