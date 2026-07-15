import unittest

from scripts.classifica_articole import clasifica_denumirea


class ClassificationTests(unittest.TestCase):
    def test_keywords_are_case_insensitive(self):
        self.assertEqual(clasifica_denumirea("International CONFERENCE on Energy"), "C")
        self.assertEqual(clasifica_denumirea("Annual symposium"), "C")
        self.assertEqual(clasifica_denumirea("Proceedings of the IEEE"), "C")

    def test_parenthesized_acronym(self):
        self.assertEqual(
            clasifica_denumirea("Advanced Topics in Electrical Engineering (ATEE)"),
            "C",
        )

    def test_journal_and_empty_values(self):
        self.assertEqual(clasifica_denumirea("Sensors"), "J")
        self.assertEqual(clasifica_denumirea(""), "")
        self.assertEqual(clasifica_denumirea(None), "")


if __name__ == "__main__":
    unittest.main()
