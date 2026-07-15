import unittest

from scripts.pipeline import author_id_from_profile_url


class ProfileUrlTests(unittest.TestCase):
    def test_author_id_from_profile_url(self):
        url = "https://scholar.google.com/citations?hl=en&user=ABC123&view_op=list_works"
        self.assertEqual(author_id_from_profile_url(url), "ABC123")

    def test_profile_url_requires_user_parameter(self):
        with self.assertRaises(ValueError):
            author_id_from_profile_url("https://scholar.google.com/citations?hl=en")

    def test_profile_url_requires_google_scholar(self):
        with self.assertRaises(ValueError):
            author_id_from_profile_url("https://example.com/?user=ABC123")


if __name__ == "__main__":
    unittest.main()
