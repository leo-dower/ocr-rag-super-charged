

import unittest
from src.utils.json_formatter import JsonFormatter

class TestJsonFormatter(unittest.TestCase):

    def test_create_mistral_entry(self):
        text = "This is a test."
        paragraphs = [("This is a paragraph.", "normal")]
        entry = JsonFormatter.create_mistral_entry(text, paragraphs)
        self.assertIsNotNone(entry)
        self.assertIn("messages", entry)
        self.assertEqual(len(entry["messages"]), 2)
        self.assertEqual(entry["messages"][0]["role"], "user")
        self.assertEqual(entry["messages"][1]["role"], "assistant")

    def test_sanitize_text(self):
        self.assertEqual(JsonFormatter.sanitize_text("  Test\nString  "), "Test String")

if __name__ == '__main__':
    unittest.main()

