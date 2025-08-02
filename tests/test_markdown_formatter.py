
import unittest
import os
from src.utils.markdown_formatter import MarkdownFormatter

class TestMarkdownFormatter(unittest.TestCase):

    def test_save_as_markdown(self):
        output_dir = "tests/temp"
        os.makedirs(output_dir, exist_ok=True)
        file_name = "test.txt"
        text = "This is a test."
        MarkdownFormatter.save_as_markdown(output_dir, file_name, text)
        md_file_path = os.path.join(output_dir, "test.md")
        self.assertTrue(os.path.exists(md_file_path))
        with open(md_file_path, 'r', encoding='utf-8') as f:
            self.assertEqual(f.read(), text)
        os.remove(md_file_path)
        os.rmdir(output_dir)

if __name__ == '__main__':
    unittest.main()
