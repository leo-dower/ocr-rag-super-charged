
import os

class MarkdownFormatter:
    @staticmethod
    def save_as_markdown(output_dir: str, file_name: str, text: str):
        md_file_path = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}.md")
        with open(md_file_path, 'w', encoding='utf-8') as f:
            f.write(text)
