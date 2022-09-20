import docx
import sys
import os
import yaml

CONFIG_FILE = os.path.join(os.getcwd(), 'config.yaml')


def help() -> None:
    print('Usage: python3 main.py <file> <position> <company_name>')
    print('Example: python3 main.py resume.docx "software developer" Google')


def replace_text(doc: docx.Document, old: str, new: str) -> None:
    for p in doc.paragraphs:
        if old in p.text:
            inline = p.runs
            # Replace strings
            for i in range(len(inline)):
                if old in inline[i].text:
                    text = inline[i].text.replace(old, new)
                    inline[i].text = text


if __name__ == '__main__':
    if len(sys.argv) < 2:
        help()
        sys.exit(1)

    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = yaml.safe_load(f)
    else:
        config = {
            'position_placeholder': '*',
            'company_name_placeholder': '~'
        }

    document: docx.Document = docx.Document(sys.argv[1])
    position: str = sys.argv[2]
    company_name: str = sys.argv[3]

    replace_text(document, config['position_placeholder'], position)
    replace_text(document, config['company_name_placeholder'], company_name)

    document.save(sys.argv[1].replace('.docx', '_edited.docx'))
