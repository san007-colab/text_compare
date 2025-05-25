import os
import re
from flask import Flask, render_template, request, url_for
from werkzeug.utils import secure_filename
from docx import Document
from bs4 import BeautifulSoup
from difflib import SequenceMatcher

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
COMPARE_FOLDER = 'static/comparisons'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(COMPARE_FOLDER, exist_ok=True)

def split_sentences(text):
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z])', text)
    return [s.strip() for s in sentences if s.strip()]

def read_docx(path):
    doc = Document(path)
    full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    return split_sentences(full_text)

def read_html(path):
    with open(path, encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')
        for tag in soup(['script', 'style', 'noscript', 'header', 'footer', 'nav']):
            tag.decompose()
        text = soup.get_text(separator='\n')
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        return split_sentences(" ".join(lines))

def normalize_number(num_str):
    try:
        return float(num_str)
    except:
        return None

def highlight_differences(doc_sent, html_sent):
    doc_tokens = re.findall(r'\b\w+[\.\w]*\b|[^\w\s]', doc_sent)
    html_tokens = re.findall(r'\b\w+[\.\w]*\b|[^\w\s]', html_sent)

    highlighted_doc = []
    highlighted_html = []

    max_len = max(len(doc_tokens), len(html_tokens))
    for i in range(max_len):
        doc_token = doc_tokens[i] if i < len(doc_tokens) else None
        html_token = html_tokens[i] if i < len(html_tokens) else None

        if doc_token is None:
            highlighted_html.append(f'<span class="extra">{html_token}</span>')
            continue
        if html_token is None:
            highlighted_doc.append(f'<span class="missing">{doc_token}</span>')
            continue

        if doc_token == html_token:
            highlighted_doc.append(doc_token)
            highlighted_html.append(html_token)
        else:
            if doc_token.lower() == html_token.lower():
                highlighted_doc.append(f'<span class="case-diff">{doc_token}</span>')
                highlighted_html.append(f'<span class="case-diff">{html_token}</span>')
            else:
                doc_num = normalize_number(doc_token)
                html_num = normalize_number(html_token)
                if doc_num is not None and html_num is not None and doc_num == html_num:
                    highlighted_doc.append(f'<span class="decimal-diff">{doc_token}</span>')
                    highlighted_html.append(f'<span class="decimal-diff">{html_token}</span>')
                else:
                    highlighted_doc.append(f'<span class="diff">{doc_token}</span>')
                    highlighted_html.append(f'<span class="diff">{html_token}</span>')

    return " ".join(highlighted_doc), " ".join(highlighted_html)

def match_sentences_full(doc_sentences, html_sentences, threshold=0.3):
    used_html_indexes = set()
    pairs = []

    for doc_sent in doc_sentences:
        best_score = 0
        best_index = -1
        for i, html_sent in enumerate(html_sentences):
            if i in used_html_indexes:
                continue
            score = SequenceMatcher(None, doc_sent, html_sent).ratio()
            if score > best_score:
                best_score = score
                best_index = i

        if best_score >= threshold and best_index != -1:
            html_sent = html_sentences[best_index]
            used_html_indexes.add(best_index)
            h_doc, h_html = highlight_differences(doc_sent, html_sent)
            pairs.append((h_doc, h_html, False))
        else:
            pairs.append((f'<span class="missing">{doc_sent}</span>', '', True))

    for i, html_sent in enumerate(html_sentences):
        if i not in used_html_indexes:
            pairs.append(('', f'<span class="extra">{html_sent}</span>', False))

    return pairs

def save_comparison_page(filename, matched_pairs):
    compare_filename = f"{filename}_compare.html"
    compare_path = os.path.join(COMPARE_FOLDER, compare_filename)
    with open(compare_path, 'w', encoding='utf-8') as f:
        f.write(render_template('comparison.html', filename=filename, pairs=matched_pairs))
    return compare_filename

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        docx_files = request.files.getlist('docx_files')
        html_files = request.files.getlist('html_files')

        docx_map = {os.path.splitext(secure_filename(f.filename))[0]: f for f in docx_files}
        html_map = {os.path.splitext(secure_filename(f.filename))[0]: f for f in html_files}
        matched_keys = set(docx_map) & set(html_map)

        results = []
        for key in matched_keys:
            docx_file = docx_map[key]
            html_file = html_map[key]
            docx_path = os.path.join(UPLOAD_FOLDER, docx_file.filename)
            html_path = os.path.join(UPLOAD_FOLDER, html_file.filename)
            docx_file.save(docx_path)
            html_file.save(html_path)

            doc_sentences = read_docx(docx_path)
            html_sentences = read_html(html_path)
            matched_pairs = match_sentences_full(doc_sentences, html_sentences)
            compare_filename = save_comparison_page(key, matched_pairs)

            results.append({
                'name': key,
                'url': url_for('static', filename=f'comparisons/{compare_filename}')
            })

        return render_template('results.html', results=results)

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
