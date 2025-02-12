import io
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from docx import Document
from docx.shared import RGBColor
import PyPDF2
from sentence_transformers import SentenceTransformer
from rapidfuzz import fuzz

app = Flask(__name__)
CORS(app)  


semantic_model = SentenceTransformer("all-mpnet-base-v2")

def extract_text_from_doc(file, file_type):
    """Extract text from a DOCX or PDF file."""
    if file_type == "docx":
        doc = Document(io.BytesIO(file.read()))
        return " ".join([para.text for para in doc.paragraphs])
    elif file_type == "pdf":
        pdf_reader = PyPDF2.PdfReader(io.BytesIO(file.read()))
        text = "".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
        return text.strip()
    return ""

def ai_word_compare(original, edited):
    """AI-powered word-level document comparison that detects insertions, deletions, and modifications."""
    original_words = original.split()
    edited_words = edited.split()

    differences = []
    original_index = 0
    edited_index = 0

    while original_index < len(original_words) or edited_index < len(edited_words):
        if original_index < len(original_words) and edited_index < len(edited_words):
            orig_word = original_words[original_index]
            edit_word = edited_words[edited_index]

            similarity = fuzz.ratio(orig_word, edit_word)  # Check word similarity

            if similarity > 85:  # If words are similar enough, consider them the same
                differences.append(orig_word)  # No change
                original_index += 1
                edited_index += 1
            else:
                # Word has changed → Mark the original as deleted and the edited as added
                differences.append(f"<del>{orig_word}</del> <ins>{edit_word}</ins>")
                original_index += 1
                edited_index += 1

        elif original_index < len(original_words):  # Remaining words in original → Deleted
            differences.append(f"<del>{original_words[original_index]}</del>")
            original_index += 1

        elif edited_index < len(edited_words):  # Remaining words in edited → Inserted
            differences.append(f"<ins>{edited_words[edited_index]}</ins>")
            edited_index += 1

    return " ".join(differences)



def create_highlighted_docx(highlighted_differences):
    """Generate a properly formatted DOCX file with text styling."""
    doc = Document()
    # doc.add_heading("Document Comparison Result", level=1)
    # doc.add_heading("Highlighted Differences", level=2)
    para = doc.add_paragraph()

    parts = highlighted_differences.replace("<del>", "[DEL]").replace("</del>", "[/DEL]") \
                                   .replace("<ins>", "[INS]").replace("</ins>", "[/INS]").split(" ")

    for word in parts:
        if word.startswith("[DEL]") and word.endswith("[/DEL]"):
            clean_word = word.replace("[DEL]", "").replace("[/DEL]", "")
            run = para.add_run(clean_word + " ")
            run.font.color.rgb = RGBColor(255, 0, 0)  # Red for deleted words
            run.font.strike = True  # Strikethrough for deleted words
        elif word.startswith("[INS]") and word.endswith("[/INS]"):
            clean_word = word.replace("[INS]", "").replace("[/INS]", "")
            run = para.add_run(clean_word + " ")
            run.font.color.rgb = RGBColor(0, 128, 0)  # Green for added words
            run.bold = True  # Bold for added words
        else:
            para.add_run(word + " ")  # Normal text

    file_path = "highlighted_comparison.docx"
    doc.save(file_path)
    return file_path

@app.route('/compare', methods=['POST'])
def compare_documents():
    file1 = request.files.get('file1')
    file2 = request.files.get('file2')

    if not file1 or not file2:
        return jsonify({"error": "Both files are required."}), 400

    file1_type = file1.filename.split('.')[-1].lower()
    file2_type = file2.filename.split('.')[-1].lower()
    
    if file1_type not in ['docx', 'pdf'] or file2_type not in ['docx', 'pdf']:
        return jsonify({"error": "Only DOCX and PDF files are supported."}), 400

    file1_content = extract_text_from_doc(file1, file1_type)
    file2_content = extract_text_from_doc(file2, file2_type)

    # Use AI for word-level comparison instead of difflib
    ai_highlighted_differences = ai_word_compare(file1_content, file2_content)

   

    highlighted_doc_path = create_highlighted_docx(ai_highlighted_differences)

    return jsonify({
        "origin_content" : file1_content,
        "highlighted_differences": ai_highlighted_differences,        
        "download_url": "/download"
    })

@app.route('/download', methods=['GET'])
def download_highlighted_docx():
    """Provide the highlighted DOCX file for download."""
    file_path = "highlighted_comparison.docx"
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
