from flask import Blueprint, request, jsonify, send_file
import tempfile
import os
from app.core import extract_red_text, create_synonym_questions_from_red_text, HwpManager

bp = Blueprint("main", __name__)

@bp.route('/api/synonym', methods=['POST'])
def process_docx():
    if 'file' not in request.files:
        return jsonify({"error": "파일이 없습니다"}), 400

    file = request.files['file']
    if not file.filename.lower().endswith(('.hwp', '.hwpx', '.docx')):
        return jsonify({"error": "hwp, hwpx, docx 파일만 지원합니다"}), 400

    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1]) as tmp_input:
        file.save(tmp_input.name)

    try:
        hwp_path = tmp_input.name
        with HwpManager() as hwp:
            hwp.Open(hwp_path)
            all_results = extract_red_text(hwp)
            create_synonym_questions_from_red_text(hwp, all_results)
            output_file_path = os.path.splitext(hwp_path)[0] + "_synonym_questions.docx"
            hwp.SaveAs(output_file_path)
        return send_file(output_file_path, as_attachment=True)
    
    except Exception as e:
        if os.path.exists(tmp_input.name):
            os.remove(tmp_input.name)
        return jsonify({"error": str(e)}), 500
