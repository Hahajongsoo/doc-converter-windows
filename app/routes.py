from flask import Blueprint, request, jsonify, send_file
import tempfile
import os
from app.core import extract_red_text, create_synonym_questions_from_red_text, save_as_hwp, hwp_to_docx

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
        # hwp/hwpx 파일인 경우 docx로 변환
        if tmp_input.name.lower().endswith(('.hwp', '.hwpx')):
            docx_path = hwp_to_docx(tmp_input.name)
            os.remove(tmp_input.name)  # 원본 임시 파일 삭제
        else:
            docx_path = tmp_input.name

        # 빨간 글자 추출 및 문제 생성
        red_text_groups = extract_red_text(docx_path)
        if not red_text_groups:
            return jsonify({"error": "빨간 글자를 찾지 못했습니다"}), 400

        # 문제 생성 및 hwp 변환
        docx_output_path = create_synonym_questions_from_red_text(docx_path, red_text_groups)
        hwp_output_path = save_as_hwp(docx_output_path)

        # 임시 파일 정리
        if docx_path != tmp_input.name:  # hwp/hwpx에서 변환된 경우
            os.remove(docx_path)
        os.remove(docx_output_path)

        return send_file(hwp_output_path, as_attachment=True)
    except Exception as e:
        # 오류 발생 시 임시 파일 정리
        if os.path.exists(tmp_input.name):
            os.remove(tmp_input.name)
        return jsonify({"error": str(e)}), 500
