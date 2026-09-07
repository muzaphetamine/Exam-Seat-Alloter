from flask import Flask, render_template, request, send_file, jsonify
import os
import tempfile
import shutil
import zipfile
from io import BytesIO
from werkzeug.utils import secure_filename
from script import (
    process_session_files,
    generate_sessions_from_centralized,
    read_rooms_from_files,
    read_centralized_students_from_files,
    read_centralized_schedules_from_files,
    ingest_session_files,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=50*1024*1024


def save_uploaded_files(files, folder):
    paths=[]
    for index, file in enumerate(files):
        if not file or not file.filename:
            continue
        filename=secure_filename(file.filename)
        if not filename:
            continue
        filename=f"{index + 1}_{filename}"
        path=os.path.join(folder, filename)
        file.save(path)
        paths.append(path)
    return paths


def make_zip(output_dir):
    zip_buffer=BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, dirs, files in os.walk(output_dir):
            for filename in files:
                file_path =os.path.join(root, filename)
                arcname =os.path.relpath(file_path, output_dir)
                zip_file.write(file_path, arcname)
    zip_buffer.seek(0)
    return zip_buffer


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    temp_dir=None
    try:
        mode=request.form.get('mode')
        temp_dir= tempfile.mkdtemp(prefix='exam_seat_alloter_')
        input_dir =os.path.join(temp_dir, 'input')
        output_dir =os.path.join(temp_dir, 'output')
        rooms_dir =os.path.join(input_dir, 'rooms')
        students_dir=os.path.join(input_dir, 'students')
        schedules_dir=os.path.join(input_dir, 'schedules')
        sessions_dir =os.path.join(input_dir, 'sessions')
        for directory in (rooms_dir, students_dir, schedules_dir, sessions_dir, output_dir):
            os.makedirs(directory, exist_ok=True)
        if mode=='session-files':
            room_files =save_uploaded_files(request.files.getlist('rooms'), rooms_dir)
            session_files =save_uploaded_files(request.files.getlist('sessions'), sessions_dir)
            if not room_files:
                return jsonify({'error': 'Please upload at least one room Excel file.'}), 400
            if not session_files:
                return jsonify({'error': 'Please upload at least one session Excel file.'}), 400
            rooms_df =read_rooms_from_files(room_files)
            ingest_session_files(session_files)
            results =process_session_files(session_files, rooms_df, output_dir)
        elif mode=='centralized':
            room_files =save_uploaded_files(request.files.getlist('rooms'), rooms_dir)
            student_files =save_uploaded_files(request.files.getlist('students'), students_dir)
            schedule_files =save_uploaded_files(request.files.getlist('schedule'), schedules_dir)
            if not room_files:
                return jsonify({'error': 'Please upload at least one room Excel file.'}), 400
            if not student_files:
                return jsonify({'error': 'Please upload at least one student Excel file.'}), 400
            if not schedule_files:
                return jsonify({'error': 'Please upload at least one schedule Excel file.'}), 400
            rooms_df =read_rooms_from_files(room_files)
            students_df =read_centralized_students_from_files(student_files)
            schedule_df =read_centralized_schedules_from_files(schedule_files)
            session_paths =generate_sessions_from_centralized(students_df, schedule_df, sessions_dir)
            if not session_paths:
                return jsonify({
                    'error': 'No sessions could be generated from the supplied schedule and student files.'
                }), 400
            results =process_session_files(session_paths, rooms_df, output_dir)
        else:
            return jsonify({'error': 'Invalid mode'}), 400
        zip_buffer=make_zip(output_dir)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='exam_allocations.zip'
        )
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': f'Processing failed: {exc}'}), 500
    finally:
        if temp_dir and os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
