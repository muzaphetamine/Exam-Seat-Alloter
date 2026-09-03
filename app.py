from flask import Flask, render_template, request, send_file, jsonify
import os
import tempfile
import shutil
from werkzeug.utils import secure_filename
import zipfile
from io import BytesIO
from script import process_session_files, generate_sessions_from_centralized
import pandas as pd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        mode = request.form.get('mode')
        temp_dir =tempfile.mkdtemp()
        input_dir =os.path.join(temp_dir, 'input')
        output_dir =os.path.join(temp_dir, 'output')
        os.makedirs(input_dir)
        os.makedirs(output_dir)
        
        if mode =='session-files':
            rooms_file =request.files.get('rooms')
            session_files =request.files.getlist('sessions')
            if not rooms_file or not session_files:
                return jsonify({'error': 'Missing required files'}), 400
            
            rooms_path = os.path.join(temp_dir, 'Rooms.xlsx')
            rooms_file.save(rooms_path)
            rooms_df = pd.read_excel(rooms_path)
            
            session_paths= []
            for f in session_files:
                if f.filename:
                    filepath = os.path.join(input_dir, secure_filename(f.filename))
                    f.save(filepath)
                    session_paths.append(filepath)
            results =process_session_files(session_paths, rooms_df, output_dir)
            
        elif mode =='centralized':
            rooms_file =request.files.get('rooms')
            students_file =request.files.get('students')
            schedule_file =request.files.get('schedule')

            if not all([rooms_file, students_file, schedule_file]):
                return jsonify({'error': 'Missing required centralized files'}), 400

            rooms_path =os.path.join(temp_dir, 'Rooms.xlsx')
            students_path =os.path.join(temp_dir, 'Students.xlsx')
            schedule_path =os.path.join(temp_dir, 'Schedule.xlsx')

            rooms_file.save(rooms_path)
            students_file.save(students_path)
            schedule_file.save(schedule_path)

            rooms_df= pd.read_excel(rooms_path)
            students_df =pd.read_excel(students_path)
            schedule_df = pd.read_excel(schedule_path)

            session_paths= generate_sessions_from_centralized(
                students_df,
                schedule_df,
                input_dir
            )
            if not session_paths:
                return jsonify({'error': 'No sessions generated from schedule'}), 400

            results= process_session_files(
                session_paths,
                rooms_df,
                output_dir
            )
  
        else:
            return jsonify({'error': 'Invalid mode'}), 400
        

        zip_buffer =BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path =os.path.join(root, file)
                    arcname =os.path.relpath(file_path, output_dir)
                    zip_file.write(file_path, arcname)
        zip_buffer.seek(0)
        
        shutil.rmtree(temp_dir)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='exam_allocations.zip'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)