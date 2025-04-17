import os
from subprocess import run
from flask import Flask, render_template, request, send_from_directory
from ad_utils import process_employee_data, create_ad_user, get_ad_departments
from excel_utils import generate_excel_files, generate_obaz_excel, copy_obaz_docx

app = Flask(__name__)

@app.route('/')
def main():
    return render_template('main.html')

@app.route('/create', methods=['GET', 'POST'])
def create():
    if request.method == 'POST':
        raw = request.form.to_dict()
        data = process_employee_data(raw)
        data['department_dn'] = raw.get('department_dn','')
        if create_ad_user(data):
            eisz, kmis = generate_excel_files(data)
            obaz_excel = generate_obaz_excel(data)
            #copy_obaz_docx(data)
            message = "Пользователь успешно создан в AD."
            files = [eisz, kmis, obaz_excel]
        else:
            message = "Ошибка при создании пользователя в AD."
            files = []
        return render_template('result.html',
                               message=message,
                               user=data,
                               file_links=files)
    else:
        deps = get_ad_departments()
        return render_template('create.html', departments=deps)

@app.route('/reset', methods=['GET','POST'])
def reset():
    if request.method == 'POST':
        fn = request.form.get('first_name')
        ln = request.form.get('last_name')
        # вызываем powershell
        cmd = ["powershell.exe","-ExecutionPolicy","Bypass","-File","./scripts/reset_password.ps1","-FirstName",fn,"-LastName",ln]
        res = run(cmd, capture_output=True, text=True, encoding='cp866', errors='replace')
        ok = (res.returncode==0)
        msg = res.stdout if ok else res.stdout or res.stderr
        return render_template('result.html', message=msg)
    return render_template('reset.html')

@app.route('/block', methods=['GET','POST'])
def block():
    if request.method=='POST':
        rd = request.form.to_dict()
        cmd = ["powershell.exe","-ExecutionPolicy","Bypass","-File","./scripts/block_user.ps1"]
        if rd.get('login'):
            cmd += ["-Login", rd['login']]
        else:
            cmd += ["-FirstName", rd.get('first_name',''), "-LastName", rd.get('last_name','')]
        res = run(cmd, capture_output=True, text=True, encoding='cp866', errors='replace')
        msg = res.stdout
        return render_template('result.html', message=msg)
    return render_template('block.html')

@app.route('/download_doc/<path:filename>')
def download_doc(filename):
    directory = os.path.join(app.root_path, 'templates', 'templates_doc')
    return send_from_directory(directory, filename, as_attachment=True)

@app.errorhandler(500)
def internal_server_error(error):
    app.logger.error(error)
    return render_template('500.html'), 500

@app.errorhandler(404)
def page_not_found(error):
    app.logger.warning(f"404: {request.path}")
    return render_template('404.html'), 404

if __name__ == '__main__':
    # слушать на всех интерфейсах, порт 80
    app.run(host="0.0.0.0", port=80, debug=False)
