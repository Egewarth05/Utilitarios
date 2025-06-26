from flask import Flask, render_template, request, send_file, redirect, url_for, flash, session
from flask_session import Session
import os
from utils.nf_comparador import processar_comparacao_nf
from utils.combustivel_processador import processar_combustivel 
from utils.ofx_processador import processar_ofx  
from utils.nf_comparador import extrair_notas_zip, extrair_relatorio, comparar_nfs

app = Flask(__name__)
app.config['SESSION_TYPE']      = 'filesystem'
app.config['SESSION_FILE_DIR']  = './flask_session'
app.config['SESSION_PERMANENT'] = False
Session(app)

app.secret_key = 'Ic04854@'
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route('/nf-comparador', methods=['GET','POST'])
def nf_comparador():
    if request.method == 'POST':
        rar_file = request.files['zip_file']
        pdf_file = request.files['relatorio_pdf']

        # 1) salva os uploads em disco
        rar_path = os.path.join(app.config['UPLOAD_FOLDER'], rar_file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        rar_file.save(rar_path)
        pdf_file.save(pdf_path)

        # 2) processa TUDO num só passo
        resultado, pdf_path_saida = processar_comparacao_nf(
            rar_path, pdf_path, app.config['UPLOAD_FOLDER']
        )

        # 3) guarda só o que precisa na sessão
        session['resultado']    = resultado
        session['pdf_filename'] = os.path.basename(pdf_path)

        return redirect(url_for('nf_comparador'))

    # ==== GET ====
    return render_template(
      'nf_comparador.html',
      resultado    = session.get('resultado'),
      pdf_filename = session.get('pdf_filename'),
    )

@app.route('/relatorio-nf-pdf')
def relatorio_nf_pdf():
    filename = session.get('pdf_filename')
    return send_file(
        os.path.join(app.config['UPLOAD_FOLDER'], filename),
        as_attachment=True,
        download_name="relatorio_validacao.pdf"
    )

def processar_comparacao_nf_from_lists(notas_zip, relatorio_formatado, output_dir):
    resultado = comparar_nfs(notas_zip, relatorio_formatado, output_dir)
    pdf_path  = resultado.pop("pdf")
    return resultado, pdf_path

@app.route('/download/<filename>')
def download_relatorio(filename):
    caminho = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(caminho, as_attachment=True)

@app.route("/ofx-processador", methods=["GET", "POST"])
def ofx_processador():
    resultado = None
    if request.method == 'POST':
        banco = request.form.get('banco')
        arquivo = request.files.get('ofx_file')
        if not banco or not arquivo:
            return render_template('ofx_processador.html', erro="Preencha todos os campos.")

        ofx_path = os.path.join(app.config['UPLOAD_FOLDER'], arquivo.filename)
        arquivo.save(ofx_path)

        output_path = ofx_path.replace('.ofx', '_modificado.ofx')
        processar_ofx(banco, ofx_path, output_path)
        return send_file(output_path, as_attachment=True)

    return render_template('ofx_processador.html')

if __name__ == '__main__':
    app.run(debug=True)
