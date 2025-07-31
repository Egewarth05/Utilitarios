from flask import Flask, render_template, request, send_file, send_from_directory, redirect, url_for, flash, session
from flask_session import Session
import os
import tempfile
import traceback
from werkzeug.utils import secure_filename
from utils.nf_comparador import processar_comparacao_nf
from utils.combustivel_processador import processar_combustivel 
from utils.ofx_processador import processar_ofx  
from utils.nf_comparador import extrair_notas_zip, extrair_relatorio, comparar_nfs
import json

app = Flask(__name__)
SETTINGS_PATH = os.path.join(app.root_path, 'combustivel_settings.json')
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

        # 1) salva uploads em disco
        rar_path = os.path.join(app.config['UPLOAD_FOLDER'], rar_file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        rar_file.save(rar_path)
        pdf_file.save(pdf_path)

        # valida
        if not os.path.isfile(rar_path):
            flash(f"Não encontrei o arquivo de notas “{rar_file.filename}”.")
            return redirect(url_for("nf_comparador"))
        if not os.path.isfile(pdf_path):
            flash(f"Não encontrei o relatório “{pdf_file.filename}”.")
            return redirect(url_for("nf_comparador"))

        # ** NOVO **: cria um diretório de saída próprio
        result_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'resultados_nf')
        os.makedirs(result_dir, exist_ok=True)

        try:
            resultado, pdf_path_saida = processar_comparacao_nf(
                rar_path,
                pdf_path,
                result_dir
            )
        except FileNotFoundError as e:
            flash(str(e))
            return redirect(url_for("nf_comparador"))
        except Exception as e:
            flash(f"Erro inesperado: {e}")
            return redirect(url_for("nf_comparador"))

        session['resultado']    = resultado
        # salva só o nome do PDF de saída, não do upload original
        session['pdf_filename'] = os.path.basename(pdf_path_saida)
        session['result_dir']   = 'resultados_nf'
        return redirect(url_for('nf_comparador'))

    return render_template(
        'nf_comparador.html',
        resultado    = session.get('resultado'),
        pdf_filename = session.get('pdf_filename'),
        result_dir   = session.get('result_dir'),
    )

@app.route('/relatorio-nf-pdf')
def relatorio_nf_pdf():
    filename   = session.get('pdf_filename')
    result_dir = session.get('result_dir', '')
    caminho    = os.path.join(app.config['UPLOAD_FOLDER'], result_dir, filename)
    return send_file(
        caminho,
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
    debug = None

    if request.method == "POST":
        banco   = request.form.get("banco")
        arquivo = request.files.get("ofx_file")

        # monte a string de debug
        debug = f"DEBUG → banco: {banco!r} | filename: {getattr(arquivo, 'filename', None)!r}"

        # validação: banco e nome do arquivo não podem estar vazios
        if not banco or not arquivo or arquivo.filename == "":
            return render_template(
                "ofx_processador.html",
                erro="Preencha todos os campos.",
                debug=debug
            )

        # salva e processa
        ofx_path = os.path.join(app.config["UPLOAD_FOLDER"], arquivo.filename)
        arquivo.save(ofx_path)
        base, ext       = os.path.splitext(arquivo.filename)
        output_filename = f"{base}_modificado{ext}"
        output_path     = os.path.join(app.config["UPLOAD_FOLDER"], output_filename)
        processar_ofx(ofx_path, output_path, banco)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename
        )

    # GET — renderiza sem erro, mas ainda passa debug (None)
    return render_template(
        "ofx_processador.html",
        erro=None,
        debug=debug
    )

@app.route('/combustivel', methods=['GET', 'POST'])
def combustivel():
    # 1) Carrega defaults do JSON
    defaults = {'gasolina': '', 'diesel': ''}
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                defaults.update(json.load(f))
        except:
            pass

    if request.method == 'POST':
        vg = request.form.get('gasolina')
        vd = request.form.get('diesel')
        file = request.files.get('csv_file')

        if not vg or not vd or not file:
            flash('Preencha todos os campos e escolha um CSV.')
            return redirect(request.url)

        # 2) Salva os novos defaults
        try:
            with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump({'gasolina': vg, 'diesel': vd}, f)
        except Exception as e:
            flash(f'Não foi possível salvar as configurações: {e}')
            # mas continua o processamento mesmo assim

        # 3) Cria pasta temporária e salva CSV
        tmp_dir = tempfile.mkdtemp()
        session['tmp_dir'] = tmp_dir
        nome_csv = secure_filename(file.filename)
        csv_path = os.path.join(tmp_dir, nome_csv)
        file.save(csv_path)

        # 4) Define saída e processa
        nome_xlsx = 'relatorio_combustivel.xlsx'
        out_path  = os.path.join(tmp_dir, nome_xlsx)
        try:
            processar_combustivel(csv_path, vg, vd, out_path)
        except Exception as e:
            print(traceback.format_exc())
            flash(f'Erro no processamento: {e}')
            return redirect(request.url)

        # 5) Renderiza com sucesso, repassando os defaults para manter no form
        return render_template(
            'combustivel.html',
            resultado=True,
            arquivo_saida=nome_xlsx,
            default_gasolina=vg,
            default_diesel=vd
        )

    # GET — só renderiza, passando os defaults lidos
    return render_template(
        'combustivel.html',
        default_gasolina=defaults['gasolina'],
        default_diesel=defaults['diesel']
    )

@app.route('/combustivel/download/<filename>')
def download_combustivel(filename):
    # Serve o arquivo gerado na pasta temporária
    tmp_dir = session.get('tmp_dir')
    if not tmp_dir:
        flash('Nenhum relatório disponível para download.')
        return redirect(url_for('combustivel'))
    return send_from_directory(
        directory=tmp_dir,
        path=filename,
        as_attachment=True
    )


if __name__ == '__main__':
    app.run(debug=True)
