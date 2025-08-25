from flask import Flask, render_template, request, send_file, send_from_directory, redirect, url_for, flash, session, abort
from flask_session import Session
import sys
import tempfile
import traceback
from werkzeug.utils import secure_filename
from utils.combustivel_processador import processar_combustivel
from utils.processar_fornecedores import processar_planilha_pagamentos_separado_custom
from utils.extrato_pdf_processador import processar_extrato_pdf
from utils.ofx_processador import processar_ofx
from utils.folha_processador import process_sheet
import json
import utils.caixa_financeiro as caixa_fin
import os
import subprocess

app = Flask(__name__)
BASE_DIR = app.root_path  # ex.: /home/UtilsInatec/mysite
SETTINGS_PATH = os.path.join(BASE_DIR, 'combustivel_settings.json')

app.config['SESSION_TYPE']      = 'filesystem'
app.config['SESSION_FILE_DIR']  = os.path.join(BASE_DIR, 'flask_session')
app.config['SESSION_PERMANENT'] = False
Session(app)

app.secret_key = 'Ic04854@'

UPLOAD_DIR   = os.path.join(BASE_DIR, 'uploads')
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'uploads')  # use a mesma pasta se quiser

app.config['UPLOAD_FOLDER']   = UPLOAD_DIR
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_DIR

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route('/download_geral/<path:filename>')
def download_geral(filename):
    directory = app.config['DOWNLOAD_FOLDER']
    file_path = os.path.join(directory, filename)
    if not os.path.isfile(file_path):
        # logs úteis para depuração
        try:
            print("DEBUG download_geral: não achei", file_path)
            print("DEBUG DOWNLOAD_FOLDER:", directory)
            print("DEBUG listdir:", os.listdir(directory))
        except Exception:
            pass
        flash('Arquivo não encontrado. Gere novamente.', 'danger')
        return redirect(url_for('pagamentos_processador'))
    return send_from_directory(directory=directory, path=filename, as_attachment=True)

@app.route('/pagamentos', methods=['GET', 'POST'])
def pagamentos_processador():
    error = None
    success = None
    download_link = None

    if request.method == 'POST':
        file = request.files.get('excel')
        empresa = (request.form.get('empresa') or '').strip()

        if not file or file.filename == '':
            error = 'Nenhum arquivo selecionado.'
        elif not empresa:
            error = 'Selecione a empresa.'
        else:
            filename_secure = secure_filename(file.filename)
            input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_secure)
            file.save(input_filepath)

            output_filename = 'pagamentos_processados_final.xlsx'
            output_filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)

            # remove saída antiga se existir
            if os.path.exists(output_filepath):
                os.remove(output_filepath)

            try:
                # chamada direta (sem subprocess)
                processar_planilha_pagamentos_separado_custom(
                    input_filepath,
                    output_filepath,
                    empresa
                )

                if not os.path.exists(output_filepath):
                    error = 'O processamento terminou, mas o arquivo de saída não foi gerado.'
                else:
                    success = 'Arquivo processado com sucesso!'
                    download_link = url_for('download_geral', filename=output_filename)

            except Exception as e:
                print(traceback.format_exc())
                error = f'Ocorreu um erro no processamento: {e}'

    return render_template('pagamentos.html', error=error, success=success, download_link=download_link)

@app.route('/pagamentos/download/<filename>')
def download_pagamentos(filename):
    tmp_dir = session.get('tmp_dir_pagamentos')
    if not tmp_dir or not os.path.exists(os.path.join(tmp_dir, filename)):
        flash('Nenhum relatório disponível para download.', 'danger')
        return redirect(url_for('pagamentos_processador'))
    return send_from_directory(
        directory=tmp_dir,
        path=filename,
        as_attachment=True
    )

@app.route("/folha-pagamento", methods=["GET", "POST"])
def folha_pagamento():
    if request.method == "POST":
        csv_file = request.files.get("csv_file")
        generate_txt = request.form.get("generate_txt") == "on"

        if not csv_file or csv_file.filename == "":
            flash("Por favor, selecione um arquivo CSV.", "danger")
            return redirect(request.url)

        try:
            tmp_dir = tempfile.mkdtemp()
            session['tmp_dir_folha'] = tmp_dir
            nome_csv = secure_filename(csv_file.filename)
            csv_path = os.path.join(tmp_dir, nome_csv)
            csv_file.save(csv_path)

            base_name = os.path.splitext(nome_csv)[0]
            output_xlsx_name = f"{base_name}_processado.xlsx"
            output_xlsx_path = os.path.join(tmp_dir, output_xlsx_name)

            output_txt_name = None
            output_txt_path = None
            if generate_txt:
                output_txt_name = f"{base_name}_processado.txt"
                output_txt_path = os.path.join(tmp_dir, output_txt_name)

            process_sheet(csv_path, output_xlsx_path, output_txt_path)

            session['output_xlsx_name'] = output_xlsx_name
            session['output_txt_name'] = output_txt_name if generate_txt else None
            flash("Folha de pagamento processada com sucesso!", "success")
            return render_template(
                'folha_processador.html',
                resultado=True,
                output_xlsx_name=output_xlsx_name,
                output_txt_name=output_txt_name
            )
        except Exception as e:
            print(traceback.format_exc())
            flash(f"Erro ao processar a folha de pagamento: {e}", "danger")
            return redirect(request.url)

    return render_template("folha_processador.html")

@app.route('/folha-pagamento/download/<filename>')
def download_folha_pagamento(filename):
    tmp_dir = session.get('tmp_dir_folha')
    if not tmp_dir or not os.path.exists(os.path.join(tmp_dir, filename)):
        flash('Nenhum relatório disponível para download ou arquivo não encontrado.', 'danger')
        return redirect(url_for('folha_pagamento'))
    return send_from_directory(
        directory=tmp_dir,
        path=filename,
        as_attachment=True
    )

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

        debug = f"DEBUG → banco: {banco!r} | filename: {getattr(arquivo, 'filename', None)!r}"

        if not banco or not arquivo or arquivo.filename == "":
            return render_template("ofx_processador.html", erro="Preencha todos os campos.", debug=debug)

        # --- paths seguros ---
        in_name  = secure_filename(arquivo.filename)
        ofx_path = os.path.join(app.config["UPLOAD_FOLDER"], in_name)
        arquivo.save(ofx_path)

        base, ext = os.path.splitext(in_name)
        out_name  = secure_filename(f"{base}_modificado{ext.lower()}")
        output_path = os.path.join(app.config["UPLOAD_FOLDER"], out_name)

        # --- chama o processador; ideal que ele retorne o caminho real de saída ---
        real_out = processar_ofx(ofx_path, output_path, banco) or output_path

        # fallback: se não criou onde esperávamos, tenta localizar
        if not os.path.exists(real_out):
            # procura algo tipo base_modificado.*
            prefix = f"{base}_modificado"
            candidates = [fn for fn in os.listdir(app.config["UPLOAD_FOLDER"]) if fn.startswith(prefix)]
            if candidates:
                real_out = os.path.join(app.config["UPLOAD_FOLDER"], candidates[0])

        if not os.path.exists(real_out):
            return render_template("ofx_processador.html",
                                   erro=f"Arquivo de saída não encontrado: {real_out}. Verifique o processar_ofx.",
                                   debug=debug)

        return send_file(real_out, as_attachment=True, download_name=os.path.basename(real_out))

    return render_template("ofx_processador.html", erro=None, debug=debug)


@app.route('/combustivel', methods=['GET', 'POST'])
def combustivel():
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
            flash('Preencha todos os campos e escolha um CSV.', "danger")
            return redirect(request.url)

        try:
            with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump({'gasolina': vg, 'diesel': vd}, f)
        except Exception as e:
            flash(f'Não foi possível salvar as configurações: {e}', "danger")

        tmp_dir = tempfile.mkdtemp()
        session['tmp_dir'] = tmp_dir
        nome_csv = secure_filename(file.filename)
        csv_path = os.path.join(tmp_dir, nome_csv)
        file.save(csv_path)

        nome_xlsx = 'relatorio_combustivel.xlsx'
        out_path  = os.path.join(tmp_dir, nome_xlsx)
        try:
            processar_combustivel(csv_path, vg, vd, out_path)
        except Exception as e:
            print(traceback.format_exc())
            flash(f'Erro no processamento: {e}', "danger")
            return redirect(request.url)

        return render_template(
            'combustivel.html',
            resultado=True,
            arquivo_saida=nome_xlsx,
            default_gasolina=vg,
            default_diesel=vd
        )

    return render_template(
        'combustivel.html',
        default_gasolina=defaults['gasolina'],
        default_diesel=defaults['diesel']
    )

@app.route('/combustivel/download/<filename>')
def download_combustivel(filename):
    tmp_dir = session.get('tmp_dir')
    if not tmp_dir:
        flash('Nenhum relatório disponível para download.', "danger")
        return redirect(url_for('combustivel'))
    return send_from_directory(
        directory=tmp_dir,
        path=filename,
        as_attachment=True
    )

@app.route('/resumo-contas', methods=['GET', 'POST'])
def resumo_contas():
    resultado = None

    if request.method == 'POST':
        arquivo = request.files.get('txt_file')
        decimal = request.form.get('decimal') or 'comma'

        if not arquivo or arquivo.filename.strip() == '':
            flash('Selecione um arquivo TXT.', 'danger')
            return render_template('caixa_financeiro.html', resultado=None)

        try:
            tmp_dir = tempfile.mkdtemp()
            session['tmp_dir_resumo'] = tmp_dir

            in_name = secure_filename(arquivo.filename)
            in_path = os.path.join(tmp_dir, in_name)
            arquivo.save(in_path)

            base, _ = os.path.splitext(in_name)
            out_name = f"{base}_importacao.txt"
            out_path = os.path.join(tmp_dir, out_name)

            # ✅ chama a função correta do módulo utils.caixa_financeiro
            ret = caixa_fin.processar_resumo_contas(in_path, out_path, decimal=decimal)

            flash('Arquivo gerado com sucesso!', 'success')
            resultado = {
                'linhas': ret.get('linhas', 0),
                'filename': out_name
            }
        except Exception as e:
            print(traceback.format_exc())
            flash(f'Erro ao processar: {e}', 'danger')

    # ✅ renderiza o template correto
    return render_template('caixa_financeiro.html', resultado=resultado)

@app.route('/resumo-contas/download/<filename>')
def download_resumo_contas(filename):
    tmp_dir = session.get('tmp_dir_resumo')
    if not tmp_dir or not os.path.exists(os.path.join(tmp_dir, filename)):
        from flask import flash, redirect, url_for
        flash('Nenhum arquivo para download. Envie o TXT novamente.', 'danger')
        return redirect(url_for('resumo_contas'))
    from flask import send_from_directory
    return send_from_directory(directory=tmp_dir, path=filename, as_attachment=True)


@app.route('/extrato-pdf', methods=['GET','POST'], endpoint='extrato_pdf')
def extrato_pdf():
    resultado = None
    xlsx_name = txt_name = None
    gerar_txt = False

    if request.method == 'POST':
        pdf_file = request.files.get('pdf_file')
        gerar_txt = (request.form.get('gerar_txt') == 'sim')

        cfg = {'codigo_prefixo': '1'}

        if not pdf_file or pdf_file.filename == '':
            flash('Envie o PDF do extrato.', 'danger')
            return render_template('extrato_pdf.html', gerar_txt=gerar_txt)

        in_dir = app.config.get('UPLOAD_FOLDER', 'uploads')
        out_dir = app.config.get('DOWNLOAD_FOLDER', in_dir)
        os.makedirs(in_dir, exist_ok=True)
        os.makedirs(out_dir, exist_ok=True)

        in_path = os.path.join(in_dir, 'extrato.pdf')
        xlsx_name = 'extrato_processado.xlsx'
        xlsx_path = os.path.join(out_dir, xlsx_name)
        txt_name = 'extrato_processado.txt' if gerar_txt else None
        txt_path = os.path.join(out_dir, txt_name) if gerar_txt else None

        pdf_file.save(in_path)

        try:
            ret = processar_extrato_pdf(in_path, xlsx_path, txt_path, cfg)
            resultado = {'qtd': ret.get('quantidade_lancamentos', 0)}
            flash('Extrato processado com sucesso!', 'success')
        except Exception as e:
            flash(f'Erro ao processar extrato: {e}', 'danger')

    return render_template('extrato_pdf.html',
                           resultado=resultado,
                           xlsx_name=xlsx_name,
                           txt_name=txt_name,
                           gerar_txt=gerar_txt)

if __name__ == '__main__':
    app.run(debug=True)


