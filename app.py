from flask import Flask, render_template, request, send_file, send_from_directory, redirect, url_for, flash, session
from flask_session import Session
import sys
import tempfile
import traceback
from werkzeug.utils import secure_filename
from utils.nf_comparador import processar_comparacao_nf
from utils.combustivel_processador import processar_combustivel
from utils.extrato_pdf_processador import processar_extrato_pdf
from utils.ofx_processador import processar_ofx
from utils.folha_processador import process_sheet 
import json
import os
import subprocess
from utils.nf_comparador import processar_comparacao_nf

app = Flask(__name__)
SETTINGS_PATH = os.path.join(app.root_path, 'combustivel_settings.json')
app.config['SESSION_TYPE']      = 'filesystem'
app.config['SESSION_FILE_DIR']  = './flask_session'
app.config['SESSION_PERMANENT'] = False
Session(app)

app.secret_key = 'Ic04854@'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route('/download_geral/<filename>') 
def download_geral(filename):

    return send_from_directory(
        directory=app.config['DOWNLOAD_FOLDER'],
        path=filename,
        as_attachment=True
    )

@app.route('/pagamentos', methods=['GET', 'POST'])
def pagamentos_processador():
    error = None
    success = None
    download_link = None

    if request.method == 'POST':
        # O template usa name="excel"
        file = request.files.get('excel')
        empresa = request.form.get('empresa')
        if not file or file.filename == '':
            error = 'Nenhum arquivo selecionado.'
        else:
            filename_secure = secure_filename(file.filename)
            input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_secure)
            file.save(input_filepath)

            output_filename = 'pagamentos_processados_final.xlsx'
            output_filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)

            # caminho absoluto customizado (ajuste conforme necessário)
            script_path = os.path.join(app.root_path, 'utils', 'processar_pagamentos.py')

            if not os.path.isfile(script_path):
                error = f"Script de processamento não encontrado em: {script_path}"
                print(error)
                return render_template('pagamentos.html', error=error, success=None, download_link=None)

            if os.path.exists(output_filepath):
                os.remove(output_filepath)

            try:
                result = subprocess.run(
                    [sys.executable, script_path, input_filepath, output_filepath, empresa],
                    capture_output=True,
                    text=True
                )

                print("stdout>", result.stdout)
                print("stderr>", result.stderr)
                
                # Se o script retornou erro ou não gerou o arquivo, trata como falha
                if result.returncode != 0:
                    detalhes = (result.stderr or result.stdout).strip()
                    error = f'Erro ao processar o arquivo: {detalhes}'
                    print("processar_pagamentos.py falhou:", detalhes)
                elif not os.path.exists(output_filepath):
                    error = 'O script rodou mas não gerou o arquivo de saída.' 
                    print("Aviso: saída esperada não encontrada em", output_filepath)
                else:
                    success = 'Arquivo processado com sucesso!'
                    download_link = url_for('download_geral', filename=output_filename)
                    print("processar_pagamentos.py saída:", result.stdout)
            except FileNotFoundError:
                error = 'Erro: o script "processar_pagamentos.py" não foi encontrado. Verifique o caminho.'
            except Exception as e:
                error = f'Ocorreu um erro inesperado: {e}'

    return render_template('pagamentos.html', error=error, success=success, download_link=download_link)

@app.route("/folha-pagamento", methods=["GET", "POST"])
def folha_pagamento():
    if request.method == "POST":
        csv_file = request.files.get("csv_file")
        generate_txt = request.form.get("generate_txt") == "on"

        if not csv_file or csv_file.filename == "":
            flash("Por favor, selecione um arquivo CSV.", "danger") # Adicionado categoria para o flash
            return redirect(request.url)

        try:
            # 1) Salva o arquivo CSV temporariamente
            tmp_dir = tempfile.mkdtemp()
            session['tmp_dir_folha'] = tmp_dir
            nome_csv = secure_filename(csv_file.filename)
            csv_path = os.path.join(tmp_dir, nome_csv)
            csv_file.save(csv_path)

            # 2) Define os caminhos de saída
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

@app.route('/nf-comparador', methods=['GET','POST'])
def nf_comparador():
    if request.method == 'POST':
        # 1. Busca os arquivos do form (names do template)
        rar_file = request.files.get('zip_file')
        pdf_file = request.files.get('relatorio_pdf')

        # 2. Validação imediata
        if not rar_file or rar_file.filename == '' or not pdf_file or pdf_file.filename == '':
            flash('Envie o RAR de Notas e o PDF de Relatório.', 'danger')
            return render_template('nf_comparador.html')

        # 3. Gera nomes seguros e paths
        filename_rar = secure_filename(rar_file.filename)
        filename_pdf = secure_filename(pdf_file.filename)
        rar_path = os.path.join(app.config['UPLOAD_FOLDER'], filename_rar)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename_pdf)

        # 4. Salva no disco
        rar_file.save(rar_path)
        pdf_file.save(pdf_path)

        # 5. Cria pasta de resultados
        result_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'resultados_nf')
        os.makedirs(result_dir, exist_ok=True)

        # 6. Chama o processador
        try:
            resultado, pdf_path_saida = processar_comparacao_nf(
                rar_path,
                pdf_path,
                result_dir
            )
            
            session['pdf_filename'] = os.path.basename(pdf_path_saida)
            session['result_dir']   = 'resultados_nf'
            
        except FileNotFoundError as e:
            flash(str(e), 'danger')
            return redirect(url_for('nf_comparador'))
        except Exception as e:
            flash(f'Erro inesperado: {e}', 'danger')
            return render_template('nf_comparador.html')

        # 7. Renderiza direto com o resultado
        return render_template(
            'nf_comparador.html',
            resultado    = resultado,
            pdf_filename = os.path.basename(pdf_path_saida),
            result_dir   = 'resultados_nf'
        )
    return render_template('nf_comparador.html')

@app.route('/relatorio-nf-pdf')
def relatorio_nf_pdf():
    filename   = session.get('pdf_filename')
    result_dir = session.get('result_dir', '')
    if not filename:
        flash('Nenhum relatório encontrado para download. Refaça a comparação.', 'danger')
        return redirect(url_for('nf_comparador'))
    caminho = os.path.join(app.config['UPLOAD_FOLDER'], result_dir, filename)
    if not os.path.exists(caminho):
        flash('Arquivo de relatório não está mais disponível. Refaça a comparação.', 'danger')
        return redirect(url_for('nf_comparador'))
    return send_file(
        caminho,
        as_attachment=True,
        download_name="relatorio_validacao.pdf"
    )

def processar_comparacao_nf_from_lists(notas_zip, relatorio_formatado, output_dir):
    # usa a função real que você importou acima
    resultado, pdf_path = processar_comparacao_nf(notas_zip, relatorio_formatado, output_dir)
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
            flash('Preencha todos os campos e escolha um CSV.', "danger")
            return redirect(request.url)

        # 2) Salva os novos defaults
        try:
            with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump({'gasolina': vg, 'diesel': vd}, f)
        except Exception as e:
            flash(f'Não foi possível salvar as configurações: {e}', "danger")
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
            flash(f'Erro no processamento: {e}', "danger")
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
        flash('Nenhum relatório disponível para download.', "danger")
        return redirect(url_for('combustivel'))
    return send_from_directory(
        directory=tmp_dir,
        path=filename,
        as_attachment=True
    )

@app.route('/extrato-pdf', methods=['GET','POST'], endpoint='extrato_pdf')
def extrato_pdf():
    resultado = None
    xlsx_name = txt_name = None
    gerar_txt = False

    if request.method == 'POST':
        pdf_file = request.files.get('pdf_file')
        gerar_txt = (request.form.get('gerar_txt') == 'sim')

        # Prefixo fixo "1". Os demais códigos seguem o padrão do processador
        # (meio="5" e cc="337") até você passar via planilha.
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
