from flask import Flask, render_template, request, send_file
import openpyxl
from reportlab.pdfgen import canvas
import io
from datetime import datetime, timedelta

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/atep_form.html', methods=['GET', 'POST'])
def atep_form():
    if request.method == 'POST':
        # Obtenha os dados do formulário de cada aba
        stop6 = request.form.getlist('stop6')
        zero6 = request.form.getlist('zero6')
        perigos_riscos = request.form.getlist('perigos_riscos')
        epis = request.form.getlist('epis')
        equipamentos = request.form.getlist('equipamentos')
        data_inicio = request.form.get('data_inicio')
        hora_inicio = request.form.get('hora_inicio')
        produtos_quimicos = request.form.get('produtos_quimicos')

        # Converta a data de início
        data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d').date()

        # Verifique se hora_inicio não é None antes de tentar convertê-lo
        try:
            hora_inicio = datetime.strptime(hora_inicio, '%H:%M').time()
        except ValueError:
            return "Erro: Formato de hora inválido. Use HH:MM."
        except TypeError:
            return "Erro: O campo hora_inicio é obrigatório."

        # Calcule a data de término (15 dias após a data de início)
        data_termino = data_inicio + timedelta(days=15)

        # Calcule a hora de término (10 horas após a hora de início)
        hora_termino = (datetime.combine(datetime.min, hora_inicio) + timedelta(hours=10)).time()

        # Crie o arquivo XLSX com os dados
        caminho_arquivo_xlsx = criar_arquivo_xlsx(
            stop6, zero6, perigos_riscos, epis, equipamentos,
            data_inicio, data_termino, hora_inicio, hora_termino,
            produtos_quimicos
        )

        # Gere o PDF a partir do XLSX (opcional, se você ainda quiser gerar o PDF)
        caminho_arquivo_pdf = gerar_pdf(caminho_arquivo_xlsx)

        # Envie o XLSX como resposta para o navegador
        return send_file(caminho_arquivo_xlsx, as_attachment=True, download_name='atep.xlsx')

    return render_template('atep_form.html')

def criar_arquivo_xlsx(stop6, zero6, perigos_riscos, epis, equipamentos, data_inicio, data_termino, hora_inicio, hora_termino, produtos_quimicos):
    caminho_modelo = 'arquivos/modelo_atep.xlsx'
    workbook = openpyxl.load_workbook(caminho_modelo)
    sheet = workbook.active

    if 'prensagem' in stop6:
        sheet['C16'] = 'X'
    if 'carga_pesada' in stop6:
        sheet['C17'] = 'X'
    if 'colisao' in stop6:
        sheet['C18'] = 'X'
    if 'queda_altura' in stop6:
        sheet['C19'] = 'X'
    if 'choque' in stop6:
        sheet['C20'] = 'X'
    if 'queimadura' in stop6:
        sheet['C21'] = 'X'

    if 'corte_perfuracao' in zero6:
        sheet['O16'] = 'X'
    if 'queda_material' in zero6:
        sheet['O17'] = 'X'
    if 'prensagem_zero' in zero6:
        sheet['O18'] = 'X'
    if 'desnivel' in zero6:
        sheet['O19'] = 'X'
    if 'batidas_contra' in zero6:
        sheet['O20'] = 'X'
    if 'projecoes' in zero6:
        sheet['O21'] = 'X'

    if 'pressoes_temperatura' in perigos_riscos:
        sheet['Y16'] = 'X'
    if 'fumos_poeira' in perigos_riscos:
        sheet['Y17'] = 'X'
    if 'produtos_inflamaveis' in perigos_riscos:
        sheet['Y18'] = 'X'
    if 'produtos_toxicos' in perigos_riscos:
        sheet['Y19'] = 'X'
    if 'ruido' in perigos_riscos:
        sheet['Y20'] = 'X'
    if 'incendio' in perigos_riscos:
        sheet['Y21'] = 'X'
    if 'contusao' in perigos_riscos:
        sheet['AJ16'] = 'X'
    if 'excesso_peso' in perigos_riscos:
        sheet['AJ17'] = 'X'
    if 'vazamento' in perigos_riscos:
        sheet['AJ18'] = 'X'
    if 'derramamento' in perigos_riscos:
        sheet['AJ19'] = 'X'
    if 'produtos_corrosivos' in perigos_riscos:
        sheet['AJ20'] = 'X'
    if 'outros_perigos' in perigos_riscos:
        sheet['AJ21'] = 'X'

    if 'mascara' in epis:
        sheet['B39'] = 'X'
    if 'mascara_solda' in epis:
        sheet['B40'] = 'X'
    if 'creme_protecao' in epis:
        sheet['B42'] = 'X'
    if 'luvas' in epis:
        sheet['N41'] = 'X'
    if 'avental' in epis:
        sheet['N42'] = 'X'
    if 'macacao' in epis:
        sheet['X39'] = 'X'
    if 'protetor_facial' in epis:
        sheet['X40'] = 'X'
    if 'cinto_seg_paraquedista' in epis:
        sheet['X42'] = 'X'
    if 'trava_quedas' in epis:
        sheet['AJ39'] = 'X'
    if 'cordas_trabalho_altura' in epis:
        sheet['AJ40'] = 'X'
    if 'mangotes' in epis:
        sheet['AJ41'] = 'X'
    if 'outros_epi' in epis:
        sheet['AJ42'] = 'X'
    

    if 'solda' in equipamentos:
        sheet['C23'] = 'X'
    if 'macarico' in equipamentos:
        sheet['C24'] = 'X'
    if 'compactadora' in equipamentos:
        sheet['C25'] = 'X'
    if 'lixadeira' in equipamentos:
        sheet['L23'] = 'X'
    if 'serra_marmore' in equipamentos:
        sheet['L24'] = 'X'
    if 'serra_circular' in equipamentos:
        sheet['L25'] = 'X'
    if 'furadeira' in equipamentos:
        sheet['T23'] = 'X'
    if 'serra_tico_tico' in equipamentos:
        sheet['T24'] = 'X'
    if 'ferramentas_manuais' in equipamentos:
        sheet['T25'] = 'X'
    if 'escada' in equipamentos:
        sheet['AB23'] = 'X'
    if 'andaime' in equipamentos:
        sheet['AB24'] = 'X'
    if 'plataforma' in equipamentos:
        sheet['AB25'] = 'X'
    if 'talha' in equipamentos:
        sheet['AK23'] = 'X'
    if 'serra_cliper' in equipamentos:
        sheet['AK24'] = 'X'
    if 'outros_equipamentos' in equipamentos:
        sheet['AK25'] = 'X'

    # Preenche as células com os dados dos responsáveis
    sheet['C75'] = resp_int
    sheet['O75'] = set_resp_int
    sheet['X75'] = ramal_resp
    sheet['C73'] = sv
    sheet['O73'] = set_sv
    sheet['X73'] = ramal_sv
    
    sheet['AG12'] = data_inicio.strftime('%d/%m/%Y')
    sheet['AN12'] = data_termino.strftime('%d/%m/%Y')
    sheet['X12'] = hora_inicio.strftime('%H:%M')
    sheet['AB12'] = hora_termino.strftime('%H:%M')
    sheet['P8'] = request.form.get('empresa')  
    sheet['O77'] = request.form.get('empresa') 
    sheet['AC8'] = request.form.get('descricao') 
    sheet['B12'] = request.form.get('responsavel_ext') 
    sheet['T12'] = request.form.get('quantidade_tb')
    setor = request.form.get('setor') 
    sheet['B8'].value = request.form.get('setor') 

    # Lógica para preencher os responsáveis com base no setor
    if setor == "(P) Prensas":
        resp_int = 'Gustavo Henrique'
        set_resp_int = 'Eng. Manut.'
        ramal_resp = '1570'
        sv = 'Rafael C. Brossi'
        set_sv = 'Supervisor Manut.'
        ramal_sv = '1622'

    elif setor == "(W) Funilaria":
        resp_int = 'Gustavo Henrique'  
        set_resp_int = 'Eng. Manut.'
        ramal_resp = '1570'
        sv = 'Rafael C. Brossi'
        set_sv = 'Supervisor Manut.'
        ramal_sv = '1622'
    
    elif setor == "(T) Pintura":
        resp_int = 'Henrique Nascimento'  
        set_resp_int = 'Eng. Manut.'
        ramal_resp = '1570'
        sv = 'Alexandre Pereira'
        set_sv = 'Supervisor Manut.'
        ramal_sv = '1532'
    
    elif setor == "(A) Montagem":
        resp_int = 'Gustavo Henrique' 
        set_resp_int = 'Eng. Manut.'
        ramal_resp = '1570'
        sv = 'Marcelo Barbosa'
        set_sv = 'Supervisor Manut.'
        ramal_sv = '1534'
    else:
        # Define valores padrão caso o setor não corresponda a nenhum dos esperados
        resp_int = 'Error 404'
        set_resp_int = 'Error 404'
        ramal_resp = 'Error 404'
        sv = 'Error 404'
        set_sv = 'Error 404'
        ramal_sv = 'Error 404'




    workbook.save('atep.xlsx')
    return 'atep.xlsx'

def gerar_pdf(caminho_arquivo_xlsx):
    workbook = openpyxl.load_workbook(caminho_arquivo_xlsx)
    sheet = workbook.active

    # Crie um PDF com o mesmo layout do XLSX
    c = canvas.Canvas('atep.pdf')
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                c.drawString(cell.col_idx * 50, 800 - cell.row * 20, str(cell.value))
    c.save()
    return 'atep.pdf'

@app.route('/download/<nome_arquivo>')
def download_arquivo(nome_arquivo):
    return send_file(nome_arquivo, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
