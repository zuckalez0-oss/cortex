# app.py

from flask import Flask, render_template, request, json, session, make_response
from calculo_cortes import calcular_plano_de_corte
import random
import pandas as pd
from weasyprint import HTML
from io import BytesIO
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui_pode_ser_qualquer_coisa'

CORES_PECAS = ["#FF6347", "#4682B4", "#32CD32", "#FFD700", "#6A5ACD", "#40E0D0", "#FF69B4", "#DAA520", "#8A2BE2", "#00BFFF", "#7FFF00", "#DC143C", "#FF8C00", "#ADFF2F", "#BA55D3"]
def gerar_cor_aleatoria(): return f"#{random.randint(0, 0xFFFFFF):06x}"

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            chapa_largura = int(request.form['chapa_largura'])
            chapa_altura = int(request.form['chapa_altura'])

            pecas = []
            i = 0
            while f'peca_{i}_largura' in request.form:
                largura = request.form.get(f'peca_{i}_largura')
                altura = request.form.get(f'peca_{i}_altura')
                qtd = request.form.get(f'peca_{i}_quantidade')
                if largura and altura and qtd:
                    pecas.append({'largura': int(largura), 'altura': int(altura), 'quantidade': int(qtd)})
                i += 1

            if not pecas:
                return render_template('index.html', erro="Nenhuma peça foi adicionada.")

            legenda = []
            cores_por_tipo = {}
            cor_index = 0
            for peca in pecas:
                tipo_peca_key = f"{peca['largura']}x{peca['altura']}"
                if tipo_peca_key not in cores_por_tipo:
                    cor = CORES_PECAS[cor_index % len(CORES_PECAS)] if cor_index < len(CORES_PECAS) else gerar_cor_aleatoria()
                    cores_por_tipo[tipo_peca_key] = cor
                    legenda.append({"tipo": tipo_peca_key, "cor": cor})
                    cor_index += 1
            
            resultado = calcular_plano_de_corte(chapa_largura, chapa_altura, pecas)

            for plano_unico in resultado['planos_unicos']:
                for peca_plano in plano_unico['plano']:
                    tipo_key = f"{peca_plano['tipo_largura']}x{peca_plano['tipo_altura']}"
                    peca_plano['cor'] = cores_por_tipo.get(tipo_key)

            resultado['legenda'] = legenda
            chapa_info = {'largura': chapa_largura, 'altura': chapa_altura}
            
            session['last_result'] = resultado
            session['chapa_info'] = chapa_info
            session['pecas_info'] = pecas

            return render_template('index.html', resultado=resultado, chapa=chapa_info, resultado_json=json.dumps(resultado))

        except (ValueError, KeyError) as e:
            return render_template('index.html', erro=f"Erro nos dados de entrada: '{e.args[0]}'")

    return render_template('index.html', resultado=None, chapa=None, resultado_json=None)

@app.route('/export/pdf')
def export_pdf():
    resultado = session.get('last_result')
    chapa_info = session.get('chapa_info')
    if not resultado: return "Nenhum resultado para exportar.", 404
    
    timestamp = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")

    html_renderizado = render_template('report_template.html', resultado=resultado, chapa=chapa_info, timestamp=timestamp)
    
    pdf_bytes = HTML(string=html_renderizado).write_pdf()
    response = make_response(pdf_bytes)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'attachment; filename=relatorio_cortex.pdf'
    return response

@app.route('/export/excel')
def export_excel():
    resultado = session.get('last_result')
    chapa_info = session.get('chapa_info')
    pecas_info = session.get('pecas_info')
    if not resultado: return "Nenhum resultado para exportar.", 404

    timestamp = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")

    dados_pecas = [{'Quantidade': p['quantidade'], 'Dimensões (LxA)': f"{p['largura']}x{p['altura']}"} for p in pecas_info]
    df_pecas = pd.DataFrame(dados_pecas)
    
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    resumo_data = {
        'Parâmetro': ['Dimensões da Chapa (LxA)', 'Total de Chapas Utilizadas', 'Aproveitamento Geral'],
        'Valor': [f"{chapa_info['largura']}x{chapa_info['altura']}", resultado['total_chapas'], resultado['aproveitamento_geral']]
    }
    df_resumo = pd.DataFrame(resumo_data)
    df_resumo.to_excel(writer, sheet_name='Resumo', index=False, startrow=0)

    df_pecas.to_excel(writer, sheet_name='Resumo', index=False, startrow=len(df_resumo)+2, header=['Quantidade Solicitada', 'Dimensões da Peça'])
    
    worksheet = writer.sheets['Resumo']
    final_row = len(df_resumo) + len(df_pecas) + 4
    
    # MUDANÇA DO TEXTO DO CARIMBO AQUI
    worksheet.cell(row=final_row, column=1, value=f"Aproveitamento feito em {timestamp}")

    writer.close()
    output.seek(0)

    response = make_response(output.getvalue())
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response.headers['Content-Disposition'] = 'attachment; filename=relatorio_cortex.xlsx'
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
