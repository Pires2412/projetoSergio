from flask import Flask, render_template,redirect, request, send_file, url_for
##from process import process_text
import docx
from docx import Document
import os
from datetime import datetime

app = Flask(__name__)
nome_documento = None

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/form', methods=['GET','POST'])
def form():

    if request.method == 'POST':

#### Requirindo os dados do formulário e atribuindo o valor dado pelo usuario
        referencia = {
####Perícia
    "{{VF}}": request.form.get('vara_federal'),
    "{{SEJUD}}": request.form.get('sejud'),
    "{{VT}}": request.form.get('vara_trabalho'),
    "{{DTP}}": formatar_data(request.form.get('data_pericia')),
    "{{TDP}}": request.form.get('tipo_pericia'),
        
    "{{ENDERECO}}": request.form.get('end'),

    "{{DIPR}}": formatar_data(request.form.get('data_ing_re')),
    "{{DTSA}}": formatar_data(request.form.get('data_setor_a')),
    "{{CH}}": request.form.get('carga_horaria'),

#### Atividade
    "{{ATV}}": request.form.get('atividade1'),
#### Tecnico


#### Medições

    "{{ABIO}}": request.form.get('agente_bio'),
    "{{TABIO}}": request.form.get('t_agente_bio'),
    "{{QUALABIO}}": request.form.get('agente_bio_qual'),

    "{{AFIS}}": request.form.get('agente_fis'),
    "{{TAFIS}}": request.form.get('t_agente_fis'),
    "{{QUALAFIS}}": request.form.get('agente_fis_qual'),
    "{{QUANAFIS}}": request.form.get('agente_fis_quan'),


    "{{AQUI}}": request.form.get('agente_quim'),
    "{{TAQUI}}": request.form.get('t_agente_quim'),
    "{{QUALAQUI}}": request.form.get('agente_quim_qual'),
    "{{QUANAQUI}}":request.form.get('agente_quim_quant'),
    "{{PERIC}}":request.form.get('periculosodade'),
    "{{TPERIC}}":request.form.get('t_periculosodade'),

#### Equipamento de Proteção

    "{{MASCC}}":request.form.get('masc_cirg'),
    "{{MASCCCA}}":request.form.get('masc_cirg_ca'),

    "{{MASCN95}}": request.form.get('masc_n95'),
    "{{MASCN95CA}}": request.form.get('masc_n95_ca'),

    "{{FS}}": request.form.get('face_shield'),
    "{{FSCA}}": request.form.get('face_shield_ca'),

    "{{ODP}}": request.form.get('oculos_pro'),
    "{{ODPCA}}": request.form.get('oculos_pro_ca'), 

    "{{LDP}}":request.form.get('luva_proc'),
    "{{LDPCA}}":request.form.get('luva_proc_ca'),

    "{{LE}}":request.form.get('luva_esteril'),
    "{{LECA}}":request.form.get('luva_esteril_ca'),

    "{{CP}}": request.form.get('capote_plumb'),
    "{{CPCA}}": request.form.get('apote_plumb_ca'),

    "{{PDT}}": request.form.get('protetor_tireoide'),
    "{{PDTCA}}": request.form.get('protetor_tireoide_ca'),

    "{{ADR}}": request.form.get('abaf_ruido'),
    "{{ADRCA}}": request.form.get('abaf_ruido_ca'),

    "{{PAR}}": request.form.get('pro_aur'),
    "{{PARCA}}": request.form.get('pro_aur_ca'),

    "{{AIMP}}": request.form.get('avental_imperm'),
    "{{AIMPCA}}":request.form.get('avental_imperm_ca'),

    "{{LLCAMR}}": request.form.get('luv_ltx_c_amar'),
    "{{LLCAMRCA}}": request.form.get('luv_ltx_c_amar_ca'),

    "{{LLLAMR}}": request.form.get('uv_ltx_l_amar'),
    "{{LLLAMRCA}}": request.form.get('luv_ltx_l_amar_ca') ,

    "{{LLCVER}}": request.form.get('luv_ltx_c_ver'),
    "{{LLCVERCA}}": request.form.get('luv_ltx_c_ver_ca'), 

    "{{LLLVER}}":request.form.get('luv_ltx_l_ver'),
    "{{LLLVERCA}}": request.form.get('luv_ltx_l_ver_ca'),

    "{{LLCAZ}}": request.form.get('luv_ltx_c_az'),
    "{{LLCAZCA}}": request.form.get('luv_ltx_c_az_ca'),

    "{{LLLAZ}}": request.form.get('luv_ltx_l_az'),
    "{{LLLAZCA}}": request.form.get('luv_ltx_l_az_ca'),

    "{{BDS}}": request.form.get('bota_seg'),
    "{{BDSCA}}": request.form.get('bota_seg_ca'),

    "{{MASCPFF2}}": request.form.get('masc_pff2'),
    "{{MASCPFF2CA}}": request.form.get('masc_pff2_ca'),
    }

####Verificar todos os dados mostrando no terminal
    # print("Dados capturados do formulário:")
    # for chave, valor in referencia.items():
    #     print(f"{chave}: {valor}")

    

#### Arquivo Modelo
    documento = Document("pericia_relatorio.docx")
    novo_documento = request.form.get('novo_documento')+".docx"
    global nome_documento
    nome_documento = novo_documento

#### Substituir as variáveis no documento
    for chave, valor in referencia.items():
        for par in documento.paragraphs:
            if chave in par.text:
                par.text = par.text.replace(chave, valor if valor else '')
                documento.save(novo_documento)

    return render_template("resposta.html")

@app.route('/download', methods=['GET'])
def download():
    if request.method == 'GET':
        global nome_documento
        return send_file(nome_documento, as_attachment=True)
    
@app.route('/excluir', methods=['GET'])
def excluir_arquivo():
    global nome_documento
    os.remove(nome_documento)
    return render_template('index.html')

def formatar_data(data):
    return datetime.strptime(data, "%Y-%m-%d").strftime("%d-%m-%Y") if data else None

if __name__ == '__main__':
    app.run(debug=True, port=5500)


