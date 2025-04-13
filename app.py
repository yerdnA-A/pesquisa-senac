from flask import Flask, request, redirect, url_for, render_template
import pandas as pd
import os
from datetime import datetime  # <-- Adicionado

app = Flask(__name__)

EXCEL_FILE = 'respostas.xlsx'

# Garante que o arquivo exista com cabeçalhos
if not os.path.exists(EXCEL_FILE):
    colunas = [
        "Data/Hora", "Endereço IP", "Nome", "Idade", "Curso/Função", "Computador", "Celular",
        "Antivírus", "Escaneamento", "Atualizações", "Reutiliza Senhas",
        "Usa Símbolos", "Verifica Domínio"
    ]
    pd.DataFrame(columns=colunas).to_excel(EXCEL_FILE, index=False)

@app.route('/')
def formulario():
    return render_template('formulario.html')

@app.route('/enviar', methods=['POST'])
def enviar():
    ip = request.remote_addr
    data_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    dados = {
        "Data/Hora": data_hora,
        "Endereço IP": ip,
        "Nome": request.form.get('nome'),
        "Idade": request.form.get('idade'),
        "Curso/Função": request.form.get('funcao'),
        "Computador": request.form.get('computador'),
        "Celular": request.form.get('celular'),
        "Antivírus": request.form.get('antivirus'),
        "Escaneamento": request.form.get('scan'),
        "Atualizações": request.form.get('atualiza'),
        "Reutiliza Senhas": request.form.get('reutiliza'),
        "Usa Símbolos": request.form.get('simbolos'),
        "Verifica Domínio": request.form.get('dominio')
    }

    df = pd.read_excel(EXCEL_FILE)
    df = df._append(dados, ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    return redirect(url_for('obrigado'))

@app.route('/obrigado')
def obrigado():
    return render_template('obrigado.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
