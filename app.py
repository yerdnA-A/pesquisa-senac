from flask import Flask, request, redirect, url_for, render_template
from openpyxl import Workbook, load_workbook
from openpyxl.utils import *
import os

app = Flask(__name__)
EXCEL_FILE = 'respostas.xlsx'

# Garante que o arquivo Excel existe
def inicializa_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Respostas'
        ws.append([
            'Nome', 'Idade', 'Curso/Função', 'Possui Computador',
            'Possui Celular', 'Possui Antivírus', 'Escaneamento',
            'Atualizações', 'Reutiliza Senhas', 'Senhas com Símbolos',
            'Verifica Domínio', 'IP'
        ])
        wb.save(EXCEL_FILE)

@app.route('/enviar', methods=['POST'])
def enviar():
    data = request.form
    ip_usuario = request.remote_addr


    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    ws.append([
        data.get('nome'),
        data.get('idade'),
        data.get('funcao'),
        data.get('computador'),
        data.get('celular'),
        data.get('antivirus'),
        data.get('scan'),
        data.get('atualiza'),
        data.get('reutiliza'),
        data.get('simbolos'),
        data.get('dominio'),
        ip_usuario 
    ])

    wb.save(EXCEL_FILE)
    return redirect(url_for('obrigado'))

@app.route('/obrigado')
def obrigado():
    return render_template('obrigado.html')

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    inicializa_excel()
    app.run(debug=True)