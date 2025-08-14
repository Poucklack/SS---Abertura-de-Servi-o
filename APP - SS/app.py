from flask import Flask, render_template, request, redirect, url_for, session
import pandas as pd
import os
from datetime import datetime
from functools import wraps

app = Flask(__name__)
app.secret_key = 'uma_chave_secreta_aqui'

# Arquivos Excel
ARQUIVO_SERVICOS = 'servicos.xlsx'
ARQUIVO_USUARIOS = 'usuarios.xlsx'

# Criar Excel de serviços se não existir
if not os.path.exists(ARQUIVO_SERVICOS):
    df = pd.DataFrame(columns=['ID', 'Data', 'Setor', 'Descrição', 'Solicitante', 'Prioridade'])
    df.to_excel(ARQUIVO_SERVICOS, index=False)

# Criar Excel de usuários se não existir
if not os.path.exists(ARQUIVO_USUARIOS):
    df_usuarios = pd.DataFrame([{
        'Usuario': 'admin',
        'Senha': '1234',  # altere depois para senha segura
        'Permissao': 'admin'
    }])
    df_usuarios.to_excel(ARQUIVO_USUARIOS, index=False)

# Decorators
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get('permissao') != 'admin':
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

# Código admin para criação de usuários
CODIGO_ADMIN = "7410"

# Login
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']
        df_usuarios = pd.read_excel(ARQUIVO_USUARIOS)
        user = df_usuarios[(df_usuarios['Usuario'] == usuario) & (df_usuarios['Senha'] == senha)]
        if not user.empty:
            session['usuario'] = usuario
            session['permissao'] = user.iloc[0]['Permissao']
            return redirect(url_for('index'))  # vai para index.html
        else:
            return render_template('login.html', erro="Usuário ou senha incorretos")
    return render_template('login.html')

# Logout
@app.route('/logout')
@login_required
def logout():
    session.clear()
    return redirect(url_for('login'))

# Página inicial (index)
@app.route('/index')
@login_required
def index():
    return render_template('index.html', usuario=session.get('usuario'))

# Abrir novo serviço
@app.route('/abrir_servico', methods=['POST'])
@login_required
def abrir_servico():
    setor = request.form['setor']
    descricao = request.form['descricao']
    solicitante = request.form['solicitante']
    prioridade = request.form['prioridade']

    df = pd.read_excel(ARQUIVO_SERVICOS)
    novo_id = len(df) + 1
    data = datetime.now().strftime("%d/%m/%Y %H:%M")

    novo_servico = {
        'ID': novo_id,
        'Data': data,
        'Setor': setor,
        'Descrição': descricao,
        'Solicitante': solicitante,
        'Prioridade': prioridade
    }

    df = pd.concat([df, pd.DataFrame([novo_servico])], ignore_index=True)
    df.to_excel(ARQUIVO_SERVICOS, index=False)

    return redirect(url_for('index'))

# Listar serviços
@app.route('/lista')
@login_required
def lista_servicos():
    df = pd.read_excel(ARQUIVO_SERVICOS)
    prioridade_ordem = {'Alta': 1, 'Média': 2, 'Baixa': 3}
    df['Prioridade_Ordem'] = df['Prioridade'].map(prioridade_ordem)
    df = df.sort_values(by='Prioridade_Ordem')
    df = df.drop(columns=['Prioridade_Ordem'])
    return render_template('lista.html', servicos=df.to_dict(orient='records'))

# Apagar serviço (admin)
@app.route('/apagar/<int:id>', methods=['POST'])
@login_required
@admin_required
def apagar_servico(id):
    df = pd.read_excel(ARQUIVO_SERVICOS)
    df = df[df['ID'] != id]
    df.to_excel(ARQUIVO_SERVICOS, index=False)
    return redirect(url_for('lista_servicos'))

# Editar serviço (admin)
@app.route('/editar/<int:id>')
@login_required
@admin_required
def editar_servico(id):
    df = pd.read_excel(ARQUIVO_SERVICOS)
    servico = df[df['ID'] == id].to_dict(orient='records')
    if not servico:
        return redirect(url_for('lista_servicos'))
    return render_template('editar.html', servico=servico[0])

# Atualizar serviço (admin)
@app.route('/atualizar/<int:id>', methods=['POST'])
@login_required
@admin_required
def atualizar_servico(id):
    df = pd.read_excel(ARQUIVO_SERVICOS)
    setor = request.form['setor']
    descricao = request.form['descricao']
    solicitante = request.form['solicitante']
    prioridade = request.form['prioridade']

    df.loc[df['ID'] == id, ['Setor', 'Descrição', 'Solicitante', 'Prioridade']] = [setor, descricao, solicitante, prioridade]
    df.to_excel(ARQUIVO_SERVICOS, index=False)

    return redirect(url_for('lista_servicos'))

# Criar usuário
@app.route('/criar_usuario', methods=['GET', 'POST'])
def criar_usuario():
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']
        codigo_admin = request.form.get('codigo_admin', '')

        permissao = 'admin' if codigo_admin == CODIGO_ADMIN else 'user'

        df = pd.read_excel(ARQUIVO_USUARIOS)
        if usuario in df['Usuario'].values:
            return render_template('criar_usuario.html', erro="Usuário já existe")

        novo_usuario = {'Usuario': usuario, 'Senha': senha, 'Permissao': permissao}
        df = pd.concat([df, pd.DataFrame([novo_usuario])], ignore_index=True)
        df.to_excel(ARQUIVO_USUARIOS, index=False)

        return redirect(url_for('login'))

    return render_template('criar_usuario.html')


# Rodar app
if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=5001)
