from flask import Flask, request, jsonify
from flask_cors import CORS
from email_service import enviar_informe_tecnico, enviar_tentativa_contato
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Diretório onde os modelos de e-mail estão armazenados por departamento
EMAIL_TEMPLATES_DIR = "email_templates"

# Função para carregar os modelos de e-mail por departamento
def carregar_modelos():
    modelos = {}
    if not os.path.exists(EMAIL_TEMPLATES_DIR):
        return modelos

    for dept in os.listdir(EMAIL_TEMPLATES_DIR):
        dept_path = os.path.join(EMAIL_TEMPLATES_DIR, dept)
        if os.path.isdir(dept_path):
            modelos[dept] = {}  # Mudando de lista para dicionário
            for arquivo in os.listdir(dept_path):
                if arquivo.endswith(".html"):
                    nome_modelo = os.path.splitext(arquivo)[0]
                    caminho_arquivo = os.path.join(dept_path, arquivo)

                    # Lendo o conteúdo do arquivo HTML
                    with open(caminho_arquivo, "r", encoding="utf-8") as f:
                        modelos[dept][nome_modelo] = f.read()  # Guardando o conteúdo
    return modelos

# Carregar os modelos no dicionário global
EMAIL_MODELOS = carregar_modelos()

def definir_saudacao():
    hora_atual = datetime.now().hour
    if hora_atual < 12:
        return "bom dia"
    elif hora_atual < 18:
        return "boa tarde"
    else:
        return "boa noite"  

@app.route("/enviar-email", methods=["POST"])
def enviar_email():
    try:
        dados = request.json
        destinatarios = dados.get("destinatarios", [])  # Lista de destinatários
        copia = dados.get("copia", [])
        assunto = dados.get("assunto")
        departamento = dados.get("departamento")
        modelo = dados.get("modelo")
        empresa = dados.get("empresa")

        if not destinatarios or not assunto or not departamento or not modelo:
            return jsonify({"erro": "Campos obrigatórios faltando"}), 400

        corpo_email = EMAIL_MODELOS.get(departamento, {}).get(modelo, "")
        
        if not corpo_email:
            return jsonify({"erro": "Modelo de e-mail não encontrado"}), 404

        corpo_email = corpo_email.replace("{saudacao}", definir_saudacao())

        # Enviar e-mail para cada destinatário separadamente
        if modelo == "informe_tecnico":
            for destinatario in destinatarios:
                enviar_informe_tecnico(destinatario, copia, assunto, corpo_email)
        if modelo == "tentativa_contato":
            enviar_tentativa_contato(destinatarios, copia, assunto, corpo_email, empresa)


        return jsonify({"mensagem": "E-mails enviados com sucesso!"})

    except Exception as e:
        print(f"Erro ao enviar e-mail: {str(e)}")
        return jsonify({"erro": str(e)}), 500

@app.route("/email-templates", methods=["GET"])
def listar_modelos_email():
    try:
        return jsonify(EMAIL_MODELOS)
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)