from flask import Flask, request, jsonify
from flask_cors import CORS, cross_origin
import smtplib
from email.message import EmailMessage
import datetime
import os
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

load_dotenv()

app = Flask(__name__)

# CONFIGURAÇÃO CORS COMPLETA
CORS(app, 
     origins=["http://127.0.0.1:5500", "http://localhost:5500", "http://127.0.0.1:5000"],
     methods=["GET", "POST", "OPTIONS"],
     allow_headers=["Content-Type", "Authorization"],
     supports_credentials=True)

# ================== GOOGLE SHEETS ==================
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

try:
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build("sheets", "v4", credentials=credentials)
    sheet = service.spreadsheets()
    print("✅ Conectado ao Google Sheets com sucesso!")
except Exception as e:
    print(f"❌ Erro ao conectar ao Google Sheets: {e}")
    sheet = None

# ================== EMAIL ==================
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
EMAIL_SENHA = os.getenv("EMAIL_SENHA")
EMAIL_DESTINATARIO = os.getenv("EMAIL_DESTINATARIO")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

def enviar_email_notificacao(dados):
    """Envia email de notificação sobre novo contato"""
    try:
        msg = EmailMessage()
        msg["Subject"] = f"📩 Novo contato: {dados['Nome Completo']}"
        msg["From"] = EMAIL_REMETENTE
        msg["To"] = EMAIL_DESTINATARIO

        corpo = f"""
        Novo contato recebido pelo site:

        👤 Nome: {dados['Nome Completo']}
        📧 Email: {dados['Email']}
        📞 Telefone: {dados['Telefone']}
        🔧 Serviço: {dados['Servico de Interesse']}
        💬 Mensagem: {dados.get('Mensagem', 'Não informada')}
        📅 Data: {dados['Data de Envio']}
        """
        
        msg.set_content(corpo)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)
        
        print("✅ Email enviado com sucesso!")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar email: {e}")
        return False

@app.route("/")
def home():
    """Endpoint raiz para verificar se a API está funcionando"""
    return jsonify({
        "mensagem": "API de contato está funcionando!",
        "status": "online",
        "endpoints": {
            "/api/excel/salvar": "POST - Envia dados para o Google Sheets"
        }
    })

@app.route("/teste", methods=["GET"])
def teste():
    """Endpoint para testar conexão"""
    return jsonify({
        "status": "conectado",
        "mensagem": "API está funcionando!",
        "timestamp": datetime.datetime.now().isoformat()
    })

@app.route("/api/excel/salvar", methods=["POST", "OPTIONS"])
@cross_origin()
def salvar_documento():
    """Salva os dados do formulário no Google Sheets e envia email"""
    # Tratar requisição OPTIONS para CORS
    if request.method == 'OPTIONS':
        return jsonify({}), 200
    
    dados = request.get_json()

    if not dados:
        return jsonify({"erro": "JSON vazio ou inválido"}), 400

    # Campos obrigatórios
    campos_obrigatorios = [
        "Nome Completo", "Email", "Telefone", "Servico de Interesse"
    ]

    # Verificar campos obrigatórios
    campos_faltantes = []
    for campo in campos_obrigatorios:
        if not dados.get(campo):
            campos_faltantes.append(campo)
    
    if campos_faltantes:
        return jsonify({
            "erro": "Campos obrigatórios ausentes",
            "campos_faltantes": campos_faltantes
        }), 400

    try:
        # Adicionar data/hora atual
        agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        dados["Data de Envio"] = agora

        # Preparar valores para o Google Sheets
        valores = [[
            dados["Nome Completo"],
            dados["Email"],
            dados["Telefone"],
            dados["Servico de Interesse"],
            dados.get("Mensagem", ""),
            dados["Data de Envio"]
        ]]

        # Salvar no Google Sheets
        if sheet:
            result = sheet.values().append(
                spreadsheetId=SPREADSHEET_ID,
                range="Página1!A1",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": valores}
            ).execute()
            
            print(f"✅ Dados salvos no Sheets: {result.get('updates', {}).get('updatedCells', 0)} células atualizadas")
        else:
            print("⚠️ Google Sheets não disponível, simulando salvamento...")

        # Enviar email de notificação
        email_enviado = enviar_email_notificacao(dados)

        return jsonify({
            "mensagem": "Mensagem enviada com sucesso!",
            "email_enviado": email_enviado,
            "timestamp": agora
        }), 201

    except Exception as e:
        print(f"❌ ERRO INTERNO: {str(e)}")
        return jsonify({
            "erro": "Erro interno no servidor",
            "detalhes": str(e) if app.debug else None
        }), 500

@app.route("/health")
def health_check():
    """Endpoint para verificar a saúde da API"""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.datetime.now().isoformat(),
        "sheets_connected": sheet is not None,
        "email_config": {
            "remetente": bool(EMAIL_REMETENTE),
            "destinatario": bool(EMAIL_DESTINATARIO)
        }
    })

if __name__ == "__main__":
    print("🚀 Iniciando servidor Flask...")
    print(f"📧 Email remetente configurado: {bool(EMAIL_REMETENTE)}")
    print(f"📧 Email destinatário configurado: {bool(EMAIL_DESTINATARIO)}")
    print(f"🌐 CORS configurado para: 127.0.0.1:5500 e localhost:5500")
    app.run(host="127.0.0.1", port=5000, debug=True)

    # Adicione este endpoint para verificar dados enviados
@app.route("/api/dados/ultimos", methods=["GET"])
def ultimos_dados():
    """Retorna os últimos dados enviados (útil para debugging)"""
    try:
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Página1!A:F",
            majorDimension="ROWS"
        ).execute()
        
        valores = result.get('values', [])
        ultimos_5 = valores[-5:] if len(valores) > 5 else valores
        
        return jsonify({
            "total_registros": len(valores) - 1 if valores else 0,
            "ultimos_registros": ultimos_5
        })
    except Exception as e:
        return jsonify({"erro": str(e)}), 500