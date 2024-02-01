import gspread
from google.oauth2 import service_account

# Função para fazer login e retornar uma instância autorizada do Google Sheets
def login():
    escopos = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    arquivo_json = "desafiotuntsrocks2024.json"
    credenciais = service_account.Credentials.from_service_account_file(arquivo_json)
    credenciais_escopos = credenciais.with_scopes(escopos)
    gc = gspread.authorize(credenciais_escopos)
    return gc


def calcular_situacao(media, faltas, total_aulas):
    if faltas > total_aulas * 0.25:
        return "Reprovado por Falta"
    elif media < 50:
        return "Reprovado por Nota"
    elif media < 70:
        return "Exame Final"
    else:
        return "Aprovado"


gc = login()
planilha = gc.open("Engenharia de Software - Desafio MATHEUS MORETE ESPINOSO").worksheet("engenharia_de_software")

# Definindo o número total de aulas
total_aulas = 60  # Supondo que há 60 aulas no total

# Começando da quarta linha (índice 3)
for i, linha in enumerate(planilha.get_all_values()[3:], start=4):
    matricula, aluno, faltas, P1, P2, P3, situacao, nota_aprovacao = linha

    # Convertendo para inteiros
    faltas = int(faltas)
    P1 = int(P1)
    P2 = int(P2)
    P3 = int(P3)

    # Calculando média
    media = (P1 + P2 + P3) / 3

    # Verificando situação do aluno
    situacao_final = calcular_situacao(media, faltas, total_aulas)

    # Se a situação for "Exame Final", calcular a Nota para Aprovação Final
    if situacao_final == "Exame Final":
        nota_aprovacao = max(0, 100 - media)
    else:
        nota_aprovacao = 0

    # Arredondando para o próximo número inteiro
    nota_aprovacao = round(nota_aprovacao)

    # Atualizando a planilha
    planilha.update_cell(i, 8, nota_aprovacao)

    # Atualizando a coluna "Situação"
    planilha.update_cell(i, 7, situacao_final)

    # Registro de log
    print(f"Aluno: {aluno}, Situação: {situacao_final}, Nota para Aprovação Final: {nota_aprovacao}")

print("Processamento concluído.")
