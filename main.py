import gspread
from google.oauth2 import service_account

# Function to login and return an authorized instance of Google Sheets
def login():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    json_file = "desafiotuntsrocks2024.json"
    credentials = service_account.Credentials.from_service_account_file(json_file)
    scoped_credentials = credentials.with_scopes(scopes)
    gc = gspread.authorize(scoped_credentials)
    return gc

def calculate_status(average, absences, total_classes):
    if absences > total_classes * 0.25:
        return "Reprovado por Falta"
    elif average < 50:
        return "Reprovado por Nota"
    elif average < 70:
        return "Exame Final"
    else:
        return "Aprovado"

gc = login()
spreadsheet = gc.open("Engenharia de Software - Desafio MATHEUS MORETE ESPINOSO").worksheet("engenharia_de_software")

# Defining the total number of classes
total_classes = 60  # Assuming there are 60 classes in total

# Starting from the fourth row (index 3)
for i, row in enumerate(spreadsheet.get_all_values()[3:], start=4):
    registration, student, absences, P1, P2, P3, status, approval_score = row

    # Converting to integers
    absences = int(absences)
    P1 = int(P1)
    P2 = int(P2)
    P3 = int(P3)

    # Calculating average
    average = (P1 + P2 + P3) / 3

    # Checking student's status
    final_status = calculate_status(average, absences, total_classes)

    # If status is "Exame Final", calculate the Approval Score
    if final_status == "Exame Final":
        approval_score = max(0, 100 - average)
    else:
        approval_score = 0

    # Rounding to the next integer
    approval_score = round(approval_score)

    # Updating the spreadsheet
    spreadsheet.update_cell(i, 8, approval_score)

    # Updating the "Status" column
    spreadsheet.update_cell(i, 7, final_status)

    # Logging
    print(f"Aluno: {student}, Situação: {final_status}, Nota para Aprovação Final: {approval_score}")

print("Processing completed.")
