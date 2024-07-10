from datetime import date
from excelreport import TExcelReport

# Exemplo de uso:
dataset = [
    {"docnum": "1001", "serial": "A123", "instnum": 1, "docdate": date(2023, 1, 1), "duedate":date(2023, 1, 1),"doctotal": 1000.00, "status": "Pago", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    {"docnum": "1002", "serial": "A124", "instnum": 1, "docdate": date(2023, 1, 2), "duedate":date(2023, 1, 1),"doctotal": 1500.00, "status": "Pendente", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    {"docnum": "1003", "serial": "A125", "instnum": 1, "docdate": date(2023, 1, 3), "duedate":date(2023, 1, 1),"doctotal": 2000.00, "status": "Pago", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    {"docnum": "1004", "serial": "B123", "instnum": 1, "docdate": date(2023, 2, 1), "duedate":date(2023, 1, 1),"doctotal": 2500.00, "status": "Pago", "cardname": "Cliente B","cardcode":"C002","phone":"(11) 256-0873"},
    {"docnum": "1005", "serial": "B124", "instnum": 1, "docdate": date(2023, 2, 2), "duedate":date(2023, 1, 1),"doctotal": 3000.00, "status": "Pendente", "cardname": "Cliente B","cardcode":"C002","phone":"(11) 256-0873"},
    {"docnum": "1006", "serial": "B125", "instnum": 1, "docdate": date(2023, 2, 3), "duedate":date(2023, 1, 1),"doctotal": 3500.00, "status": "Pago", "cardname": "Cliente B","cardcode":"C002","phone":"(11) 256-0873"},
    {"docnum": "1007", "serial": "C123", "instnum": 1, "docdate": date(2023, 3, 1), "duedate":date(2023, 1, 1),"doctotal": 4000.00, "status": "Pago", "cardname": "Cliente C","cardcode":"C003","phone":"(11) 256-0873"},
    {"docnum": "1008", "serial": "C124", "instnum": 1, "docdate": date(2023, 3, 2), "duedate":date(2023, 1, 1),"doctotal": 4500.00, "status": "Pendente", "cardname": "Cliente C","cardcode":"C003","phone":"(11) 256-0873"},
    {"docnum": "1009", "serial": "C125", "instnum": 1, "docdate": date(2023, 3, 3), "duedate":date(2023, 1, 1),"doctotal": 5000.00, "status": "Pago", "cardname": "Cliente C","cardcode":"C003","phone":"(11) 256-0873"},
    {"docnum": "1010", "serial": "A126", "instnum": 2, "docdate": date(2023, 4, 1), "duedate":date(2023, 1, 1),"doctotal": 5500.00, "status": "Pendente", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    {"docnum": "1011", "serial": "B126", "instnum": 2, "docdate": date(2023, 4, 2), "duedate":date(2023, 1, 1),"doctotal": 6000.00, "status": "Pago", "cardname": "Cliente B","cardcode":"C002","phone":"(11) 256-0873"},
    {"docnum": "1012", "serial": "C126", "instnum": 2, "docdate": date(2023, 4, 3), "duedate":date(2023, 1, 1),"doctotal": 6500.00, "status": "Pendente", "cardname": "Cliente C","cardcode":"C003","phone":"(11) 256-0873"},
    {"docnum": "1013", "serial": "A127", "instnum": 3, "docdate": date(2023, 5, 1), "duedate":date(2023, 1, 1),"doctotal": 7000.00, "status": "Pago", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    {"docnum": "1014", "serial": "B127", "instnum": 3, "docdate": date(2023, 5, 2), "duedate":date(2023, 1, 1),"doctotal": 7500.00, "status": "Pendente", "cardname": "Cliente B","cardcode":"C002","phone":"(11) 256-0873"},
    {"docnum": "1015", "serial": "C127", "instnum": 3, "docdate": date(2023, 5, 3), "duedate":date(2023, 1, 1),"doctotal": 8000.00, "status": "Pago", "cardname": "Cliente C","cardcode":"C003","phone":"(11) 256-0873"},
]

dataset.sort(key=lambda x: x['cardcode'])

yaml_file = "config.yaml"

report = TExcelReport('Por Cliente',dataset, yaml_file)
date_of_issue = date.today().strftime('%d/%m/%Y')
report.setCell("A1",f"Títulos em Aberto até {date_of_issue}")
report.setCell("A2",f"Emissão")
report.setCell("B2",f"{date_of_issue}")
report.build()
report.save("relatorio.xlsx")
