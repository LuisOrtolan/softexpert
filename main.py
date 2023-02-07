from suds.client import Client
import openpyxl
import datetime

wb = openpyxl.load_workbook("planilha.xlsx")
sheet = wb["sheet"]

#esses eram parametros do sistemas que seriam os mesmos, independente da linha que eu ia subir na planilha
typeid = "23"
phase = "3"
amount = "1"

def api_go(projectid, name, date, unitvalue):
    api_gateway_h = 'codigo para acesso da api'
    url = "https://acesso.softexpert.com/se/ws/pr_ws.php?wsdl"

    headers = {
        'Authorization': api_gateway_h
    }
    
#nesse caso estava usando para subir muitos custos de projetos de uma vez (newProjectCost)

    client = Client(url, headers=headers)
    method = client.wsdl.services[0].ports[0].methods["newProjectCost"]
    date2 = date.strftime('%d/%m/%Y')
    response = client.service.newProjectCost(
        ProjectId=projectid,
        Name=name,
        TypeId=typeid,
        Phase=phase,
        Date=date2,
        Amount=amount,
        UnitValue=unitvalue
    )
    print(response)

for i in range(1, sheet.max_row):
    linha = i + 1
    localP = f"A{linha}"
    localN = f"G{linha}"
    localD = f"L{linha}"
    localU = f"T{linha}"

    projectid = sheet[localP].value
    name = sheet[localN].value
    date = sheet[localD].value
    unitvalue = sheet[localU].value

    print(f"ProjectId={projectid}, Name={name}, TypeId={typeid}, Phase={phase}, Date={date2}, Amount={amount}, UnitValue={unitvalue}")

    api_go(projectid, name, date, unitvalue)
