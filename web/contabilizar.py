import os
import re
from openpyxl import Workbook
from openpyxl.styles import numbers

# Función auxiliar para obtener los impuestos IVA e INC para cada ítem
def obtener_impuestos(item):
    iva_tax_amount = 0
    inc_tax_amount = 0
    item_tax_total = re.findall(r'<cac:TaxTotal>(.*?)</cac:TaxTotal>', item, re.DOTALL)
    for total in item_tax_total:
        item_iva = re.search(r'<cbc:TaxAmount currencyID="COP">(.*?)</cbc:TaxAmount>.*?<cbc:Name>IVA</cbc:Name>', total, re.DOTALL)
        item_inc = re.search(r'<cbc:TaxAmount currencyID="COP">(.*?)</cbc:TaxAmount>.*?<cbc:Name>INC</cbc:Name>', total, re.DOTALL)
        if item_iva:
            iva_tax_amount = float(item_iva.group(1)) if item_iva.group(1) else 0
        if item_inc:
            inc_tax_amount = float(item_inc.group(1)) if item_inc.group(1) else 0
    return iva_tax_amount, inc_tax_amount

# Ruta al directorio que contiene los archivos XML
ruta_directorio = 'C:/Users/Admin/OneDrive/Desktop/Golden/XML'

# Obtener la lista de archivos XML en el directorio
archivos_xml = [archivo for archivo in os.listdir(ruta_directorio) if archivo.endswith('.xml')]

# Crear un nuevo archivo de Excel
libro_excel = Workbook()
hoja_excel = libro_excel.active

# Definir los encabezados de las columnas
encabezados = ['Archivo', 'Número','factura', 'Empresa', 'NIT',
               'Fecha de emisión', 'Fecha de vencimiento', 'Número de Item', 'Cantidad', 'Descripción',
               'Valor IVA Item', 'Valor INC Item', 'Subtotal', 'Total con Impuestos']
hoja_excel.append(encabezados)

# Obtener la cantidad de archivos XML
cantidad_archivos = len(archivos_xml)

# Agregar los nombres de los archivos al archivo de Excel con numeración consecutiva y contenido en nuevas columnas
for i, archivo in enumerate(archivos_xml, start=1):
    ruta_archivo = os.path.join(ruta_directorio, archivo)
    with open(ruta_archivo, 'r', encoding='utf-8') as file:
        contenido = file.read()

    party_name_value = ''
    company_id_value = ''
    issue_date_value = ''
    due_date_value = ''

    tamano = re.search(r'<cbc:LineCountNumeric.*?>(.*?)</cbc:LineCountNumeric>', contenido, re.DOTALL)
    tamano_valor = tamano.group(1) if tamano else ''

    accounting_supplier_party = re.search(r'<cac:AccountingSupplierParty>(.*?)</cac:AccountingSupplierParty>', contenido, re.DOTALL)
    party_name = re.search(r'<cbc:RegistrationName>(.*?)</cbc:RegistrationName>', accounting_supplier_party.group(1), re.DOTALL)
    party_name_value = party_name.group(1) if party_name else ''
    company_id = re.search(r'<cbc:CompanyID.*?>(.*?)</cbc:CompanyID>', contenido, re.DOTALL)
    company_id_value = company_id.group(1) if company_id else ''
    issue_date = re.search(r'<cbc:IssueDate>(.*?)</cbc:IssueDate>', contenido)
    issue_date_value = issue_date.group(1) if issue_date else ''
    due_date = re.search(r'<cbc:DueDate>(.*?)</cbc:DueDate>', contenido)
    due_date_value = due_date.group(1) if due_date else ''
    item_description = re.search(r'<cbc:Description>(.*?)</cbc:Description>', contenido, re.DOTALL)
    item_description_value = item_description.group(1) if item_description else ''
    id_value = re.search(r'<cbc:ID>(.*?)</cbc:ID>', contenido, re.DOTALL)
    id_value = id_value.group(1) if id_value else ''
    
    if int(tamano_valor) == 1:
        legal_monetary_total = re.search(r'<cac:LegalMonetaryTotal>(.*?)</cac:LegalMonetaryTotal>', contenido, re.DOTALL)
        line_extension_amount = re.search(r'<cbc:LineExtensionAmount.*?>(.*?)</cbc:LineExtensionAmount>', legal_monetary_total.group(1))
        tax_exclusive_amount = re.search(r'<cbc:TaxExclusiveAmount.*?>(.*?)</cbc:TaxExclusiveAmount>', legal_monetary_total.group(1))
        tax_inclusive_amount = re.search(r'<cbc:TaxInclusiveAmount.*?>(.*?)</cbc:TaxInclusiveAmount>', legal_monetary_total.group(1))
        
        iva_tax_amount, inc_tax_amount = obtener_impuestos(contenido)
        
        fila = [archivo, i, id_value, party_name_value, company_id_value,
                issue_date_value,
                due_date_value,
                '',
                '',
                item_description_value,
                iva_tax_amount,
                inc_tax_amount,
                tax_exclusive_amount.group(1) if tax_exclusive_amount else '',
                tax_inclusive_amount.group(1) if tax_inclusive_amount else '',
                ]
        hoja_excel.append(fila)
    else:
        # Obtener detalles de los items
        items = re.findall(r'<cac:InvoiceLine>(.*?)</cac:InvoiceLine>', contenido, re.DOTALL)
        for index, item in enumerate(items, start=1):
            item_quantity = re.search(r'<cbc:InvoicedQuantity unitCode="(.*?)">(.*?)</cbc:InvoicedQuantity>', item, re.DOTALL)
            item_quantity_unitCode = item_quantity.group(1) if item_quantity else ''
            item_quantity_value = item_quantity.group(2) if item_quantity else ''

            item_line_extension_amount = re.search(r'<cbc:LineExtensionAmount currencyID="(.*?)">(.+?)</cbc:LineExtensionAmount>', item, re.DOTALL)
            item_line_extension_amount_currencyID = item_line_extension_amount.group(1) if item_line_extension_amount else ''
            item_line_extension_amount_value = item_line_extension_amount.group(2) if item_line_extension_amount else ''

            iva_tax_amount, inc_tax_amount = obtener_impuestos(item)
            item_description = re.search(r'<cbc:Description>(.*?)</cbc:Description>', item, re.DOTALL)
            item_description_value = item_description.group(1) if item_description else ''
            iva = float(iva_tax_amount)
            inc = float(inc_tax_amount)
            subtotal = float(item_line_extension_amount_value) if item_line_extension_amount_value else 0
            total_tax_amount = iva + inc
            neto = float(subtotal + total_tax_amount)

            fila = [archivo, i,id_value ,party_name_value, company_id_value,
                    issue_date_value,
                    due_date_value,
                    f'Item {index}',
                    f'{float(item_quantity_value)}',
                    item_description_value,
                    f'{iva}',
                    f'{inc}',
                    f'{subtotal}',
                    f'{neto}'
                    ]
            hoja_excel.append(fila)

# Guardar el archivo de Excel
nombre_archivo_excel = 'archivos.xml.xlsx'
ruta_archivo_excel = os.path.join('C:/Users/Admin/OneDrive/Desktop/Golden/excel', nombre_archivo_excel)
libro_excel.save(ruta_archivo_excel)

print('Archivo de Excel generado exitosamente:', ruta_archivo_excel)
