# -*- coding:UTF-8 -*-
import os.path
import pyodbc
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, colors, Font
from colorama import init, Fore, Style

init()

#Read configuration info
if not os.path.exists('AppSettings.xml'):
    print(Fore.RED + "Cannot find configuration file.")
    exit()

databaseConnectionPool = {}

def getDatabaseConfiguration(databaseName):
    result = {'driver':'','server':'', 'userName':'', 'password':''}
    
    configurationContent = '\n'.join(open('AppSettings.xml').readlines())
    configurationDocument = BeautifulSoup(configurationContent, features='xml')
    databaseSettingsNode = configurationDocument.find('databaseSettings')
    if not databaseSettingsNode is None:
        databaseConfigurationNode = databaseSettingsNode.find('database', databaseName=databaseName)
        if not databaseConfigurationNode is None:
            if not databaseConfigurationNode.get('driver') is None:
                result['driver'] = databaseConfigurationNode.get('driver')
            if not databaseConfigurationNode.get('server') is None:
                result['server'] = databaseConfigurationNode.get('server')
            if not databaseConfigurationNode.get('userName') is None:
                result['userName'] = databaseConfigurationNode.get('userName')
            if not databaseConfigurationNode.get('password') is None:
                result['password'] = databaseConfigurationNode.get('password')

    return result

eudDatabaseConfiguration = getDatabaseConfiguration('EUD')
eudConn = pyodbc.connect('Driver={driver};'
            'Server={server};'
            'Database={database};'
            'UID={username};'
            'PWD={password}'.format(driver = eudDatabaseConfiguration['driver'], server = eudDatabaseConfiguration['server'], database = 'EUD',
                        username = eudDatabaseConfiguration['userName'], password = eudDatabaseConfiguration['password']))

def getTruckCenterConnection(truckCenterID):
    sql = '''select 
                tc.SqlDb,
                s.IPAddress
            from adtbTruckCenters as tc
                inner join adtbServers as s on tc.ServerName = s.ServerName
            where TruckCenterID = {tcID}'''.format(tcID = truckCenterID)
    dbCursor = eudConn.cursor()
    dbCursor.execute(sql)
    dbRow = dbCursor.fetchone()

    truckCenterDatabaseConfiguration = getDatabaseConfiguration('TruckCenter')
    
    if dbRow != None:
        if truckCenterID in databaseConnectionPool.keys():
            return databaseConnectionPool[truckCenterID]
        else:
            newConnection = pyodbc.connect('Driver={driver};'
                                'Server={server};'
                                'Database={database};'
                                'UID={username};'
                                'PWD={password}'.format(driver = truckCenterDatabaseConfiguration['driver'], server = dbRow[1], database = dbRow[0],
                                            username = truckCenterDatabaseConfiguration['userName'], password = truckCenterDatabaseConfiguration['password']))
            databaseConnectionPool[truckCenterID] = newConnection
            return newConnection
    else:
        return None

checkSql = '''select top 3 * from
(select 
s.FileID,
s.ProcessDate,
s.ProcessResult,
s.MessageID,
s.FileLength,
s.Data.value('declare namespace n0="http://finance.group.volvo.com/vendorinvoicedetails/1_0";
(/n0:SyncVendorInvoiceDetails/VendorInvoiceDetails/VendorInvoiceHeader/CompanyCode)[1]', 'nvarchar(20)') AS CompanyCode,
s.Data
from satbSourceData s
where s.FileType = 'VendorInvoiceDetails') as t
where ProcessResult = 'Error'
    and MessageID <> '414D512046524F444F5F423120202020AA90F15B24C61FEF'
order by T.ProcessDate desc'''

sapDatabaseConfiguration = getDatabaseConfiguration('SAP')

conn = pyodbc.connect('Driver={driver};'
        'Server={server};'
        'Database={database};'
        'UID={username};'
        'PWD={password}'.format(driver = sapDatabaseConfiguration['driver'], server = sapDatabaseConfiguration['server'], database = 'SAP',
                    username = sapDatabaseConfiguration['userName'], password = sapDatabaseConfiguration['password']))
print(Fore.GREEN + 'Connected to SAP database.\n')
print(Style.RESET_ALL)
cursor = conn.cursor()
cursor.execute(checkSql)

wb = openpyxl.workbook.Workbook()
overallSheet = wb.active
overallSheet.title = 'Overall'
overallSheet['A1'] = 'FileID'
overallSheet['B1'] = 'ProcessDate'
overallSheet['C1'] = 'ProcessResult'
overallSheet['D1'] = 'MessageID'
overallSheet['E1'] = 'FileLength'
overallSheet['F1'] = 'CompanyCode'

#Read configuration
configurationContent = '\n'.join(open('AppSettings.xml').readlines())
configurationDocument = BeautifulSoup(configurationContent, features='xml')

headerRowDisplayStyle = {'fontColor':'FFFFFF', 'backgroundColor':'000080'}
highLightRowDisplayStyle = {'fontColor':'FFFFFF', 'backgroundColor':'FFFF00'}
uiConfiguration = configurationDocument.find('uiSettings')
if not uiConfiguration is None:
    headerRowConfiguration = uiConfiguration.find('headerRow')
    if not headerRowConfiguration is None:
        if not headerRowConfiguration.get('fontColor') is None:
            headerRowDisplayStyle['fontColor'] = headerRowConfiguration.get('fontColor')
        if not headerRowConfiguration.get('backgroundColor') is None:
            headerRowDisplayStyle['backgroundColor'] = headerRowConfiguration.get('backgroundColor')
    highLightRowConfiguration = uiConfiguration.find('highLightRow')
    if not highLightRowConfiguration is None:
        if not highLightRowConfiguration.get('fontColor') is None:
            highLightRowDisplayStyle['fontColor'] = highLightRowConfiguration.get('fontColor')
        if not highLightRowConfiguration.get('backgroundColor') is None:
            highLightRowDisplayStyle['backgroundColor'] = highLightRowConfiguration.get('backgroundColor')

validTypeOfInvoiceValues = []
validDocumentTypeValues = []
validationConfiguration = configurationDocument.find('validationSettings')
if not validationConfiguration is None:
    typeOfInvoiceConfiguration = validationConfiguration.find('typeOfInvoice')
    if not typeOfInvoiceConfiguration is None:
        for item in typeOfInvoiceConfiguration.find_all('item'):
            if not item.get('value') is None:
                validTypeOfInvoiceValues.append(item.get('value'))
    
    documentTypeConfiguration = validationConfiguration.find('documentType')
    if not documentTypeConfiguration is None:
        for item in documentTypeConfiguration.find_all('item'):
            if not item.get('value') is None:
                validDocumentTypeValues.append(item.get('value'))

headerRowFill = PatternFill(start_color=headerRowDisplayStyle['backgroundColor'],
                   end_color=headerRowDisplayStyle['backgroundColor'],
                   fill_type='solid')
highLightRowFill = PatternFill(start_color=highLightRowDisplayStyle['backgroundColor'],
                   end_color=highLightRowDisplayStyle['backgroundColor'],
                   fill_type='solid')
headerRowFont = Font(color=headerRowDisplayStyle['fontColor'])
highLightRowFont = Font(color=highLightRowDisplayStyle['fontColor'])

maxColumn = overallSheet.max_column
for i in range(1, maxColumn + 1):
    overallSheet.cell(row = 1, column = i).fill = headerRowFill
    overallSheet.cell(row = 1, column = i).font = headerRowFont

rowIndex = 2
for row in cursor.fetchall():
    fileID = row[0]
    processDate = row[1]
    processResult = row[2]
    messageID = row[3]
    fileLength = row[4]
    companyCode = row[5]
    message = row[6]

    overallSheet['A' + str(rowIndex)] = fileID
    overallSheet['B' + str(rowIndex)] = processDate
    overallSheet['C' + str(rowIndex)] = processResult
    overallSheet['D' + str(rowIndex)] = messageID
    overallSheet['E' + str(rowIndex)] = fileLength
    overallSheet['F' + str(rowIndex)] = companyCode

    print('Analyzing content of File #{fID}\n'.format(fID = fileID))

    rowIndex += 1
    fileSheet = wb.create_sheet(title=str(fileID), index=1)
    fileSheet['A1'] = 'Accounting Document' # DocumentNumber
    fileSheet['B1'] = 'Year' # FiscalYear
    fileSheet['C1'] = 'Doc. Date' # InvoiceDate
    fileSheet['D1'] = 'Pstng Date' # PostingDate
    fileSheet['E1'] = 'Doc. No.' # InvoiceDocumentNumber
    fileSheet['F1'] = 'PO Number' # PurchasingDocumentNumber
    fileSheet['G1'] = 'P++' # OrderReference
    fileSheet['H1'] = 'PO line item' # PurchasingDocumentNumberItem
    fileSheet['I1'] = 'Vendor #' # VendorNumber
    fileSheet['J1'] = 'PO Currency' # CurrencyCode
    fileSheet['K1'] = 'Type of Invoice' # TypeOfInvoice
    fileSheet['L1'] = 'Document Type' # DocumentType
    fileSheet['M1'] = 'Quantity' # Quantity
    fileSheet['N1'] = 'GrossValue' # GrossValue
    fileSheet['O1'] = 'Authorised Value'
    fileSheet['P1'] = 'Invoice Value'

    maxColumn = fileSheet.max_column
    for i in range(1, maxColumn + 1):
        fileSheet.cell(row = 1, column = i).fill = headerRowFill
        fileSheet.cell(row = 1, column = i).font = headerRowFont

    fileRowIndex = 2

    soup = BeautifulSoup(message, features='xml')
    for header in soup.find_all('VendorInvoiceHeader'):
        documentNumber = header.find('DocumentNumber').get_text()
        fiscalYear = header.find('FiscalYear').get_text()
        invoiceDate = header.find('InvoiceDate').get_text()
        postingDate = header.find('PostingDate').get_text()
        invoiceDocumentNumber = header.find('InvoiceDocumentNumber').get_text()
        orderReference = header.find('OrderReference').get_text()
        vendorNumber = header.find('VendorNumber').get_text()
        currencyCode = header.find('CurrencyCode').get_text()
        typeOfInvoice = header.find('TypeOfInvoice').get_text()
        documentType = header.find('DocumentType').get_text()
        invoiceNumber = header.find('InvoiceNumber').get_text()

        print('\tFound DocumentNumber {docNumber} with TypeOfInvoice as {typeOfInv}, DocumentType as {docType}\n'.format(docNumber = documentNumber, typeOfInv = typeOfInvoice, docType = documentType))

        for item in header.parent.find_all('VendorInvoiceItem'):
            purchasingDocumentNumber = item.find('PurchasingDocumentNumber').get_text()
            purchasingDocumentNumberItem = item.find('PurchasingDocumentNumberItem').get_text()
            quantity = item.find('Quantity').get_text()
            grossValue = item.find('GrossValue').get_text()
            
            fileSheet['A' + str(fileRowIndex)] = documentNumber
            fileSheet['B' + str(fileRowIndex)] = fiscalYear
            fileSheet['C' + str(fileRowIndex)] = invoiceDate
            fileSheet['D' + str(fileRowIndex)] = postingDate
            fileSheet['E' + str(fileRowIndex)] = invoiceDocumentNumber
            fileSheet['F' + str(fileRowIndex)] = purchasingDocumentNumber
            fileSheet['G' + str(fileRowIndex)] = orderReference
            fileSheet['H' + str(fileRowIndex)] = purchasingDocumentNumberItem
            fileSheet['I' + str(fileRowIndex)] = vendorNumber
            fileSheet['J' + str(fileRowIndex)] = currencyCode
            fileSheet['K' + str(fileRowIndex)] = typeOfInvoice
            if typeOfInvoice not in validTypeOfInvoiceValues:
                fileSheet['K' + str(fileRowIndex)].fill = highLightRowFill
                fileSheet['K' + str(fileRowIndex)].font = highLightRowFont
            fileSheet['L' + str(fileRowIndex)] = documentType
            if documentType not in validDocumentTypeValues:
                fileSheet['L' + str(fileRowIndex)].fill = highLightRowFill
                fileSheet['L' + str(fileRowIndex)].font = highLightRowFont
            fileSheet['M' + str(fileRowIndex)] = quantity
            fileSheet['N' + str(fileRowIndex)] = grossValue

            #Check Invoice from Truck Center Database
            truckCenterID = ''
            checkInvoiceSql = """select 
                                    item.TruckCenterID
                                from satbVendorInvoices as inv
                                    left join satbVendorInvoiceItems as item on inv.InvoiceID = item.InvoiceID
                                where inv.OrderReference = '{orderRef}'
                                    and inv.InvoiceNumber = '{invNo}'
                                    and item.PurchasingDocumentNumberItem = '{purDocNoItem}'""".format(orderRef = orderReference, invNo = invoiceNumber, purDocNoItem = purchasingDocumentNumberItem)
            invoiceCursor = conn.cursor()
            invoiceCursor.execute(checkInvoiceSql)
            invoiceRow = invoiceCursor.fetchone()
            if invoiceRow != None:
                truckCenterID = invoiceRow[0]

                checkTruckCenterSql = """select 
                                                AuthorisedValue, InvoiceValue
                                        from ppvwAllOrderSummary
                                        where ActualOrderNo = '{poNumber}'""".format(poNumber = orderReference)
                truckCenterConn = getTruckCenterConnection(truckCenterID)
                truckCenterCursor = truckCenterConn.cursor()
                truckCenterCursor.execute(checkTruckCenterSql)
                truckCenterRow = truckCenterCursor.fetchone()
                if truckCenterRow != None:
                    fileSheet['O' + str(fileRowIndex)] = truckCenterRow[0]
                    fileSheet['P' + str(fileRowIndex)] = truckCenterRow[1]

            fileRowIndex += 1

wb.save('xmlMessage.xlsx')
print(Fore.GREEN + 'finished.\n')
conn.close()
print('Disconnected from SAP database.\n')
eudConn.close()
print('Disconnected from EUD database.\n')
for connectionKey in databaseConnectionPool.keys():
    databaseConnectionPool[connectionKey].close()
    print('Disconnected from Truck Center {truckCenterID} database.\n'.format(truckCenterID = connectionKey))
print(Style.RESET_ALL)

