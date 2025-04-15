import xlwings as xw

def transfer_contents():
    app = xw.App(visible=False)
    
    wb_mfq1 = app.books.open('MFQ1.xlsx')
    wb_sf9 = app.books.open('SF9.xlsb')
    wb_sf10 = app.books.open('sf10.xlsx')
    
    front_sf9 = wb_sf9.sheets['FRONT']
    front_sf10 = wb_sf10.sheets['FRONT']
    mfq1_sheet = wb_mfq1.sheets['Sheet1']
    
    transfers = [
        ('B1', 'Q26', ['AS23', 'AS66']),
        ('B2', 'T26', ['AS25', 'AS68']),
        ('B3', 'R29', ['G25', 'G68']),
        ('F1', 'P40', []),
        ('F2', 'S40', ['A49', 'A92']),
        ('F3', 'R28', ['BA23', 'BA66'])
    ]
    
    for source, sf9_dest, sf10_dests in transfers:
        value = mfq1_sheet.range(source).value
        front_sf9.range(sf9_dest).value = value
        
        for sf10_dest in sf10_dests:
            front_sf10.range(sf10_dest).value = value
    
    wb_sf9.save()
    wb_sf10.save()
    
    wb_mfq1.close()
    wb_sf9.close()
    wb_sf10.close()
    app.quit()

if __name__ == '__main__':
    transfer_contents()