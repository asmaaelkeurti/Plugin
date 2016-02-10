application = WIN32OLE.new('Excel.Application')
workbook = application.Workbooks.Open("C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a\\test")
worksheet=workbook.Worksheets("Sheet2")
            
worksheet.Range("A200:Z250").Clear