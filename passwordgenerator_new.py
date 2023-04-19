import xlsxwriter as xlsx
import random

colors={
    "orange":{"Dark":"#ED7D31", "Light":"#FCE4D6"},
    "gold":{"Dark":"#FFC000", "Light": "#FFF2CC"},
    "blue":{"Dark":"#4472C4", "Light":"#DDEBF7"},
    "green":{"Dark":"#70AD47", "Light":"#E2EFDA"},
    "black":{"Dark":"#000000", "Light":"#D9D9D9"},
    "red":{"Dark":"#990000", "Light":"#FFC8C8"},
    "purple":{"Dark":"#9900CC", "Light":"#E1CDFF"},
    "navy":{"Dark":"#103D7E", "Light":"#CDD5E4"},
    "brown":{"Dark":"#664229", "Light":"#E5D3B3"},
    "pink":{"Dark":"#FF64AA", "Light":"#FFD7F0"}
}

def genPasses(numPassGenerated, numWordsPerPass):
    file = open("words.txt", "r")
    wordList=[word.strip() for word in file]
    passes=[]
    for x in range (0, numPassGenerated):
        p=""
        for y in range (0, numWordsPerPass):
            p+=str(wordList[random.randrange(len(wordList))])+str(random.randrange(10))+"-"
        passes.append(p[:-1])
    return passes

def writeSeparator(workbook, worksheet, color, row):
    merge_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_name': 'Arial',
        'font_size': 20,
        'bold': 1,
        'font_color': colors[color]["Dark"]})
    r="A"+str(row)+":E"+str(row+1)
    worksheet.merge_range(r,color.capitalize(),merge_format)

def writeHeader(wb, ws, color, row):
    format = wb.add_format({
        'border': 1,
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': 1,
        'font_color': "white",
        'bg_color': colors[color]["Dark"]})
    header=["Name","Password","User","System","Service"]
    start="A"+str(row)
    ws.write_row(start,header,format)

def writePass(wb, ws, color, row, name, password, f):
    format_color = wb.add_format({
        'border': 1,
        'font_name': 'Calibri',
        'font_size': 11,
        'bg_color': colors[color]["Light"]})
    format_blank = wb.add_format({
        'border': 1,
        'font_name': 'Calibri',
        'font_size': 11})
    bold = wb.add_format({'bold': True})
    cell="A"+str(row)
    pRow=[password, "", "", ""]
    if f:
        ws.write_rich_string(cell, bold, name[0], name[1:], format_color)
        cell="B"+str(row)
        ws.write_row(cell, pRow, format_color)
    else:
        ws.write_rich_string(cell, bold, name[0], name[1:], format_blank)
        cell="B"+str(row)
        ws.write_row(cell, pRow, format_blank)

def writePasswords(wb, ws, color):
    rows=["Alpha","Bravo","Charlie","Delta","Echo","Foxtrot","Golf","Hotel","India","Juliett","Kilo","Lima","Mike","November","Oscar","Papa","Quebec","Romeo","Sierra","Tango","Uniform","Victor","Whiskey","Xray","Yankee","Zulu","One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten"]
    passes=genPasses(36, 2)
    writeSeparator(wb, ws, color, 1)
    writeHeader(wb, ws, color, 3)
    for x in range(18):
        writePass(wb, ws, color, x+4, rows[x], passes[x], x%2==0)
    writeSeparator(wb, ws, color, 22)
    writeHeader(wb, ws, color, 24)
    for x in range(18):
        writePass(wb, ws, color, x+25, rows[x+18], passes[x+18], x%2==0)
    writeSeparator(wb, ws, color, 43)
    
    return max([len(p) for p in passes])

def fitColumns(ws, maxLenPass):
    pixels=548
    ws.set_column_pixels(0, 0, 66)
    pixels-=66
    ws.set_column_pixels(1, 1, int(7.7*maxLenPass))
    pixels-=int(7.7*maxLenPass)
    ws.set_column_pixels(2, 2, 130)
    pixels-=130
    pixels//=1.484
    ws.set_column_pixels(3, 3, pixels)
    ws.set_column_pixels(4, 4, pixels)


def main():
    colorList=list(colors.keys())
    workbook=xlsx.Workbook("passes.xlsx")
    worksheets=[workbook.add_worksheet(c.capitalize()) for c in colorList]
    for x in range(len(colorList)):
        maxLen = writePasswords(workbook, worksheets[x], colorList[x])
        fitColumns(worksheets[x], maxLen)
    workbook.close()

if __name__=="__main__":
    main()