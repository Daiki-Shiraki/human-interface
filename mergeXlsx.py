import pandas as pd
import pandas.io.formats.excel
import openpyxl
from openpyxl.styles.fonts import Font
import methods as mt


def main():
    file_path = input("得点の入ったエクセルファイルパスを指定してください")
    name =mt.GetOutputFileName(file_path)
    if(file_path == None):
        print("パスが指定されていません")
        return
    score_wb = pd.ExcelFile(file_path)
    score_wb_name = score_wb.sheet_names
    score_sheet_df = score_wb.parse(score_wb_name[0])
    
    output_wb = pd.ExcelFile('./ヒューマンインタフェース特論評価まとめ.xlsx')
    output_wb_name = output_wb.sheet_names
    output_sheet_df = output_wb.parse(output_wb_name[0])
    
    for index, row in score_sheet_df.iterrows():
        number = mt.Number2String(row['学籍番号（半角）']) 
        row = row[2:12]
        index = output_sheet_df.query('学生ID == \"' + number + '\"').index[0]       
        for i, out_row in enumerate(row):
            output_sheet_df.iloc[index, i + 1] = row[i]
        output_sheet_df.iloc[index, 12] = '=AVERAGE(B'+ str(index + 2) + ':K' + str(index + 2) + ')'
    for i in range(12):
        output_sheet_df.iloc[77, i + 1] = '=AVERAGE('+ chr(66 + i) +'2:' + chr(66 + i) + '77)'
        
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    output_sheet_df.to_excel(name, sheet_name = '出席簿', index = False)
    
    
    wb1 = openpyxl.load_workbook(filename = name)
    ws1 = wb1.worksheets[0]
    ws1.column_dimensions['A'].width = 16

    font = Font(name='游ゴシック Regular')
    
    for row in ws1:
        for cell in row:
            ws1[cell.coordinate].font = font
    wb1.save(name)
    return


if __name__ == "__main__":
    main()
    
