import json
import os

import xlrd3
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

OUTPUT_POSITION_MAPPING = dict()
INPUT_COLUMN_MAPPING = list()
OUTPUT_PATH = "./output"
TEMPLATE_PATH = ""
addtion_infos = dict()

"""
    预处理每一行数据，符合填充要求，然后填充
"""
def process_line(sheet, row):
    print("process line {}".format(row))
    line = []
    for col in range(sheet.ncols):
        # 获取单元格的值
        value = sheet.cell_value(row, col)
        # 第一列确保序号为整数
        if col == 0:
            value = int(value)
        if col == 2:
            value = '收货单位：' + value
        if col == 3:
            value = '收货地址：' + value
        # 拼接收件人+ 地址
        if col == 4:
            next_vaule = str(sheet.cell_value(row, col + 1)).replace(' ', '')
            value = '收货联系人：' + value + ' ' + str(int(float(next_vaule)))
        if col == 6:
            value = str(int(float(value)))
        line.append(value)
    process_docx(line)


def process_file(input_file):
    global OUTPUT_POSITION_MAPPING
    global INPUT_COLUMN_MAPPING
    with xlrd3.open_workbook(input_file) as wb:
        for sheet in wb.sheets():
            print("开始处理文件[{}-{}]".format(input_file, sheet.name))
            if sheet.ncols < len(INPUT_COLUMN_MAPPING):
                print("文件{}-{}列数少于预期，跳过处理".format(input_file, sheet.name))
            global addtion_infos
            # hard code here
            addtion_infos['item'] = sheet.cell_value(0, 6)
            for row in range(1, sheet.nrows):
                process_line(sheet, row)
            break


def process():
    check_output_path()
    for dirpath, dirnames, filenames in os.walk('./input'):
        for filename in filenames:
            if filename.find("~$") != -1:
                continue
            process_file(dirpath + '/' + filename)


def check_output_path():
    if not os.path.exists(OUTPUT_PATH):
        os.makedirs(OUTPUT_PATH)
    else:
        files = os.listdir(OUTPUT_PATH)
        # 遍历所有的文件和文件夹
        for file in files:
            # 拼接完整的路径
            file_path = os.path.join(OUTPUT_PATH, file)
            # 判断是否是文件
            if os.path.isfile(file_path):
                # 删除文件
                os.remove(file_path)


def process_docx(line):
    global OUTPUT_POSITION_MAPPING
    global INPUT_COLUMN_MAPPING
    document = Document(TEMPLATE_PATH)
    table = document.tables[0]
    # 补充excel中每一行的信息
    for idx in range(len(line)):
        pos = OUTPUT_POSITION_MAPPING.get(INPUT_COLUMN_MAPPING[idx])
        if pos is None:
            continue
        cell = table.cell(pos[0], pos[1])
        cell.text = str(line[idx])
        if idx not in (2, 3, 4):
            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell.paragraphs[0].runs[0].bold = True
    # 补充额外信息
    for key in addtion_infos:
        pos = OUTPUT_POSITION_MAPPING.get(key)
        if pos is None:
            continue
        cell = table.cell(pos[0], pos[1])
        cell.text = addtion_infos.get(key)
        if key != 'three':
            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell.paragraphs[0].runs[0].bold = True

    # document.add_picture("./1.jpg")  # 等同于doc.add_paragraph().add_run().add_picture()
    # document.save('./test.docx')

    output_file_name = OUTPUT_PATH + '/' + '{}、{}.docx'.format(line[0], line[1])
    document.save(output_file_name)


def load_conf():
    with open('./setting.json', 'r', encoding='utf-8') as f:
        conf = json.load(f)
    global INPUT_COLUMN_MAPPING
    global OUTPUT_POSITION_MAPPING
    global OUTPUT_PATH
    global TEMPLATE_PATH
    global addtion_infos
    INPUT_COLUMN_MAPPING = conf.get("INPUT_COLUMN_MAPPING")
    OUTPUT_POSITION_MAPPING = conf.get("OUTPUT_POSITION_MAPPING")
    OUTPUT_PATH = conf.get("OUTPUT_PATH")
    TEMPLATE_PATH = conf.get("TEMPLATE_PATH")
    addtion_infos = conf.get("ADDTIONAL_INFO")
    pass

def main():
    load_conf()
    process()


if __name__ == "__main__":
    main()
