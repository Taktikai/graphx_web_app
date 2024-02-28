from tkinter import Tk, filedialog
import xlsxwriter
from Parser import *


def create_xls(filename):
    root = Tk()
    root.withdraw()
    default_filename = filename[:-8].replace('_PROXY', '') + '.xlsx'
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_filename, filetypes=[("Excel Workbook", "*.xlsx")])
    if not file_path:
        return
    
    root_group = get_all_devices(filename)
    counts = count_devices_by_type(root_group, filename)
    workbook = xlsxwriter.Workbook(file_path)
    worksheet1 = workbook.add_worksheet('Összeszerelés')
    worksheet2 = workbook.add_worksheet('Anyaglista')

    attr_name_format = workbook.add_format({'border': 2, 'bold': True, 'bg_color': '#7AEC92'})
    group_format = workbook.add_format({'border': 2, 'bold': True, 'bg_color': '#EDED79'})
    red_bg_format = workbook.add_format({'bg_color': 'red'})
    attr_value_format = workbook.add_format({'border': 1, 'bg_color': '#BDF5C9'})


    row = 0

    def traverse(group, level=0):
        nonlocal row
        if group.closed:
            return
        if level > 0:
            group_info = f"{group.típus or ''} || {group.elnevezés or ''}"
            worksheet1.merge_range(row, 0, row, 10 - min(level, 3), group_info, group_format)
            row += 1
            if any(isinstance(item, Device) for item in group.content.values()):
                headers = ['Típus', 'Elnevezés', 'SN', 'Saját Cím', 'Fogadó Eszköz', 'Kommunikáció', 'Fogadó Interface', 'Hálózatazonosító Frekvencia']
                worksheet1.write_row(row, 0, headers, attr_name_format)
                row += 1
        for item in group.content.values():
            if isinstance(item, Device):
                worksheet1.write_row(row, 0, [item.típus, item.elnevezés, item.SN, item.saját_cím, item.fogadó_eszköz, item.kommunikáció, item.fogadó_interface, item.hálózatazonosító_frekvencia], attr_value_format)
                row += 1
        for item in group.content.values():
            if isinstance(item, Group):
                traverse(item, level+1)

    traverse(root_group)

    row = 0
    col = 0

    worksheet2.write(row, col, 'Eszköz', group_format)
    worksheet2.write(row, col + 1, 'Darabszám', group_format)

    row += 1

    for key, value in counts.items():
        worksheet2.write(row, col, key, attr_value_format)
        worksheet2.write(row, col + 1, value, attr_value_format)
        row += 1


    workbook.close()
    return file_path