# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os
import json
import openpyxl
import time
import subprocess

delete_params = {
    "Yes": ["да", "yes", "y", "д", "удаляй нахой все"],
    "No": ["нет", "no", "n", "н", "не", "не трожь бля"]
}


def not_empty_list(lst):
    return max([_ is not None for _ in lst])


def read_excel(file_name):
    xlsx_data = openpyxl.load_workbook(file_name)
    sheet_data = xlsx_data.active
    max_row_count = sheet_data.max_row
    max_col_count = sheet_data.max_column
    rows_list = []
    for row in range(1, max_row_count + 1):
        row_data = []
        for col in range(1, max_col_count + 1):
            row_data.append(sheet_data.cell(row=row, column=col).value)
        if not_empty_list(row_data):
            rows_list.append(row_data)
    return rows_list


def create_json(log, passw):
    data_dict = {
        "Enabled": True,
        "SteamLogin": f"{log}",
        "SteamPassword": f"{passw}",
    }
    with open (f'{log}_.json', "w+") as file:
        json.dump(data_dict, file)


# Press the green button in the gutter to run the script.
def delete_old_files():
    now = time.time()
    folder = os.getcwd()

    files = [os.path.join(folder, filename) for filename in os.listdir(folder)]
    for filename in files:
        if ".json" in filename:
            if (now - os.stat(filename).st_mtime) > 100:
                command = "rm {0}".format(filename)
                subprocess.call(command, shell=True)


if __name__ == '__main__':
    directory = os.getcwd()
    file_name = input('Enter xlsx file name: ')
    delete_old = input('Delete old files? Enter y/n: ')
    print(1)
    if delete_old in delete_params["Yes"]:
        delete_old_files()
    elif delete_old in delete_params["No"]:
        print("ну и хуй с тобой")
    else:
        print("ты че делаешь дядя алло")

    excel_data = read_excel(file_name=file_name)
    for row in excel_data:
        create_json(row[2], row[3])


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
