#!/usr/bin/python3

import os
import argparse
import xlsxwriter

def get_content(fname): 
    with open(fname) as f:
        return f.readlines()

def check_file_existence(fname):
    answer = True
    if not os.path.isfile(fname):
        print("Input file does not exists. Please check %s" %fname)
        answer = False
    return answer

def get_human(human_data):
    tmp = {}
    human_list = human_data.split()
    tmp['name'] = human_list[0].strip()
    tmp['surname'] = human_list[1].strip()
    tmp['age'] = int(human_list[2].strip())
    tmp['profession'] = human_list[3].strip()
    return tmp

def get_humans_list(ml):
    return [get_human(line) for line in ml if len(line) > 1]

def sort_by_criteria(humans, criteria):
    if criteria == "s":
	    humans.sort(key=lambda el: el["surname"])
    if criteria == "n":
	    humans.sort(key=lambda el: el["name"])
    if criteria == "a":
	    humans.sort(key=lambda el: el["age"])
    if criteria == "p":
	    humans.sort(key=lambda el: el["profession"])
    return humans

# def write_into_file(fname, content):
#     with open(fname, "w") as f:
#         for human in content:
#             line = human['name'] + " " \
#                 + human['surname'] + " " \
#                 + str(human['age']) + " " \
#                 + human['profession']
#             f.write(line + "\n")

def write_into_xlsx_file(fname, content):
    workbook = xlsxwriter.Workbook(fname)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    yellowbg = workbook.add_format({'bg_color': 'yellow'})
    greenbg = workbook.add_format({'bg_color': 'green'}) 
    worksheet.write('A1', "Name", bold)
    worksheet.write('B1', "Surname", bold)
    worksheet.write('C1', "Age", bold)
    worksheet.write('D1', "Profession", bold)
    row = 1
    for el in content:
        if int(el['age']) > 20:
            worksheet.write(row, 0, el['name'], greenbg)
            worksheet.write(row, 1, el['surname'], greenbg)
            worksheet.write(row, 2, el['age'], greenbg)
            worksheet.write(row, 3, el['profession'], greenbg)
            row+=1
        else:
            worksheet.write(row, 0, el['name'], yellowbg)
            worksheet.write(row, 1, el['surname'], yellowbg)
            worksheet.write(row, 2, el['age'], yellowbg)
            worksheet.write(row, 3, el['profession'], yellowbg)
            row+=1
    workbook.close()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--filename")
    parser.add_argument("-o", "--outputfile")
    parser.add_argument("criteria", choices = ['n','s','a','p'], nargs = '?', const=1, default = 'n')
    args = parser.parse_args()
    checking = check_file_existence(args.filename)
    if not checking:
        exit()
    cnt = get_content(args.filename)
    humans_list = get_humans_list(cnt)	
    humans_list = sort_by_criteria(humans_list, args.criteria)
    write_into_xlsx_file(args.outputfile, humans_list)

main()
