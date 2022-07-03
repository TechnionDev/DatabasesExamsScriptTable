import openpyxl
import os
import re
from urllib.request import pathname2url
filename = "examsLinksTemplate.xlsx"
prev_exams_dir = "./prevExams"


def path_to_year_semester_moed_is_solution(path):
    if 'DS_Store' in path or 'idterm' in path:
        return
    if "Sp" in path:
        term = 'Spring'
    elif 'W' in path:
        term = 'Winter'
    else:
        raise Exception("Term not found")

    try:
        # extract year in the form of [0-9]{2,4}
        year = "20"+re.search(r'\d{2,4}', path).group(0)[-2:]

        # Extract moed from path (follows the work Moed ignore case)
        # Replace all hebrew letters with coresponding latin
        path = path.replace('חורף', 'Winter')
        path = path.replace('אביב', 'Spring')
        path = path.replace('מועד', 'Moed')
        path = path.replace('א', 'A')
        path = path.replace('ב', 'B')

        moed = re.search(
            r'(Moed.?|/|Exam|\d\d|Maman)([AB])', path, re.IGNORECASE).group(2).upper()
        # Check if "sol" apear in the path case insensitive
        is_sol = re.search(r'sol', path, re.IGNORECASE) is not None
        print(f'extracted {year=} {term=} {moed=} {is_sol} \t:\t{path=}')
        return year, term, moed, is_sol
    except Exception as e:
        print(f'Failed to extract from {path=}.')
        print('Please fix the file name to the form SpringMoedA2022')
        raise e


def main():

    print(f'Iterating over {prev_exams_dir=}')

    # Open excel file
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    row_offset_from_term_moed = {
        'Winter': {'A': 0, 'B': 1},
        'Spring': {'A': 2, 'B': 3},
    }

    # Fill in missing data
    FIRST_DATA_ROW = 16
    last_row = sheet[FIRST_DATA_ROW]
    not_found_count = 0
    for row in sheet.iter_rows(min_row=FIRST_DATA_ROW):
        if not_found_count >= 4:
            break

        if row[0].value == None:
            not_found_count += 1
            for i in range(2):
                row[i].value = last_row[i].value
                row[i].style = last_row[i].style
                row[i]._style = last_row[i]._style
                continue

        not_found_count = 0
        last_row = row

    # Iterate files over the prev_exams_dir recursively
    for subdir, dirs, files in os.walk(prev_exams_dir):
        for file in files:
            path = os.path.join(subdir, file)
            ret = path_to_year_semester_moed_is_solution(path)
            if not ret:
                continue
            # encode path to url file:// path
            url = f'{pathname2url(path[2:])}'

            year, semester, moed, is_sol = ret
            cell_value = moed + (" Sol" if is_sol else "")
            # cell_value = f'=HYPERLINK("{url}", "'+(moed + " Sol" if is_sol else "") + '")'
            year = int(year)
            #  Find the cell with `year`
            for row in sheet.iter_rows():
                if row[0].value == year:
                    link_cell_row = row[0].row + \
                        row_offset_from_term_moed[semester][moed]
                    link_cell_col = 5 if is_sol else 3
                    link_cell = sheet[link_cell_row][link_cell_col]
                    link_cell.value = cell_value
                    link_cell.hyperlink = url
                    link_cell.style = 'Hyperlink'
                    # Put todo (" ") value in the "done" cell
                    sheet[link_cell_row][4].value = ' '
                    # Excel conditional formatting number > 2015
                    conditional_formula = '=IF(AND(ISNUMBER(A{row}),A{row}>2015),"",A{row})'

    wb.save('examsLinks.xlsx')


if __name__ == "__main__":
    main()
