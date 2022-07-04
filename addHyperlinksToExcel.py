import openpyxl
import os
import re
from urllib.request import pathname2url
TEMPLATE_FILENAME = "examsLinksTemplate.xlsx"
PREV_EXAMS_DIR = "./prevExams"
OUTPUT_FILENAME = 'examsLinks.xlsx'

warning_text = ""


def path_to_year_semester_moed_is_solution(path):
    global warning_text
    # Replace all hebrew letters with coresponding latin
    path = path.replace('חורף', 'Winter')
    path = path.replace('אביב', 'Spring')
    path = path.replace('מועד', 'Moed')
    path = path.replace('פתרון', 'Solution')
    path = path.replace('בוחן אמצע', 'Midterm')
    path = path.replace('בוחן', 'Midterm')
    path = path.replace('אמצע', 'Midterm')
    path = path.replace('א', 'A')
    path = path.replace('ב', 'B')

    try:
        # Ignore the PREV_EXAMS_DIR part of the path
        path = path[path.find(PREV_EXAMS_DIR)+len(PREV_EXAMS_DIR):]
        if re.search(r'midterm', path, re.IGNORECASE) is not None or 'skip' in path:
            # Skip midterms and skips
            return None

        if 'DS_Store' in path or 'idterm' in path:
            return
        if "Sp" in path:
            term = 'Spring'
        elif 'W' in path:
            term = 'Winter'
        else:
            raise Exception("Term not found")

        # extract year in the form of [0-9]{2,4}
        year = "20"+re.search(r'\d{2,4}', path).group(0)[-2:]

        # Extract moed from path (follows the work Moed ignore case)
        moed = re.search(
            r'(Moed.?|/|Exam|\d\d|Maman)([ABC])', path, re.IGNORECASE).group(2).upper()
        # Check if "sol" apear in the path case insensitive
        is_sol = re.search(r'sol', path, re.IGNORECASE) is not None
        print(f'extracted {year=} {term=} {moed=} {is_sol} \t:\t{path=}')
        return year, term, moed, is_sol
    except Exception as e:
        print(f'Failed to extract from {path=}.')
        print('Please fix the file name to the form SpringMoedA2022')
        raise e


def main():
    global warning_text

    print(f'Iterating over {PREV_EXAMS_DIR=}')

    # Open excel file
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    sheet = wb.active

    # row_offset_from_term_moed = {
    #     'Winter': {'A': 0, 'B': 1},
    #     'Spring': {'A': 2, 'B': 3},
    # }
    heb_term_to_eng = {
        "חורף": 'Winter',
        "אביב": 'Spring',
    }
    heb_moed_to_eng = {
        "א": 'A',
        "ב": 'B',
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
                if row[i].value == None:
                    row[i].value = last_row[i].value
                    row[i]._style = last_row[i]._style

            last_row = row
            continue

        not_found_count = 0
        last_row = row

    # Iterate files over the prev_exams_dir recursively
    for subdir, dirs, files in os.walk(PREV_EXAMS_DIR):
        for file in files:
            path = os.path.join(subdir, file)
            ret = path_to_year_semester_moed_is_solution(path)
            if not ret:
                warning_text += f'Skipped path: {path}\n'
                continue
            # encode path to url file:// path
            url = f'{pathname2url(path[2:])}'

            year, semester, moed, is_sol = ret
            if (moed == 'C'):
                warning_text += f'Skipped adding C term exam: {path}\n'
            cell_value = "Moed " + moed + (" Sol" if is_sol else "")
            # cell_value = f'=HYPERLINK("{url}", "'+(moed + " Sol" if is_sol else "") + '")'
            year = int(year)
            #  Find the cell with `year`
            for row in sheet.iter_rows(min_row=FIRST_DATA_ROW):
                if row[0].value == year and \
                    row[1].value and heb_term_to_eng[row[1].value] == semester and \
                        row[2].value and heb_moed_to_eng[row[2].value] == moed:
                    link_cell_col = 5 if is_sol else 3
                    link_cell = row[link_cell_col]
                    link_cell.value = cell_value
                    link_cell.hyperlink = path[2:]
                    link_cell.style = 'Hyperlink'
                    # Excel conditional formatting number > 2015
                    conditional_formula = '=IF(AND(ISNUMBER(A{row}),A{row}>2015),"",A{row})'

    wb.save(OUTPUT_FILENAME)
    if warning_text:
        print("WARNING:\n" + warning_text)


if __name__ == "__main__":
    main()
