# open csf314-2023-Marks.xlsx

import openpyxl

# open the A1 sheet
workbook = openpyxl.load_workbook('csf314-2023-Marks.xlsx')


# get the sheet "Registered List"
sheet_registered_list = workbook['Registered List']


def longest_common_substring_length(s1, s2):
    m = [[0] * (1 + len(s2)) for i in range(1 + len(s1))]

    for x in range(1, 1 + len(s1)):
        for y in range(1, 1 + len(s2)):
            if s1[x - 1] == s2[y - 1]:
                m[x][y] = m[x - 1][y - 1] + 1
            else:
                m[x][y] = max(m[x - 1][y], m[x][y - 1])

    return m[len(s1)][len(s2)]


def search_in_registered_list(name):
    max_lcs_length = 0
    ans_row = 0

    # find the student1 in the registered list name = Column B(first name) + Column C(last name)
    for registered_list_row in range(2, 99):
        registered_list_name = sheet_registered_list['A' + str(
            registered_list_row)].value + " " + sheet_registered_list['B' + str(
                registered_list_row)].value

        # if spaces is there after the name remove it
        # registered_list_name = registered_list_name.strip()

        # find the longest common substring length
        lcs_length = longest_common_substring_length(
            name, registered_list_name)

        # if the lcs length is greater than max_lcs_length
        if lcs_length > max_lcs_length:
            max_lcs_length = lcs_length
            ans_row = registered_list_row

    return ans_row


def transfer_to_registered_list(sheet_name, max_row, input_marks_col_name, output_marks_col_name):
    sheet = workbook[sheet_name]
    # iterate over rows 2 to 48 of Column B C Q
    for row in range(2, max_row):
        # get the student1 name
        student1_name = sheet['B' + str(row)].value
        # get the student2 name
        student2_name = sheet['C' + str(row)].value

        # get the marks of both in Q
        marks = sheet[input_marks_col_name + str(row)].value

        if student1_name != "-" and student1_name != "":
            # search for student1 in the registered list
            registered_list_row = search_in_registered_list(student1_name)

            if registered_list_row == 0:
                print(student1_name)
            else:
                # set the marks of student1 in the registered list
                sheet_registered_list[output_marks_col_name + str(
                    registered_list_row)].value = marks

        if student2_name != "-" and student2_name != "":
            # search for student2 in the registered list
            registered_list_row = search_in_registered_list(student2_name)

            if registered_list_row == 0:
                print(student2_name)
            else:
                # set the marks of student2 in the registered list
                sheet_registered_list[output_marks_col_name + str(
                    registered_list_row)].value = marks


# call main function
if __name__ == "__main__":
    transfer_to_registered_list("A1", 50, 'Q', 'G')
    transfer_to_registered_list("A2", 46, 'S', 'H')

    # save the workbook
    workbook.save('csf314-2023-Marks.xlsx')
