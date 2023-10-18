# open csf314-2023-Marks.xlsx

import openpyxl

# open the A1 sheet
workbook = openpyxl.load_workbook('csf314-2023-Marks.xlsx')

# get the sheet A1
sheet_a1 = workbook['A1']

# get the sheet A2
sheet_a2 = workbook['A2']

# get the sheet "Registered List"
sheet_registered_list = workbook['Registered List']


def search_in_registered_list(name):
    # find the student1 in the registered list name = Column B(first name) + Column C(last name)
    for registered_list_row in range(2, 99):
        registered_list_name = sheet_registered_list['A' + str(
            registered_list_row)].value

        if sheet_registered_list['B' + str(registered_list_row)].value != ".":
            registered_list_name += " " + sheet_registered_list['B' + str(
                registered_list_row)].value

        # if spaces is there after the name remove it
        registered_list_name = registered_list_name.strip()

        # print(student1_name, registered_list_name)
        if registered_list_name == name:
            return registered_list_row

    return 0


def transfer_a1_to_registered_list():
    # iterate over rows 2 to 48 of Column B C Q
    for row in range(2, 50):
        # get the student1 name
        student1_name = sheet_a1['B' + str(row)].value
        # get the student2 name
        student2_name = sheet_a1['C' + str(row)].value

        # if spaces is there after the name remove it
        student1_name = student1_name.strip()
        student2_name = student2_name.strip()

        # get the marks of both in Q
        marks = sheet_a1['Q' + str(row)].value

        if student1_name != "-":
            # search for student1 in the registered list
            registered_list_row = search_in_registered_list(student1_name)

            if registered_list_row == 0:
                print(student1_name)
            else:
                # set the marks of student1 in the registered list
                sheet_registered_list['G' + str(
                    registered_list_row)].value = marks

        if student2_name != "-":
            # search for student2 in the registered list
            registered_list_row = search_in_registered_list(student2_name)

            if registered_list_row == 0:
                print(student2_name)
            else:
                # set the marks of student2 in the registered list
                sheet_registered_list['G' + str(
                    registered_list_row)].value = marks

    # save the workbook
    workbook.save('csf314-2023-Marks.xlsx')


def transfer_a2_to_registered_list():
    # iterate over rows 2 to 48 of Column B C Q
    for row in range(2, 46):
        # get the student1 name
        student1_name = sheet_a2['B' + str(row)].value
        # get the student2 name
        student2_name = sheet_a2['C' + str(row)].value

        # if spaces is there after the name remove it
        student1_name = student1_name.strip()
        student2_name = student2_name.strip()

        # get the marks of both in S
        marks = sheet_a2['S' + str(row)].value

        if student1_name != "-":
            # search for student1 in the registered list
            registered_list_row = search_in_registered_list(student1_name)

            if registered_list_row == 0:
                print(student1_name)
            else:
                # set the marks of student1 in the registered list
                sheet_registered_list['H' + str(
                    registered_list_row)].value = marks

        if student2_name != "-":
            # search for student2 in the registered list
            registered_list_row = search_in_registered_list(student2_name)

            if registered_list_row == 0:
                print(student2_name)
            else:
                # set the marks of student2 in the registered list
                sheet_registered_list['H' + str(
                    registered_list_row)].value = marks

    # save the workbook
    workbook.save('csf314-2023-Marks.xlsx')


# call main function
if __name__ == "__main__":
    # transfer_a1_to_registered_list()
    transfer_a2_to_registered_list()
