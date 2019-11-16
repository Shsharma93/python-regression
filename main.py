import csv
import re
import textract
import xlsxwriter


# csv regular expression pattern
name_pattern_csv = r'[A-Z][a-z]\s+[A-Z][a-z]+\s+[A-Z][a-z]+'
address_pattern_csv = r'[A-Z0-9]+[a-z]*\s[A-Z][a-z]*\s[A-Z0-9]+[a-z]*\s{4}[A-Z][a-z]+\s[A-Z]{2,3}\s[0-9]{4}\s[A-Z]{2,3}'
tfn_pattern_csv = r'[0-9]{9}'
dob_pattern_csv = r'[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}'
email_pattern_csv = r'[0-9]*[a-z]+[0-9]*@[a-z]+\.[a-z]+'
mobile_pattern_csv = r'[0-9]{4} [0-9]{3} [0-9]{3}'


# excel regular expression pattern
address_pattern_excel = r'[A-Z0-9]+[a-z]*\s[A-Z][a-z]*\s[A-Z0-9]+[a-z]*\s+[A-Z][a-z]+\s[A-Z]{2,3}\s[0-9]{4}\.0\s*[A-Z]{2,3}'
dob_pattern_excel = r'\s[0-9]{5}\.0'


# regular expression for name pattern
name_pattern_pdf = r'\s{8}[A-Z][a-z]+\s[A-Z][a-z]+\s{2}'
name_pattern_word = r'\s{4}[A-Z][a-z]+\s[A-Z][a-z]+\s{2}'

# regular expression for address pattern
address_pattern = r'\s+[A-Z0-9]+[a-z]*\s+[A-Z][a-z]+\s+[A-Z0-9]+[a-z]*\s+[A-Z]+\s+[A-Z]{3}\s+\d{4}'

# regular expression for name pattern
tfn_pattern = r'\d{6}\s*\d{2}\s*\d'


# decode byte objects to produce string
csv_text = textract.process("./source-files/csv2.csv").decode("utf-8")
excel_text = textract.process(
    "./source-files/spreadsheet1.xlsx").decode("utf-8")
word_text = textract.process("./source-files/word1.docx").decode("utf-8")
pdf_text = textract.process("./source-files/pdf1.pdf").decode("utf-8")


# first row of output file
first_row = ['Full Name', 'Full Address', 'TFN Number',
             'Date of Birth', 'Email Address', 'Mobile Number']


# initial data string
name_data = []
address_data = []
tfn_data = []
dob_data = []
email_data = []
mobile_data = []


# convert 5 digit numbers to date format
def convExcelDate(inp):
    inp = float(inp)
    Yearconv = str(1900+int(inp/365.25))
    DaysRemconv = inp-((int(inp/365.25))*365.25)
    Month = 1
    for M in [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]:
        if DaysRemconv > M:
            DaysRemconv = DaysRemconv - M
            Month = Month + 1
    if int(Yearconv) % 4 == 0:
        DaysRemconv -= 1
    returnVal = str(int(DaysRemconv))+'/'+str(Month) + \
        '/' + str(int(float(Yearconv)))
    return returnVal


# match the regular expression pattern
def scan_data(pattern, file, data):
    text = re.compile(pattern)
    matches = text.finditer(file)    # find pattern from text

    for match in matches:
        match = re.sub('\s+', ' ', match[0])  # remove spaces from string
        # remove .0 from number data from excel
        match = match.replace('.0', '')
        if(data == name_data and file == csv_text or excel_text):
            match = re.sub('M(s|r)', '', match)  # remove Mr Ms from name
        if(data == dob_data and file == excel_text):
            # call function to conver 5 digit number to date of birth
            match = convExcelDate(match)
        data.append(match.strip())  # remove any spaces from string


# write output data to csv file
def write_to_csv(first_row):
    with open('csv_output.csv', mode='w') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(first_row)
        for i in range(len(name_data)):
            if len(mobile_data) > i:
                csv_writer.writerow([name_data[i], address_data[i],
                                     tfn_data[i], dob_data[i], email_data[i], mobile_data[i]])
            else:
                csv_writer.writerow([name_data[i], address_data[i],
                                     tfn_data[i]])


# read csv file as a tuple to write to spreadhseet
def read_csv_output():
    with open("csv_output.csv") as f:
        return tuple(csv.reader(f))


# write eperson data to spreadsheet
def write_to_spreadsheet(first_row, person_data):
    workbook = xlsxwriter.Workbook('person_data.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for item in (person_data):
        worksheet.write(row, col, item[0])
        worksheet.write(row, col + 1, item[1])
        worksheet.write(row, col + 2, item[2])
        if len(item) > 3:
            worksheet.write(row, col + 3, item[3])
            worksheet.write(row, col + 4, item[4])
            worksheet.write(row, col + 5, item[5])
        row += 1

    workbook.close()


scan_data(name_pattern_csv, csv_text, name_data)
scan_data(address_pattern_csv, csv_text, address_data)
scan_data(tfn_pattern_csv, csv_text, tfn_data)
scan_data(dob_pattern_csv, csv_text, dob_data)
scan_data(email_pattern_csv, csv_text, email_data)
scan_data(mobile_pattern_csv, csv_text, mobile_data)

scan_data(name_pattern_csv, excel_text, name_data)
scan_data(address_pattern_excel, excel_text, address_data)
scan_data(tfn_pattern_csv, excel_text, tfn_data)
scan_data(dob_pattern_excel, excel_text, dob_data)
scan_data(email_pattern_csv, excel_text, email_data)
scan_data(mobile_pattern_csv, excel_text, mobile_data)

scan_data(name_pattern_pdf, pdf_text, name_data)
scan_data(address_pattern, pdf_text, address_data)
scan_data(tfn_pattern, pdf_text, tfn_data)

scan_data(name_pattern_word, word_text, name_data)
scan_data(address_pattern, word_text, address_data)
scan_data(tfn_pattern, word_text, tfn_data)

write_to_csv(first_row)
person_data = read_csv_output()
write_to_spreadsheet(first_row, person_data)
