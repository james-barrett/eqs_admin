import os
import time

import fitz
import logging
import xlsxwriter
import random
from win32com.client import Dispatch
import Levenshtein
import re
import shutil


def is_pdf(file):
    split_file_name = os.path.splitext(file)
    file_extension = split_file_name[1]
    file_extension = file_extension.lower()
    if file_extension == ".pdf":
        return True
    else:
        return False

# helper.print_rect_coords(False, active_directory, file)

def get_timestamp():
    t = time.localtime()
    timestamp = time.strftime('%b-%d-%Y_%H%M', t)
    return timestamp


def get_pdf_data(file, file_type):

    if file_type == "EIC":
        job_no_rect = (400.0, 135.17999267578125, 460.8079833984375, 146.1719970703125)
        uprn_rect = (674.0, 148.17999267578125, 740.68798828125, 159.1719970703125)
        date_rect = (110.0, 232.17999267578125, 240.031982421875, 243.1719970703125)
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (578.0, 164.0, 800.0, 173.1719970703125)
        address_line_2_rect = (553.0, 176.5, 800.0, 186.1719970703125)
        postcode_rect = (582.0, 189.17999267578125, 650, 200.0)
        engineer_rect = (97.0, 479.17999267578125, 300.0, 490.1719970703125)
        supervisor_rect = (97.0, 543.1800537109375, 300.0, 554.1720581054688)
    elif file_type == "EICR":
        status_rect = (598.0, 374.2200012207031, 656.6959838867188, 385.2359924316406)
        job_no_rect = (400.0, 135.17999267578125, 460.8079833984375, 146.1719970703125)
        uprn_rect = (582.0, 150.17999267578125, 650.68798828125, 161.1719970703125)
        date_rect = (190.0, 276.17999267578125, 277.17596435546875, 287.1719970703125)
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (582.0, 162.17999267578125, 800.0, 174.1719970703125)
        address_line_2_rect = (552.0, 177.17999267578125, 800.0, 188.1719970703125)
        postcode_rect = (588.0, 189.17999267578125, 650.0, 200.1719970703125)
        engineer_rect = (218.0, 464.17999267578125, 400.0, 475.1719970703125)
        supervisor_rect = (218.0, 540.1800537109375, 400.0, 551.1720581054688)
    elif file_type == "MW":
        address_line_1_rect = (578.0, 168.17999267578125, 790.0, 178.0)
        address_line_2_rect = (555.0, 182.0, 790.0, 191.0)
        job_no_rect = (400.0, 141.17999267578125, 460.0, 152.1719970703125)
        uprn_rect = (570.0, 155.17999267578125, 640.68798828125, 165.1719970703125)
        date_rect = (90.0, 251.17999267578125, 200.031982421875, 262.1719970703125)
        cert_num_rect = (610.0, 41.220001220703125, 680.0, 52.23600387573242)
        postcode_rect = (583.0, 195.17999267578125, 650.0, 206.1719970703125)
        engineer_rect = (472.0, 449.17999267578125, 600.0, 460.1719970703125)
        supervisor_rect = (468.0, 509.1799621582031, 600.0, 520.1719970703125)
    elif file_type == "VIS":
        job_no_rect = (398.0, 136.17999267578125, 460.8079833984375, 147.1719970703125)
        uprn_rect = (572.0, 152.0, 640.68798828125, 161.0)
        date_rect = (696.0, 429.17999267578125, 760.031982421875, 440.1719970703125)
        cert_num_rect = (611.0, 39.220001220703125, 660.583984375, 50.23600387573242)
        address_line_1_rect = (578.0, 163.17999267578125, 800.0, 174.1719970703125)
        address_line_2_rect = (552.0, 177.17999267578125, 800.0, 188.1719970703125)
        postcode_rect = (580.0, 191.0, 630.0, 200.1719970703125)
        engineer_rect = (256.0, 429.17999267578125, 600.0, 440.1719970703125)
        supervisor_rect = (256.0, 451.17999267578125, 500, 462.1719970703125)
    else:
        return

    with (fitz.open(file) as doc):

        pdf_page = doc[0]

        if not pdf_page.is_wrapped:
            pdf_page.clean_contents()

        job_no = clean_text(doc[0].get_textbox(job_no_rect))
        uprn = clean_text(doc[0].get_textbox(uprn_rect))
        date = format_date(clean_text(doc[0].get_textbox(date_rect)))
        cert_num = clean_text(doc[0].get_textbox(cert_num_rect))
        engineer = clean_text(doc[0].get_textbox(engineer_rect))
        supervisor = clean_text(doc[0].get_textbox(supervisor_rect))

        if file_type == "EICR":
            status_rect_data = doc[0].get_textbox(status_rect)
            status = clean_text(status_rect_data)
            if 'XXXXXXXXXXX' in status:
                status = 'UNSATISFACTORY'
            else:
                status = "SATISFACTORY"
        else:
            status = "N/A"

        address = \
            clean_text(doc[0].get_textbox(address_line_1_rect)) + " " + clean_text(
                doc[0].get_textbox(address_line_2_rect))

        post_code = clean_text(doc[0].get_textbox(postcode_rect))

    return [uprn, date, address, cert_num, job_no, status, post_code, engineer, supervisor]


def rename_pdf_file(uprn, date, type, status):
    naming_convention = ""
    clean_uprn = clean_text(uprn)

    if type == "EICR":
        if any(c.isalpha() for c in uprn):
            naming_convention = "C"

            if 'dw' in clean_uprn.lower():
                naming_convention = "D"

        else:
            naming_convention = "D"
    else:
        naming_convention = ''

    if len(clean_uprn) < 1:
        clean_uprn = "MISSING"

    if status == "UNSATISFACTORY":
        new_file = clean_uprn + "_" + naming_convention + type + "_" + date + "_UNSAT.pdf"
    else:
        new_file = clean_uprn + "_" + naming_convention + type + "_" + date + ".pdf"

    return new_file


def move_processed_file(working_dir, file, sub, cert_num):
    processed_dir = os.path.join(working_dir, '_PROCESSED\\' + sub)

    file_name = os.path.basename(file)

    try:
        logging.info(f'Moving {file_name} to _PROCESSED {sub} directory.')
        print(f'Moving {file_name} to _PROCESSED {sub} directory.')
        os.rename(file, os.path.join(processed_dir, file_name))
    except WindowsError as e:
        logging.debug(e)
        try:
            logging.info(f'FILE MOVE ERROR: Trying to append certificate number to {file_name}, {sub}')
            print(f'FILE MOVE ERROR: Trying to append certificate number to {file_name}, {sub}')

            if 'UNSAT' in file:
                amended_file_name = os.path.splitext(file_name)[0].replace('_UNSAT', '_' + cert_num + '_UNSAT.pdf')
            else:
                amended_file_name = os.path.splitext(file_name)[0] + '_' + cert_num + '.pdf'
            try:
                os.rename(file, os.path.join(processed_dir, amended_file_name))
            except WindowsError as e:
                logging.info(f'FILE MOVE ERROR: Failed to move file on second attempt, try to append random number.')
                print(f'FILE MOVE ERROR: Failed to move file on second attempt, try to append random number.')
                logging.debug(e)
                try:
                    amended_file_name = (os.path.splitext(file_name)[0] + '_' +
                                         cert_num + str(random.randrange(1,1000)) + '.pdf')
                    os.rename(file, os.path.join(processed_dir, amended_file_name))
                except Exception as e:
                    pass
        except WindowsError as e:
            logging.info(f'FILE MOVE ERROR: Failed to move file on second attempt.')
            print(f'FILE MOVE ERROR: Failed to move file on second attempt.')
            logging.debug(e)


def clean_text(item):
    special_characters = ['!', '#', '$', '%', '&', '@', '[', ']', ']', ',', '(', ')']

    for i in special_characters:
        item = item.replace(i, '')
        item = item.strip()

    return item


def create_accuserv_list(working_dir, data, timestamp):
    # Create file
    # data = [uprn, date, address, cert_num, job_no]
    file = open(os.path.join(working_dir, 'accuserv' + '_' + timestamp + '.txt'), 'a+')
    file.write(data[1] + " : " +
               data[0] + " : " +
               data[2] + " : " +
               data[3] + " : " +
               data[4] + " : " +
               data[5] + '\r\n\n')
    file.close()


def format_date(d):
    d = d.replace('/', '')

    if '-' in d:
        # date_1 = d[:8]
        date = d[-8:]
        date = date[:4] + date[-2:]
    else:
        if len(d) == 8:
            date = d[:4] + d[-2:]
        else:
            date = 'MISSING'

    return date


def get_file_type(file: str) -> str:

    try:
        with fitz.open(file) as doc:
            page = doc[0]

            if not page.is_wrapped:
                page.wrap_contents()

            page_0 = str(page.get_text()).upper()

            if "FIRE DETECTION" in page_0:
                return "UNSUPPORTED"
            elif "CERTIFICATE OF COMPLIANCE" in page_0:
                return "UNSUPPORTED"
            if "ELECTRICAL INSTALLATION CONDITION" in page_0:
                return "EICR"
            elif "ELECTRICAL INSTALLATION CERTIFICATE" in page_0:
                return "EIC"
            elif "MINOR ELECTRICAL INSTALLATION" in page_0:
                return "MW"
            elif "DOMESTIC VISUAL CONDITION" in page_0:
                return "VIS"
            else:
                return "UNSUPPORTED"

    except Exception as e:
        logging.error(e)
        return 'ERROR'


def create_excel_accu_sheet(excel_dir, data, file_path):
    table_length = str(len(data) + 1)

    workbook = xlsxwriter.Workbook(os.path.join(excel_dir, f'{file_path}.xlsx'))
    worksheet = workbook.add_worksheet()
    format_test = workbook.add_format({'text_wrap': True})
    format_num = workbook.add_format({'num_format': 'dd/mm/yy'})
    worksheet.set_column('A:A', None, format_num)
    worksheet.set_column('A:F', None, format_test)
    worksheet.set_column('A:B', 20)
    worksheet.set_column('C:C', 75)
    worksheet.set_column('D:F', 20)

    worksheet.add_table('A1:F' + table_length, {'data': data,
                                                'columns': [{'header': 'TEST DATE'},
                                                            {'header': 'UPRN'},
                                                            {'header': 'ADDRESS'},
                                                            {'header': 'CERT NUMBER'},
                                                            {'header': 'JOB NO'},
                                                            {'header': 'RESULT'}
                                                            ]})

    workbook.close()


def create_excel_sheet(excel_dir, data, timestamp):
    table_length = str(len(data) + 1)

    workbook = xlsxwriter.Workbook(os.path.join(excel_dir, f'address_score_{timestamp}.xlsx'))
    worksheet = workbook.add_worksheet()
    format_test = workbook.add_format({'text_wrap': True})
    format_num = workbook.add_format({'num_format': 'dd/mm/yy'})
    worksheet.set_column('C:C', None, format_num)
    worksheet.set_column('A:H', None, format_test)
    worksheet.set_column('A:E', 15)
    worksheet.set_column('F:G', 75)
    worksheet.set_column('H:H', 10)
    worksheet.set_column('I:J', 20)

    worksheet.add_table('A1:J' + table_length, {'data': data,
                                                'columns': [{'header': 'UPRN'},
                                                            {'header': 'JOB NO'},
                                                            {'header': 'CERT TYPE'},
                                                            {'header': 'TEST DATE'},
                                                            {'header': 'RESULT'},
                                                            {'header': 'ADDRESS - CERT'},
                                                            {'header': 'ADDRESS - TGP'},
                                                            {'header': 'ADDRESS SCORE'},
                                                            {'header': 'ENGINEER'},
                                                            {'header': 'SUPERVISOR'}
                                                            ]})
    workbook.close()


def address_match_check(value_1, value_2):
    score = Levenshtein.ratio(value_1, value_2)
    data = [value_1, value_2, score]
    return data


def pdf_text_finder(file):
    doc = fitz.open(file)
    for page in doc:
        wlist = page.get_text_words()
        return wlist


def print_rect_coords(enable: bool, directory: str, file: str) -> None:
    if enable:
        print(pdf_text_finder(os.path.join(directory, file)))


def low_score_email(data):
    supervisor_email = data[9].split(' ')
    supervisor_email = supervisor_email[0] + '.' + supervisor_email[1]
    supervisor_email = supervisor_email + '@guinness.org.uk'

    subject = 'Please review the included information.'
    content = (
        f'Property UPRN: {data[0]} \n'
        f'Job Number: {data[1]} \n'
        f'Certificate Type: {data[2]} \n'
        f'Test Date: {data[3]} \n'
        f'Certificate Number: {data[4]} \n'
        f'Property Address - Certificate: {data[5]} \n'
        f'Property Address - TGP Database: {data[6]} \n'
        f'Match Score: {data[7]} \n'
        f'Approving EQS: {data[9]}'
    )

    outlook = Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = supervisor_email
    message.Subject = subject
    message.Body = content
    message.Send()

    time.sleep(2)


def email_pdf(file, subject, file_data, receivers):
    outlook = Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = "".join(receivers)
    message.Subject = subject
    message.Attachments.Add(Source=file)
    message.Body = (f'Please Find Attached Your Certificate \n'
                    f'{file_data}')
    message.Send()

    time.sleep(2)


def copy_file(source, destination):
    if os.path.isfile(source):
        try:
            shutil.copy(source, destination)
        except WindowsError as e:
            print(e)

def delete_file():
    pass


def rename_file(current_name, new_name):
    pass
