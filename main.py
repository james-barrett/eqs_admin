import logging
import helper
import os
import time
import pandas as pd
import re
import yaml
from yaml.loader import SafeLoader
import random


def main():

    # Set current working directory same as this python script
    current_dir = os.getcwd()

    # Set up logging
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%a, %d %b %Y %H:%M:%S',
                        filename=os.path.join(current_dir, 'renamer_process.log'),
                        filemode='a+')

    logging.info(f'Program Started {helper.get_timestamp()}.')
    print('Starting PDF processor.')

    time.sleep(1)

    # attempt to load the config file, quit app if unsuccessful
    config_path = os.path.join(current_dir, "config.yaml")
    if not os.path.exists(config_path):
        logging.error('Main config file missing, please create.')
        print('Main config file missing, please create.')
        time.sleep(1)
        quit()

    with open(config_path) as f:
        config = yaml.load(f, Loader=SafeLoader)
        logging.info('Config file loaded successfully.')
        print('Config file loaded successfully.')

    time.sleep(1)

    # check mode set in config for needed folder structure
    mode = config['operation_mode']

    # needed directory structure if for USER/MANAGER set in config
    if mode == 'USER':
        dirs = [
            'Certificates',
            'Certificates\\EH',
            'Certificates\\FWT',
            'Certificates\\RR',
            'Certificates\\KB',
            'Certificates\\EH\\FILE_ERROR',
            'Certificates\\FWT\\FILE_ERROR',
            'Certificates\\RR\\FILE_ERROR',
            'Certificates\\KB\\FILE_ERROR',
            'Certificates\\EH\\UPRN_ERROR',
            'Certificates\\FWT\\UPRN_ERROR',
            'Certificates\\RR\\UPRN_ERROR',
            'Certificates\\KB\\UPRN_ERROR',
            'Certificates\\EH\\UNSUPPORTED',
            'Certificates\\FWT\\UNSUPPORTED',
            'Certificates\\RR\\UNSUPPORTED',
            'Certificates\\KB\\UNSUPPORTED',
            'Certificates\\EH\\FAILED',
            'Certificates\\FWT\\FAILED',
            'Certificates\\RR\\FAILED',
            'Certificates\\KB\\FAILED',
            'Certificates\\EH\\PROCESSED',
            'Certificates\\FWT\\PROCESSED',
            'Certificates\\RR\\PROCESSED',
            'Certificates\\KB\\PROCESSED'
        ]
    elif mode == 'MANAGER':
        dirs = [
            'Certificates',
            'Certificates\\AUDIT',
            'Certificates\\AUDIT\\FILE_ERROR',
            'Certificates\\AUDIT\\UPRN_ERROR',
            'Certificates\\AUDIT\\PASSED',
            'Certificates\\AUDIT\\FAILED',
            'Certificates\\AUDIT\\UNSUPPORTED'
        ]
    else:
        logging.info('No Valid profile found, quitting.....')
        print('No Valid profile found, quitting.....')
        quit()

    time.sleep(1)

    # check if needed folder paths exist
    if config['move_successfully_processed_to_tgp'] == 'YES':
        if not os.path.exists(config['tgp_eh_file_path']):
            logging.error(f'Error with TGP file save path: {config['tgp_eh_file_path']}.')
            print(f'Error with TGP file save path: {config['tgp_eh_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid TGP file save path: {config['tgp_eh_file_path']}.')
            print(f'Found valid TGP file save path: {config['tgp_eh_file_path']}')

        if not os.path.exists(config['tgp_fwt_file_path']):
            logging.error(f'Error with TGP file save path: {config['tgp_fwt_file_path']}.')
            print(f'Error with TGP file save path: {config['tgp_fwt_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid TGP file save path: {config['tgp_fwt_file_path']}.')
            print(f'Found valid TGP file save path: {config['tgp_fwt_file_path']}')

        if not os.path.exists(config['tgp_rr_file_path']):
            logging.error(f'Error with TGP file save path: {config['tgp_rr_file_path']}.')
            print(f'Error with TGP file save path: {config['tgp_rr_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid TGP file save path: {config['tgp_rr_file_path']}.')
            print(f'Found valid TGP file save path: {config['tgp_rr_file_path']}')

        if not os.path.exists(config['tgp_unsat_file_path']):
            logging.error(f'Error with TGP file save path: {config['tgp_unsat_file_path']}.')
            print(f'Error with TGP file save path: {config['tgp_unsat_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid TGP file save path: {config['tgp_unsat_file_path']}.')
            print(f'Found valid TGP file save path: {config['tgp_unsat_file_path']}')

    if config['move_successfully_processed_to_gp'] == 'YES':
        if not os.path.exists(config['gp_unsat_file_path']):
            logging.error(f'Error with GP file save path: {config['gp_unsat_file_path']}.')
            print(f'Error with GP file save path: {config['gp_unsat_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid GP file save path: {config['gp_unsat_file_path']}.')
            print(f'Found valid GP file save path: {config['gp_unsat_file_path']}')

        if not os.path.exists(config['gp_cert_file_path']):
            logging.error(f'Error with GP file save path: {config['gp_cert_file_path']}.')
            print(f'Error with GP file save path: {config['gp_cert_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid GP file save path: {config['gp_cert_file_path']}.')
            print(f'Found valid GP file save path: {config['gp_cert_file_path']}')

        if not os.path.exists(config['gp_eqs_fwt_logs_file_path']):
            logging.error(f'Error with GP file save path: {config['gp_eqs_fwt_logs_file_path']}.')
            print(f'Error with GP file save path: {config['gp_eqs_fwt_logs_file_path']}')
            time.sleep(1)
            quit()
        else:
            logging.info(f'Found valid GP file save path: {config['gp_eqs_fwt_logs_file_path']}.')
            print(f'Found valid GP file save path: {config['gp_eqs_fwt_logs_file_path']}')

    if not len(config['eqs_name']) > 5:
        logging.info(f'EQS name issue, must be greater than 5 characters {config['eqs_name']}, quitting.')
        print(f'EQS name issue, must be greater than 5 characters {config['eqs_name']}, quitting.')
        time.sleep(1)
        quit()

    dirs_created = False

    # check for directory structure, create if missing, if created quit
    for d in dirs:
        if not os.path.exists(os.path.join(current_dir, d)):
            print(f'Missing directory found, creating {os.path.join(current_dir, d)}.')
            os.mkdir(os.path.join(current_dir, d))
            dirs_created = True

    # quit if directories were missing to allow files to be added
    if dirs_created:
        logging.info('Directories created, quitting to allow files to be added.')
        print('Directories created, quitting to allow files to be added.')
        quit()

    # Load address excel file
    print('Attempting to load very large address database file, please wait....')
    try:
        property_list_path = os.path.join(current_dir, 'PROPERTIES.xlsx')
        # noinspection PyTypeChecker
        property_list_df = pd.read_excel(property_list_path,
                                         usecols={'Property Reference', 'Property Address'})

        property_list_df = property_list_df.astype(str)

        logging.info('Successfully loaded address database')
        print('Successfully loaded address database')
    except Exception as e:
        logging.info(f'Error loading database: {e}')
        logging.info('An error occurred loading the property database, exiting.....')
        print(f'Error loading database: {e}')
        print('An error occurred loading the property database, exiting.....')
        time.sleep(1)
        quit()

    # set working directory to be root of folder structure
    working_dir = os.path.join(current_dir, dirs[0])
    
    sub_folders = []

    # set folder structure dependent on user
    if mode == 'USER':
        sub_folders = ['EH', 'FWT', 'RR', 'KB']
    elif mode == 'MANAGER':
        sub_folders = ['AUDIT']

    # empty array to store processed certificate data for Excel sheet creation
    excel_data = []
    accu_data = []

    # iterate oer sub directories
    for sub in sub_folders:
        active_directory = os.path.join(working_dir, sub)
        logging.info(f'Processing files in {active_directory} folder.')
        print(f'Processing files in {active_directory} folder.')

        # get files in directory
        files = []
        for entry in os.scandir(active_directory):
            if entry.is_dir():
                continue
            # use entry.path to get the full path of this entry, or use
            # entry.name for the base filename
            files.append(entry.path)

        logging.info(f'Files found: {len(files)}')
        print(f'Files found: {len(files)}')

        # get timestamp once, used for writing accuserv files
        timestamp = helper.get_timestamp()

        # Iterate through files
        for file in files:
            print(f'{file}')

            # Check if we have a pdf
            if not helper.is_pdf(file):
                print(f'File error, moving {file} to the unsupported directory.')
                try:
                    os.rename(file, active_directory + '\\UNSUPPORTED\\' + os.path.basename(file))
                except WindowsError as e:
                    logging.error(f'File error, moving {file} to the unsupported directory, retrying')
                    print(f'File error, moving {file} to the unsupported directory, retrying')
                    try:
                        amended_file_name = (os.path.splitext(os.path.basename(file))[0] + '_' +
                                             str(random.randrange(1, 1000)) + '.pdf')
                        os.rename(file, active_directory + '\\UNSUPPORTED\\' + amended_file_name)
                    except WindowsError as e:
                        logging.error(f'File error, moving {file} to the unsupported directory, skipping')
                        print(f'File error, moving {file} to the unsupported directory, skipping')
                        continue
                continue
            else:
                file_type = helper.get_file_type(os.path.join(active_directory, file))

                # if we suffer an error trying to get file type
                if file_type == 'ERROR':
                    logging.error(f'File error, moving {file} to the error directory.')
                    print(f'File error, moving {file} to the error directory.')
                    try:
                        os.rename(file, active_directory + '\\FILE_ERROR\\' + os.path.basename(file))
                    except WindowsError as e:
                        logging.error(f'File error, moving {file} to the error directory, retrying')
                        print(f'File error, moving {file} to the error directory, retrying')
                        try:
                            amended_file_name = (os.path.splitext(os.path.basename(file))[0] + '_' +
                                                 str(random.randrange(1, 1000)) + '.pdf')
                            os.rename(file, active_directory + '\\FILE_ERROR\\' + amended_file_name)
                        except WindowsError as e:
                            logging.error(f'File error, moving {file} to the error directory, skipping')
                            print(f'File error, moving {file} to the error directory, skipping')
                            continue
                    continue

                # Check we have a known certificate type
                if file_type == "UNSUPPORTED":
                    logging.info(f'Found unsupported file type, moving {file} to the unsupported directory.')
                    print(f'Found unsupported file type, moving {file} to the unsupported directory.')
                    try:
                        os.rename(file, active_directory + '\\UNSUPPORTED\\' + os.path.basename(file))
                    except WindowsError as e:
                        try:
                            logging.error(f'File error, moving {file} to the unsupported directory, retrying')
                            print(f'File error, moving {file} to the error directory, retrying')
                            amended_file_name = (os.path.splitext(os.path.basename(file))[0] + '_' +
                                                 str(random.randrange(1, 1000)) + '.pdf')
                            os.rename(file, active_directory + '\\UNSUPPORTED\\' + amended_file_name)
                        except WindowsError as e:
                            logging.error(f'File error, moving {file} to the unsupported directory, skipping')
                            print(f'File error, moving {file} to the unsupported directory, skipping')
                            continue
                    continue

                logging.info(f'Found {file_type} start processing {file}.')
                print(f'Found {file_type} start processing {file}.')

                # get needed data from pdf file
                pdf_data = helper.get_pdf_data(os.path.join(active_directory, file), file_type)

                try:
                    property_list_df = property_list_df.sort_index()
                    df_rows = property_list_df.loc[property_list_df['Property Reference'] == pdf_data[0].upper()]
                    df_row = df_rows.iloc[0]
                    o_address_full = df_row['Property Address']
                    address_split = o_address_full.split(',')
                    o_postcode = helper.clean_text(address_split[-1])
                    o_address = helper.clean_text(o_address_full.replace(o_postcode, ''))
                    o_address_numbers = re.findall(r'\d+', o_address)
                except Exception as e:
                    logging.error(f'UPRN Error: {e}')
                    print(f'Moving {file} to the UPRN_ERROR directory.')
                    try:
                        os.rename(file, active_directory + '\\UPRN_ERROR\\' + os.path.basename(file))
                    except WindowsError as e:
                        try:
                            logging.error(f'Error moving {file} to the UPRN_ERROR directory, retrying')
                            print(f'Error moving {file} to the UPRN_ERROR directory, retrying')
                            amended_file_name = (os.path.splitext(os.path.basename(file))[0] + '_' +
                                                 str(random.randrange(1, 1000)) + '.pdf')
                            os.rename(file, active_directory + '\\UPRN_ERROR\\' + amended_file_name)
                        except WindowsError as e:
                            logging.error(f'Error moving {file} to the UPRN_ERROR directory, skipping')
                            print(f'Error moving {file} to the UPRN_ERROR directory, skipping')
                            continue
                    continue

                c_address = pdf_data[2]
                c_postcode = pdf_data[6]
                c_address_numbers = re.findall(r'\d+', c_address)

                number_check = helper.address_match_check(o_address_numbers, c_address_numbers)
                address_check = helper.address_match_check(o_address, c_address)

                number_o = o_address_numbers
                number_c = c_address_numbers

                postcode_o = helper.clean_text(o_postcode).replace(' ', '')
                postcode_c = helper.clean_text(c_postcode).replace(' ', '')

                number_score = 0
                postcode_score = 0

                if postcode_o == postcode_c:
                    postcode_score = 1

                if number_o == number_c:
                    number_score = 1

                file_score = min(float(number_check[2]), float(address_check[2]), float(postcode_score))

                # 'UPRN' 'JOB NO' 'CERT TYPE' 'TEST DATE' 'RESULT' 'ADDRESS - CERT' 'ADDRESS - TGP' 'ADDRESS SCORE'}
                test_date = pdf_data[1][:2] + '/' + pdf_data[1][2:4] + '/' + pdf_data[1][4:6]

                # [uprn, date, address, cert_num, job_no, status, postcode]
                # data structure to complete Excel sheet
                rename_detail = [
                    pdf_data[0],
                    pdf_data[4],
                    file_type,
                    test_date,
                    pdf_data[5],
                    pdf_data[2] + ' ' + pdf_data[6],
                    o_address_full,
                    file_score,
                    pdf_data[7],
                    pdf_data[8]
                ]

                excel_data.append(rename_detail)

                minimum_score = config['address_minimum_score']
                if file_score < float(minimum_score):
                    print(f"Address matching issues, "
                          f"{file_score} is less than minimum {float(minimum_score) * 100}%, skipping file.")
                    print(f"Please manually check {file} for address or uprn errors.")
                    logging.info(f'Moving {file} to the FAILED directory.')
                    print(f'Moving {file} to the FAILED directory.')
                    try:
                        os.rename(file, active_directory + '\\FAILED\\' + os.path.basename(file))
                    except WindowsError as e:
                        try:
                            logging.error(f'Error moving {file} to the FAILED directory, retrying')
                            print(f'Error moving {file} to the FAILED directory, retrying')
                            amended_file_name = (os.path.splitext(os.path.basename(file))[0] + '_' +
                                                 str(random.randrange(1, 1000)) + '.pdf')
                            os.rename(file, active_directory + '\\FAILED\\' + amended_file_name)
                        except WindowsError as e:
                            logging.error(f'Error moving {file} to the FAILED directory, skipping')
                            print(f'Error moving {file} to the FAILED directory, skipping')

                    if mode == 'MANAGER':
                        if config['address_low_score_email_eqs'] == 'YES':
                            print('Sending email to EQS in relation to low score address match')
                            helper.low_score_email(rename_detail)

                    continue

                if mode == 'USER':
                    # Rename the pdf
                    file_new = helper.rename_pdf_file(pdf_data[0], pdf_data[1], file_type, pdf_data[5])
                    # End of renaming function config specific actions follow.
                    # Get post renaming actions from config
                    if sub == 'FWT' and file_type == 'EICR':
                        # helper.create_accuserv_list(working_dir, pdf_data, timestamp)

                        acc_test_date = pdf_data[1][:2] + '/' + pdf_data[1][2:4] + '/' + pdf_data[1][4:6]

                        accu_detail = [
                            acc_test_date,
                            pdf_data[0],
                            pdf_data[2],
                            pdf_data[3],
                            pdf_data[4],
                            pdf_data[5]
                        ]

                        accu_data.append(accu_detail)

                    logging.info(F'Moving {file_new} to processed folder')
                    print(F'Moving {file_new} to processed folder')

                    file_path = None

                    try:
                        os.rename(file, active_directory + '\\PROCESSED\\' + file_new)
                        file_path = os.path.join(active_directory + '\\PROCESSED\\', file_new)
                    except WindowsError as e:
                        try:
                            logging.error(f'Error moving {file} to the PROCESSED directory, retrying')
                            print(f'Error moving {file} to the PROCESSED directory, retrying')
                            amended_file_name = (os.path.splitext(os.path.basename(file_new))[0] + '_' +
                                                 str(random.randrange(1, 1000)) + '.pdf')
                            os.rename(file, active_directory + '\\PROCESSED\\' + amended_file_name)
                            file_path = os.path.join(active_directory + '\\PROCESSED\\', amended_file_name)
                        except WindowsError as e:
                            logging.error(f'Error moving {file} to the PROCESSED directory, skipping')
                            print(f'Error moving {file} to the PROCESSED directory, skipping')

                    if config['move_successfully_processed_to_tgp'] == 'YES':
                        if file_path:
                            if 'UNSAT' in os.path.basename(file_path):
                                helper.copy_file(file_path, config['tgp_unsat_file_path'])
                            elif sub == 'EH':
                                helper.copy_file(file_path, config['tgp_eh_file_path'])
                            elif sub == 'FWT':
                                helper.copy_file(file_path, config['tgp_fwt_file_path'])
                            elif sub == 'RR':
                                helper.copy_file(file_path, config['tgp_rr_file_path'])

                    if config['move_successfully_processed_to_gp'] == 'YES':
                        if file_path:
                            if 'UNSAT' in os.path.basename(file_path):
                                helper.copy_file(file_path, config['gp_unsat_file_path'])
                            else:
                                helper.copy_file(file_path, config['gp_cert_file_path'])

                    if config['remove_from_processed_on_completion'] == 'YES':
                        os.remove(file_path)
                else:
                    logging.info(f'Moving {file} to PASSED folder')
                    print(f'Moving {file} to PASSED folder')
                    try:
                        os.rename(file, active_directory + f'\\PASSED\\' + os.path.basename(file))
                    except WindowsError as e:
                        try:
                            logging.error(f'Error moving {file} to the PASSED directory, retrying')
                            print(f'Error moving {file} to the PASSED directory, retrying')
                            amended_file_name = (os.path.splitext(os.path.basename(file))[0] + '_' +
                                                 str(random.randrange(1, 1000)) + '.pdf')
                            os.rename(file, active_directory + f'\\PASSED\\' + amended_file_name)
                        except WindowsError as e:
                            logging.error(f'Error moving {file} to the PASSED directory, skipping')
                            print(f'Error moving {file} to the PASSED directory, skipping')

    else:
        print('All done.')
        if excel_data:
            print('writing excel spreadsheet with file naming scores')
            helper.create_excel_sheet(working_dir, excel_data, str(helper.get_timestamp()))
        if accu_data:
            print('writing accuserv list to text document for FWT processing')
            accu_sheet_name = f'{config['eqs_name']}_{str(helper.get_timestamp())}'
            helper.create_excel_accu_sheet(working_dir, accu_data, accu_sheet_name)
            helper.copy_file(os.path.join(working_dir, accu_sheet_name + '.xlsx'),
                                          os.path.join(config['gp_eqs_fwt_logs_file_path'], accu_sheet_name + '.xlsx'))


if __name__ == "__main__":
    main()
