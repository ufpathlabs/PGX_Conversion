import csv
import pandas as pd
import openpyxl

def find_start_line(file_to_run, string_to_find):
    num = 0
    with open(file_to_run) as f_obj:
        reader = csv.reader(f_obj, delimiter=',')
        for line in reader:  # Iterates through the rows of your csv
            if string_to_find in str(line):  # If the string you want to search is in the row
                break
               # print(line)  # line here refers to a row in the csv
               # return line
            else:
                num = num+1
    return num


def cnv_data_frame():
    string_to_find = 'Sample Name'
    cnv_file_name = 'C:\\Users\\nlark\\PycharmProjects\\Candlebox\\Test_Sample_Files\\CN GPGX-QS-23-27 Results.csv'
    # cnv_file_name = '/Users/kja3/PycharmProjects/Candlebox/Test_Sample_Files/CN GPGX-QS-23-27 Results.csv'
    num = find_start_line(cnv_file_name, string_to_find)
    cnv_df = pd.read_csv(cnv_file_name, skiprows=num, header=0)
    print(cnv_df)
    return cnv_df


def snp_data_frame():
    string_to_find = 'Assay Name'
    snp_file_name = 'C:\\Users\\nlark\\PycharmProjects\\Candlebox\\Test_Sample_Files\\GPGX-QS-23-27_20230321_060018_Export.csv'
    # snp_file_name = '/Users/kja3/PycharmProjects/Candlebox/Test_Sample_Files/GPGX-QS-23-27_20230321_060018_Export.csv'
    num = find_start_line(snp_file_name, string_to_find)
    snp_df = pd.read_csv(snp_file_name, skiprows=num, header=0, encoding='unicode_escape')
    snp_df.dropna(subset=['ROX'], inplace=True)
    print(snp_df)
    return snp_df


def print_to_file(cnv_df, snp_df):
    cn_file_name = 'C:\\Users\\nlark\\PycharmProjects\\PGX_Conversion\\new_file.xlsx'
    # cn_file_name = '/Users/kja3/PycharmProjects/Candlebox/new_file.xlsx'
    with pd.ExcelWriter(cn_file_name) as writer:
        snp_df.to_excel(writer, sheet_name='SNP', index=False)
        cnv_df.to_excel(writer, sheet_name='CNV', index=False)


def main():
    snp_df = snp_data_frame()
    cnv_df = cnv_data_frame()
    print_to_file(cnv_df, snp_df)


if __name__ == '__main__':
    main()
