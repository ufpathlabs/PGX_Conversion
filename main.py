import pandas as pd
import openpyxl


def cnv_data_frame():
    #cnv_file_name = 'C:\\Users\\nlark\\PycharmProjects\\Candlebox\\Test_Sample_Files\\CN GPGX-QS-23-27 Results.csv'
    cnv_file_name = '/Users/kja3/PycharmProjects/Candlebox/Test_Sample_Files/CN GPGX-QS-23-27 Results.csv'
    cnv_df = pd.read_csv(cnv_file_name, skiprows=58, header=0)
    print(cnv_df)
    return cnv_df


def snp_data_frame():
    #snp_file_name = 'C:\\Users\\nlark\\PycharmProjects\\Candlebox\\Test_Sample_Files\\'GPGX-QS-23-27_20230321_060018_Export.csv'
    snp_file_name = '/Users/kja3/PycharmProjects/Candlebox/Test_Sample_Files/GPGX-QS-23-27_20230321_060018_Export.csv'
    snp_df = pd.read_csv(snp_file_name, skiprows=16, header=0, encoding='unicode_escape')
    snp_df.dropna(subset=['ROX'], inplace=True)
    print(snp_df)
    return snp_df


def print_to_file(cnv_df, snp_df):
    #cn_file_name = 'C:\\Users\\nlark\\PycharmProjects\\Candlebox\\new_file.xlsx'
    cn_file_name = '/Users/kja3/PycharmProjects/Candlebox/new_file.xlsx'
    with pd.ExcelWriter(cn_file_name) as writer:
        snp_df.to_excel(writer, sheet_name='SNP', index=False)
        cnv_df.to_excel(writer, sheet_name='CNV', index=False)


def main():
    snp_df = snp_data_frame()
    cnv_df = cnv_data_frame()
    print_to_file(cnv_df, snp_df)


if __name__ == '__main__':
    main()
