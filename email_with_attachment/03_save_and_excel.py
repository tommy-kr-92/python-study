import unicodedata
from os import listdir
import pandas as pd
pd.set_option('display.max_columns', None)
import openpyxl


def make_email_list(data_path, partners_filename, title, target_filename='이메일발송목록.xlsx'):
    filenames = listdir(data_path)

    df_partners = pd.read_excel(partners_filename)

    r = []
    for filename in filenames:
        # 확장자 제거
        filename = filename.replace('.xlsx', '')
        # 파일명과 패턴을 NFC 형식으로 정규화
        filename_normalized = unicodedata.normalize('NFC', filename)
        pattern_normalized = unicodedata.normalize('NFC', '[패스트몰] ')
        # 패턴 제거
        partner_name = filename_normalized.replace(pattern_normalized, '')
        print(f"파트너 이름: '{partner_name}'")

        found_row = df_partners[df_partners['업체명'].str.contains(partner_name)]
        email1 = str(found_row['이메일 1'].values[0])
        partner_manager_name = str(found_row['컨택담당자'].values[0])
        email_cc = str(found_row['참조이메일'].values[0])
        # print(found_row)
        print(email1, partner_manager_name, email_cc)

        if email_cc == 'nan':
            email_cc = ''
        info = {'담당자메일': email1, '참조': email_cc, '컨텍담당자': partner_manager_name, '제목': title, '첨부파일명': filename}
        r.append(info)

        email_list = pd.DataFrame(r)
        email_list.to_excel(target_filename, index=False)
    print(f'엑셀로 저장 완료 되었습니다. 파일명: {target_filename}')


if __name__ == '__main__':
    make_email_list('data/', '파트너목록.xlsx', '[패스트몰] 금일 20240918 발주 목록 입니다. 확인 부탁드립니다.', 'email_list.xlsx')