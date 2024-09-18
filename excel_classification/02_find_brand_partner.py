import pandas as pd


class ClassificationExcel:
    def __init__(self, order_xlsx_file, partner_info_xlsx_filename):
        # 주문 목록
        df = pd.read_excel(order_xlsx_file)
        df = df.rename(columns=df.iloc[1])  # 첫번째 줄 열 제목으로 설정
        '''
        0                                      NaN        NaN  ...         NaN         NaN
        1                                     주문번호       상품번호  ...          주소     주문시요구사항
        '''
        '''
                                   주문번호    상품번호  ...      주소  주문시요구사항
        0                   NaN     NaN  ...     NaN      NaN
        1                  주문번호    상품번호  ...      주소  주문시요구사항
        '''
        # print(df['상품명'])
        df = df.drop([df.index[0], df.index[1]])    # 0번째 index 삭제 - NaN 삭제
        # print(df.count())   # 개수 확인

        df = df.reset_index(drop=True)
        self.order_list = df
        # print(df)

        # 파트너 목록
        df_partners = pd.read_excel(partner_info_xlsx_filename)

        self.brands = df_partners['브랜드'].to_list()  # 브랜드명 리스트화
        self.partners = df_partners['업체명'].to_list()    # 업체명 리스트화

    def classify(self):
        for i, row in self.order_list.head(5).iterrows():  # iterrow()는 두가지를 가져와야 함
            brand_name = ''
            idx_partners = 0
            for j in range(len(self.brands)):
                # print(self.brands[j])
                if self.brands[j] in row['상품명']:
                    print(f'{self.brands[j]} 이(가) {j}번째에 포함되어 있습니다.')
                    brand_name = self.brands[j]
                    idx_partners = j
                    break

            print(f'{row["상품명"]}은 {brand_name} 브랜드 입니다. {j}번째')
            print(f'업체명: {self.partners[idx_partners]}')
            print('------------------------------------')
            # print(row['상품명'])

        print(len(self.brands), self.brands)
        print(len(self.partners), self.partners)


if __name__ == '__main__':
    ce = ClassificationExcel('주문목록20221112.xlsx', '파트너목록.xlsx')
    ce.classify()
