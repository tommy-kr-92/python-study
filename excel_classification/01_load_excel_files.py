import pandas as pd


class ClassificationExcel:
    def __init__(self, order_xlsx_file, partner_info_xlsx_filename):
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
        df = df.drop([df.index[0], df.index[1]])
        # print(df.count())   # 개수 확인

        df = df.reset_index(drop=True)
        self.order_list = df
        # print(df)

        # 파트너 목록
        df_partners = pd.read_excel(partner_info_xlsx_filename)

        self.brands = df_partners['브랜드'].to_list()
        self.partners = df_partners['업체명'].to_list()

        print(len(self.brands), self.brands)
        print(len(self.partners), self.partners)

        print(self.brands[0], self.partners[0])

    def classify(self):
        for i, row in self.order_list.iterrows():  # iterrow()는 두개를 받아야 함
            print(i)

        print(self.order_list['상품명'].head())


if __name__ == '__main__':
    ce = ClassificationExcel('주문목록20221112.xlsx', '파트너목록.xlsx')
    ce.classify()
