import pandas as pd
import numpy as np
import os

class Create_Message:
    def __init__(self, date):
        self.date = date
        
    # 데이터프레임 앞뒤 공백 제거
    def clean_space_df(self, df):
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        return df

    # 특정 날짜, 데이터프레임 반환
    def get_name_date_df(self, path, date):
        df = pd.read_excel(path, header=[1, 2])
        df = self.clean_space_df(df)
        hw_col = [col for col in df.columns if (col[0].startswith(str(date[0][0])+'월') and  str(date[0][1]) in col[0]) or (col[0].startswith(str(date[1][0])+'월') and  str(date[1][1]) in col[0])]
        # 이름 공백 버리기
        df.iloc[:,0] = df.iloc[:,0].fillna('0')
        # 이름 아닌 부분 자르기
        df_name_date = df[df.iloc[:,0].apply(lambda x: 2<=len(x)<=4)][[df.columns[0]]+hw_col]
        return df_name_date

    def find_hw(self, path, date):
        df_hw = self.get_name_date_df(path, date)
        df_hw.columns = df_hw.columns.droplevel()
        df_hw = df_hw.drop(columns=df_hw.columns[2:])
        # A, B, C, D 외에 존재하는 것들 다 D로 변경
        hw_score_list = ['A', 'B', 'C', 'D']
        df_hw.loc[~df_hw['숙제 제출 여부'].isin(hw_score_list),'숙제 제출 여부'] = 'D'
        return df_hw

    def find_daily(self, path,date):
        df_daily = self.get_name_date_df(path, date)
        df_daily.columns = df_daily.columns.droplevel()
        df_daily['daily_done'] = df_daily.iloc[:,1:].apply(self.get_daily_done, axis=1)
        df_daily['daily_not_done'] = df_daily.iloc[:,1:].apply(self.get_daily_not_done, axis=1)
        return df_daily

    def find_test(self, path, date):
        df_test = self.get_name_date_df(path, date)
        df_test.columns = df_test.columns.droplevel()
        
        # 평균 점수
        mean_score = df_test['점수'].mean()
        mean_score = int(np.round(mean_score))
        
        # 결측치 점수 => 미응시로 채우기
        df_test = df_test.fillna('미응시')
        return df_test, mean_score

    def get_daily_done(self, df_daily_row):
        done_list = df_daily_row[df_daily_row.notnull()].index
        return list(done_list)

    def get_daily_not_done(self, df_daily_row):
        done_list = df_daily_row[df_daily_row.isnull()].index
        return list(done_list)

    def merge_df(self, df_daily, df_hw, df_test, col_name='이름'):
        df_merge = pd.merge(df_daily, df_hw, on=col_name)
        df_merge = pd.merge(df_merge, df_test, on=col_name)
        return df_merge
    
    def merge_df2(self, df_list, col_name):
        for i in range(len(df_list)):
            df_list[i] = df_list[i].set_index(col_name)
        df_merge = pd.concat(df_list, axis=1)
        df_merge = df_merge.reset_index()
        return df_merge

    # 일일학습지 문자
    def mk_daily_str(self, student):
        # 일일학습지 존재하면:
        if 'daily_done' in student.index:
            # 일일학습지
            if len(student['daily_done']) == 0:
                daily_str = f"일일학습지는 {' / '.join(student['daily_not_done'])} 를 미제출하였습니다."
            elif len(student['daily_done']) <= 3:
                daily_str = f"일일학습지는 {' / '.join(student['daily_done'])} 를 제출하였고, {' / '.join(student['daily_not_done'])} 를 미제출하였습니다."
            else:
                daily_str = f"일일학습지는 {' / '.join(student['daily_done'])} 를 제출하였습니다."
            return daily_str+'\n'
        else:
            return ''

    # 테스트 점수 문자
    def mk_test_str(self, student, mean_score):
        if '점수' in student.index:        
            if student['점수'] == '미응시':
                test_str = '미응시'
            else:
                test_str = str(int(round(student['점수'])))+'점'
            result_str = f'주간테스트 점수는 {test_str} / 평균은 {mean_score}점 입니다.'
            return result_str+'\n'
        else:
            return ''
    
    def mk_hw_str(self, student):
        if '숙제 제출 여부' in student.index:
            hw_str = f'숙제 제출 상태는 {student["숙제 제출 여부"]} 입니다.'
            return hw_str+'\n'
        else:
            return ''
            
    # 문자 생성
    def make_text(self, student, date, mean_score, next_hw_str):
        
        hw_str = self.mk_hw_str(student)
        daily_str = self.mk_daily_str(student)
        test_str = self.mk_test_str(student, mean_score)
        
        text = f'''<대치에스학원 박래혁수학>

{student['이름']}학생의 {date[0][0]}월 {date[0][1]}일 ~ {date[1][0]}월 {date[1][1]}일 숙제 제출 및 테스트 결과입니다.

{test_str}{hw_str}{daily_str}
다음주 숙제는 {next_hw_str} 입니다.

감사합니다.

(A : 완벽 / B : 미흡 / C : 불량 / D : 미제출 )'''
        return text


    def save_text_files(self, folder_path, df_students):
        # 폴더 경로 생성
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        else:
            for file in os.scandir(folder_path):
                os.remove(file.path)

        # 데이터프레임 루프
        for name, text_content in zip(df_students['이름'],df_students['text']):
            # 파일 이름 생성
            filename = f"{folder_path}/{name}.txt"
            
            # 파일 생성 및 값 작성
            with open(filename, "w") as f:
                f.write(text_content)
