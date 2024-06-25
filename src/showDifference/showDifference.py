import pandas as pd
import openpyxl

TARGET_EXCEL_PATH = 'target.xlsx'
COMPARE_SHEET_NAME1 = '공종별 내역서(변경전)'
COMPARE_SHEET_NAME2 = '공종별내역서(변경후)'
RESULT_SHEET_NAME = 'Sheet3' # 시트명
LARGE_CATEGORY_INDEX = 0 # 대분류가 위치하는 인덱스
TARGET_ID_LIST = [0, 1, 2] # 구분을 위한 인덱스 리스트
IDENTIFIER_NAME = 'identifier' # 식별자 이름 설정

# 엑셀 파일에서 특정 시트를 읽고 전처리를 수행하는 함수
def preprocess_excel_sheet(sheet_name):
    df = pd.read_excel(TARGET_EXCEL_PATH, sheet_name=sheet_name, header=None)
    df.columns = range(df.shape[1]) # 칼럼명 변경
    df.fillna('', inplace=True) # 빈칸 처리
    df[IDENTIFIER_NAME] = combine_columns(df, TARGET_ID_LIST) # 비교용 구분자 생성
    return df

# 데이터프레임의 각 행에서 target_ids에 해당하는 인덱스의 열 값을 문자열로 결합하여 반환
def combine_columns(df, target_column_index_list):
    combined_columns = []
    for index, row in df.iterrows():
        if all(row[idx] != '' for idx in target_column_index_list):
            combined_str = ''.join(str(row[idx]) for idx in target_column_index_list)
            combined_columns.append(combined_str)
        else:
            combined_columns.append('')
    return combined_columns

def get_large_categories(df):
    column_values = df.iloc[:, LARGE_CATEGORY_INDEX].tolist()
    distinct_values = list(set(value for value in column_values if value != ''))
    return distinct_values

def get_compare_result(df1, df2, large_categories1, large_categories2):
    results = [] # 결과를 담을 리스트 초기화

    # TO-DO
    # large_categories1, large_categories2 순으로 순회하고, 일치하는 요소가 있을 경우 비교한다.
    # 각각의 large_categories 는 df[IDENTIFIER_NAME] 로 비교한다.
    # large_categories1 에만 요소가 있을 경우 맨 처음 칼럼에 "제거" 후 동일한 데이터 대로 row 를 한줄 추가한다.
    # large_categories2 에만 요소가 있을 경우 맨 처음 칼럼에 "추가" 후 동일한 데이터 대로 row 를 한줄 추가한다.
    # large_categories1, large_categories2 가 다를 경우 각각 동일한 데이터 대로 row 를 한줄씩 총 두줄을 추가한다.
    # 첫번째 row sms "당초" 두번째 row 는 "변경"
    return results

# 결과를 데이터프레임으로 변환하고 엑셀 파일에 추가하는 함수
def append_results_to_excel(results, df1_columns):
    book = openpyxl.load_workbook(TARGET_EXCEL_PATH)

    # 결과 데이터프레임 생성
    df3 = pd.DataFrame(results, columns=[f"{col}" for col in df1_columns if col not in TARGET_ID_LIST + [IDENTIFIER_NAME]])

    # ExcelWriter를 사용하여 기존 파일에 시트 추가
    with pd.ExcelWriter(TARGET_EXCEL_PATH, engine='openpyxl', mode='a') as writer:
        writer.book = book
        df3.to_excel(writer, sheet_name=RESULT_SHEET_NAME, index=False)

    writer.save()

def main():
    # 엑셀 파일에서 각 시트 읽기
    df1 = preprocess_excel_sheet(COMPARE_SHEET_NAME1)
    df2 = preprocess_excel_sheet(COMPARE_SHEET_NAME2)

    # 대분야 리스트
    large_categories1 = get_large_categories(df1)
    large_categories2 = get_large_categories(df2)

    # 비교 및 결과 추가 함수 호출
    difference = get_compare_result(df1, df2, large_categories1, large_categories2)

    # append_results_to_excel

if __name__ == "__main__":
    main()