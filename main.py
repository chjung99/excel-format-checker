import pandas as pd
from collections import Counter
from hanspell import spell_checker


# Excel 파일을 읽어오는 함수 정의
def check_format_in_excel(file_path, save_file_path):
    # Excel 파일을 DataFrame으로 읽어오기
    print("1. 스페이스 2개 확인을 진행합니다.")
    df = pd.read_excel(file_path)
    cnt = 0
    id_list = []
    # 각 행의 데이터를 확인하고 스페이스 공백이 있는 ID를 출력
    for index, row in df.iterrows():
        for column in df.columns:
            if column == "ID":
                id_list.append(str(row[column]))
            if column == "FEEDBACK" or column == "FIXED_OUTPUT":
                data = str(row[column])  # 각 셀의 데이터를 문자열로 변환하여 확인
                if column == "FIXED_OUTPUT":
                    d = data.split("```")
                    tmp = ""
                    for i in range(len(d)):
                        if i % 2 == 0:
                            tmp += d[i]
                    data = tmp
                for d in data.split("\n"):
                    if "  " in d:  # 스페이스 공백 두 개가 있는 경우
                        dc = d.replace("  ", "^^")
                        cnt += 1
                        print(
                            f"INDEX: {int(index) + 2}, ID: {row['ID']}, COLUMN: {column}, CONTENTS: {dc}")  # 해당 행의 ID를 출력

    print(f"1. 최종 검출된 개수: {cnt}")
    print("2. 중복 ID 확인을 진행합니다.")
    element_count = Counter(id_list)
    cnt2 = 0
    for k, v in element_count.items():
        if v >= 2:
            print(f"ID: {k}, 중복 개수: {v}")
            cnt2 += 1
    print(f"2. 최종 중복된 ID 개수: {cnt2}")
    df['FEEDBACK_CHECKED'] = None
    df['FEEDBACK_ERROR'] = None
    df['FIXED_OUTPUT_CHECKED'] = None
    df['FIXED_OUTPUT_ERROR'] = None
    for index, row in df.iterrows():
        for column in df.columns:
            data = str(row[column])  # 각 셀의 데이터를 문자열로 변환하여 확인
            if column == "FEEDBACK":
                result = spell_checker.check(data).as_dict()
                df.at[index, 'FEEDBACK_CHECKED'] = result['checked']
                if result['errors'] > 0:
                    df.at[index, 'FEEDBACK_ERROR'] = result['errors']
            elif column == "FIXED_OUTPUT":
                d = data.split("```")
                tmp = []
                error_sum = 0
                for i in range(len(d)):
                    if i % 2 == 0:
                        out = ""
                        splited_string = d[i].split("\n")
                        for d_idx, dd in enumerate(splited_string):
                            if dd == '':
                                out += "\n"
                            else:
                                result = spell_checker.check(dd).as_dict()
                                error_sum += result['errors']
                                out += result['checked']

                        tmp.append(out)

                new_data = ""
                k = 0
                for i in range(len(d)):
                    if i % 2 == 0:
                        new_data += tmp[k]
                        k += 1
                    else:
                        new_data += "```" + d[i] + "```"
                df.at[index, 'FIXED_OUTPUT_CHECKED'] = new_data
                df.at[index, 'FIXED_OUTPUT_ERROR'] = error_sum

    df.to_excel(save_file_path, index=False)
    with pd.ExcelWriter(save_file_path, engine='xlsxwriter') as writer:
        # 데이터프레임을 Excel 파일로 쓰기
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # ExcelWriter 객체에서 Workbook과 Worksheet 객체 가져오기
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # 셀의 자동 줄 바꿈 설정
        cell_format = workbook.add_format({'text_wrap': True})
        worksheet.set_column('B:J', 30, cell_format)


# Excel 파일 경로 설정 (파일 경로에 맞게 수정해주세요.)
drive_path = ""
excel_file_path = ''
new_excel_save_file_path = ''

# 함수 호출

check_format_in_excel(excel_file_path, new_excel_save_file_path)
