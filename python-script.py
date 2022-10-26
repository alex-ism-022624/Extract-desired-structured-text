
import pandas as pd
pd.set_option("display.max_rows", 100000,"display.max_columns", 100)
import re

input = pd.read_excel('/content/drive/MyDrive/2022 Projects/VN Finance SKU Split Tool/Input/lazada_split.xlsx')


input['SKU'] = input['instruction'].apply(lambda x: re.findall(re.compile(r'\((\d+_\D+\d+)\) x\d+ \(\d+VND\)'),x))
input['SKU_2'] = input['SKU'].apply(lambda x: ', '.join(x))

input['No.item'] = input['instruction'].apply(lambda x: re.findall(re.compile(r'\(\d+_\D+\d+\) x(\d+) \(\d+VND\)'),x))
input['No.item_2'] = input['No.item'].apply(lambda x: sum([int(_i) for _i in x]))

input = input[['tracking_id','instruction','SKU_2','No.item_2']].reset_index (drop=True)


output_sheet = gc.open_by_key('13baPim6ZNPKyuVjPmv1wY1Ath0R_FYmUa2r1OYllAxU')

final_output = output_sheet.sheet1

final_output.update('A1',[input.columns.tolist()] + input.values.tolist())