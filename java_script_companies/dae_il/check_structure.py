# -*- coding: utf-8 -*-
import pandas as pd
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
out_path = os.path.join(script_dir, 'check_output.txt')
coa_path = os.path.join(script_dir, 'coa_dae_il_25.xlsx')

with open(out_path, 'w', encoding='utf-8') as f:
    df = pd.read_excel(coa_path)
    f.write('계정성격(col7) 고유값:\n')
    for v in df['계정성격'].unique():
        f.write(f'  {repr(v)}\n')

    f.write('\n감가상각 관련 계정 계정성격 샘플:\n')
    contra = df[df['계정과목'].str.contains('감가상각|충당|평가|손실충당', na=False)]
    f.write(contra[['계정코드','계정과목','대분류 (Level 1)','리드스케줄 (Level 4)','계정성격']].to_string() + '\n')

print('완료')
