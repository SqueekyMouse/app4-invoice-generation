import pandas as pd
import glob
#commit: load excel files into dataframes Sec25

filepaths=glob.glob('invoices/*.xlsx')
# print(filepaths)

for filepath in filepaths:
    df=pd.read_excel(filepath,sheet_name='Sheet 1') # excel files need sheet name!!!
    print(df)