pathy="P:\\Revizyon\\veri\\xxxxx.xlsx"
df=pd.read_excel(pathy, 'raw_data')

df=df.fillna(0)
df['tar'] = pd.to_datetime(df['TARİH']).dt.month #.astype(str).str.zfill(2)
df_sigara=df.loc[df['ALT GRUP'].isin(['SIGARALAR'])]

table = pd.pivot_table(df_sigara, values=['TUTAR'], index=['DEPO NO','DEPO ADI','tar'], aggfunc=np.sum)


file = "P:\Revizyon\rapor çalışmaları\spg_rk\pivot.xlsm"
xl = EnsureDispatch('Excel.Application')
wb=xl.Workbooks.Open(file)
ws=wb.Worksheets("tablo")



wb.RefreshAll()
