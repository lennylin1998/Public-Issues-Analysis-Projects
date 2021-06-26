import pandas as pd
import openpyxl
#資料來源：CDC["https://data.cdc.gov.tw/dataset"]
df = pd.read_csv('Day_Confirmation_Age_County_Gender_19CoV.csv')
df.loc[9269] = ['','',"金門縣",'','','','',0]
df[["縣市","鄉鎮"]]=df[["縣市","鄉鎮"]].replace(["空值"],"境外移入")


#增加各縣市疫苗施打比率
#資料來源：CDC["https://www.cdc.gov.tw/Category/Page/9jFXNbCe-sFK9EImRRi2Og"]
df["疫苗施打比率"] = df["縣市"].replace(["南投縣",'台中市','台北市','台南市','台東縣',
'嘉義市','嘉義縣','基隆市','宜蘭縣','屏東縣','彰化縣','新北市','新竹市','新竹縣','桃園市','澎湖縣','花蓮縣',
'苗栗縣','連江縣','雲林縣','高雄市','金門縣'],[0.0280, 0.0340,0.0565,0.0247,0.0258,0.0509,0.0220,0.0288,
0.0280, 0.0231, 0.0242,0.0291,0.0314,0.0249,0.0351,0.0318,0.0428,0.0182,0.1413,0.0211,0.0316,0.0165
])


#增加累積確診人數
cum_df = df.groupby(["縣市","個案研判日"])["確定病例數"].sum().reset_index()
cum_df["累積病例數(縣市)"] = cum_df.groupby("縣市")["確定病例數"].cumsum()
print(cum_df)

final_df = pd.merge(df, cum_df, on=["縣市", "個案研判日"], how="outer")
final_df[final_df["縣市"]=="金門縣"]["累積病例數"] = 0
print(final_df)
#轉換成excel檔案以匯入tableau
final_df.to_excel("確診個案資料總表.xlsx")


#清理出境內個案
local = df[df["是否為境外移入"]=="否"]
#個案數量(依鄉鎮)execl檔
case_per_town = local.groupby(["縣市", "鄉鎮"])["確定病例數"].sum()
df_town = pd.DataFrame(case_per_town)
df_town.to_excel("確診人數(鄉鎮).xlsx")
#個案數量(依縣市)excel檔
case_per_city = local.groupby(["縣市"])["確定病例數"].sum()
df_city = pd.DataFrame(case_per_city)
df_city.to_excel("確診人數(縣市).xlsx")
