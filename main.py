import os
import xlwings as xw
import pandas as pd

dataDIR = r'C:\Users\921722\OneDrive - Royal HaskoningDHV\IOM Stage 2 Monitoring\Data organised'
vwpList = os.listdir(dataDIR)

Dateandtime_list = []
CH1_AtmP_list = []
CH2_AtmP_list = []
CH1_Temp_list = []
Ch2_Temp_list = []
Ch1_Freq_list = []
Ch2_Freq_list = []


for vwp in vwpList:
    raw_data_path = f'C:/Users/921722/OneDrive - Royal HaskoningDHV/IOM Stage 2 Monitoring/Data organised/{vwp}'
    raw_data_file_list = os.listdir(raw_data_path)
    print(vwp, raw_data_file_list)
    for raw_data_file in raw_data_file_list:
        raw_data_file_path = f'C:/Users/921722/OneDrive - Royal HaskoningDHV/IOM Stage 2 Monitoring/Data organised/{vwp}/{raw_data_file}'
        rawdataWB = xw.Book(raw_data_file_path)
        sheetName = raw_data_file.replace(".csv", "")
        rawdataWS = rawdataWB.sheets[f'{sheetName}']
        for i in range(11, 9999):
            if rawdataWS.range(f'A{i}').value is None:
                rawdataWB.close()
                break
            print(rawdataWS.range(f'A{i}').value)
            Dateandtime_list.append(rawdataWS.range(f'A{i}').value)
            CH1_AtmP_list.append(rawdataWS.range(f'B{i}').value)
            CH2_AtmP_list.append(rawdataWS.range(f'B{i}').value)
            CH1_Temp_list.append(rawdataWS.range(f'E{i}').value)
            Ch2_Temp_list.append(rawdataWS.range(f'H{i}').value)
            Ch1_Freq_list.append(rawdataWS.range(f'D{i}').value)
            Ch2_Freq_list.append(rawdataWS.range(f'G{i}').value)

outputList = list(zip(Dateandtime_list, CH1_AtmP_list, CH2_AtmP_list, CH1_Temp_list, Ch2_Temp_list, Ch1_Freq_list, Ch2_Freq_list))

df = pd.DataFrame(outputList)
pd.set_option('display.max_rows', 99999)
pd.set_option('display.max_columns', 99999)
pd.set_option('display.width', 99999)
print(df)
#df.to_csv(r'csv\VWP raw data combined 24-05-2022.csv')