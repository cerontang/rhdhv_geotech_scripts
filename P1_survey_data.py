import xlwings as xw
import datetime
import pandas as pd

source_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Pier 1 Level surveys\20230127\Ground Level Monitoring (GH 01-GH02) 27.01.2023_NACO.xlsx"
WB = xw.Book(source_path)
WS = WB.sheets['21.03.22 to 12.08.22 master']

ID_list = []
start_date_list = []
end_date_list = []

for i in range(2, 168):
    point_ID = WS.range(f'B{i}').value
    level_switch = "N"
    local_date_list = []
    for j in range(0, 192):
        date = WS.range(f'I1').offset(0, j).value
        level = WS.range(f'I{i}').offset(0, j).value
        if level is not None:
            local_date_list.append(str(date))

    start_date_string = local_date_list[0]
    end_date_string = local_date_list[-1]
    date_format = '%d/%m/%Y'
    start_formatted_date = datetime.datetime.strptime(start_date_string, '%Y-%m-%d %H:%M:%S').strftime(date_format)
    end_formatted_date = datetime.datetime.strptime(end_date_string, '%Y-%m-%d %H:%M:%S').strftime(date_format)
    print(point_ID, start_formatted_date, end_formatted_date)

    ID_list.append(point_ID)
    start_date_list.append(start_formatted_date)
    end_date_list.append(end_formatted_date)

outputList = list(zip(ID_list, start_date_list, end_date_list))
dataframe = pd.DataFrame(outputList)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(dataframe)
dataframe.to_csv(f'csv\Pier 1 Survey Chart Timeframe Asessment.csv')
