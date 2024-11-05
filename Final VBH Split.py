import shutil
import os

finalLogs_dir = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\01 Copy of all verification BHs\01 Final Logs to Check\Done"
finalLogs_list = os.listdir(finalLogs_dir)

AHAM_logs_dir = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\01 Copy of all verification BHs\01 Final Logs to Check\Done (SPT Check)\AHAM"
ERI_logs_dir = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\01 Copy of all verification BHs\01 Final Logs to Check\Done (SPT Check)\ERI"

for finalLog in finalLogs_list:
    if not finalLog.endswith(".pdf"):
        continue
    finalLog_path = f"{finalLogs_dir}\{finalLog}"
    if "- Final" in finalLog:
        finalLog_name = finalLog.replace(" - Final", "")
        finalLog_check_path = f"{AHAM_logs_dir}\{finalLog_name}"
        shutil.copyfile(finalLog_path, finalLog_check_path)
        continue
    if "_F" in finalLog:
        finalLog_name = finalLog.replace("_F", "")
        finalLog_check_path = f"{ERI_logs_dir}\{finalLog_name}"
        shutil.copyfile(finalLog_path, finalLog_check_path)
        continue



