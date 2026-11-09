import tempfile

import pandas as pd
import regex as re, os
from datetime import datetime, time, timedelta, date
from stqdm import stqdm
import numpy as np   # used for np.select (first-match-wins logic)
import streamlit as st

validFunnels = ["AI", "PU","SMAI","DRF","PYTHON","EXCEL","10XTECHIES"]
FunnelGrouping = {"Python": ["PYTHON"], 
                  "Excel" : ["EXCEL"], 
                  "AI" : [ "AI", 'BSAI'], 
                  "10xTechies" :[ "10XTECHIES"], 
                  "SMAI Exotic": ["SMAI"], 
                  "AI Exotic": [ "AI", 'BSAI'], 
                  "AI Bootcamp": [ "AI", 'BSAI'], 
                  "SMAI" : ["SMAI"], 
                  "DRF" : ["DRF", "DR-FINANCE"], 
                  "AI TV": [ "AI", 'BSAI'], 
                  "PU" : ["PU"], 
                  "AI BootcampPaid": [ "AI", 'BSAI'] , 
                  "Python BootcampPaid" : ["PYTHON"],
                  "AI_AC" : ["AI", "EXCEL", "PYTHON"],
                  "PYTHON_AC" : ["PYTHON"],
                  "EXCEL_AC" : ["EXCEL"]}

# Funnel code -> list of keywords that identify that funnel in a topic name
FunnelAttendance = {
    "AI":   ["AI-Powered", "AI Tools"],
    "PU":   ["ProfitUni Stock Market"],
    "SMAI": ["Stock Market Using AI"],
}


FunnelMapping = {'AI BootcampPaid': "AI", 
                 'AI Bootcamp':"AI", 
                 'AI Exotic': "AI", 
                 'AI' : "AI", 
                 'Excel': "AI", 
                 'Python': "AI", 
                 'SMAI Exotic' :"SMAI", 
                 'SMAI' :"SMAI", 
                 "PU" : "PU"   }



def generatePaymentReport(filePath):
    paymentReport = pd.read_csv(filePath, sep=",", date_format="%Y-%m-%d %H:%M:%S", dayfirst=True,  low_memory=False)

    paymentReport = paymentReport[(~paymentReport["Payment Slug"].str.contains("-0-2"))]
    #Filter the rows
    #paymentReport = paymentReport[(paymentReport["Tags"].fillna("empty").isin(["l2_seat_booking", "l3_seat_booking"]))]

    #Changing the data type of phone number column
    paymentReport["Phone Number"] = paymentReport["Phone Number"].astype(str).str.replace(r"\D", "", regex=True)

    #Filtering the records according to the condition -- captured and PaymentAmountThreshold
    paymentReport = paymentReport[(paymentReport["Status"].str.strip().str.lower()  == "captured")]

    #Formatting the column as date time
    paymentReport["CreatedAt"] = pd.to_datetime(paymentReport["CreatedAt"], format = "%Y-%m-%d %H:%M:%S", exact=True, dayfirst=True, yearfirst=False)

    _date = paymentReport["CreatedAt"].unique()[0].date()
    print(_date)
    startTime = datetime.combine(_date, time(00, 00, 00))
    endTime = datetime.combine(_date, time(23, 59, 00))
    paymentReport = paymentReport[paymentReport["CreatedAt"].between(startTime, endTime, inclusive="both")]

    paymentReport.sort_values(by=["CreatedAt"], ascending=True, inplace=True, ignore_index=True)
    paymentReport["Payment Method"] = paymentReport["Payment Slug"].map(lambda x: str(x).split("-0")[0].upper() if x != "" else pd.NA)

    paymentReport.rename(columns={"Payment Method": "PaymentFunnel",}, inplace=True)
    paymentReport = paymentReport[ (paymentReport["Payment Slug"].str.contains("-0-"))] #(paymentReport["PaymentFunnel"].isin(validFunnels) ) &
    #Removing 71K BSAI
    paymentReport = paymentReport[ (~paymentReport["Payment Slug"].str.contains("71"))]
    print(paymentReport["PaymentFunnel"].unique())
    paymentReport = paymentReport.sort_values(by=["Amount"], ascending=[False]).drop_duplicates(subset=["Email", "Phone Number"], keep="first").reset_index(drop=True)
    return paymentReport

def CountIf(Main_File, Current_File, MFCol, CFCol, ValueCol = None):
    NewColName = MFCol + "s" 

    # ✅ OPTIMIZED: Build a lookup set for O(1) membership check instead of O(n) row-by-row scan
    lookup_set =  Current_File[CFCol].astype(str).str.lower().str.strip()
    if NewColName in Main_File.columns:
       Main_File = Main_File.drop(columns = NewColName )

    MaxLookup = Current_File[ValueCol].groupby(lookup_set).max()

    Main_File.insert(loc=len(Main_File.columns), column=NewColName,
                     #value=Main_File[MFCol].astype(str).str.lower().str.strip().isin(lookup_set).astype(int),
                     value = Main_File[MFCol].astype(str).str.lower().str.strip().map(MaxLookup).fillna(0), 
                     allow_duplicates=False)
      # NEW: max value per key (handles people who joined more than once)
      #value =   # NEW: look up each main-file key's max value, 0 if not found
    
    return Main_File, NewColName

def ValidatePayment(ExcelFile, Current_File):
    validatedSheets  = {}
    info = {}

    with pd.ExcelFile(ExcelFile) as f:
        sheet_names = f.sheet_names

    df = pd.read_excel(ExcelFile, sheet_names)

    for sheet in stqdm(sheet_names):
        try:
            final_df = df.get(sheet)
        
            MFCombinations = ['Email', 'Phone Number']
            AttendeeCombinations = ['Email', 'Phone Number']

            #inal_df.loc[:, 'Phone Number'] = final_df['Phone Number'] #Converts the data type to string datatype
            final_df.loc[: , "Email"] = final_df["Email"].astype("str").str.strip() #Converts the data type to string datatype
            
            if "PaymentFunnel" in final_df.columns:
                Funnel = final_df["PaymentFunnel"].unique()[0]
            else:
                final_df["PaymentFunnel"] = sheet
                Funnel = sheet

            getFunnel = FunnelGrouping.get(Funnel)
            Payment_File = Current_File[Current_File["PaymentFunnel"].isin(getFunnel)]
            CurrentFileSumColumns = [] #Stores all the columns for the summation
            for MFCol, CFCol  in zip(MFCombinations, AttendeeCombinations):
                final_df, NewColName = CountIf(final_df, Payment_File, MFCol, CFCol, "Amount")
                CurrentFileSumColumns.append(NewColName)

            #Assigns the webinarId if the countif sum is greater than 0 else "Absent"
            if "MaxAmount" in final_df.columns:
                final_df = final_df.drop(columns = "MaxAmount")

            final_df["MaxAmount"] = final_df.apply(lambda x: x[CurrentFileSumColumns].max(), axis=1)
            columns = list(final_df.columns[:6])+CurrentFileSumColumns+["MaxAmount"]+list(final_df.columns[6:-3])
            final_df = final_df[columns]

            validatedSheets.update({sheet: final_df})
            TotalCount = len(final_df)
            PaidCount = len(final_df[final_df["MaxAmount"]>0])
            info.update({sheet: [sheet, TotalCount, PaidCount]})
        except Exception as e:
            print(e)

    AVP = pd.DataFrame(columns = ["Funnel", "TotalCount", "Paid"])

    for _, values in info.items():
        AVP.loc[len(AVP), :] = values

    return validatedSheets, AVP

 

def matchAttendance(concatDF, Attendance):
    # Funnel code -> list of keywords that identify that funnel in a topic name
    FunnelAttendance = {
        "AI":   ["AI-Powered", "AI Tools"],
        "PU":   ["ProfitUni Stock Market"],
        "SMAI": ["Stock Market Using AI"],
    }

    # Build one compiled-style regex pattern per funnel, e.g. "AI\-Powered|AI\ Tools"
    # re.escape makes sure characters like "-" or "." are treated as plain text
    patterns = {
        funnel: "|".join(re.escape(keyword) for keyword in keywords)
        for funnel, keywords in FunnelAttendance.items()
    }

    # Make sure the column has no NaN problems and is treated as text
    topics = Attendance["TopicName"].astype("string")

    # One True/False column of matches per funnel, evaluated on the whole column at once
    # na=False means an empty topic name counts as "no match" instead of raising an error
    conditions = [
        topics.str.contains(pattern, regex=True, na=False)
        for pattern in patterns.values()
    ]

    # Funnel codes in the same order as the conditions above
    choices = list(patterns.keys())

    # For each row, pick the code of the first condition that is True
    # If nothing matches, use pd.NA as the default (same as your original intent)
    Attendance["Funnel"] = np.select(conditions, choices, default=pd.NA)

    Attendance.sort_values(by= "OptimizedSessionDuration", ascending=False, inplace=True)

    Attendance["Email"] = Attendance["ZoomEmail"].str.lower().str.strip()
    Attendance["Phone Number"] = pd.to_numeric(Attendance["Phone Number"], errors="coerce")


    EmailAttendance = Attendance[["Email", "OptimizedSessionDuration",  'TopicName', 'ZoomAccount', 'Funnel']]
                        
    PhoneAttendance = Attendance[["Phone Number", "OptimizedSessionDuration",  'TopicName', 'ZoomAccount', 'Funnel']] 
                        
    FunnelMapping = {'AI BootcampPaid': "AI", 
                    'AI Bootcamp':"AI", 
                    'AI Exotic': "AI", 
                    'AI' : "AI", 
                    'Excel': "AI", 
                    'Python': "AI", 
                    'SMAI Exotic' :"SMAI", 
                    'SMAI' :"SMAI", 
                    "PU" : "PU"   }


    mergedEmail = concatDF.merge(EmailAttendance, left_on=["Funnel", "Email"], right_on=["Funnel", "Email"] , how = "left")
    mergedPhone = concatDF.merge(PhoneAttendance, left_on=["Funnel", "Phone Number"], right_on=["Funnel", "Phone Number"] , how = "left")

    AttendancePayment = pd.concat([mergedEmail, mergedPhone ], axis = "rows", ignore_index=True).sort_values(by = "OptimizedSessionDuration", ascending = False).drop_duplicates(subset = 'Payment Id')

    AttendancePayment["CreatedAt"] = AttendancePayment["CreatedAt"].astype('M8[s]')
    AttendancePayment.sort_values(by = ["PaymentFunnel", "CreatedAt"], ascending = True,inplace=True, ignore_index=True)
    AttendancePayment.rename(columns={"OptimizedSessionDuration" : "SessionDuration"}, inplace=True)

    return AttendancePayment


def fetchPA(PA, AttendancePayment,isAttendancePath ):
    PA["Phone Number"] = pd.to_numeric(PA["Phone Number"].str.replace(r"\D+", "", regex=True), "coerce")
    PA["Email"] = PA["Email"].str.lower().str.strip()
    PA.sort_values(by= ["Instacred Limit Amount"] , ascending=False , inplace=True, ignore_index=True)

    PA_Email = PA[["Email", "Instacred Limit Amount"]].rename(columns={"Instacred Limit Amount" : "InstacredLimitAmount_Email"})
    PA_Phone = PA[["Phone Number", "Instacred Limit Amount"]].rename(columns={"Instacred Limit Amount" : "InstacredLimitAmount_Phone"})
    merged_PA = AttendancePayment.merge(PA_Email, on = ["Email"], how="left").merge(PA_Phone,  on= ["Phone Number"],  how = "left")

    merged_PA["InstacredLimitAmount"] = merged_PA[["InstacredLimitAmount_Email", "InstacredLimitAmount_Phone"]].fillna(0).max(axis = 1)
    merged_PA.drop(columns=["InstacredLimitAmount_Email", "InstacredLimitAmount_Phone"], inplace=True)

    if isAttendancePath == True:
        merged_PA = merged_PA.sort_values(by =['InstacredLimitAmount', "SessionDuration"], ascending = [False, False]).drop_duplicates(subset = 'Payment Id').reset_index(drop=True)
    else:
        merged_PA = merged_PA.sort_values(by =['InstacredLimitAmount'], ascending = [False, False]).drop_duplicates(subset = 'Payment Id').reset_index(drop=True)

    merged_PA.sort_values(by = ["PaymentFunnel", "CreatedAt"], ascending = True,inplace=True, ignore_index=True)


    merged_PA = merged_PA.drop(columns=["Status", "Tags", "woocommerce OrderID", "Abandon Cart", "Emails", "Phone Numbers", "Funnel"])

    return merged_PA

def next_sunday():
    """
    Returns the date of the upcoming Sunday in 'YYYY-MM-DD' format.
    If today is already a Sunday, returns today's date.
    """
    today = date.today()
    # Monday=0 ... Sunday=6
    days_until_sunday = (6 - today.weekday()) % 7
    result = today + timedelta(days=days_until_sunday)
    return result.strftime("%Y-%m-%d")

def save_upload(uploaded_file):
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.read())
    return tmp_path

st.set_page_config("📥 MEGA File Validator", layout="wide")
st.title("📥 MEGA File Validator", text_alignment="center")

st.divider()

st.info("Upload at least the MEGA File and Payment File to process.")

WSDate =str(st.date_input("Select the Next Sunday date",value=next_sunday()))

col1, col2 = st.columns(2)

with col1:
    MEGAFilePath = st.file_uploader(label = "Upload MEGA File", accept_multiple_files=False, type=["xlsx"], width="stretch")
with col2:  
    PaymentmentFilePath = st.file_uploader(label = "Upload Payment File", accept_multiple_files=False, type=["csv"], width="stretch")

st.subheader("Optional Files")
st.divider()
col3, col4 = st.columns(2)

with col3:  
    AttendanceFilePath = st.file_uploader(label = "Upload Attendance File", accept_multiple_files=False, type=["xlsx"], width="stretch")
with col4:
    PAFilePath = st.file_uploader(label = "Upload Instacred Limit Amount File", accept_multiple_files=False, type=["csv", "zip"], width="stretch")


st.divider()
outputExcel = rf"MEGA_Merged_Validated_{WSDate}.xlsx"

if MEGAFilePath and PaymentmentFilePath:

    btn = st.button("Validate Data", type="primary")

    if btn:
        with st.status("Processing..", expanded=True) as status:
            Current_File = generatePaymentReport(save_upload(PaymentmentFilePath))
            validatedSheets, AVP = ValidatePayment(save_upload(MEGAFilePath), Current_File) 
    
    
        with pd.ExcelWriter(outputExcel) as file:
            for sheet, data in validatedSheets.items():
                data["Phone Number"] = data["Phone Number"].astype(float, errors = "ignore")
                data.to_excel(file, sheet_name = sheet, index=False)
            
            Current_File.to_excel(file, sheet_name="Payments", index=False)
            AVP.to_excel(file, sheet_name="MatchedPaymentSummary", index=False)


        st.dataframe(AVP, hide_index=True)

        st.dataframe(pd.pivot_table(Current_File, index="PaymentFunnel", values="CreatedAt", aggfunc="count", sort=True , margins=True ).rename(columns={"CreatedAt" : "Count"}) )

        with open(outputExcel, "rb") as f:
                st.download_button(
                    label="Save MEGA Validated file",
                    data=f.read(),
                    file_name=outputExcel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click="ignore"
                )


        concatDF = pd.concat(validatedSheets.values(), axis="rows")
        outputParquet = rf"MEGA_Merged_Validated_{WSDate}.parquet"
        if AttendanceFilePath:
            
            concatDF["Funnel"] = concatDF["PaymentFunnel"].map(lambda x: FunnelMapping[x] )
            concatDF["Phone Number"] = pd.to_numeric(concatDF["Phone Number"], errors="coerce")


            Attendance = pd.read_excel(save_upload(AttendanceFilePath))

            concatDF = matchAttendance(concatDF, Attendance)


        if PAFilePath:
            isAttendancePath = True if AttendanceFilePath else False
            PA  = pd.read_csv(save_upload(PAFilePath), low_memory=False)

            concatDF = fetchPA(PA, concatDF , isAttendancePath)
            concatDF.to_parquet(outputParquet, index=False, compression = "zstd")


        if AttendanceFilePath or PAFilePath:
            with open(outputParquet, "rb") as f:
                            st.download_button(
                                label="Save Detailed MEGA Validated file",
                                data=f.read(),
                                file_name=outputParquet,
                                mime="application/vnd.apache.parquet",
                                on_click="ignore"
                            )
