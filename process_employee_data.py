import warnings
warnings.filterwarnings("ignore", message="Parsing dates in %Y-%m-%d %H:%M:%S format when dayfirst=True was specified.")
warnings.filterwarnings("ignore", category=FutureWarning)

import pandas as pd
import os
import calendar
from datetime import datetime, timedelta, date
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter import simpledialog

def detect_columns(df, mapping):
    df.columns = df.columns.str.lower().str.strip()
    detected = {}
    for key, possible_names in mapping.items():
        for name in possible_names:
            if name.lower().strip() in df.columns:
                detected[key] = name.lower().strip()
                break
    return detected

def parse_time(value):
    if pd.isna(value) or value == '':
        return None
    value = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(value, fmt).time()
        except ValueError:
            pass
    return None

def parse_date_both_ways(series):
    s_no = pd.to_datetime(series, dayfirst=False, errors="coerce")
    s_yes = pd.to_datetime(series, dayfirst=True, errors="coerce")
    return s_no, s_yes

def pick_dominant_month(s_no, s_yes):
    df_no = pd.DataFrame({"date": s_no.dropna()})
    df_yes = pd.DataFrame({"date": s_yes.dropna()})
    df_both = pd.concat([df_no, df_yes], ignore_index=True)
    if df_both.empty:
        return None, None
    df_both["year"] = df_both["date"].dt.year
    df_both["month"] = df_both["date"].dt.month
    grp = df_both.groupby(["year","month"]).size().reset_index(name="count")
    if grp.empty:
        return None, None
    row = grp.loc[grp["count"].idxmax()]
    return int(row["year"]), int(row["month"])

def pick_best_parse(series):
    s_no, s_yes = parse_date_both_ways(series)
    y_no, m_no = pick_dominant_month(s_no, s_no)
    y_yes, m_yes = pick_dominant_month(s_yes, s_yes)
    count_no = ((s_no.dt.year == y_no) & (s_no.dt.month == m_no)).sum()
    count_yes = ((s_yes.dt.year == y_yes) & (s_yes.dt.month == m_yes)).sum()
    if count_yes >= count_no:
        return s_yes.dt.date, y_yes, m_yes
    else:
        return s_no.dt.date, y_no, m_no

def january_like_week(d):
    if pd.isna(d):
        return None
    day = d.day
    if day <= 7:
        return 1
    elif day <= 14:
        return 2
    elif day <= 21:
        return 3
    elif day <= 28:
        return 4
    else:
        return 5


def parse_punch_report(df):
    s_no, s_yes = parse_date_both_ways(df["punch date"])
    y_no, m_no = pick_dominant_month(s_no, s_no)
    y_yes, m_yes = pick_dominant_month(s_yes, s_yes)
    count_no = ((s_no.dt.year == y_no) & (s_no.dt.month == m_no)).sum()
    count_yes = ((s_yes.dt.year == y_yes) & (s_yes.dt.month == m_yes)).sum()
    if count_yes >= count_no:
        df["parsed_date"] = s_yes.dt.date
        best_year, best_month = y_yes, m_yes
    else:
        df["parsed_date"] = s_no.dt.date
        best_year, best_month = y_no, m_no

    
    df["parsed_time"] = df["punch time"].apply(parse_time)
    def combine_dt(row):
        if pd.isna(row["parsed_date"]) or pd.isna(row["parsed_time"]):
            return None
        return datetime.combine(row["parsed_date"], row["parsed_time"])
    df["parsed_datetime"] = df.apply(combine_dt, axis=1)
    df.dropna(subset=["parsed_datetime"], inplace=True)

  
    df = df[df["parsed_datetime"].apply(lambda d: d.year==best_year and d.month==best_month)]

    return df, best_year, best_month

def calc_daily_work_multi_in_out(group):
    group = group.sort_values("parsed_datetime")
    times = []
    for _, row in group.iterrows():
        times.append((row["parsed_datetime"], row["directionality"]))

    
    total_work = 0.0
    total_break = 0.0
    last_out = None
    last_in = None
    for i in range(len(times)):
        dt, direct = times[i]
        if direct.lower() == "in":
            if last_out is not None:
                diff = (dt - last_out).total_seconds()/3600
                if diff>0:
                    total_break += diff
            last_in = dt
        elif direct.lower() == "out":
            if last_in is not None:
                diff = (dt - last_in).total_seconds()/3600
                if diff>0:
                    total_work += diff
            last_out = dt


    all_in = [t[0] for t in times if t[1].lower()=="in"]
    all_out= [t[0] for t in times if t[1].lower()=="out"]
    if len(all_in)>0:
        first_in = min(all_in)
    else:
        first_in = group["parsed_datetime"].iloc[0]
    if len(all_out)>0:
        final_out = max(all_out)
    else:
        final_out = group["parsed_datetime"].iloc[-1]


    wn = january_like_week(first_in.date())
    return pd.Series({
        "daily hours": total_work,
        "resting hours": total_break,
        "clock_in_dt": first_in,
        "clock_out_dt": final_out,
        "week number": wn
    })

def main():
    root = Tk()
    root.withdraw()

    input_file = askopenfilename(title="Select the Raw Data File",
                                 filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not input_file:
        print("[ERROR] No input file selected. Exiting.")
        return

    try:
        df = pd.read_excel(input_file, sheet_name=0, dtype=str)
    except Exception as e:
        print("[ERROR] Could not load the file:", e)
        return

   
    col_lower = [c.lower().strip() for c in df.columns]
    if "punch date" in col_lower and "punch time" in col_lower and "directionality" in col_lower:
        print("[INFO] Det ligner en Punch-format rapport. Parser derefter.")
        punch_date_col = [c for c in df.columns if c.lower().strip()=="punch date"][0]
        punch_time_col = [c for c in df.columns if c.lower().strip()=="punch time"][0]
        direction_col  = [c for c in df.columns if c.lower().strip()=="directionality"][0]

        df.rename(columns={
            punch_date_col: "punch date",
            punch_time_col: "punch time",
            direction_col:  "directionality"
        }, inplace=True)

        df, best_year, best_month = parse_punch_report(df)
        employee_id = simpledialog.askstring("Employee ID", "Enter the Employee ID to extract data for:")
        if not employee_id:
            print("[ERROR] No Employee ID entered. Exiting.")
            return
        if "employee id" in df.columns.str.lower():
            df.columns = df.columns.str.lower().str.strip()
        df = df[df["employee id"]==employee_id.lower()]
        if df.empty:
            print(f"[WARNING] No data found for Employee ID '{employee_id}'. Exiting.")
            return

      
        df["parsed_datetime"] = pd.to_datetime(df["parsed_datetime"], errors="coerce")
        df = df.sort_values("parsed_datetime")

        
        grouped = df.groupby(["employee id", df["parsed_datetime"].dt.date]).apply(calc_daily_work_multi_in_out).reset_index()
        

        print("[INFO] Punch-format parse complete. daily hours + resting hours computed for multiple In/Out.")
        return
    else:
        print("[INFO] Kører gammelt flow, fordi vi ikke fandt 'punch date/punch time/directionality' kolonner.")
      

if __name__=="__main__":
    main()
