import pandas as pd
from datetime import datetime

def raw_report(df, name):
    df = df[["Film", "Admits", "Shows"]]
    df.columns = ["TITLE", "LIVE ADMISSIONS", "no. SCREENINGS"]
    df["CINEMA"] = name
    df["YEAR"] = datetime.today().strftime('%Y-%m-%d')

    order = ["YEAR", "TITLE", "CINEMA", "LIVE ADMISSIONS", "no. SCREENINGS"]
    df = df[order]
    df = df.sort_values("TITLE")

    return df


def show_by_time_report(df, name):
    df = df.drop(df.index[0: 7])
    header = df.iloc[0]
    df.columns = header
    df = df.iloc[1:]

    df = df[["FILM", "START"]]

    df = df.groupby(["FILM"]).count().reset_index()

    df.columns = ["TITLE", "no. SCREENINGS"]
    df["CINEMA"] = name
    df["YEAR"] = datetime.today().strftime('%Y-%m-%d')
    df["LIVE ADMISSIONS"] = 0

    order = ["YEAR", "TITLE", "CINEMA", "LIVE ADMISSIONS", "no. SCREENINGS"]
    df = df[order]

    df = df.sort_values("TITLE")

    return df


def marketing_report(df, name):
    # find where to start
    search_films = "Film"
    start_point = df.loc[df.isin([search_films]).any(axis=1)].index.tolist()
    start = start_point[0]
    df = df.drop(df.index[0: start])

    # finds where to stop 
    search = "Payment Method Analysis - Box Office"
    stop_point = df.loc[df.isin([search]).any(axis=1)].index.tolist()
    stop = stop_point[0] - 29

    # drop everything after that point
    df = df.drop(df.index[stop:])
    df = df.reset_index(drop=True)

    # add cinema name
    cinema = name
    df["CINEMA"] = cinema

    # reorder
    order = ["Unnamed: 10", "Unnamed: 2", "CINEMA", "Unnamed: 33", "Unnamed: 25"]
    df = df[order]

    # rename columns
    df.columns = ["YEAR", "TITLE", "CINEMA", "LIVE ADMISSIONS", "no. SCREENINGS"]
    df = df.drop(df.index[0])
    df = df.sort_values("TITLE")

    return df


def format(file, name):
    df = pd.read_excel(file)
    raw = False
    show_by_time = False

    if "Film" in df.columns:
        # Means it's a raw data sheet, not a report
        df = raw_report(df, name)
        raw = True

    if "Shows by Time" in df.columns:
        # Means it's shows by time report
        df = show_by_time_report(df, name)
        show_by_time = True
        raw = False

    if not raw and not show_by_time:
        # Means it's a marketing report
        df = marketing_report(df, name)

    df.to_csv(f"{name}.csv", index=False)
        

    # if st.button('Download converted data as CSV'):
	#     tmp_download_link = download_link(df, name+'_venue_data.csv', 'Click here to download your converted data.')
	#     st.markdown(tmp_download_link, unsafe_allow_html=True)


format('test1.xlsx', "Oban")
format('test2.xlsx', "CPH")
format('test3.xlsx', "birks")
print("Done")