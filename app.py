import base64
import pandas as pd
import streamlit as st

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

    download = st.button("Convert data")

    if download:
        tmp_download_link = download_link(df, name+'_venue_data.csv', 'Click here to download your converted data.')
        st.markdown(tmp_download_link, unsafe_allow_html=True)


def download_link(object_to_download, download_filename, download_link_text):
    """
    Generates a link to download the given object_to_download.

    object_to_download (str, pd.DataFrame):  The object to be downloaded.
    download_filename (str): filename and extension of file. e.g. mydata.csv, some_txt_output.txt
    download_link_text (str): Text to display for download link.

    Examples:
    download_link(YOUR_DF, 'YOUR_DF.csv', 'Click here to download data!')
    download_link(YOUR_STRING, 'YOUR_STRING.txt', 'Click here to download your text!')

    """
    if isinstance(object_to_download,pd.DataFrame):
        object_to_download = object_to_download.to_csv(index=False)

    # some strings <-> bytes conversions necessary here
    b64 = base64.b64encode(object_to_download.encode()).decode()

    return f'<a href="data:file/txt;base64,{b64}" download="{download_filename}">{download_link_text}</a>'


title = "Veezi Convertor"

st.set_page_config(page_title=title, page_icon=None, layout='centered', initial_sidebar_state='auto')

st.title(title)
st.subheader("For the Cinema Incentive Scheme and Film Hub Scotland.")

st.markdown("---")

st.subheader("What it does:")

st.write("This tool converts a Veezi marketing report into usable data for the Cinema Incentive Scheme.")

st.subheader("How to use it:")

st.write("Works best with Google Chrome, not with Safari")

st.write("Enter the name of the venue below, upload the marketing report, and a button will appear below to download the converted data.")

s = st.text_input('Enter the venue name here:')

uploaded_file = st.file_uploader("Choose a file")

st.markdown("---")

st.write("*Not working? Email: hello@rae.li*")

st.markdown("---")

if uploaded_file is not None:
    format(uploaded_file, s)