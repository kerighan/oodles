import oodles as oo
import pandas as pd
import numpy as np


def text_ops():
    slide_id = "1D9Boj66ia_nu2KKMysIFLuMUvK0LO3AyXp5s8JhmXCY"
    slides = oo.Slides(slide_id)

    # get slide
    slide = slides[1]
    # change text
    slide["Title"] = "Title2"

    # change picture
    slide = slides[2]
    slide.img[0] = "/home/maixent/Téléchargements/DrC7FHBX4AAFAFM.png"

    # find the first block text that contains "yoyo"
    blocks = slide["Yoyo"]
    # add to it the first block text that contains "test"
    blocks += slide["test"]
    # in all blocks, replace the substring Yoyo by Yuyu
    blocks.replace("Yoyo", "Yuyu")


def add_df_to_sheet():
    from pprint import pprint

    # create a dummy dataframe
    df = pd.DataFrame()
    df["name"] = [f"entry {i}" for i in range(10)]
    df["data"] = range(10)
    df["data2"] = range(10)
    df["data3"] = np.linspace(0, 1, 10)

    # create sheets obj
    sheets = oo.Sheets("10x3vhlNYUMeP_aBEgn9cAqpET8eUHrKi5YoMb_yGhDk")
    # insert new values
    sheets["Data"] = df
    # get new values
    sheets["Data"].values()
    # create chart from data
    sheets["Data"].create_chart(
        x="data", # absciss data
        y=["data3", "data2"],  # series to compare
        colors=["#456123", "#616178"],  # colors to use
        chart_type="AREA"  # can be of many types
    )


def replace_chart():
    slide_id = "1D9Boj66ia_nu2KKMysIFLuMUvK0LO3AyXp5s8JhmXCY"
    slides = oo.Slides(slide_id)

    sheets = oo.Sheets("10x3vhlNYUMeP_aBEgn9cAqpET8eUHrKi5YoMb_yGhDk")
    # .chart attribute is a list of all charts
    new_chart = sheets["Data"].chart[0]
    # replace the 0th chart in the 3rd slide by the chart on the Google Sheet

    slides[3].chart[0] = sheets["Data"].chart[0]


replace_chart()
