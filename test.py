import oodles as oo
import pandas as pd
import numpy as np


def text_ops():
    slide_id = "1D9Boj66ia_nu2KKMysIFLuMUvK0LO3AyXp5s8JhmXCY"
    slides = oo.Slides(slide_id)

    slide["Title"].text = "Title"

    blocks = slide["Yoyo"]
    blocks += slide["test"]

    blocks.replace("Yoyo", "Yaya")
    blocks.text = "salut"


def replace_chart():
    slide_id = "1D9Boj66ia_nu2KKMysIFLuMUvK0LO3AyXp5s8JhmXCY"
    slides = oo.Slides(slide_id)

    sheets = oo.Sheets("10x3vhlNYUMeP_aBEgn9cAqpET8eUHrKi5YoMb_yGhDk")
    new_chart = sheets["Focus Commerce - Sectors"].chart[0]

    slides[3].chart[0].replace(new_chart)


def add_df_to_sheet():
    from pprint import pprint

    df = pd.DataFrame()
    df["name"] = [f"entry {i}" for i in range(10)]
    df["data"] = range(10)
    df["data2"] = range(10)
    df["data3"] = np.linspace(0, 1, 10)

    sheets = oo.Sheets("10x3vhlNYUMeP_aBEgn9cAqpET8eUHrKi5YoMb_yGhDk")
    sheets["Data"] = df
    sheets["Data"].values()


def create_chart():
    sheets["Data"].create_chart(
        x="data", y=["data3", "data2"],
        colors=["#456123", "#616178"],
        chart_type="AREA")
