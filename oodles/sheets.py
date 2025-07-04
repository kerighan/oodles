#!/usr/bin/env python
# -*- coding: utf-8 -*-
from typing import Union

from googleapiclient import errors
from googleapiclient.errors import HttpError

from .element import SheetChart
from .utils import hex_to_rgb, GoogleAuthorizationError
from .config import config


class Sheets:
    def __init__(self, doc_id: str):
        self.doc_id = doc_id
        self.url = "https://docs.google.com/spreadsheets/d/" + self.doc_id
        self.load()

    @staticmethod
    def create(title: str) -> "Sheets":
        body = {"properties": {"title": title}}
        spreadsheet = (
            config.SHEETS.spreadsheets()
            .create(body=body, fields="spreadsheetId")
            .execute()
        )
        return Sheets(spreadsheet.get("spreadsheetId"))

    def copy(self, title: str = None) -> "Sheets":
        if title is None:
            title = f"Copy of {self.title}"
        data = {"name": title}
        try:
            new_id = (
                config.DRIVE.files().copy(body=data, fileId=self.doc_id).execute().get("id")
            )
        except errors.HttpError as e:
            print("This error can happen if the file is in a shared drive")
            raise e
        return Sheets(new_id)

    def share_with(self, email: str, as_admin: bool = False):
        user_permission = {
            "type": "user",
            "role": "owner" if as_admin else "writer",
            "emailAddress": email,
        }
        config.DRIVE.permissions().create(
            fileId=self.doc_id,
            body=user_permission,
            fields="id",
            transferOwnership=as_admin,
        ).execute()

    def load(self):
        try:
            sheet = config.SHEETS.spreadsheets()
            result = sheet.get(spreadsheetId=self.doc_id).execute()
            self.document = result
        except HttpError as e:
            raise GoogleAuthorizationError(e, config.service_email)
        self.sheets = self.document.get("sheets")
        self.title = self.document["properties"]["title"]

    def __getitem__(self, name: Union[int, str]) -> "Sheet":
        if isinstance(name, int):
            return Sheet(self, self.doc_id, self.sheets[name])
        elif isinstance(name, str):
            for sheet in self.sheets:
                if name == sheet["properties"]["title"]:
                    return Sheet(self, self.doc_id, sheet)
            raise NameError(f"Sheet '{name}' not found")

    def __setitem__(self, name: str, df):
        try:
            sheet = self.__getitem__(name)
        except NameError:
            rows, cols = df.shape
            cols = max(cols, 26)
            rows = max(rows, 1000)
            requests = [
                {
                    "addSheet": {
                        "properties": {
                            "title": name,
                            "gridProperties": {"rowCount": rows, "columnCount": cols},
                            "tabColor": {"red": 1.0, "green": 0.3, "blue": 0.4},
                        }
                    }
                }
            ]
            config.SHEETS.spreadsheets().batchUpdate(
                spreadsheetId=self.doc_id, body={"requests": requests}
            ).execute()
            self.load()
            sheet = self.__getitem__(name)

        sheet.clear()
        sheet.value = df

    def __delitem__(self, name: str):
        sheet = self.__getitem__(name)
        requests = [{"deleteSheet": {"sheetId": sheet.sheet_id}}]
        config.SHEETS.spreadsheets().batchUpdate(
            spreadsheetId=self.doc_id, body={"requests": requests}
        ).execute()
        self.load()


class Sheet:
    def __init__(self, parent: Sheets, doc_id: str, content: dict):
        self.parent = parent
        self.doc_id = doc_id
        self.content = content
        self.title = self.content["properties"]["title"]
        self.sheet_id = self.content["properties"]["sheetId"]
        self.parse()

    def load(self):
        sheets = Sheets(self.doc_id)
        sheet = sheets[self.title]
        self.content = sheet.content
        self.title = sheet.title
        self.sheet_id = sheet.sheet_id
        self.parse()

    def clear(self):
        requests = [
            {
                "updateCells": {
                    "range": {"sheetId": self.sheet_id},
                    "fields": "userEnteredValue",
                }
            }
        ]
        config.SHEETS.spreadsheets().batchUpdate(
            spreadsheetId=self.doc_id, body={"requests": requests}
        ).execute()

    def parse(self):
        charts = []
        for obj in self.content.get("charts", []):
            spreadsheet_id = self.doc_id
            chart_id = obj["chartId"]
            charts.append(SheetChart(spreadsheet_id, chart_id))
        self.chart = charts

    def __setattr__(self, name: str, value):
        if name == "value":
            data = []
            for col in value.columns:
                # datetime64
                if value[col].dtype == "datetime64[ns]":
                    items = [col] + value[col].apply(
                        lambda x: x.strftime("%Y-%m-%d %H:%M:%S")
                    ).tolist()
                    data.append(items)
                else:
                    data.append([col] + value[col].tolist())
            resource = {"majorDimension": "COLUMNS", "values": data}

            data_range = f"{self.title}!A:A"
            config.SHEETS.spreadsheets().values().append(
                spreadsheetId=self.doc_id,
                range=data_range,
                body=resource,
                valueInputOption="USER_ENTERED",
            ).execute()
        else:
            super(Sheet, self).__setattr__(name, value)

    def values(self, _range=""):
        import pandas as pd

        if len(_range) > 0:
            sheet_title = f"{self.title}!{_range}"
        else:
            sheet_title = self.title
        request = (
            config.SHEETS.spreadsheets()
            .values()
            .get(spreadsheetId=self.doc_id, range=sheet_title)
        )
        values = request.execute()["values"]
        df = pd.DataFrame(values[1:], columns=values[0])
        return df

    def create_serie(self, col_id, size, colors=None, i=0, chart_type="COLUMN"):
        serie = {
            "series": {
                "sourceRange": {
                    "sources": [
                        {
                            "sheetId": self.sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": size + 1,
                            "startColumnIndex": col_id,
                            "endColumnIndex": col_id + 1,
                        }
                    ]
                }
            },
            "targetAxis": "BOTTOM_AXIS" if chart_type == "BAR" else "LEFT_AXIS",
        }
        if colors is not None:
            r, g, b = hex_to_rgb(colors[i])
            serie["color"] = {"red": r / 255, "green": g / 255, "blue": b / 255}
        return serie

    def create_chart(
            self,
            x,
            y,
            colors=None,
            label_color="#000000",
            background_color="#FFFFFF",
            chart_type="COLUMN",
            stacked_type="NOT_STACKED",
            legend_position="TOP_LEGEND",
            smooth=False
    ):
        if legend_position is None:
            legend_position = "NO_LEGEND"
        # get values
        df = self.values()
        size = df.shape[0]

        # create serie
        series = []
        cols = list(df.columns)

        # define x axis
        index_id = cols.index(x)

        # define y axis
        if isinstance(y, list):
            for i, y_ in enumerate(y):
                col_id = cols.index(y_)
                series.append(self.create_serie(col_id, size, colors, i))
        else:
            col_id = cols.index(y)
            series.append(self.create_serie(col_id, size, colors, 0, chart_type))

        # create background color
        r, g, b = hex_to_rgb(background_color)
        background_color = {"red": r / 255, "green": g / 255, "blue": b / 255}

        # create font color
        r, g, b = hex_to_rgb(label_color)
        text_format = {
            "foregroundColor": {"red": r, "green": g, "blue": b},
            "foregroundColorStyle": {"rgbColor": {"red": r, "green": g, "blue": b}},
            "bold": True,
            "italic": True,
            "fontSize": 24,
        }

        # create chart request batchupdate
        domains = [
            {
                "domain": {
                    "sourceRange": {
                        "sources": [
                            {
                                "sheetId": self.sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": size + 1,
                                "startColumnIndex": index_id,
                                "endColumnIndex": index_id + 1,
                            }
                        ]
                    }
                }
            }
        ]
        requests = [
            {
                "addChart": {
                    "chart": {
                        "spec": {
                            "backgroundColor": background_color,
                            "title": "",
                            "titleTextFormat": text_format,
                            "basicChart": {
                                "chartType": chart_type,
                                "legendPosition": legend_position,
                                "domains": domains,
                                "series": series,
                                "headerCount": 1,
                                "stackedType": (
                                    # stacked_type if chart_type == "BAR" else None
                                    stacked_type
                                ),

                            },
                        },
                        "position": {
                            "overlayPosition": {
                                "anchorCell": {
                                    "sheetId": self.sheet_id,
                                    "rowIndex": 0,
                                    "columnIndex": 0,
                                },
                                "offsetXPixels": 500,
                                "offsetYPixels": 100,
                            }
                        },
                    }
                }
            }
        ]

        # Additional graph options
        if chart_type == "LINE":
            requests[0]["addChart"]["chart"]["spec"]["basicChart"].update({
                "lineSmoothing": smooth
            })

        # Execute request
        config.SHEETS.spreadsheets().batchUpdate(
            spreadsheetId=self.doc_id, body={"requests": requests}
        ).execute()
        self.parent.load()
