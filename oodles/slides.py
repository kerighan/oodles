#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import time

import addict
from googleapiclient.errors import HttpError

from .element import (Chart, Charts, Image, Images, SheetChart, TextBlock,
                      TextBlocks)
from .utils import hex_to_rgb
from .variables import DRIVE, SHEETS, SLIDES, service_email


class Slides:
    def __init__(self, doc_id):
        self.doc_id = doc_id
        try:
            self.document = SLIDES.presentations().get(
                presentationId=doc_id).execute()
        except HttpError:
            print(f"remember to share your slides with {service_email}\n")
            raise HttpError
        self.pages = self.document.get('slides')
        self.title = self.document["title"]
        self.url = "https://docs.google.com/presentation/d/" + self.doc_id

    def copy(self, title=None):
        if title is None:
            title = f"Copy of {self.title}"
        data = {"name": title}
        new_id = DRIVE.files().copy(
            body=data, fileId=self.doc_id
        ).execute().get("id")
        return Slides(new_id)

    def share_with(self, email):
        user_permission = {
            "type": "user",
            "role": "writer",
            "emailAddress": email
        }
        DRIVE.permissions().create(
            fileId=self.doc_id,
            body=user_permission,
            fields="id").execute()

    def __getitem__(self, page):
        if page == 0:
            raise KeyError("Slides are 1-indexed to avoid confusion.")
        content = self.pages[page - 1]
        return Slide(self.doc_id, page, content)


class Slide:
    def __init__(self, doc_id, page, content):
        self.doc_id = doc_id
        self.page = page
        self.content = addict.Dict(content)
        self.slide_id = self.content.objectId
        self.parse()

    def parse(self):
        texts = []
        img = Images()
        charts = Charts()
        for obj in self.content.pageElements:
            obj_id = obj["objectId"]
            if "shape" in obj and "text" in obj.shape:
                text = obj["shape"]["text"]["textElements"]
                for item in text:
                    if "textRun" in item:
                        style = item["textRun"]["style"]
                        content = item["textRun"]["content"].strip()
                        if len(content) > 1:
                            texts.append(TextBlock(obj_id, content, style))
            elif "image" in obj:
                transform = obj["transform"]
                size = obj["size"]
                src = obj["image"]["contentUrl"]
                img.add(Image(obj_id, self.doc_id, src, transform, size))
            elif "sheetsChart" in obj:
                transform = obj["transform"]
                size = obj["size"]
                charts.add(
                    Chart(obj_id, self.doc_id, self.slide_id, size, transform)
                )

        self.text = texts
        self.img = img
        self.chart = charts

    def find(self, query):
        _res = TextBlocks(self.doc_id)
        for block in self.text:
            if block.match(query):
                _res.add(block)
                return _res
        return None

    def find_all(self, query):
        res = TextBlocks(self.doc_id)
        for block in self.text:
            if block.match(query):
                res.add(block)
        return res

    def __setitem__(self, query, value):
        blocks = self.find(query)
        blocks.change_text(value)

    def __getitem__(self, query):
        return self.find(query)
