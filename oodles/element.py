from .config import config
from datetime import datetime, timedelta


# =============================================================================
# Text management - GOOGLE SLIDES SIDE
# =============================================================================

class TextBlocks(object):
    def __init__(self, doc_id):
        self.doc_id = doc_id
        self.elements = []
        self.obj_ids = set()

    def add(self, block):
        self.elements.append(block)
        self.obj_ids.add(block.obj_id)

    def __repr__(self):
        return self.elements.__repr__()

    def __iadd__(self, block):
        if isinstance(block, TextBlock):
            self.add(block)
        elif isinstance(block, TextBlocks):
            for element in block.elements:
                if element.obj_id in self.obj_ids:
                    continue
                self.elements.append(block)
                self.obj_ids.add(element.obj_id)
        return self

    def change_text(self, value):
        requests = []
        for element in self.elements:
            requests += element._fill(value)
        config.SLIDES.presentations().batchUpdate(
            body={"requests": requests},
            presentationId=self.doc_id).execute()

    def __setattr__(self, name, value):
        if name == "text":
            self.change_text(value)
        else:
            super(TextBlocks, self).__setattr__(name, value)

    def __assign__(self, value):
        self.change_text(value)

    def replace(self, key, value):
        requests = []
        for element in self.elements:
            requests += element._replace(key, value)

        config.SLIDES.presentations().batchUpdate(
            body={"requests": requests},
            presentationId=self.doc_id).execute()


class TextBlock:
    def __init__(self, obj_id, text, style={}):
        self.obj_id = obj_id
        self.text = text
        self.style = style

    def match(self, query):
        return query in self.text

    def __repr__(self):
        return f"{self.text}"

    def _fill(self, value):
        new_text = value
        return [
            {"deleteText": {"objectId": self.obj_id}},
            {
                "insertText": {
                    "text": new_text,
                    "objectId": self.obj_id
                }
            },
            {"updateTextStyle": {
                "style": self.style,
                "objectId": self.obj_id,
                "fields": "*"
            }}
        ]

    def _replace(self, key, value):
        new_text = self.text.replace(key, value)
        return [
            {"deleteText": {"objectId": self.obj_id}},
            {
                "insertText": {
                    "text": new_text,
                    "objectId": self.obj_id
                }
            },
            {
                "updateTextStyle": {
                    "style": self.style,
                    "objectId": self.obj_id,
                    "fields": "*"
                }
            }
        ]


# =============================================================================
# Images management - GOOGLE SLIDES SIDE
# =============================================================================

class Images:
    def __init__(self):
        self.elements = []

    def add(self, block):
        self.elements.append(block)

    def __setitem__(self, key, value):
        self.elements[key].replace_image(value)


class Image:
    def __init__(self, obj_id, doc_id, src, transform, size):
        self.obj_id = obj_id
        self.doc_id = doc_id
        self.src = src
        self.transform = transform
        self.size = size

    def replace_image(self, value):
        if value[:4] == "http":
            self.url = value
        else:
            self.file = value

    def __setattr__(self, name, value):
        if name == "url":
            requests = [
                {
                    'replaceImage': {
                        'imageObjectId': self.obj_id,
                        'imageReplaceMethod': 'CENTER_INSIDE',
                        'url': value
                    }
                }
            ]
            config.SLIDES.presentations().batchUpdate(
                body={"requests": requests},
                presentationId=self.doc_id).execute()

        elif name == "file":
            filename = value.split("/")[-1]
            expire_in = datetime.today() + timedelta(hours=1)
            # use google cloud storage client
            blob = config.STORAGE_CLIENT.bucket("data-studies").blob(
                f"img/{filename}")
            blob.upload_from_filename(value)
            url = blob.generate_signed_url(expire_in)
            self.url = url
        else:
            super(Image, self).__setattr__(name, value)

    def __repr__(self):
        return f"<{self.src}>"


# =============================================================================
# Chart management - GOOGLE SLIDES SIDE
# =============================================================================

class Charts:
    def __init__(self):
        self.elements = []

    def add(self, chart):
        self.elements.append(chart)

    def __setitem__(self, key, value):
        self.elements[key].replace_chart(value)

    def __getitem__(self, key):
        return self.elements[key]


class Chart:
    def __init__(self, obj_id, doc_id, slide_id, size, transform):
        self.obj_id = obj_id
        self.slide_id = slide_id
        self.doc_id = doc_id
        self.size = size
        self.transform = transform

    def replace_chart(self, sheetchart):
        ss_id = sheetchart.spreadsheet_id
        chart_id = sheetchart.chart_id
        requests = [
            {
                "deleteObject": {
                    "objectId": self.obj_id
                }
            },
            {
                "createSheetsChart": {
                    "spreadsheetId": ss_id,
                    "chartId": chart_id,
                    "objectId": self.obj_id,
                    'linkingMode': 'LINKED',
                    "elementProperties": {
                        'size': self.size,
                        'transform': self.transform,
                        "pageObjectId": self.slide_id
                    }
                }
            },
            {
                'refreshSheetsChart': {
                    'objectId': self.obj_id
                }
            }
        ]
        config.SLIDES.presentations().batchUpdate(
            body={"requests": requests},
            presentationId=self.doc_id).execute()

    def __repr__(self):
        return f"{self.obj_id}"


# =============================================================================
# Chart management - GOOGLE SHEETS SIDE
# =============================================================================

class SheetChart:
    def __init__(self, spreadsheet_id, chart_id):
        self.spreadsheet_id = spreadsheet_id
        self.chart_id = chart_id

    def __repr__(self):
        return f"{self.chart_id}"
