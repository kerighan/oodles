from .variables import (
    DRIVE, SLIDES, SHEETS, BUCKET,
    GOOGLE_APPLICATION_CREDENTIALS, service_email)
from googleapiclient.http import MediaFileUpload
from subprocess import check_output


class TextBlocks:
    elements = []
    obj_ids = set()

    def __init__(self, doc_id):
        self.doc_id = doc_id
    
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
    
    def __setattr__(self, name, value):
        if name == "text":
            requests = []
            for element in self.elements:
                requests += element._fill(value)
            SLIDES.presentations().batchUpdate(
                        body={"requests": requests},
                        presentationId=self.doc_id).execute()
        else:
            super(TextBlocks, self).__setattr__(name, value)

    
    def replace(self, key, value):
        requests = []
        for element in self.elements:
            requests += element._replace(key, value)

        SLIDES.presentations().batchUpdate(
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


class Image:
    def __init__(self, obj_id, doc_id, src, transform, size):
        self.obj_id = obj_id
        self.doc_id = doc_id
        self.src = src
        self.transform = transform
        self.size = size
    
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
            SLIDES.presentations().batchUpdate(
                body={"requests": requests},
                presentationId=self.doc_id).execute()

        elif name == "file":
            filename = value.split("/")[-1]
            check_output(
                f"gsutil -q cp {value} gs://data-studies/img/{filename}",
                shell=True)
            url = check_output(
                f"gsutil -q signurl -d 5m "
                f"-m GET {GOOGLE_APPLICATION_CREDENTIALS} "
                f"{BUCKET}/{filename}",
                shell=True)
            url = "https://" + str(url, "utf8").split("https://")[-1]
            self.url = url
        else:
            super(Image, self).__setattr__(name, value)

    def __repr__(self):
        return self.src


class SheetChart:
    def __init__(self, spreadsheet_id, chart_id):
        self.spreadsheet_id = spreadsheet_id
        self.chart_id = chart_id
    
    def __repr__(self):
        return f"{self.chart_id}"


class Chart:
    def __init__(self, obj_id, doc_id, slide_id, size, transform):
        self.obj_id = obj_id
        self.slide_id = slide_id
        self.doc_id = doc_id
        self.size = size
        self.transform = transform

    def replace(self, sheetchart):
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
        SLIDES.presentations().batchUpdate(
            body={"requests": requests},
            presentationId=self.doc_id).execute()

    def __repr__(self):
        return f"{self.obj_id}"
