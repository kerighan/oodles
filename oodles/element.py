from typing import Union

from .config import config
from datetime import datetime, timedelta


# =============================================================================
# Text management - GOOGLE SLIDES SIDE
# =============================================================================


class TextBlocks(object):
    def __init__(self, doc_id: str):
        self.doc_id = doc_id
        self.elements = []
        self.obj_ids = set()
        self._batch_requests = None

    def add(self, block: "TextBlock"):
        self.elements.append(block)
        self.obj_ids.add(block.obj_id)

    def __repr__(self):
        return self.elements.__repr__()

    def __iadd__(self, block: Union["TextBlock", "TextBlocks"]):
        if isinstance(block, TextBlock):
            self.add(block)
        elif isinstance(block, TextBlocks):
            for element in block.elements:
                if element.obj_id in self.obj_ids:
                    continue
                self.elements.append(block)
                self.obj_ids.add(element.obj_id)
        return self

    def _execute_requests(self, requests):
        """Execute requests immediately or add to batch if in batch mode"""
        if hasattr(self, "_batch_requests") and self._batch_requests is not None:
            self._batch_requests.extend(requests)
        else:
            config.SLIDES.presentations().batchUpdate(
                body={"requests": requests}, presentationId=self.doc_id
            ).execute()

    def change_text(self, value: str):
        requests = []
        for element in self.elements:
            requests += element._fill(value)
        self._execute_requests(requests)

    def __setattr__(self, name: str, value: str):
        if name == "text":
            self.change_text(value)
        elif name == "_batch_requests":
            super(TextBlocks, self).__setattr__(name, value)
            # Propagate batch requests to all text blocks
            if hasattr(self, "elements"):
                for block in self.elements:
                    block._batch_requests = value
        else:
            super(TextBlocks, self).__setattr__(name, value)

    def __assign__(self, value: str):
        self.change_text(value)

    def replace(self, key: str, value: str):
        requests = []
        for element in self.elements:
            requests += element._replace(key, value)
        self._execute_requests(requests)


class TextBlock:
    def __init__(self, obj_id: str, text: str, style: dict = None):
        if style is None:
            style = {}
        self.obj_id = obj_id
        self.text = text
        self.style = style
        self._batch_requests = None

    def match(self, query: str) -> bool:
        return query in self.text

    def __repr__(self) -> str:
        return f"{self.text}"

    def _fill(self, value: str) -> list:
        new_text = value
        return [
            {"deleteText": {"objectId": self.obj_id}},
            {"insertText": {"text": new_text, "objectId": self.obj_id}},
            {
                "updateTextStyle": {
                    "style": self.style,
                    "objectId": self.obj_id,
                    "fields": "*",
                }
            },
        ]

    def _replace(self, key: str, value: str) -> list:
        new_text = self.text.replace(key, value)
        return [
            {"deleteText": {"objectId": self.obj_id}},
            {"insertText": {"text": new_text, "objectId": self.obj_id}},
            {
                "updateTextStyle": {
                    "style": self.style,
                    "objectId": self.obj_id,
                    "fields": "*",
                }
            },
        ]

    def _execute_requests(self, requests):
        """Execute requests immediately or add to batch if in batch mode"""
        if hasattr(self, "_batch_requests") and self._batch_requests is not None:
            self._batch_requests.extend(requests)
        else:
            config.SLIDES.presentations().batchUpdate(
                body={"requests": requests}, presentationId=self.doc_id
            ).execute()


# =============================================================================
# Images management - GOOGLE SLIDES SIDE
# =============================================================================


class Images:
    def __init__(self):
        self.elements = []
        self.title_map = {}  # Map titles to indices
        self._batch_requests = None

    def add(self, block: "Image"):
        self.elements.append(block)
        if block.title:  # If image has a title, add it to the map
            self.title_map[str(block.title)] = len(self.elements) - 1

    def __setattr__(self, name: str, value):
        if name == "_batch_requests":
            super(Images, self).__setattr__(name, value)
            # Propagate batch requests to all images
            if hasattr(self, "elements"):
                for img in self.elements:
                    img._batch_requests = value
        else:
            super(Images, self).__setattr__(name, value)

    def __setitem__(self, key, value: str):
        if isinstance(key, str):
            # If key is a string, treat it as a title
            if key not in self.title_map:
                raise KeyError(f"No image found with title '{key}'")
            self.elements[self.title_map[key]].replace_image(value)
        else:
            # Otherwise use the original index-based behavior
            self.elements[key].replace_image(value)

    def __len__(self):
        return len(self.elements)

    def __getitem__(self, key):
        if isinstance(key, str):
            # If key is a string, treat it as a title
            if key not in self.title_map:
                raise KeyError(f"No image found with title '{key}'")
            return self.elements[self.title_map[key]]
        return self.elements[key]


class Image:
    def __init__(
        self, obj_id: str, doc_id: str, src: str, transform, size, title: str = None
    ):
        self.obj_id = obj_id
        self.doc_id = doc_id
        self.src = src
        self.transform = transform
        self.size = size
        self.title = title
        self._batch_requests = None

    def _execute_requests(self, requests):
        """Execute requests immediately or add to batch if in batch mode"""
        if hasattr(self, "_batch_requests") and self._batch_requests is not None:
            self._batch_requests.extend(requests)
        else:
            config.SLIDES.presentations().batchUpdate(
                body={"requests": requests}, presentationId=self.doc_id
            ).execute()

    def replace_image(self, value):
        requests = []
        if isinstance(value, tuple):
            # Handle (image, link) tuple
            image_value, link_url = value
            # First collect the image replacement request
            if image_value.startswith("http"):
                requests.append(
                    {
                        "replaceImage": {
                            "imageObjectId": self.obj_id,
                            "imageReplaceMethod": "CENTER_INSIDE",
                            "url": image_value,
                        }
                    }
                )
            else:
                filename = image_value.split("/")[-1]
                expire_in = datetime.now() + timedelta(minutes=10)
                # use google cloud storage client
                bucket = config.BUCKET.removeprefix("gs://").split("/")[0]
                blob = "/".join(config.BUCKET.removeprefix("gs://").split("/")[1:])
                blob = config.STORAGE_CLIENT.bucket(bucket).blob(blob + "/" + filename)
                blob.upload_from_filename(image_value)
                url = blob.generate_signed_url(expire_in)
                requests.append(
                    {
                        "replaceImage": {
                            "imageObjectId": self.obj_id,
                            "imageReplaceMethod": "CENTER_INSIDE",
                            "url": url,
                        }
                    }
                )

            # Then add the link request
            requests.append(
                {
                    "updateImageProperties": {
                        "objectId": self.obj_id,
                        "imageProperties": {
                            "link": {"url": link_url} if link_url else None
                        },
                        "fields": "link",
                    }
                }
            )
        else:
            # Handle single image replacement
            if value.startswith("http"):
                requests.append(
                    {
                        "replaceImage": {
                            "imageObjectId": self.obj_id,
                            "imageReplaceMethod": "CENTER_INSIDE",
                            "url": value,
                        }
                    }
                )
            else:
                filename = value.split("/")[-1]
                expire_in = datetime.now() + timedelta(minutes=10)
                # use google cloud storage client
                bucket = config.BUCKET.removeprefix("gs://").split("/")[0]
                blob = "/".join(config.BUCKET.removeprefix("gs://").split("/")[1:])
                blob = config.STORAGE_CLIENT.bucket(bucket).blob(blob + "/" + filename)
                blob.upload_from_filename(value)
                url = blob.generate_signed_url(expire_in)
                requests.append(
                    {
                        "replaceImage": {
                            "imageObjectId": self.obj_id,
                            "imageReplaceMethod": "CENTER_INSIDE",
                            "url": url,
                        }
                    }
                )

        # Execute all requests
        self._execute_requests(requests)

    def __setattr__(self, name: str, value: str):
        if name == "url":
            requests = [
                {
                    "replaceImage": {
                        "imageObjectId": self.obj_id,
                        "imageReplaceMethod": "CENTER_INSIDE",
                        "url": value,
                    }
                }
            ]
            self._execute_requests(requests)
        elif name == "file":
            filename = value.split("/")[-1]
            expire_in = datetime.now() + timedelta(minutes=10)
            # use google cloud storage client
            bucket = config.BUCKET.removeprefix("gs://").split("/")[0]
            blob = "/".join(config.BUCKET.removeprefix("gs://").split("/")[1:])
            blob = config.STORAGE_CLIENT.bucket(bucket).blob(blob + "/" + filename)
            blob.upload_from_filename(value)
            url = blob.generate_signed_url(expire_in)
            self.url = url
        else:
            super(Image, self).__setattr__(name, value)

    def __repr__(self):
        return f"{self.obj_id}"


# =============================================================================
# Chart management - GOOGLE SLIDES SIDE
# =============================================================================


class Charts:
    def __init__(self):
        self.elements = []

    def add(self, chart: "Chart"):
        self.elements.append(chart)

    def __setitem__(self, key: int, value: str):
        self.elements[key].replace_chart(value)

    def __getitem__(self, key):
        return self.elements[key]


class Chart:
    def __init__(self, obj_id: str, doc_id: str, slide_id: str, size, transform):
        self.obj_id = obj_id
        self.slide_id = slide_id
        self.doc_id = doc_id
        self.size = size
        self.transform = transform

    def replace_chart(self, sheet_chart: "SheetChart"):
        ss_id = sheet_chart.spreadsheet_id
        chart_id = sheet_chart.chart_id
        requests = [
            {"deleteObject": {"objectId": self.obj_id}},
            {
                "createSheetsChart": {
                    "spreadsheetId": ss_id,
                    "chartId": chart_id,
                    "objectId": self.obj_id,
                    "linkingMode": "LINKED",
                    "elementProperties": {
                        "size": self.size,
                        "transform": self.transform,
                        "pageObjectId": self.slide_id,
                    },
                }
            },
            {"refreshSheetsChart": {"objectId": self.obj_id}},
        ]
        config.SLIDES.presentations().batchUpdate(
            body={"requests": requests}, presentationId=self.doc_id
        ).execute()

    def __repr__(self) -> str:
        return f"{self.obj_id}"


# =============================================================================
# Chart management - GOOGLE SHEETS SIDE
# =============================================================================


class SheetChart:
    def __init__(self, spreadsheet_id: str, chart_id: str):
        self.spreadsheet_id = spreadsheet_id
        self.chart_id = chart_id

    def __repr__(self) -> str:
        return f"{self.chart_id}"
