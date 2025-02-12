import os.path

import addict
import requests
from google.auth.transport.requests import AuthorizedSession
from googleapiclient.errors import HttpError

from .element import Chart, Charts, Image, Images, TextBlock, TextBlocks
from .config import config
from .utils import GoogleAuthorizationError


class Slides:
    def __init__(self, doc_id: str):
        self.doc_id = doc_id
        try:
            self.document = (
                config.SLIDES.presentations().get(presentationId=doc_id).execute()
            )
        except HttpError as e:
            raise GoogleAuthorizationError(e, config.service_email)
        self.pages = self.document.get("slides")
        self.title = self.document["title"]
        self.url = "https://docs.google.com/presentation/d/" + self.doc_id

    def copy(self, title: str = None):
        if title is None:
            title = f"Copy of {self.title}"
        data = {"name": title}
        new_id = (
            config.DRIVE.files().copy(body=data, fileId=self.doc_id).execute().get("id")
        )
        return Slides(new_id)

    def share_with(self, email: str):
        user_permission = {"type": "user", "role": "writer", "emailAddress": email}
        config.DRIVE.permissions().create(
            fileId=self.doc_id, body=user_permission, fields="id"
        ).execute()

    def __getitem__(self, page: int) -> "Slide":
        if page == 0:
            raise KeyError("Slides are 1-indexed to avoid confusion.")
        content = self.pages[page - 1]
        return Slide(self.doc_id, page, content)

    def screenshot(self, slide_num: int, path: str):
        """
        Takes a screenshot of the n-th slide and save it as png in the provided path.
        :param slide_num: Number of the slide, starts at 1
        :param path: Where to save the file
        :return:
        """
        self[slide_num].screenshot(path)


class BatchUpdate:
    def __init__(self, slide):
        self.slide = slide
        self.requests = []

    def __enter__(self):
        self.slide._batch_requests = self.requests
        # Make sure to propagate to img
        if hasattr(self.slide, "img"):
            self.slide.img._batch_requests = self.requests
            # Also propagate to each image
            for img in self.slide.img.elements:
                img._batch_requests = self.requests
        return self.slide

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.requests:  # Only make the API call if we have requests
            config.SLIDES.presentations().batchUpdate(
                body={"requests": self.requests}, presentationId=self.slide.doc_id
            ).execute()
        self.slide._batch_requests = None
        if hasattr(self.slide, "img"):
            self.slide.img._batch_requests = None
            for img in self.slide.img.elements:
                img._batch_requests = None


class Slide:
    def __init__(self, doc_id: str, page: int, content: dict):
        self.doc_id = doc_id
        self.page = page
        self.content = addict.Dict(content)
        self.slide_id = self.content.objectId
        self._batch_requests = None  # Will store requests during batch update
        self.parse()

    def parse(self):
        texts = []
        imgs = Images()
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
                # Extract title from alt text if available
                title = None
                if "description" in obj:
                    title = obj["description"]
                imgs.add(Image(obj_id, self.doc_id, src, transform, size, title))
            elif "sheetsChart" in obj:
                transform = obj["transform"]
                size = obj["size"]
                charts.add(Chart(obj_id, self.doc_id, self.slide_id, size, transform))

        self.text = texts
        self.img = imgs
        self.chart = charts

    def find(self, query: str):
        _res = TextBlocks(self.doc_id)
        for block in self.text:
            if block.match(query):
                _res.add(block)
                return _res
        return None

    def find_all(self, query: str):
        res = TextBlocks(self.doc_id)
        for block in self.text:
            if block.match(query):
                res.add(block)
        return res

    def __setitem__(self, query, value):
        blocks = self.find(query)
        if blocks is not None:
            blocks._batch_requests = self._batch_requests
            blocks.change_text(value)

    def __getitem__(self, query):
        return self.find(query)

    def __setattr__(self, name: str, value):
        if name == "_batch_requests":
            super(Slide, self).__setattr__(name, value)
            # Propagate batch requests to images and text blocks
            if hasattr(self, "img"):
                self.img._batch_requests = value
            if hasattr(self, "text"):
                for block in self.text:
                    block._batch_requests = value
        else:
            super(Slide, self).__setattr__(name, value)

    def batch_update(self):
        """Context manager for batching multiple updates into a single API call"""
        return BatchUpdate(self)

    def screenshot(self, path: str):
        """
        Takes a screenshot of this slide and saves it as a PNG file at the provided path.
        :param path: The file path or directory where to save the screenshot.
        """
        if not path.endswith(".png"):
            if not os.path.exists(path):
                raise ValueError(
                    "Provided path is not a PNG file nor an existing folder."
                )
            path = os.path.join(path, f"slide_{self.page}.png")

        http = config.SLIDES._http
        credentials = http.credentials
        authed_session = AuthorizedSession(credentials)

        url = f"https://slides.googleapis.com/v1/presentations/{self.doc_id}/pages/{self.slide_id}/thumbnail"
        params = {
            "thumbnailProperties.mimeType": "PNG",  # 'PNG' or 'JPEG'
            "thumbnailProperties.thumbnailSize": "LARGE",  # 'LARGE', 'MEDIUM', 'SMALL'
        }

        response = authed_session.get(url, params=params)
        if response.status_code == 200:
            thumbnail = response.json()
            thumbnail_url = thumbnail.get("contentUrl")
            if not thumbnail_url:
                raise Exception("Thumbnail URL not found in the response.")

            image_response = requests.get(thumbnail_url)
            if image_response.status_code == 200:
                with open(path, "wb") as f:
                    f.write(image_response.content)
            else:
                raise Exception(
                    f"Failed to download thumbnail image. Status code: {image_response.status_code}"
                )
        else:
            raise Exception(
                f"Failed to get thumbnail. Status code: {response.status_code}, Response: {response.text}"
            )
