import os.path

import addict
from googleapiclient import errors
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
        try:
            new_id = (
                config.DRIVE.files().copy(body=data, fileId=self.doc_id).execute().get("id")
            )
        except errors.HttpError as e:
            print("This error can happen if the file is in a shared drive")
            raise e
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

    @classmethod
    def list_all(cls):
        """List all Google Slides presentations."""
        presentations = []
        page_token = None
        slides_mime_type = 'application/vnd.google-apps.presentation'

        while True:
            try:
                results = config.DRIVE.files().list(
                    q=f"mimeType='{slides_mime_type}'",
                    pageSize=100,
                    fields="nextPageToken, files(id, name, modifiedTime, createdTime, webViewLink)",
                    pageToken=page_token
                ).execute()

                items = results.get('files', [])
                # Convert to Slides objects if desired
                for item in items:
                    presentations.append({
                        'id': item['id'],
                        'name': item['name'],
                        'url': item.get('webViewLink'),
                        'modified': item.get('modifiedTime'),
                        'created': item.get('createdTime')
                    })

                page_token = results.get('nextPageToken', None)
                if page_token is None:
                    break

            except HttpError as error:
                raise GoogleAuthorizationError(error, config.service_email)

        return presentations

    @classmethod
    def delete_by_id(cls, doc_id: str):
        """
        Delete a presentation by its ID without instantiating a Slides object.
        WARNING: This permanently deletes the presentation!

        Args:
            doc_id: The ID of the presentation to delete

        Returns:
            bool: True if successful, False if not found
        """
        try:
            config.DRIVE.files().delete(fileId=doc_id).execute()
            print(f"Successfully deleted presentation with ID: {doc_id}")
            return True
        except HttpError as e:
            if e.resp.status == 404:
                raise FileNotFoundError(f"Presentation with ID {doc_id} not found.")
            else:
                raise GoogleAuthorizationError(e, config.service_email)

    def delete(self):
        """
        Delete this presentation from Google Drive.
        WARNING: This permanently deletes the presentation!
        """
        try:
            config.DRIVE.files().delete(fileId=self.doc_id).execute()
            print(f"Successfully deleted presentation: {self.title} (ID: {self.doc_id})")
            # Clear the instance data since it's deleted
            self.document = None
            self.pages = None
        except HttpError as e:
            if e.resp.status == 404:
                print(f"Presentation with ID {self.doc_id} not found.")
            else:
                raise GoogleAuthorizationError(e, config.service_email)

    def get_permissions(self) -> list[dict]:
        """
        Get list of all users/permissions for this presentation.

        Returns:
            list: List of permission objects with user info
        """
        try:
            permissions = config.DRIVE.permissions().list(
                fileId=self.doc_id,
                fields="permissions(id, type, emailAddress, role, displayName)"
            ).execute()

            return permissions.get('permissions', [])
        except HttpError as e:
            raise GoogleAuthorizationError(e, config.service_email)

    def get_shared_users(self):
        """
        Get list of users (excluding owner) who have access to this presentation.

        Returns:
            list: List of dicts with user email and role
        """
        permissions = self.get_permissions()

        shared_users = []
        for perm in permissions:
            if perm.get('role') != 'owner':
                shared_users.append({
                    'email': perm.get('emailAddress', 'Unknown'),
                    'name': perm.get('displayName', 'Unknown'),
                    'role': perm.get('role'),
                    'type': perm.get('type'),
                    'id': perm.get('id')
                })

        return shared_users

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
