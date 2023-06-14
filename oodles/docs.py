#!/usr/bin/env python
# -*- coding: utf-8 -*-
from googleapiclient.errors import HttpError

from .element import SheetChart
from .utils import hex_to_rgb
from .variables import DOCS, DRIVE, service_email


class Document:
    def __init__(self, doc_id):
        self.doc_id = doc_id
        self.url = "https://docs.google.com/document/d/" + self.doc_id
        self.load()

    @staticmethod
    def create(title):
        body = {
            "title": title
        }
        document = DRIVE.files().create(
            body=body, fields="id").execute()
        return Document(document.get("id"))

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

    def load(self):
        try:
            result = DOCS.documents().get(documentId=self.doc_id).execute()
            print(result)
            self.document = result
        except HttpError:
            print(f"remember to share your document with {service_email}\n")
            raise HttpError
        self.title = self.document["title"]

    def change_title(self, new_title):
        body = {"name": new_title}
        request = DRIVE.files().update(fileId=self.doc_id, body=body)
        request.execute()
        self.title = new_title

    def set_content(self, text_list):
        """Update the content of the document. 'text_list' should be a list of strings."""

        # Fetch the document again to get the latest endIndex
        self.load()

        # Empty the document and execute the request
        requests = [{
            'deleteContentRange': {
                'range': {
                    'startIndex': 1,
                    'endIndex': self.document['body']['content'][-1]['endIndex'] - 1,
                }
            }
        }]
        result = DOCS.documents().batchUpdate(
            documentId=self.doc_id, body={'requests': requests}).execute()

        # Find the current end of the document
        current_end = self.document['body']['content'][-1]['endIndex']
        current_end = 1

        # Start a list of requests
        requests = []

        # For each string in text_list
        for text in text_list:
            # Check if text is to be bold
            if '<b>' in text and '</b>' in text:
                # Remove HTML tags from text
                clean_text = text.replace('<b>', '').replace('</b>', '')
                clean_text += '\n'

                # Add an insert text request
                requests.append({
                    'insertText': {
                        'location': {
                            'index': current_end,
                        },
                        'text': clean_text
                    }
                })

                # Add a bold request
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': current_end,
                            'endIndex': current_end + len(clean_text),
                        },
                        'textStyle': {
                            'bold': True,
                        },
                        'fields': 'bold',
                    }
                })
                # Update the current end of the document
                current_end += len(clean_text)

            else:
                text += '\n'

                # Add an insert text request
                requests.append({
                    'insertText': {
                        'location': {
                            'index': current_end,
                        },
                        'text': text
                    }
                })
                # Update the current end of the document
                current_end += len(text)

        # Send the batchUpdate request
        result = DOCS.documents().batchUpdate(
            documentId=self.doc_id, body={'requests': requests}).execute()
        return result
