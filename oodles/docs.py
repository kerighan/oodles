#!/usr/bin/env python
# -*- coding: utf-8 -*-
from googleapiclient.errors import HttpError

from .config import config


class Document:
    def __init__(self, doc_id):
        self.doc_id = doc_id
        self.url = "https://docs.google.com/document/d/" + self.doc_id
        self.load()

    @staticmethod
    def create(title):
        body = {
            "title": title,
            'mimeType': 'application/vnd.google-apps.document'
        }
        document = config.DRIVE.files().create(
            body=body, fields="id").execute()
        return Document(document.get("id"))

    def share_with(self, email):
        user_permission = {
            "type": "user",
            "role": "writer",
            "emailAddress": email
        }
        config.DRIVE.permissions().create(
            fileId=self.doc_id,
            body=user_permission,
            fields="id").execute()

    def load(self):
        try:
            result = config.DOCS.documents().get(
                documentId=self.doc_id).execute()
            self.document = result
        except HttpError:
            print("remember to share your document with "
                  f"{config.service_email}\n")
            raise HttpError
        self.title = self.document["title"]

    def empty_document(self):
        try:
            # Fetch the document again to get the latest endIndex
            self.load()

            # Empty the document and execute the request
            end_index = self.document['body']['content'][-1]['endIndex'] - 1
            requests = [{
                'deleteContentRange': {
                    'range': {
                        'startIndex': 1,
                        'endIndex': end_index,
                    }
                }
            }]
            config.DOCS.documents().batchUpdate(
                documentId=self.doc_id, body={'requests': requests}).execute()

            # Fetch the document again to get the latest endIndex
            self.load()
        except HttpError:
            pass

    def set_title(self, new_title):
        body = {"name": new_title}
        request = config.DRIVE.files().update(fileId=self.doc_id, body=body)
        request.execute()
        self.title = new_title

    def set_content(self, text_list):
        """Set the content of the document."""
        # Empty the document
        self.empty_document()

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
        result = config.DOCS.documents().batchUpdate(
            documentId=self.doc_id, body={'requests': requests}).execute()
        return result
