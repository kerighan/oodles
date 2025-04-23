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
    """Represents a text block element in a Google Slides presentation.

    This class provides methods to manipulate text content in Google Slides.
    You can use HTML-like tags to format parts of the text:
    - <b>text</b>: Makes the enclosed text bold
    - <i>text</i>: Makes the enclosed text italic
    - <a href="https://example.com">text</a>: Creates a hyperlink
    - <a href="https://example.com" color="red">text</a>: Creates a hyperlink with custom color

    Examples:
        text_block.text = "This is <b>bold</b> and <i>italic</i> text"
        slide["Title"] = "Regular and <b>bold</b> and <i>italic</i> text"
        slide["Link"] = "Visit our <a href='https://example.com'>website</a>"
        slide["ColoredLink"] = "Visit our <a href='https://example.com' color='red'>website</a>"
    """

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
        new_text = str(value)

        # Check if there are any formatting tags in the text
        has_formatting = (
            ("<b>" in new_text and "</b>" in new_text)
            or ("<i>" in new_text and "</i>" in new_text)
            or ("<a " in new_text and "</a>" in new_text)
        )

        if has_formatting:
            # Extract the parts to be formatted
            text_parts = []
            current_index = 0
            bold_ranges = []
            italic_ranges = []
            link_ranges = []

            # Stack to keep track of open tags
            bold_start_indices = []
            italic_start_indices = []
            link_data = []  # Store tuples of (start_index, url, color)

            while current_index < len(new_text):
                if new_text[current_index : current_index + 3] == "<b>":
                    # Start of bold text
                    bold_start_indices.append(len("".join(text_parts)))
                    current_index += 3  # Skip the <b> tag
                elif new_text[current_index : current_index + 4] == "</b>":
                    # End of bold text
                    if bold_start_indices:  # Check if we have a matching start tag
                        start_index = bold_start_indices.pop()
                        end_index = len("".join(text_parts))
                        bold_ranges.append((start_index, end_index))
                    current_index += 4  # Skip the </b> tag
                elif new_text[current_index : current_index + 3] == "<i>":
                    # Start of italic text
                    italic_start_indices.append(len("".join(text_parts)))
                    current_index += 3  # Skip the <i> tag
                elif new_text[current_index : current_index + 4] == "</i>":
                    # End of italic text
                    if italic_start_indices:  # Check if we have a matching start tag
                        start_index = italic_start_indices.pop()
                        end_index = len("".join(text_parts))
                        italic_ranges.append((start_index, end_index))
                    current_index += 4  # Skip the </i> tag
                elif new_text[current_index : current_index + 2] == "<a":
                    # Start of link text - find the href attribute
                    href_start = new_text.find("href=", current_index)
                    if href_start != -1:
                        # Find the quote character used (single or double)
                        quote_char = new_text[href_start + 5]
                        href_end = new_text.find(quote_char, href_start + 6)
                        if href_end != -1:
                            # Extract the URL
                            url = new_text[href_start + 6 : href_end]

                            # Check for color attribute
                            color = None
                            color_start = new_text.find("color=", current_index)
                            if color_start != -1 and color_start < new_text.find(
                                ">", href_end
                            ):
                                # Find the quote character used (single or double)
                                color_quote_char = new_text[color_start + 6]
                                color_end = new_text.find(
                                    color_quote_char, color_start + 7
                                )
                                if color_end != -1:
                                    # Extract the color
                                    color = new_text[color_start + 7 : color_end]

                            # Find the end of the opening tag
                            tag_end = new_text.find(
                                ">", max(href_end, color_end if color else href_end)
                            )
                            if tag_end != -1:
                                # Record the start position, URL, and color
                                link_start = len("".join(text_parts))
                                link_data.append((link_start, url, color))

                                # Move past the opening tag
                                current_index = tag_end + 1
                            else:
                                # Malformed tag, treat as regular text
                                text_parts.append(new_text[current_index])
                                current_index += 1
                        else:
                            # Malformed href, treat as regular text
                            text_parts.append(new_text[current_index])
                            current_index += 1
                    else:
                        # No href found, treat as regular text
                        text_parts.append(new_text[current_index])
                        current_index += 1
                elif new_text[current_index : current_index + 4] == "</a>":
                    # End of link text
                    if link_data:  # Check if we have a matching start tag with URL
                        start_index, url, color = link_data.pop()
                        end_index = len("".join(text_parts))
                        link_ranges.append((start_index, end_index, url, color))
                    current_index += 4  # Skip the </a> tag
                else:
                    # Regular text
                    text_parts.append(new_text[current_index])
                    current_index += 1

            # Join all text parts without the tags
            clean_text = "".join(text_parts)

            # Create the basic requests to delete and insert text
            requests = [
                {"deleteText": {"objectId": self.obj_id}},
                {"insertText": {"text": clean_text, "objectId": self.obj_id}},
            ]

            # First apply the base style to the entire text
            requests.append(
                {
                    "updateTextStyle": {
                        "style": self.style,
                        "objectId": self.obj_id,
                        "fields": "*",
                    }
                }
            )

            # Then add style update requests for each bold range
            for start_index, end_index in bold_ranges:
                requests.append(
                    {
                        "updateTextStyle": {
                            "objectId": self.obj_id,
                            "textRange": {
                                "type": "FIXED_RANGE",
                                "startIndex": start_index,
                                "endIndex": end_index,
                            },
                            "style": {"bold": True},
                            "fields": "bold",
                        }
                    }
                )

            # Add style update requests for each italic range
            for start_index, end_index in italic_ranges:
                requests.append(
                    {
                        "updateTextStyle": {
                            "objectId": self.obj_id,
                            "textRange": {
                                "type": "FIXED_RANGE",
                                "startIndex": start_index,
                                "endIndex": end_index,
                            },
                            "style": {"italic": True},
                            "fields": "italic",
                        }
                    }
                )

            # Add link update requests for each link range
            for start_index, end_index, url, color in link_ranges:
                # Create the style with link
                link_style = {"link": {"url": url}}
                fields = "link"

                # Add color if specified
                if color:
                    # Handle color names
                    if color.lower() == "red":
                        rgb = {"red": 1.0, "green": 0.0, "blue": 0.0}
                    elif color.lower() == "green":
                        rgb = {"red": 0.0, "green": 0.8, "blue": 0.0}
                    elif color.lower() == "blue":
                        rgb = {"red": 0.0, "green": 0.0, "blue": 1.0}
                    elif color.lower() == "black":
                        rgb = {"red": 0.0, "green": 0.0, "blue": 0.0}
                    elif color.lower() == "white":
                        rgb = {"red": 1.0, "green": 1.0, "blue": 1.0}
                    elif color.lower() == "yellow":
                        rgb = {"red": 1.0, "green": 1.0, "blue": 0.0}
                    elif color.lower() == "purple":
                        rgb = {"red": 0.5, "green": 0.0, "blue": 0.5}
                    elif color.lower() == "orange":
                        rgb = {"red": 1.0, "green": 0.65, "blue": 0.0}
                    elif color.lower() == "gray" or color.lower() == "grey":
                        rgb = {"red": 0.5, "green": 0.5, "blue": 0.5}
                    # Handle hex colors
                    elif color.startswith("#") and len(color) == 7:
                        try:
                            r = int(color[1:3], 16) / 255.0
                            g = int(color[3:5], 16) / 255.0
                            b = int(color[5:7], 16) / 255.0
                            rgb = {"red": r, "green": g, "blue": b}
                        except ValueError:
                            # Invalid hex color, use default
                            rgb = None
                    else:
                        # Unknown color, don't apply color
                        rgb = None

                    if rgb:
                        link_style["foregroundColor"] = {
                            "opaqueColor": {"rgbColor": rgb}
                        }
                        fields += ",foregroundColor"

                # Add the request
                requests.append(
                    {
                        "updateTextStyle": {
                            "objectId": self.obj_id,
                            "textRange": {
                                "type": "FIXED_RANGE",
                                "startIndex": start_index,
                                "endIndex": end_index,
                            },
                            "style": link_style,
                            "fields": fields,
                        }
                    }
                )

            return requests
        else:
            # No formatting tags, use the original implementation
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

        # Use the _fill method which now handles bold and italic tags
        return self._fill(new_text)

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
