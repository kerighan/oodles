from googleapiclient.errors import HttpError

def hex_to_rgb(h):
    if h is None:
        return None
    elif isinstance(h, tuple):
        return h
    else:
        h = h.lstrip("#")
        if len(h) == 3:
            h += h
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

class GoogleAuthorizationError(HttpError):
    def __init__(self,error: HttpError,  email: str, *args: object) -> None:
        super().__init__(error.resp, error.content, error.uri)
        self.service_account = email
        self.message = ("Oodles cannot access your document, please share "
            "it with account " + email +
            " and that slides are not in .xlsx format")

    def __str__(self):
        return self.message
