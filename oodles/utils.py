
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
