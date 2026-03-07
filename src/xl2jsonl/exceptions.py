class Xl2JsonlError(Exception):
    pass


class LoaderError(Xl2JsonlError):
    pass


class EmptySheetError(Xl2JsonlError):
    pass


class NoHeaderError(Xl2JsonlError):
    pass
