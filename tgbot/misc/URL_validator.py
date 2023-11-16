import validators


def is_string_an_url(url_string: str) -> bool:
    """Проверяет является ли строка ссылкой."""
    return validators.url(url_string)


