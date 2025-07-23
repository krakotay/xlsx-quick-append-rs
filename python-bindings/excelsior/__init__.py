"""
excelsior: тонкая обёртка над нативным расширением, собранным через maturin.

Файл нужен только для editable-установок (`maturin develop`): без него каталог
`excelsior/` считается namespace-package и Python не подхватывает лежащий
внутри бинарный модуль `excelsior*.pyd`.
"""

from importlib import import_module as _import_module
import sys as _sys

# Импортируем бинарник как подпакет:  excelsior.excelsior
try:
    _ext = _import_module("." + __name__.split(".")[-1], package=__name__)
except ModuleNotFoundError as exc:          # pragma: no cover
    raise ImportError(
        "Не удалось загрузить нативное расширение для 'excelsior'. "
        "Убедитесь, что вы запускали `maturin develop` или установили wheel."
    ) from exc

# Экспортируем все публичные имена
globals().update(
    {k: v for k, v in _ext.__dict__.items() if not k.startswith("_")}
)

# Чтобы `importlib.reload(excelsior)` работал предсказуемо
_sys.modules[__name__] = _ext

# Чистим внутренние ссылки
del _import_module, _sys, _ext
