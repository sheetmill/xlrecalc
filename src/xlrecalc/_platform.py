"""Which Excel computes the workbook, and the viewer's regional settings (``platform=``).

The platform is ``"mac"`` (Excel for Mac) or ``"windows"`` (Excel for Windows), given by name or
as an ``xlformat.Platform``, whose name is what counts here. The regional settings are
``xlformat``'s (``xlformat.RegionalSettings``, en-US only); by default the platform's en-US
preset. What depends on them (the code page of ``CHAR`` and ``CODE``, the order of text, the
names Excel knows, ``TEXT``'s output) is decided from the two together, inside the package.
"""

from __future__ import annotations

from typing import Literal

import xlformat

PlatformName = Literal["mac", "windows"]

DEFAULT_PLATFORM: PlatformName = "mac"
"""What ``platform=None`` means."""


def resolve(
    platform: object, regional_settings: object
) -> tuple[PlatformName, xlformat.RegionalSettings]:
    """The platform's name and the regional settings a call means.

    Takes what the caller gave: ``str | xlformat.Platform | None`` and
    ``xlformat.RegionalSettings | None``.

    Raises:
        ValueError: ``platform`` is not ``"mac"``, ``"windows"`` or an ``xlformat.Platform``.
        TypeError: ``regional_settings`` is not an ``xlformat.RegionalSettings``.
    """
    if platform is None:
        name: str = DEFAULT_PLATFORM
    elif isinstance(platform, xlformat.Platform):
        name = platform.name
    elif isinstance(platform, str):
        name = platform
    else:
        msg = f"platform must be 'mac', 'windows' or an xlformat.Platform, not {platform!r}"
        raise ValueError(msg)
    if name not in ("mac", "windows"):
        msg = f"platform must be 'mac', 'windows' or an xlformat.Platform, not {platform!r}"
        raise ValueError(msg)
    if regional_settings is None:
        regional_settings = (
            xlformat.RegionalSettings.EN_US_MAC
            if name == "mac"
            else xlformat.RegionalSettings.EN_US_WINDOWS
        )
    elif not isinstance(regional_settings, xlformat.RegionalSettings):
        msg = f"regional_settings must be an xlformat.RegionalSettings, not {regional_settings!r}"
        raise TypeError(msg)
    return name, regional_settings  # type: ignore[return-value]  # checked above
