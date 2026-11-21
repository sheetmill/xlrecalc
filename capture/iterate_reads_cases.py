"""Capture ``iterate_reads``'s file form, read by its probe and by ``tests/test_iteration.py``.

Kept apart from the probe so that reading it imports neither openpyxl nor Excel's tooling.
"""

from __future__ import annotations

# every formula cached 0 ("cached_0"), with fullCalcOnLoad: Excel's opening run
FORMS = {"cached_0": '<calcPr calcId="124519" iterate="1" {} fullCalcOnLoad="1"/>'}
