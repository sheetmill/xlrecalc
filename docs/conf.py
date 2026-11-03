"""Sphinx configuration for the xlrecalc documentation.

Build: ``uv run --group docs sphinx-build -W docs docs/_build/html``, then
``npx --yes pagefind@1.5.2 --site docs/_build/html`` for the search index (``pagefind.yml``).
"""

from __future__ import annotations

import xlrecalc

project = "xlrecalc"
author = "Ethan Post"
copyright = "2026, Ethan Post"  # noqa: A001 - the name Sphinx reads
release = xlrecalc.__version__
version = ".".join(release.split(".")[:2])

extensions = [
    "myst_parser",
    "sphinx.ext.autodoc",
    "sphinx.ext.napoleon",
    "sphinx_autodoc_typehints",
    "sphinx.ext.intersphinx",
    "sphinx.ext.viewcode",
]

source_suffix = {".md": "markdown", ".rst": "restructuredtext"}
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]
templates_path = ["_templates"]

myst_enable_extensions = ["colon_fence", "deflist"]
myst_heading_anchors = 3

napoleon_google_docstring = True
napoleon_numpy_docstring = False
napoleon_use_rtype = False
autodoc_member_order = "bysource"
autodoc_typehints = "description"
autodoc_class_signature = "separated"
autodoc_default_options = {"exclude-members": "__init__, __new__"}
napoleon_use_ivar = True
typehints_use_signature = False
always_document_param_types = False
always_use_bars_union = True

intersphinx_mapping = {
    "python": ("https://docs.python.org/3", None),
    "xlformat": ("https://xlformat.readthedocs.io/en/stable/", None),
}

html_theme = "furo"
html_title = f"xlrecalc {release}"
html_theme_options = {
    "source_repository": "https://github.com/sheetmill/xlrecalc/",
    "source_branch": "main",
    "source_directory": "docs/",
}
