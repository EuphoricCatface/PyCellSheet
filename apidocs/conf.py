# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

import os
import sys

try:
    from pycellsheet.settings import VERSION
except ImportError:
    sys.path.insert(0, os.path.abspath('../'))
    from pycellsheet.settings import VERSION

project = 'pycellsheet'
copyright = 'Martin Manns and the pycellsheet team '
author = 'Martin Manns and the pycellsheet team'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.autosummary',
    'sphinx.ext.intersphinx',
    'sphinx.ext.todo',
    'sphinx.ext.viewcode',
    'sphinx.ext.graphviz',
    'sphinx.ext.napoleon',
    'sphinx_autodoc_typehints',
    'recommonmark'
]

templates_path = ['_templates']

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store', "requirements.txt"]

source_suffix = {
    '.rst': 'restructuredtext',
    '.txt': 'markdown',
    '.md': 'markdown',
}

primary_domain = 'py'
highlight_language = 'py'

release = VERSION

# Were autodoc so set stuff aphabetically
autodoc_member_order = "alphabetical"

# Flags for the stuff to document..
autodoc_default_flags = [
    'members',
    'private-members',
    #'special-members',
    'inherited-members',
    #"exclude-members ",
    'show-inheritance'
]

autodoc_default_options = {
    "exclude-members": "cell_to_update,gui_update,colorChanged,fontChanged,fontSizeChanged",
}

autoclass_content = "both"
autosummary_generate = True



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'nature'

html_static_path = ['_static']
html_css_files = [
    'css/w3.css',
]

html_title = 'pycellsheet API docs'
html_short_title = "pycellsheet"
html_favicon = "_static/pyspread.png"


html_sidebars = {
   '**': ['globaltoc.html', 'sourcelink.html', 'searchbox.html'],
   'using/windows': ['windowssidebar.html', 'searchbox.html'],
}


# If true, links to the reST sources are added to the pages.
#
html_show_sourcelink = False

# If true, "Created using Sphinx" is shown in the HTML footer. Default is True.
#
html_show_sphinx = False

# If true, "(C) Copyright ..." is shown in the HTML footer. Default is True.
#
html_show_copyright = False

html_context = {"git_repos_url": "https://github.com/EuphoricCatface/PyCellSheet"}

# Normalize legacy docstring roles/links from upstream qimage2ndarray/matplotlib
# docs so API builds stay stable while we keep source behavior unchanged.
rst_epilog = """
.. role:: mpltype(emphasis)
.. _qimage: https://doc.qt.io/qt-6/qimage.html
.. _numpy.ndarray: https://numpy.org/doc/stable/reference/generated/numpy.ndarray.html
.. _loading and saving images: https://doc.qt.io/qt-6/qimagereader.html
"""

# Keep API build signal-to-noise stable while we incrementally normalize
# legacy type-hint forward references in vendored packaging helpers.
suppress_warnings = [
    "sphinx_autodoc_typehints.forward_reference",
]

intersphinx_mapping = {
    'python': ('http://docs.python.org/3', None),
    'numpy': ('http://docs.scipy.org/doc/numpy', None),
    'scipy': ('http://docs.scipy.org/doc/scipy/reference', None),
    'matplotlib': ('https://matplotlib.org/', None)
}
