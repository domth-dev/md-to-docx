from markdown_it import MarkdownIt
from mdit_py_plugins.tasklists import tasklists_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.texmath import texmath_plugin

class MarkdownParser:
    def __init__(self, *, breaks=False, linkify=True):
        self.md = (
            MarkdownIt("gfm-like", {"breaks": breaks, "linkify": linkify})
            .enable("table")
            .use(tasklists_plugin, enabled=True)
            .use(footnote_plugin)
            .use(texmath_plugin)
        )

    def parse(self, text: str):
        return self.md.parse(text)