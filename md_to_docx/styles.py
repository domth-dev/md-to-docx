from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE

class StyleManager:
    def __init__(self, doc, base_font="Calibri", base_size_pt=11, apply_base=False):
        self.doc = doc
        self.base_font = base_font
        self.base_size_pt = base_size_pt
        self.apply_base = apply_base

    def ensure(self):
        styles = self.doc.styles
        if self.apply_base:
            for s in styles:
                try:
                    s.font.name = self.base_font
                    s.font.size = Pt(self.base_size_pt)
                except Exception:
                    pass
        if "Code Block" not in styles:
            st = styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
            st.font.name = "Consolas"; st.font.size = Pt(self.base_size_pt)
        if "Inline Code" not in styles:
            st = styles.add_style("Inline Code", WD_STYLE_TYPE.CHARACTER)
            st.font.name = "Consolas"; st.font.size = Pt(self.base_size_pt)