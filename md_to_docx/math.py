from typing import Optional, List
from io import BytesIO
from PIL import Image
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from .errors import MathConversionError

class MathEngine:
    def __init__(self, *, strict=False, dpi=220):
        self.strict = strict
        self.dpi = dpi

    # ---------- public ----------
    def add_math(self, paragraph, latex_code: str, *, display: bool = False):
        try:
            omml = self._latex_to_omml(latex_code)
        except Exception as e:
            if self.strict:
                raise MathConversionError(f"Parsing fehlgeschlagen: {latex_code!r}") from e
            omml = None

        if omml is not None:
            if display:
                para = OxmlElement('m:oMathPara'); para.append(omml); paragraph._p.append(para)
            else:
                paragraph._p.append(omml)
            return

        if self.strict:
            raise MathConversionError(f"Formel nicht unterstützt: {latex_code!r}")
        self._add_png_fallback(paragraph, latex_code, display=display)

    # ---------- intern ----------
    def _add_png_fallback(self, paragraph, tex: str, *, display: bool):
        fig = matplotlib.figure.Figure()
        ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
        t = ax.text(0.5, 0.5, f"${tex}$", ha='center', va='center',
                    fontsize=14 if not display else 16)
        fig.canvas = matplotlib.backends.backend_agg.FigureCanvasAgg(fig)
        fig.canvas.draw()
        bbox = t.get_window_extent(fig.canvas.get_renderer()).expanded(1.15, 1.25)
        fig.set_size_inches(bbox.width/self.dpi, bbox.height/self.dpi)
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=self.dpi, transparent=True, bbox_inches=bbox)
        buf.seek(0)
        run = paragraph.add_run()
        img = Image.open(buf); w,h = img.size
        target_pt = 20 if display else 14
        target_px = int(target_pt * self.dpi / 72)
        scale = target_px / max(h, 1)
        width_inches = (w * scale) / self.dpi
        buf.seek(0)
        run.add_picture(buf, width=Inches(width_inches))

    def _m_text(self, txt: str):
        r = OxmlElement('m:r'); t = OxmlElement('m:t'); t.text = txt; r.append(t); return r

    def _latex_to_omml(self, src: str) -> Optional[OxmlElement]:
        r"""
        80/20 LaTeX → OMML: ^, _, \frac{a}{b}, \sqrt{...}, \sum_{..}^{..} body, \int_{..}^{..} body
        """
        src = src.strip(); i = 0

        def parse_group():
            nonlocal i
            if i < len(src) and src[i] == '{':
                i += 1; nodes: List[OxmlElement] = []
                while i < len(src) and src[i] != '}':
                    nodes.extend(parse_atom())
                if i < len(src) and src[i] == '}': i += 1
                return nodes
            m = re.match(r'[A-Za-z\u0370-\u03FF0-9]+', src[i:])
            if m:
                i += len(m.group(0)); return [self._m_text(m.group(0))]
            ch = src[i]; i += 1; return [self._m_text(ch)]

        def apply_sup_sub(base_nodes: List[OxmlElement]):
            nonlocal i
            sub_nodes = sup_nodes = None
            while i < len(src) and src[i] in ['_', '^']:
                kind = src[i]; i += 1
                nodes = parse_group()
                if kind == '_': sub_nodes = nodes
                else: sup_nodes = nodes
            if sub_nodes and sup_nodes:
                node = OxmlElement('m:sSubSup')
                e = OxmlElement('m:e');  [e.append(n) for n in base_nodes]
                sub = OxmlElement('m:sub'); [sub.append(n) for n in sub_nodes]
                sup = OxmlElement('m:sup'); [sup.append(n) for n in sup_nodes]
                node.extend([e, sub, sup]); return [node]
            if sub_nodes:
                node = OxmlElement('m:sSub')
                e = OxmlElement('m:e');  [e.append(n) for n in base_nodes]
                sub = OxmlElement('m:sub'); [sub.append(n) for n in sub_nodes]
                node.extend([e, sub]); return [node]
            if sup_nodes:
                node = OxmlElement('m:sSup')
                e = OxmlElement('m:e');  [e.append(n) for n in base_nodes]
                sup = OxmlElement('m:sup'); [sup.append(n) for n in sup_nodes]
                node.extend([e, sup]); return [node]
            return base_nodes

        def parse_atom():
            nonlocal i
            if src.startswith(r'\frac', i):
                i += 5
                num = parse_group(); den = parse_group()
                f = OxmlElement('m:f')
                num_e = OxmlElement('m:num'); [num_e.append(n) for n in num]
                den_e = OxmlElement('m:den'); [den_e.append(n) for n in den]
                f.extend([num_e, den_e]); return apply_sup_sub([f])
            if src.startswith(r'\sqrt', i):
                i += 5
                rad = parse_group()
                r = OxmlElement('m:rad')
                deg = OxmlElement('m:deg')
                e = OxmlElement('m:e'); [e.append(n) for n in rad]
                r.extend([deg, e]); return apply_sup_sub([r])
            if src.startswith(r'\sum', i) or src.startswith(r'\int', i):
                is_sum = src.startswith(r'\sum', i); i += 4
                lower = upper = None
                tmp = i
                if tmp < len(src) and src[tmp] in ['_', '^']:
                    while tmp < len(src) and src[tmp] in ['_', '^']:
                        kind = src[tmp]; tmp += 1; i = tmp
                        grp = parse_group()
                        if kind == '_': lower = grp
                        else: upper = grp
                        tmp = i
                body = parse_group()
                nary = OxmlElement('m:nary')
                npr = OxmlElement('m:naryPr')
                chr_el = OxmlElement('m:chr'); chr_el.set(qn('w:val'), '∑' if is_sum else '∫')
                limLoc = OxmlElement('m:limLoc'); limLoc.set(qn('w:val'), 'undOvr')
                npr.extend([chr_el, limLoc]); nary.append(npr)
                if lower is not None:
                    sub = OxmlElement('m:sub'); [sub.append(n) for n in lower]; nary.append(sub)
                if upper is not None:
                    sup = OxmlElement('m:sup'); [sup.append(n) for n in upper]; nary.append(sup)
                e = OxmlElement('m:e'); [e.append(n) for n in body]; nary.append(e)
                return apply_sup_sub([nary])
            base = parse_group()
            return apply_sup_sub(base)

        out = []
        while i < len(src):
            if src[i].isspace(): i += 1; continue
            out.extend(parse_atom())
        if not out: return None
        oMath = OxmlElement('m:oMath'); [oMath.append(n) for n in out]
        return oMath
