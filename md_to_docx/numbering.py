from typing import Optional
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .errors import TemplateError, NumberingError

class NumberingManager:
    def __init__(self, doc):
        self.doc = doc
        self.bullet_num_id: Optional[str] = None
        self.decimal_num_id: Optional[str] = None

    def initialize(self):
        self.bullet_num_id  = self._get_num_id(ordered=False)
        self.decimal_num_id = self._get_num_id(ordered=True)

    def _get_num_id(self, ordered: bool) -> Optional[str]:
        numbering_part = getattr(self.doc.part, "numbering_part", None)
        if numbering_part is None:
            raise TemplateError("Template hat keinen numbering_part (keine Listen-Definitionen).")

        root = numbering_part.element
        want = "decimal" if ordered else "bullet"
        target_abs = None
        for abs_ in root.xpath(".//*[local-name()='abstractNum']"):
            lvl0 = None
            for lvl in abs_.xpath("./*[local-name()='lvl']"):
                if lvl.get(qn("w:ilvl")) == "0":
                    lvl0 = lvl; break
            if lvl0 is None:
                continue
            fmts = lvl0.xpath("./*[local-name()='numFmt']")
            if fmts and fmts[0].get(qn("w:val")) == want:
                target_abs = abs_; break

        if target_abs is None:
            return None

        abs_id = target_abs.get(qn("w:abstractNumId"))
        if not abs_id:
            raise NumberingError("abstractNumId fehlt in der Nummerierungsdefinition.")

        for num in root.xpath(".//*[local-name()='num']"):
            absRef = num.xpath("./*[local-name()='abstractNumId']")
            if absRef and absRef[0].get(qn("w:val")) == abs_id:
                nid = num.get(qn("w:numId"))
                if nid: return nid

        max_id = 0
        for num in root.xpath(".//*[local-name()='num']"):
            try: max_id = max(max_id, int(num.get(qn("w:numId"))))
            except Exception: pass
        new_id = str(max_id + 1)
        num_el = OxmlElement("w:num"); num_el.set(qn("w:numId"), new_id)
        abs_ref = OxmlElement("w:abstractNumId"); abs_ref.set(qn("w:val"), abs_id)
        num_el.append(abs_ref); root.append(num_el)
        return new_id

    def apply(self, paragraph, *, ordered: bool, level: int):
        num_id = self.decimal_num_id if ordered else self.bullet_num_id
        if not num_id:
            return  # kein Stil verfügbar -> plain paragraph
        pPr = paragraph._p.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl"); ilvl.set(qn("w:val"), str(level)); numPr.append(ilvl)
        nId  = OxmlElement("w:numId"); nId.set(qn("w:val"), str(num_id)); numPr.append(nId)
        pPr.append(numPr)
