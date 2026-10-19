from openpyxl.formula.translate import Translator

from .types import Formula


def readdress(formula: str, origin: str, dest: str) -> Formula:
    """Move uma fórmula de `origin` para `dest` (ex.: 'L4' -> 'L2'), ajustando
    referências relativas e mantendo as absolutas."""
    return Formula(Translator(formula, origin=origin).translate_formula(dest))
