"""
Temas modernos e sofisticados para apresentações.

Cada tema define paleta de cores, tipografia e estilo visual.
"""

from dataclasses import dataclass, field
from pptx.util import Pt
from pptx.dml.color import RGBColor


@dataclass
class Theme:
    """Define um tema visual completo para apresentações."""

    name: str
    # Cores principais
    primary: str
    secondary: str
    accent: str
    background: str
    surface: str
    text_primary: str
    text_secondary: str
    text_on_primary: str
    # Cores extras para gráficos e destaques
    chart_colors: list = field(default_factory=list)
    # Tipografia
    font_title: str = "Calibri"
    font_body: str = "Calibri"
    font_mono: str = "Consolas"
    # Tamanhos de fonte
    size_title: int = 40
    size_subtitle: int = 22
    size_heading: int = 28
    size_body: int = 16
    size_caption: int = 12
    size_small: int = 10

    def rgb(self, hex_color: str) -> RGBColor:
        """Converte cor hex para RGBColor."""
        h = hex_color.lstrip("#")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    @property
    def primary_rgb(self) -> RGBColor:
        return self.rgb(self.primary)

    @property
    def secondary_rgb(self) -> RGBColor:
        return self.rgb(self.secondary)

    @property
    def accent_rgb(self) -> RGBColor:
        return self.rgb(self.accent)

    @property
    def background_rgb(self) -> RGBColor:
        return self.rgb(self.background)

    @property
    def surface_rgb(self) -> RGBColor:
        return self.rgb(self.surface)

    @property
    def text_primary_rgb(self) -> RGBColor:
        return self.rgb(self.text_primary)

    @property
    def text_secondary_rgb(self) -> RGBColor:
        return self.rgb(self.text_secondary)

    @property
    def text_on_primary_rgb(self) -> RGBColor:
        return self.rgb(self.text_on_primary)

    def chart_colors_rgb(self) -> list:
        return [self.rgb(c) for c in self.chart_colors]


# ─── Temas disponíveis ────────────────────────────────────────────────

CORPORATE_DARK = Theme(
    name="Corporate Dark",
    primary="#1A73E8",
    secondary="#0D47A1",
    accent="#FF6D00",
    background="#1E1E2E",
    surface="#2D2D3F",
    text_primary="#FFFFFF",
    text_secondary="#B0B0C0",
    text_on_primary="#FFFFFF",
    chart_colors=["#1A73E8", "#FF6D00", "#00C853", "#AA00FF", "#FF1744", "#00B8D4"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=44,
    size_heading=30,
)

CORPORATE_LIGHT = Theme(
    name="Corporate Light",
    primary="#0055B8",
    secondary="#003D82",
    accent="#E85D00",
    background="#FFFFFF",
    surface="#F5F7FA",
    text_primary="#1A1A2E",
    text_secondary="#5A5A7A",
    text_on_primary="#FFFFFF",
    chart_colors=["#0055B8", "#E85D00", "#00875A", "#6554C0", "#DE350B", "#00A3BF"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=44,
    size_heading=30,
)

TECH_STARTUP = Theme(
    name="Tech Startup",
    primary="#6C63FF",
    secondary="#3F3D56",
    accent="#FF6584",
    background="#0F0E17",
    surface="#1A1A2E",
    text_primary="#FFFFFE",
    text_secondary="#A7A9BE",
    text_on_primary="#FFFFFF",
    chart_colors=["#6C63FF", "#FF6584", "#43E97B", "#F9D423", "#38F9D7", "#FA709A"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=46,
    size_heading=30,
)

MINIMAL_ELEGANT = Theme(
    name="Minimal Elegant",
    primary="#2D2D2D",
    secondary="#4A4A4A",
    accent="#C9A96E",
    background="#FAFAFA",
    surface="#F0F0F0",
    text_primary="#2D2D2D",
    text_secondary="#7A7A7A",
    text_on_primary="#FFFFFF",
    chart_colors=["#2D2D2D", "#C9A96E", "#7A9E7E", "#B07AA1", "#E15759", "#4E79A7"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=42,
    size_heading=28,
)

CREATIVE_BOLD = Theme(
    name="Creative Bold",
    primary="#FF2D55",
    secondary="#5856D6",
    accent="#FFD60A",
    background="#000000",
    surface="#1C1C1E",
    text_primary="#FFFFFF",
    text_secondary="#98989D",
    text_on_primary="#FFFFFF",
    chart_colors=["#FF2D55", "#5856D6", "#FFD60A", "#30D158", "#64D2FF", "#FF9F0A"],
    font_title="Calibri",
    font_body="Calibri",
    size_title=48,
    size_heading=32,
)

NATURE_GREEN = Theme(
    name="Nature Green",
    primary="#2E7D32",
    secondary="#1B5E20",
    accent="#FF8F00",
    background="#FEFDF5",
    surface="#F1F8E9",
    text_primary="#1B2E1B",
    text_secondary="#5D7A5D",
    text_on_primary="#FFFFFF",
    chart_colors=["#2E7D32", "#FF8F00", "#00838F", "#6A1B9A", "#C62828", "#283593"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=42,
    size_heading=28,
)

EDUCATION = Theme(
    name="Education",
    primary="#1565C0",
    secondary="#0D47A1",
    accent="#F9A825",
    background="#FFFFFF",
    surface="#E3F2FD",
    text_primary="#212121",
    text_secondary="#616161",
    text_on_primary="#FFFFFF",
    chart_colors=["#1565C0", "#F9A825", "#2E7D32", "#C62828", "#6A1B9A", "#00838F"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=42,
    size_heading=28,
)

MEDICAL = Theme(
    name="Medical",
    primary="#00838F",
    secondary="#006064",
    accent="#26A69A",
    background="#FFFFFF",
    surface="#E0F7FA",
    text_primary="#1A2B3C",
    text_secondary="#5A6B7C",
    text_on_primary="#FFFFFF",
    chart_colors=["#00838F", "#26A69A", "#0277BD", "#00695C", "#4DD0E1", "#80CBC4"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=42,
    size_heading=28,
)

FINANCIAL = Theme(
    name="Financial",
    primary="#1A237E",
    secondary="#0D1642",
    accent="#00897B",
    background="#FAFBFC",
    surface="#E8EAF6",
    text_primary="#1A1A2E",
    text_secondary="#5C5C7A",
    text_on_primary="#FFFFFF",
    chart_colors=["#1A237E", "#00897B", "#F57C00", "#C62828", "#6A1B9A", "#2E7D32"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=42,
    size_heading=28,
)

GRADIENT_PURPLE = Theme(
    name="Gradient Purple",
    primary="#7C4DFF",
    secondary="#651FFF",
    accent="#18FFFF",
    background="#12002E",
    surface="#1E0048",
    text_primary="#FFFFFF",
    text_secondary="#B39DDB",
    text_on_primary="#FFFFFF",
    chart_colors=["#7C4DFF", "#18FFFF", "#FF4081", "#76FF03", "#FFAB40", "#40C4FF"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=46,
    size_heading=30,
)

SUNSET_WARM = Theme(
    name="Sunset Warm",
    primary="#E64A19",
    secondary="#BF360C",
    accent="#FFC107",
    background="#FFF8F0",
    surface="#FBE9E7",
    text_primary="#3E2723",
    text_secondary="#795548",
    text_on_primary="#FFFFFF",
    chart_colors=["#E64A19", "#FFC107", "#F06292", "#7E57C2", "#26A69A", "#5C6BC0"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=42,
    size_heading=28,
)

OCEAN_BLUE = Theme(
    name="Ocean Blue",
    primary="#0288D1",
    secondary="#01579B",
    accent="#00E5FF",
    background="#0A1628",
    surface="#132238",
    text_primary="#FFFFFF",
    text_secondary="#90CAF9",
    text_on_primary="#FFFFFF",
    chart_colors=["#0288D1", "#00E5FF", "#26C6DA", "#4FC3F7", "#80DEEA", "#B3E5FC"],
    font_title="Calibri Light",
    font_body="Calibri",
    size_title=44,
    size_heading=30,
)

# Dicionário de todos os temas
THEMES = {
    "corporate_dark": CORPORATE_DARK,
    "corporate_light": CORPORATE_LIGHT,
    "tech_startup": TECH_STARTUP,
    "minimal_elegant": MINIMAL_ELEGANT,
    "creative_bold": CREATIVE_BOLD,
    "nature_green": NATURE_GREEN,
    "education": EDUCATION,
    "medical": MEDICAL,
    "financial": FINANCIAL,
    "gradient_purple": GRADIENT_PURPLE,
    "sunset_warm": SUNSET_WARM,
    "ocean_blue": OCEAN_BLUE,
}
