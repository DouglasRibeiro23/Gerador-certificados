import types
import pandas as pd
import pytest

# --------------------------------------------------------------------
# FAKES para "python-pptx"
# Por que criar fakes?
# - Para não abrir/salvar PPTX de verdade nos testes (rápido, determinístico).
# - Simulamos apenas o que o script usa: Presentation/slides/shapes/text_frame/runs.
# - Com isso, testamos a lógica sem I/O real e sem depender de MS Office.
# --------------------------------------------------------------------

class FakeFont:
    def __init__(self):
        self.name = None
        self.size = None

        # Cor em "python-pptx" é um objeto com .rgb; aqui basta um holder simples
        class _Color:
            def __init__(self):
                self.rgb = None
        self.color = _Color()

class FakeRun:
    def __init__(self):
        self.text = ""
        self.font = FakeFont()

class FakeParagraph:
    def __init__(self):
        self.runs = []

    def add_run(self):
        # O script cria "fragmentos" de texto (runs) para aplicar estilos diferentes
        run = FakeRun()
        self.runs.append(run)
        return run

class FakeTextFrame:
    def __init__(self):
        self.word_wrap = False
        self.paragraphs = []

    def add_paragraph(self):
        p = FakeParagraph()
        self.paragraphs.append(p)
        return p

class FakeTextbox:
    # Representa as caixas de texto adicionadas ao novo slide
    def __init__(self, left=0, top=0, width=0, height=0):
        self.text_frame = FakeTextFrame()

class FakeShape:
    # Representa um shape de texto do slide "modelo"
    def __init__(self, text):
        self._text = text
        self.has_text_frame = True
        # Medidas: não afetam asserts, mas o script lê esses atributos
        self.left = 0
        self.top = 0
        self.width = 100
        self.height = 20

    @property
    def text(self):
        return self._text

class FakeShapes:
    def __init__(self, initial_shapes=None):
        # Shapes do slide modelo e textboxes que forem sendo adicionados no novo slide
        self._shapes = list(initial_shapes or [])
        self.added_textboxes = []

    def __iter__(self):
        return iter(self._shapes)

    def add_textbox(self, left, top, width, height):
        # Esta chamada é o "efeito" que esperamos observar ao renderizar um certificado
        tb = FakeTextbox(left, top, width, height)
        self.added_textboxes.append(tb)
        return tb

class FakeSlide:
    def __init__(self, shapes, layout="LAYOUT"):
        self.shapes = shapes
        self.slide_layout = layout

class FakeSlides:
    def __init__(self, model_shapes):
        # slides[0] é o slide modelo (igual ao template)
        self._slides = [FakeSlide(FakeShapes(model_shapes))]
        # python-pptx mantém uma lista interna de IDs de slides; simulamos para testar remoção
        self._sldIdLst = ["ID0"]

    def __getitem__(self, idx):
        return self._slides[idx]

    def __len__(self):
        return len(self._slides)

    def add_slide(self, layout):
        slide = FakeSlide(FakeShapes(), layout)
        self._slides.append(slide)
        self._sldIdLst.append(f"ID{len(self._slides)-1}")
        return slide

class FakePresentation:
    # Presentation fake carregada a partir de um "template"
    def __init__(self, template_path):
        # Um shape de texto no slide modelo com todos os placeholders
        model_text = (
            "Aluno {NOME} concluiu {CAPACITACAO} no evento {EVENTO} "
            "em {DIA} de {MES} de {ANO} - {HORAS} horas."
        )
        self.slides = FakeSlides([FakeShape(model_text)])
        self._saved_path = None
        self.template_path = template_path

    def save(self, path):
        # Não gravamos em disco; guardamos o caminho para inspeção em testes
        self._saved_path = path


# ------------------------------ FIXTURES ------------------------------ #

@pytest.fixture
def fake_pptx(monkeypatch):
    """
    Injeta um pacote fake "pptx" completo o suficiente:
      - pptx.Presentation -> FakePresentation
      - pptx.util.Pt -> função fake (apenas retorna o valor)
      - pptx.dml.color.RGBColor -> construtor fake que devolve (r, g, b)
    Por quê?
      - O script faz `from pptx.util import Pt` e `from pptx.dml.color import RGBColor`.
        Se esses submódulos não existirem, o import quebra antes dos testes rodarem.
    """
    import sys

    # Submódulo: pptx.util com Pt
    util_mod = types.SimpleNamespace()
    def Pt_fake(value):
        # Em python-pptx real Pt cria uma unidade; aqui basta devolver o valor
        return value
    util_mod.Pt = Pt_fake

    # Submódulo: pptx.dml.color com RGBColor
    color_mod = types.SimpleNamespace()
    def RGBColor_fake(r, g, b):
        # Para nossos testes, tupla (r, g, b) é suficiente e fácil de comparar
        return (int(r), int(g), int(b))
    color_mod.RGBColor = RGBColor_fake

    # Módulo pptx.dml e raiz pptx
    dml_mod = types.SimpleNamespace(color=color_mod)
    pptx_mod = types.SimpleNamespace(Presentation=FakePresentation, util=util_mod, dml=dml_mod)

    # Registrar todos em sys.modules para suportar os imports "normais"
    sys.modules["pptx"] = pptx_mod
    sys.modules["pptx.util"] = util_mod
    sys.modules["pptx.dml"] = dml_mod
    sys.modules["pptx.dml.color"] = color_mod

    # E garantir que quando o código pedir pptx.Presentation, receba nosso Fake
    monkeypatch.setattr(pptx_mod, "Presentation", FakePresentation, raising=True)
    return pptx_mod


@pytest.fixture
def chdir_tmp(tmp_path, monkeypatch):
    """
    Isola o diretório de trabalho:
    - Mantém arquivos gerados (ex.: ./certificados) dentro de uma pasta temporária.
    - Evita poluir o repositório e permite rodadas repetíveis.
    """
    monkeypatch.chdir(tmp_path)
    return tmp_path


@pytest.fixture
def df_basico():
    """
    Classe de equivalência "dados normais":
    - Duas linhas com valores típicos, para exercitar o caminho feliz.
    """
    return pd.DataFrame({
        "Nome": ["Maria", "João"],
        "Capacitacao": ["Excel: Dashboards*", "IA/ML: Fundamentos?"],
        "Evento": ["Tech Week", "Jornada IA"],
        "Data": ["2025-11-05", "2024-02-01"],
        "Horas": [8, 4],
    })


@pytest.fixture
def mock_read_excel(monkeypatch, df_basico):
    """
    Mock de pandas.read_excel:
    - Por que mockar? Para controlar a entrada e eliminar I/O real de planilha.
    - Retornamos sempre um DataFrame cópia (evita efeitos colaterais entre testes).
    """
    monkeypatch.setattr(pd, "read_excel", lambda *a, **k: df_basico.copy(deep=True), raising=True)
    return df_basico
