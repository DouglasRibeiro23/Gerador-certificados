import os
import importlib
import importlib.util
from pathlib import Path
import pandas as pd

# ============================================================
# Por que mexemos aqui?
# - Seu script tem hífen no nome (CERTIFICADOS-MINICURSOS.py).
# - importlib.import_module("CERTIFICADOS-MINICURSOS") não funciona
#   porque hífen não é um identificador Python válido.
# - Para manter o nome do arquivo, carregamos o módulo PELO CAMINHO.
# - Isso torna os testes portáveis sem exigir renomear seu arquivo.
# ============================================================

def _import_or_reload():
    """
    Carrega (ou recarrega) CERTIFICADOS-MINICURSOS.py pelo caminho absoluto,
    calculado a partir deste arquivo de teste (não do CWD do pytest).
    """
    # tests/ -> pai é a pasta do projeto onde está o .py com hífen
    project_root = Path(__file__).resolve().parent.parent
    script_path = project_root / "CERTIFICADOS-MINICURSOS.py"
    assert script_path.exists(), f"Arquivo não encontrado: {script_path}"

    module_name = "certificados_minicursos_under_test"
    if module_name in importlib.sys.modules:
        del importlib.sys.modules[module_name]

    spec = importlib.util.spec_from_file_location(module_name, script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)  # executa o código top-level do script
    importlib.sys.modules[module_name] = module
    return module

def test_fluxo_feliz_gera_2_certificados(chdir_tmp, mock_read_excel, fake_pptx, capsys):
    """
    Por que este teste?
      Validar o “caminho feliz”: com dados válidos, o script deve processar todas as linhas,
      criar a pasta de saída, gerar um slide novo para cada participante e “salvar” o arquivo.

    Tipo de teste:
      • Caminho feliz (happy path) + oráculo por efeitos colaterais visíveis (stdout e paths).

    O que checamos?
      • O script imprime uma linha de sucesso para CADA certificado.
      • O nome de arquivo gerado é coerente com a sanitização de caracteres inválidos.
      • A pasta 'certificados' existe (sinal de que tentou escrever arquivos).
    """
    # Template "vazio" apenas para não quebrar a abertura do Presentation
    Path("Template capacitações.pptx").write_bytes(b"")

    _import_or_reload()
    out = capsys.readouterr().out

    lines = [ln for ln in out.splitlines() if ln.strip().startswith("PPTX gerado sem slide extra:")]
    assert len(lines) == 2, "Esperava duas confirmações de geração (2 participantes)."

    # Sanitização esperada (caracteres especiais -> '_')
    expected1 = os.path.join("certificados", "certificado_Maria_Excel_ Dashboards_.pptx")
    expected2 = os.path.join("certificados", "certificado_João_IA_ML_ Fundamentos_.pptx")
    assert expected1 in out
    assert expected2 in out
    assert os.path.isdir("certificados")


def test_placeholders_em_cor_azul_com_fonte(chdir_tmp, fake_pptx, capsys, monkeypatch):
    """
    Por que este teste?
      Assegurar que os placeholders especiais ({NOME}, {CAPACITACAO}, {EVENTO})
      são renderizados como “runs” com estilo azul e fonte específica.

    Tipo de teste:
      • Teste funcional de estilo/conteúdo (verificação semântica mínima).

    O que checamos?
      • No slide novo, existe pelo menos 1 textbox com parágrafo/runs.
      • Entre os runs, encontramos textos dos placeholders com:
        - cor RGB (0,112,192),
        - fonte "Arial Rounded MT Bold",
        - tamanho definido.
    """
    Path("Template capacitações.pptx").write_bytes(b"")

    df = pd.DataFrame({
        "Nome": ["Alice"],
        "Capacitacao": ["Oficina de Dados"],
        "Evento": ["TechDay"],
        "Data": ["2025-11-05"],
        "Horas": [6],
    })
    monkeypatch.setattr(pd, "read_excel", lambda *a, **k: df, raising=True)

    # 1ª execução só para “aquecer”
    _import_or_reload()
    _ = capsys.readouterr().out

    # Intercepta save() para capturar a última Presentation criada (fake)
    import sys
    SAVED = {}
    def spy_save(self, path):
        SAVED["pres"] = self
        self._saved_path = path
    # pptx foi substituído por fake em conftest; patchamos o método da classe
    sys.modules["pptx"].Presentation.save = spy_save  # type: ignore[attr-defined]

    # 2ª execução: agora capturamos a instância fake e inspecionamos
    _import_or_reload()
    _ = capsys.readouterr().out
    pres = SAVED.get("pres")
    assert pres is not None, "Não capturamos a apresentação salva."

    # Novo slide é índice 1 (0 é o modelo)
    new_slide = pres.slides._slides[1]
    assert new_slide.shapes.added_textboxes, "Esperava pelo menos 1 textbox gerado."
    tb = new_slide.shapes.added_textboxes[0]
    assert tb.text_frame.paragraphs, "Textbox sem parágrafos."
    runs = [r for p in tb.text_frame.paragraphs for r in p.runs]
    assert runs, "Parágrafo sem runs."

    wanted = {"Alice", "Oficina de Dados", "TechDay"}
    seen = {k: False for k in wanted}
    for r in runs:
        if r.text in wanted and getattr(r.font.color, "rgb", None) == (0, 112, 192):
            assert r.font.name == "Arial Rounded MT Bold"
            assert r.font.size is not None  # no script é Pt(16)
            seen[r.text] = True
    assert all(seen.values()), f"Nem todos os placeholders azuis apareceram como runs estilizados: {seen}"


def test_meses_em_portugues_e_datas_validas(chdir_tmp, fake_pptx, capsys, monkeypatch):
    """
    Por que este teste?
      Garantir que datas diferentes resultam no mapeamento correto do mês em PT-BR
      (ex.: 2025-11-05 → “novembro”; 2024-02-01 → “fevereiro”).

    Tipo de teste:
      • Classe de equivalência com variação de campo (datas diferentes).

    O que checamos?
      • O script consegue processar as duas linhas sem erro (2 logs de sucesso).
      • (Oráculo indireto) Isso indica que o mapeamento foi aplicado sem quebrar.
    """
    Path("Template capacitações.pptx").write_bytes(b"")
    df = pd.DataFrame({
        "Nome": ["X", "Y"],
        "Capacitacao": ["C1", "C2"],
        "Evento": ["E1", "E2"],
        "Data": ["2025-11-05", "2024-02-01"],
        "Horas": [2, 3],
    })
    monkeypatch.setattr(pd, "read_excel", lambda *a, **k: df, raising=True)

    _import_or_reload()
    out = capsys.readouterr().out
    lines = [ln for ln in out.splitlines() if ln.strip().startswith("PPTX gerado sem slide extra:")]
    assert len(lines) == 2, "Processamento das duas datas deveria concluir sem erros."


def test_remove_slide_modelo(chdir_tmp, mock_read_excel, fake_pptx, capsys):
    """
    Por que este teste?
      O script remove o slide “modelo” antes de salvar. Queremos comprovar esse efeito,
      já que ele impacta diretamente a qualidade do arquivo final.

    Tipo de teste:
      • Comportamento específico (pré-condição de salvamento).

    O que checamos?
      • A lista interna de IDs de slides NÃO contém mais o ID do primeiro slide (“ID0”).
    """
    Path("Template capacitações.pptx").write_bytes(b"")
    _import_or_reload()
    _ = capsys.readouterr().out

    # Reexecutamos com uma subclasse que nos permite inspecionar a última instância
    import sys
    class SpyPresentation(fake_pptx.Presentation):  # type: ignore[name-defined]
        last_instance = None
        def __init__(self, path):
            super().__init__(path)
            SpyPresentation.last_instance = self
    sys.modules["pptx"].Presentation = SpyPresentation  # type: ignore[attr-defined]

    df_one = pd.DataFrame({
        "Nome": ["Z"],
        "Capacitacao": ["C3"],
        "Evento": ["E3"],
        "Data": ["2025-01-10"],
        "Horas": [1],
    })
    pd.read_excel = lambda *a, **k: df_one  # monkeypatch simples

    _import_or_reload()
    _ = capsys.readouterr().out

    pres = SpyPresentation.last_instance
    assert pres is not None
    assert "ID0" not in pres.slides._sldIdLst, "Esperava que o slide modelo tivesse sido removido."
    assert len(pres.slides._slides) >= 1, "Deveria restar ao menos o slide final."


def test_sanitizacao_de_caracteres_invalidos_no_nome(chdir_tmp, fake_pptx, capsys, monkeypatch):
    import types
    import pandas as pd
    import pytest

    # --------------------------------------------------------------------
    # FAKES para "python-pptx"
    # Por que fazer fakes?
    # - Para não abrir/salvar PPTX de verdade (mais rápido e estável).
    # - Simulamos apenas o que o script usa: slides, shapes, text_frame, runs...
    # - Isso permite que os testes foquem na LÓGICA do nosso código.
    # --------------------------------------------------------------------

    class FakeFont:
        def __init__(self):
            self.name = None
            self.size = None

            class _Color:  # mini-objeto com atributo rgb
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
        def __init__(self, left=0, top=0, width=0, height=0):
            self.text_frame = FakeTextFrame()

    class FakeShape:
        def __init__(self, text):
            self._text = text
            self.has_text_frame = True
            self.left = 0;
            self.top = 0;
            self.width = 100;
            self.height = 20

        @property
        def text(self):
            return self._text

    class FakeShapes:
        def __init__(self, initial_shapes=None):
            self._shapes = list(initial_shapes or [])
            self.added_textboxes = []

        def __iter__(self):
            return iter(self._shapes)

        def add_textbox(self, left, top, width, height):
            tb = FakeTextbox(left, top, width, height)
            self.added_textboxes.append(tb)
            return tb

    class FakeSlide:
        def __init__(self, shapes, layout="LAYOUT"):
            self.shapes = shapes
            self.slide_layout = layout

    class FakeSlides:
        def __init__(self, model_shapes):
            self._slides = [FakeSlide(FakeShapes(model_shapes))]
            self._sldIdLst = ["ID0"]

        def __getitem__(self, idx):
            return self._slides[idx]

        def __len__(self):
            return len(self._slides)

        def add_slide(self, layout):
            slide = FakeSlide(FakeShapes(), layout)
            self._slides.append(slide)
            self._sldIdLst.append(f"ID{len(self._slides) - 1}")
            return slide

    class FakePresentation:
        def __init__(self, template_path):
            model_text = ("Aluno {NOME} concluiu {CAPACITACAO} no evento {EVENTO} "
                          "em {DIA} de {MES} de {ANO} - {HORAS} horas.")
            self.slides = FakeSlides([FakeShape(model_text)])
            self._saved_path = None
            self.template_path = template_path

        def save(self, path):
            self._saved_path = path

    # ------------------------- FIXTURES ------------------------- #

    @pytest.fixture
    def fake_pptx(monkeypatch):
        """Troca pptx.Presentation pela nossa FakePresentation (injeção de dependência)."""
        import sys
        pptx_mod = types.SimpleNamespace(Presentation=FakePresentation)
        sys.modules.setdefault("pptx", pptx_mod)
        monkeypatch.setattr(pptx_mod, "Presentation", FakePresentation, raising=True)
        return pptx_mod

    @pytest.fixture
    def chdir_tmp(tmp_path, monkeypatch):
        """Isola o CWD para não poluir sua pasta do projeto."""
        monkeypatch.chdir(tmp_path)
        return tmp_path

    @pytest.fixture
    def df_basico():
        """Classe de equivalência: dados normais (duas linhas)."""
        return pd.DataFrame({
            "Nome": ["Maria", "João"],
            "Capacitacao": ["Excel: Dashboards*", "IA/ML: Fundamentos?"],
            "Evento": ["Tech Week", "Jornada IA"],
            "Data": ["2025-11-05", "2024-02-01"],
            "Horas": [8, 4],
        })

    @pytest.fixture
    def mock_read_excel(monkeypatch, df_basico):
        """Mock do pandas.read_excel: controlamos a entrada e evitamos I/O."""
        monkeypatch.setattr(pd, "read_excel", lambda *a, **k: df_basico.copy(deep=True), raising=True)
        return df_basico
