"""Testes simples do relatório de vendas.

Rodar:  python -m unittest -v
"""
import contextlib
import io
import tempfile
import unittest
from pathlib import Path

import relatorio as rel

AQUI = Path(__file__).parent


class TestLeitura(unittest.TestCase):
    def _csv(self, conteudo):
        pasta = tempfile.TemporaryDirectory()
        self.addCleanup(pasta.cleanup)
        caminho = Path(pasta.name) / "v.csv"
        caminho.write_text(conteudo, encoding="utf-8")
        return caminho

    def test_coluna_faltando_gera_mensagem_clara(self):
        caminho = self._csv("data,produto,quantidade\n2026-10-01,Caneca,1\n")
        with self.assertRaises(rel.ErroPlanilha) as ctx:
            rel.ler(caminho)
        msg = str(ctx.exception)
        self.assertIn("coluna obrigatória: valor", msg)
        self.assertIn("Colunas encontradas: data, produto, quantidade", msg)

    def test_varias_colunas_faltando(self):
        caminho = self._csv("data,produto\n2026-10-01,Caneca\n")
        with self.assertRaises(rel.ErroPlanilha) as ctx:
            rel.ler(caminho)
        self.assertIn("quantidade, valor", str(ctx.exception))

    def test_aceita_ponto_e_virgula_e_valor_brasileiro(self):
        caminho = self._csv("Data;Produto;Quantidade;Valor\n2026-10-01;Caneca;2;1.234,56\n")
        vendas = rel.ler(caminho)
        self.assertEqual(len(vendas), 1)
        self.assertEqual(vendas[0]["valor"], 1234.56)

    def test_linha_invalida_e_ignorada_com_aviso(self):
        caminho = self._csv("data,produto,quantidade,valor\n"
                            "2026-10-01,Caneca,1,35\n"
                            "31/10/2026,Vela,1,28\n")
        saida = io.StringIO()
        with contextlib.redirect_stdout(saida):
            vendas = rel.ler(caminho)
        self.assertEqual(len(vendas), 1)
        self.assertIn("linha 3 ignorada", saida.getvalue())

    def test_arquivo_inexistente(self):
        with self.assertRaises(rel.ErroPlanilha):
            rel.ler(AQUI / "nao_existe.csv")


class TestCalculo(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.vendas = rel.ler(AQUI / "vendas.csv")

    def test_numeros_basicos(self):
        self.assertEqual(rel.num("84.00"), 84.0)
        self.assertEqual(rel.num("R$ 1.234,56"), 1234.56)

    def test_mes_anterior_na_virada_do_ano(self):
        self.assertEqual(rel.mes_anterior("2026-01"), "2025-12")
        self.assertEqual(rel.mes_anterior("2026-10"), "2026-09")

    def test_variacao(self):
        self.assertAlmostEqual(rel.variacao(150, 100), 0.5)
        self.assertIsNone(rel.variacao(10, 0))

    def test_outubro_do_exemplo(self):
        r = rel.calcular(self.vendas, "2026-10", 3)
        self.assertEqual(r["total"], 2039.0)
        self.assertEqual(r["itens"], 97)
        self.assertEqual(r["dias"], 24)
        self.assertEqual(r["melhor_dia"].isoformat(), "2026-10-10")
        self.assertEqual(r["ranking"][0][0], "Sabonete artesanal")
        self.assertAlmostEqual(r["comp"]["var_total"], (2039 - 1151) / 1151)

    def test_mes_sem_vendas(self):
        self.assertIsNone(rel.calcular(self.vendas, "2026-01", 3))

    def test_primeiro_mes_nao_tem_comparacao(self):
        r = rel.calcular(self.vendas, "2026-09", 3)
        self.assertIsNone(r["comp"])
        self.assertIn("não foi possível comparar", rel.narrativa(r, "2026-09", rel.numeracao(r))["comp"])


class TestSaidas(unittest.TestCase):
    def test_gera_excel_e_pdf(self):
        try:
            import openpyxl
            import reportlab  # noqa: F401
        except ImportError:
            self.skipTest("openpyxl/reportlab não instalados")
        vendas = rel.ler(AQUI / "vendas.csv")
        r = rel.calcular(vendas, "2026-10", 3)
        with tempfile.TemporaryDirectory() as pasta:
            xlsx, pdf = Path(pasta) / "r.xlsx", Path(pasta) / "r.pdf"
            rel.gerar_excel(r, "2026-10", xlsx, "vendas.csv")
            rel.gerar_pdf(r, "2026-10", pdf, "vendas.csv")
            self.assertEqual(openpyxl.load_workbook(xlsx).sheetnames,
                             ["Resumo", "Produtos", "Por dia", "Vendas do mês"])
            self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
