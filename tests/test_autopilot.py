import contextlib
import importlib.util
import io
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('autopilot', ROOT / 'kit/autopilot.py')
ap = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ap)


class AutopilotTests(unittest.TestCase):
    def setUp(self):
        ap.PECAS.clear()

    def test_motor_copies_equal(self):
        self.assertEqual((ROOT/'kit/autopilot.py').read_bytes(), (ROOT/'vscode/autopilot.py').read_bytes())

    def test_brazilian_numbers(self):
        for text, value in [('R$ 1.234,56',1234.56),('-R$ 1.234,56',-1234.56),('R$ -12,30',-12.3),(' 1234 ',1234),('1.234',1234),(1500.5,1500.5)]:
            self.assertEqual(ap.numero_br(text),value)

    def test_reject_ambiguous_or_missing_numbers(self):
        for text in ['1234.56','1,234.56','',None,float('nan'),True,'--1','R$ 1.23']:
            with self.subTest(text=text), self.assertRaises(ValueError): ap.numero_br(text)

    def test_rules_do_not_approve_arbitrary_text(self):
        self.assertFalse(ap.checar_regras('NUNCA use Python. Explique em passos com JavaScript.'))
        self.assertTrue(ap.checar_regras(ap.REGRAS_TUTOR))

    def test_context_sent(self):
        request=ap.preparar_pergunta('Como converto dinheiro?')
        self.assertIn('Python 3',request)
        self.assertIn('para_numero',request)
        self.assertIn('Como converto dinheiro?',request)

    def test_invalid_responses(self):
        for answer in [None,'','  ','```javascript\nparseFloat(x)\n```']:
            with self.assertRaises(ValueError): ap.validar_resposta(answer)
        self.assertEqual(ap.validar_resposta(' Use Python. '),'Use Python.')

    def test_piece_rejection_invalidates_old_function(self):
        def para_numero(x): return ap.numero_br(x)
        ap.testar(para_numero)
        self.assertIn('para_numero',ap.PECAS)
        def para_numero(x): return 0
        ap.testar(para_numero)
        self.assertNotIn('para_numero',ap.PECAS)

    def model(self):
        return ap.Modelo('Teste', fatos=pd.DataFrame({'Valor':[1.,2.]}))

    def test_reexecuted_measure_replaces(self):
        m=self.model()
        m.medidas({'Total':'SUM(Fatos[Valor])'})
        m.medidas({'Total':'SUM(Fatos[Valor]) * 2'})
        self.assertEqual(len(m.lista_medidas),1)
        self.assertIn('* 2',m.lista_medidas[0][2])

    def test_bad_measure_does_not_mutate_existing_list(self):
        m=self.model()
        m.medidas({'Total':'SUM(Fatos[Valor])'})
        before=list(m.lista_medidas)
        with self.assertRaises(ValueError): m.medidas({'Nova':'SUM(Fatos[Valor])','Invalida':'SUM(Inexistente[Valor])'})
        self.assertEqual(m.lista_medidas,before)

    def test_export_rejects_missing_model_before_deleting(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(ap,'PASTA_PROJETO',Path(folder)/'projeto'):
            ap.PASTA_PROJETO.mkdir()
            marker=ap.PASTA_PROJETO/'marker.txt';marker.write_text('keep')
            with self.assertRaises(ValueError): ap.Painel('Teste').baixar()
            self.assertTrue(marker.exists())

    def test_failed_export_preserves_previous_project(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(ap,'PASTA_PROJETO',Path(folder)/'projeto'):
            ap.PASTA_PROJETO.mkdir()
            marker=ap.PASTA_PROJETO/'marker.txt';marker.write_text('keep')
            m=self.model()
            with patch.object(m,'salvar',side_effect=RuntimeError('failure')):
                with self.assertRaises(RuntimeError): ap.Painel(m).baixar()
            self.assertTrue(marker.exists())
            self.assertEqual(ap.PASTA_PROJETO,Path(folder)/'projeto')

    def test_successful_export_preserves_other_project(self):
        cwd=Path.cwd()
        with tempfile.TemporaryDirectory() as folder:
            try:
                os.chdir(folder)
                Path('projeto').mkdir()
                Path('projeto/Outro.pbip').write_text('keep')
                m=self.model();m.medidas({'Total':'SUM(Fatos[Valor])'})
                p=ap.Painel(m);p.cartao('Total')
                output=p.baixar()
                self.assertTrue(Path(output).is_file())
                self.assertTrue(Path('projeto/Outro.pbip').exists())
                self.assertTrue(Path('projeto/Teste.pbip').exists())
            finally: os.chdir(cwd)

    def test_secret_error_does_not_show_secret(self):
        secret='gsk_'+'x'*40
        self.assertNotIn(secret,ap.erro_seguro(ValueError('Bad key '+secret)))
        self.assertNotIn(secret,ap.motivo(ValueError('Bad key '+secret)))

    def test_zip_failure_rolls_back_project(self):
        cwd=Path.cwd()
        with tempfile.TemporaryDirectory() as folder:
            try:
                os.chdir(folder)
                Path('projeto').mkdir()
                Path('projeto/Teste.pbip').write_text('previous')
                m=self.model();m.medidas({'Total':'SUM(Fatos[Valor])'})
                with patch.object(ap.os,'replace',side_effect=OSError('zip locked')):
                    with self.assertRaises(OSError):ap.Painel(m).baixar()
                self.assertEqual(Path('projeto/Teste.pbip').read_text(),'previous')
                self.assertFalse(Path('projeto/Teste.Report').exists())
            finally:os.chdir(cwd)


if __name__=='__main__': unittest.main()
