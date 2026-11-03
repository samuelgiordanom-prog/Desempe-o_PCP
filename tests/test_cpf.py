import argparse
import csv
import math
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
import cpf


class TestCPF(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def libro(self, name='datos.xlsx', fecha=datetime(2026, 1, 1), sheet='CPF Horario'):
        path = self.root / name
        w = Workbook()
        s = w.active
        s.title = sheet
        s.append(['Título'])
        s.append([None, 'Fecha', 'Hora', 'Unidad', 'Desempeño\n(DCPF)', 'Factor de Desempeño\n(Fd_CPF)', 'Respuesta CPF+', 'respuesta CPF-'])
        s.append([None, fecha, 0, '  HE ÁNGOSTURA U1 ', 0.75, 0.5])
        s.append([None, fecha, 0, 'HE ANGOSTURA U1', 'No Participó', None])
        s.append([None, fecha, 1, 'OTRA', 1, 1])
        w.save(path)
        w.close()
        return path

    def test_entrada_elige_cpf_y_deduplica(self):
        path = self.root / 'entrada.xlsx'
        w = Workbook()
        w.active.title = 'Control de Cambios'
        w.active.append(['Unidad'])
        w.active.append(['NO INCLUIR'])
        s = w.create_sheet('CPF')
        s.append(['Título'])
        s.append(['Unidad', 'Verificado'])
        s.append(['HE ÁNGOSTURA U1', 'SI'])
        s.append(['  he angostura u1  ', ' si '])
        s.append(['NO INCLUIR', 'NO'])
        s.append(['SIN VERIFICAR', None])
        s.append([None])
        w.save(path)
        self.assertEqual(cpf.leer_unidades(path), {'HE ANGOSTURA U1': 'HE ÁNGOSTURA U1'})

    def test_no_participo_en_ambas_respuestas(self):
        path = self.libro()
        from openpyxl import load_workbook
        w = load_workbook(path)
        sheet = w['CPF Horario']
        sheet.cell(3, 7, 'No Participó')
        sheet.cell(3, 8, ' no participó ')
        sheet.cell(4, 7, 'No Participó')
        sheet.cell(4, 8, 'Participó')
        w.save(path)
        w.close()
        rows = list(cpf.extraer(path, {'HE ANGOSTURA U1': 'u'}, cpf.Fuente(2026, 1, 1, 'url')))
        self.assertTrue(math.isnan(rows[0]['Desempeno']))
        self.assertTrue(math.isnan(rows[0]['Factor_desempeno']))
        csv_path = self.root / 'nan.csv'
        file, writer = cpf.escribir_csv(csv_path, ['Desempeno', 'Factor_desempeno'])
        with file:
            writer.writerow(rows[0])
        with csv_path.open(encoding='utf-8-sig') as file:
            exported = next(csv.DictReader(file))
        self.assertEqual(exported, {'Desempeno': 'nan', 'Factor_desempeno': 'nan'})
        self.assertEqual(rows[1]['Desempeno'], 'No Participó')
        self.assertIsNone(rows[1]['Factor_desempeno'])

    def test_csv_filtra_verificados_antes_de_deduplicar(self):
        path = self.root / 'entrada.csv'
        path.write_text('Unidad;Verificado\nA;NO\nA;SI\nB;NO\nC; si \nD;\n', encoding='utf-8')
        self.assertEqual(cpf.leer_unidades(path), {'A': 'A', 'C': 'C'})

    def test_preserva_metricas_horas_repetidas_y_vacios(self):
        rows = list(cpf.extraer(self.libro(), {'HE ANGOSTURA U1': 'HE ANGOSTURA U1'},
                                cpf.Fuente(2026, 1, 3, 'https://example.test/a.xlsx')))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]['Desempeno'], .75)
        self.assertEqual(rows[0]['Factor_desempeno'], .5)
        self.assertEqual(rows[1]['Desempeno'], 'No Participó')
        self.assertIsNone(rows[1]['Factor_desempeno'])
        self.assertEqual(rows[0]['Hora'], rows[1]['Hora'])

    def test_rechaza_mes_incorrecto_y_hoja_ausente(self):
        fuente = cpf.Fuente(2026, 2, 1, 'url')
        with self.assertRaisesRegex(ValueError, 'fuera del mes'):
            list(cpf.extraer(self.libro(), {'HE ANGOSTURA U1': 'u'}, fuente))
        with self.assertRaisesRegex(ValueError, 'falta la hoja'):
            list(cpf.extraer(self.libro('sin.xlsx', sheet='Otra'), {}, fuente))

    def test_descubre_ultima_version_sin_depender_del_orden(self):
        year = cpf.BASE_URL + '2026-indices-ejemplo'
        month = year + '/enero-2026'
        v1, v3 = month + '/version-1', month + '/version-3'
        class Fake:
            def pagina(self, url):
                return {cpf.BASE_URL: [year], year: [month], month: [v3, v1],
                        v3: ['https://example.test/a.zip', 'https://example.test/a.xlsx']}[url]
        self.assertEqual(cpf.descubrir(Fake(), 2025, 2026),
                         [cpf.Fuente(2026, 1, 3, 'https://example.test/a.xlsx')])

    def test_paginacion_y_duplicados(self):
        from io import BytesIO
        url = 'https://example.test/year'
        pages = {url: b'<a class="cen_btn cen_btn-primary" href="/enero">Ver</a><link rel="next" href="/year/page/2/">',
                 url + '/page/2': b'<a class="cen-text-link" href="/febrero">Ver</a><a class="cen-text-link" href="/enero">Ver</a>'}
        cliente = cpf.Cliente(self.root / 'cache')
        with patch.object(cliente, 'abrir', side_effect=lambda u: BytesIO(pages[u])):
            self.assertEqual(cliente.pagina(url), ['https://example.test/enero', 'https://example.test/febrero'])

    def test_version_mayor_vacia_elige_ultima_publicada(self):
        year = cpf.BASE_URL + '2026-indices-ejemplo'
        month = year + '/enero-2026'
        v2, v3 = month + '/version-2', month + '/version-3'
        class Fake:
            paginas_vacias = {v3}
            def pagina(self, url):
                return {cpf.BASE_URL: [year], year: [month], month: [v2, v3],
                        v3: [], v2: ['https://example.test/a.xlsx']}[url]
        self.assertEqual(cpf.descubrir(Fake(), 2026, 2026),
                         [cpf.Fuente(2026, 1, 2, 'https://example.test/a.xlsx', (3,))])

    def test_pagina_adicional_404_y_denegacion(self):
        from io import BytesIO
        from urllib.error import HTTPError
        url = 'https://example.test/year'
        cliente = cpf.Cliente(self.root / 'cache')
        html = b'<a class="cen_btn" href="/enero">Ver</a><link rel="next" href="/year/page/2">'
        with patch.object(cliente, 'abrir', side_effect=[BytesIO(html), HTTPError(url, 404, 'Not found', {}, None)]):
            self.assertEqual(cliente.pagina(url), ['https://example.test/enero'])
        with patch.object(cliente, 'abrir', side_effect=[BytesIO(html), HTTPError(url, 403, 'Forbidden', {}, None)]):
            with self.assertRaises(HTTPError):
                cliente.pagina(url)

    def test_salida_dos_series_y_reporte_ausentes(self):
        libro = self.libro()
        entrada = self.root / 'entrada.csv'
        entrada.write_text('Unidad\nHE ANGOSTURA U1\nAUSENTE\n', encoding='utf-8')
        args = argparse.Namespace(unidades=entrada, hoja_unidades='CPF', cache=self.root / 'cache',
                                  actualizar=False, anio=2026, hasta=2026, mes=None, salida=self.root / 'salida')
        fuente = cpf.Fuente(2026, 1, 3, 'https://example.test/a.xlsx')
        with patch('cpf.descubrir', return_value=[fuente]), patch.object(cpf.Cliente, 'descargar', return_value=libro):
            report = cpf.ejecutar(args)
        self.assertEqual(report['registros'], 2)
        self.assertEqual(report['unidades_sin_datos'], ['AUSENTE'])
        self.assertEqual(report['meses'][0]['horas_repetidas_conservadas'], 1)
        with (args.salida / 'factor_desempeno.csv').open(encoding='utf-8-sig') as f:
            rows = list(csv.DictReader(f))
        self.assertEqual(rows[0]['Factor_desempeno'], '0.5')
        self.assertNotIn('Desempeno', rows[0])
        self.assertEqual(len(list((args.salida / 'unidades').rglob('*.csv'))), 4)


if __name__ == '__main__':
    unittest.main()
