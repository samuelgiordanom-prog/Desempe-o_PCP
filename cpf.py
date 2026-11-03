"""Descarga y extracción de índices horarios de CPF del Coordinador Eléctrico."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import re
import shutil
import tempfile
import time
import unicodedata
from collections import OrderedDict
from dataclasses import dataclass, asdict
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urljoin, urlsplit
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo
from zipfile import BadZipFile

from openpyxl import load_workbook

BASE_URL = ('https://www.coordinador.cl/operacion/documentos/servicios-complementarios/'
            'indices-de-desempeno-y-disponibilidad-de-servicios-complementarios/')
MESES = 'enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre'.split()
LOG = logging.getLogger(__name__)


def normalizar(value):
    text = unicodedata.normalize('NFKD', str(value or ''))
    return ' '.join(''.join(c for c in text if not unicodedata.combining(c)).upper().split())


class Enlaces(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag in ('a', 'link') and attrs.get('href'):
            self.links.append(attrs)


class Cliente:
    def __init__(self, cache: Path, refresh=False):
        self.cache = cache
        self.refresh = refresh
        self.paginas_vacias = set()
        self.meses_sin_documentos = []
        cache.mkdir(parents=True, exist_ok=True)

    def abrir(self, url):
        for intento in range(3):
            try:
                return urlopen(Request(url, headers={'User-Agent': 'CPF-series/1.0'}), timeout=60)
            except HTTPError as exc:
                if exc.code not in (429, 500, 502, 503, 504) or intento == 2:
                    raise
            except (URLError, TimeoutError):
                if intento == 2:
                    raise
            time.sleep(2 ** intento)

    def pagina(self, url):
        """Lee enlaces de documentos y paginación; no reutiliza HTML obsoleto."""
        pendientes, vistos, documentos = [url], set(), set()
        root = url.rstrip('/')
        while pendientes:
            actual = pendientes.pop(0)
            if actual in vistos:
                continue
            if len(vistos) >= 100:
                raise ValueError(f'Paginación excesiva en {url}')
            vistos.add(actual)
            try:
                response = self.abrir(actual)
            except HTTPError as exc:
                # WordPress anuncia rel=next incluso en listados completos.
                # Solo un 404 de página adicional indica fin de paginación.
                if exc.code == 404 and actual != url:
                    LOG.info('Fin de paginación: %s', actual)
                    continue
                raise
            with response:
                parser = Enlaces()
                html = response.read().decode('utf-8')
                parser.feed(html)
                if actual == url and 'Sin documentos.' in html:
                    self.paginas_vacias.add(url)
            for attrs in parser.links:
                destino = urljoin(actual, attrs['href']).split('#')[0].rstrip('/')
                if (set(attrs.get('class', '').split()) & {'cen-text-link', 'cen_btn'}
                        or re.search(r'\.(xlsx|xlsm|zip)(?:\?|$)', destino, re.I)):
                    documentos.add(destino)
                # El sitio también presenta rel=next cuando todas las carpetas
                # ya están visibles. Se deduplican páginas y documentos.
                if (destino.startswith(root + '/page/')
                        and re.fullmatch(r'\d+', destino[len(root + '/page/'):])):
                    if destino not in vistos:
                        pendientes.append(destino)
        return sorted(documentos)

    def descargar(self, url):
        nombre = Path(unquote(urlsplit(url).path)).name
        target = self.cache / (hashlib.sha256(url.encode()).hexdigest()[:16] + '_' + nombre)
        if not target.exists() or self.refresh:
            LOG.info('Descargando %s', nombre)
            temp = target.with_suffix(target.suffix + '.part')
            try:
                with self.abrir(url) as response, temp.open('wb') as output:
                    shutil.copyfileobj(response, output)
                temp.replace(target)
            finally:
                temp.unlink(missing_ok=True)
        return target


@dataclass(frozen=True)
class Fuente:
    anio: int
    mes: int
    version: int
    url: str
    versiones_sin_documentos: tuple[int, ...] = ()


def descubrir(cliente, anio, hasta, meses=None):
    fuentes = []
    years = {}
    for url in cliente.pagina(BASE_URL):
        match = re.match(r'(\d{4})-indices-', urlsplit(url).path.rstrip('/').split('/')[-1])
        if match and anio <= int(match[1]) <= hasta:
            years[int(match[1])] = url
    if not years:
        raise ValueError('No se encontraron carpetas para los años solicitados.')
    for year, url in sorted(years.items()):
        months = {}
        for month_url in cliente.pagina(url):
            segment = urlsplit(month_url).path.rstrip('/').split('/')[-1]
            match = re.fullmatch(r'(' + '|'.join(MESES) + r')-' + str(year), segment)
            if match:
                month = MESES.index(match[1]) + 1
                if not meses or month in meses:
                    months[month] = month_url
        if not months:
            LOG.warning('Sin meses publicados para el filtro en %s', year)
        for month, month_url in sorted(months.items()):
            LOG.info('Buscando última versión: %s-%02d', year, month)
            versions = {}
            for version_url in cliente.pagina(month_url):
                segment = urlsplit(version_url).path.rstrip('/').split('/')[-1]
                match = re.search(r'version-(\d+)(?:-|$)', segment)
                if match:
                    versions.setdefault(int(match[1]), set()).add(version_url)
            if not versions:
                raise ValueError(f'No se encontraron versiones en {month_url}')
            vacias = []
            for version in sorted(versions, reverse=True):
                if len(versions[version]) != 1:
                    raise ValueError(f'Carpeta de versión ambigua en {month_url}')
                version_url = next(iter(versions[version]))
                documentos = cliente.pagina(version_url)
                files = [u for u in documentos
                         if urlsplit(u).path.lower().endswith(('.xlsx', '.xlsm'))]
                if version_url in getattr(cliente, 'paginas_vacias', set()):
                    LOG.warning('%s-%02d V%s: el sitio indica Sin documentos', year, month, version)
                    vacias.append(version)
                    continue
                if len(files) != 1:
                    raise ValueError(f'Se esperaba un Excel en {year}-{month:02d} V{version}; encontrados {files}')
                fuentes.append(Fuente(year, month, version, files[0], tuple(vacias)))
                break
            else:
                cliente.meses_sin_documentos.append({'anio': year, 'mes': month, 'versiones': vacias})
    if not fuentes:
        raise ValueError('No hay meses publicados que coincidan con los filtros.')
    return fuentes


def leer_unidades(path, hoja='CPF'):
    """Busca Unidad en las primeras 50 filas de la hoja elegida, sin mezclar servicios."""
    if path.suffix.lower() == '.csv':
        with path.open(encoding='utf-8-sig', newline='') as file:
            sample = file.read(4096)
            file.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
            except csv.Error:
                dialect = csv.excel
            rows = list(csv.reader(file, dialect))
    else:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if hoja not in workbook.sheetnames:
                raise ValueError(f'No existe la hoja {hoja!r}. Hojas: {workbook.sheetnames}')
            rows = list(workbook[hoja].values)
        finally:
            workbook.close()
    for i, row in enumerate(rows[:50]):
        columns = [j for j, value in enumerate(row) if normalizar(value) == 'UNIDAD']
        if columns:
            if len(columns) != 1:
                raise ValueError('Más de una columna Unidad en la cabecera.')
            col = columns[0]
            verified = [j for j, value in enumerate(row) if normalizar(value) == 'VERIFICADO']
            if len(verified) > 1 or (not verified and path.suffix.lower() != '.csv'):
                raise ValueError('Columna Verificado ausente o ambigua en la cabecera.')
            units = {}
            for values in rows[i + 1:]:
                if verified and (verified[0] >= len(values) or normalizar(values[verified[0]]) != 'SI'):
                    continue
                value = values[col] if col < len(values) else None
                if value is not None and str(value).strip() and normalizar(value) != 'UNIDAD':
                    units.setdefault(normalizar(value), str(value).strip())
            if not units:
                raise ValueError('La columna Unidad no contiene unidades elegibles (Verificado = SI cuando existe).')
            return units
    raise ValueError('No se encontró una columna Unidad en las primeras 50 filas.')


def extraer(path, unidades, fuente):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if 'CPF Horario' not in workbook.sheetnames:
            raise ValueError(f'{path.name}: falta la hoja CPF Horario')
        rows = iter(workbook['CPF Horario'].values)
        columns = None
        for i, row in enumerate(rows):
            names = [normalizar(v).split(' (')[0] for v in row]
            if all(k in names for k in ('FECHA', 'HORA', 'UNIDAD')):
                columns = {k: names.index(k) for k in ('FECHA', 'HORA', 'UNIDAD')}
                for label, prefix in [('desempeno', 'DESEMPENO'), ('factor', 'FACTOR DE DESEMPENO')]:
                    matches = [j for j, n in enumerate(names) if n == prefix or n.startswith(prefix + ' ')]
                    if len(matches) != 1:
                        raise ValueError(f'{path.name}: columna {prefix} ausente o ambigua')
                    columns[label] = matches[0]
                for label in ('RESPUESTA CPF+', 'RESPUESTA CPF-'):
                    matches = [j for j, n in enumerate(names) if n == label]
                    if len(matches) != 1:
                        raise ValueError(f'{path.name}: columna {label} ausente o ambigua')
                    columns[label] = matches[0]
                break
            if i >= 49:
                break
        if columns is None:
            raise ValueError(f'{path.name}: no se encontró la cabecera horaria')
        for row_number, row in enumerate(rows, i + 2):
            value = row[columns['UNIDAD']]
            key = normalizar(value)
            if key not in unidades:
                continue
            fecha, hora = row[columns['FECHA']], row[columns['HORA']]
            if not isinstance(fecha, (date, datetime)):
                raise ValueError(f'{path.name}, fila {row_number}: fecha inválida {fecha!r}')
            if (fecha.year, fecha.month) != (fuente.anio, fuente.mes):
                raise ValueError(f'{path.name}, fila {row_number}: fecha fuera del mes seleccionado')
            if hora is None:
                raise ValueError(f'{path.name}, fila {row_number}: hora vacía')
            # Se conservan hora original, textos y vacíos. No convertir vacíos
            # en cero ni eliminar horas repetidas por cambios de horario.
            no_participo = all(normalizar(row[columns[label]]) == 'NO PARTICIPO'
                               for label in ('RESPUESTA CPF+', 'RESPUESTA CPF-'))
            yield {'Unidad': unidades[key], 'Unidad_fuente': value,
                   'Fecha': fecha.strftime('%Y-%m-%d'), 'Hora': hora,
                   'Desempeno': float('nan') if no_participo else row[columns['desempeno']],
                   'Factor_desempeno': float('nan') if no_participo else row[columns['factor']],
                   'Version': fuente.version, 'Fuente': fuente.url}
    finally:
        workbook.close()


def escribir_csv(path, columns):
    file = path.open('w', encoding='utf-8-sig', newline='')
    writer = csv.DictWriter(file, fieldnames=columns, extrasaction='ignore')
    writer.writeheader()
    return file, writer


def ejecutar(args):
    unidades = leer_unidades(args.unidades, args.hoja_unidades)
    LOG.info('%s unidades únicas en %s', len(unidades), args.hoja_unidades)
    cliente = Cliente(args.cache, args.actualizar)
    fuentes = descubrir(cliente, args.anio, args.hasta, args.mes)
    output = args.salida
    if output.exists():
        raise ValueError(f'La salida ya existe: {output}. Elige otra carpeta para no sobrescribir resultados.')
    output.parent.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='.cpf-', dir=output.parent))
    reporte = {'anio_inicial': args.anio, 'anio_final': args.hasta,
               'meses_sin_documentos': cliente.meses_sin_documentos,
               'hoja_unidades': args.hoja_unidades, 'unidades_solicitadas': list(unidades.values()),
               'meses': [], 'archivos_por_unidad': {}}
    comunes = ['Unidad', 'Unidad_fuente', 'Fecha', 'Hora']
    columns = [comunes + [metric, 'Version', 'Fuente'] for metric in ('Desempeno', 'Factor_desempeno')]
    handles = []
    abiertos = OrderedDict()
    encontrados = set()
    try:
        generales = []
        for name, cols in zip(('desempeno.csv', 'factor_desempeno.csv'), columns):
            file, writer = escribir_csv(stage / name, cols)
            handles.append(file)
            generales.append(writer)
        # Cada unidad tiene sus propias dos series, incluso si no tiene datos.
        rutas = {}
        for index, (key, unidad) in enumerate(unidades.items(), 1):
            slug = re.sub(r'[^A-Za-z0-9_-]+', '_', normalizar(unidad))[:80]
            folder = Path('unidades') / f'{index:04d}_{slug}'
            (stage / folder).mkdir(parents=True)
            reporte['archivos_por_unidad'][unidad] = str(folder)
            rutas[key] = []
            for name, cols in zip(('desempeno.csv', 'factor_desempeno.csv'), columns):
                file, writer = escribir_csv(stage / folder / name, cols)
                file.close()
                rutas[key].append(stage / folder / name)

        def writers_unidad(key):
            # Limita los descriptores abiertos incluso con listados grandes.
            if key in abiertos:
                abiertos.move_to_end(key)
                return [writer for _, writer in abiertos[key]]
            if len(abiertos) >= 128:
                _, anteriores = abiertos.popitem(last=False)
                for file, _ in anteriores:
                    file.close()
            actuales = []
            abiertos[key] = actuales
            for path, cols in zip(rutas[key], columns):
                file = path.open('a', encoding='utf-8-sig', newline='')
                actuales.append((file, csv.DictWriter(file, fieldnames=cols, extrasaction='ignore')))
            return [writer for _, writer in actuales]
        for fuente in fuentes:
            LOG.info('Procesando %s-%02d V%s', fuente.anio, fuente.mes, fuente.version)
            path = cliente.descargar(fuente.url)
            count, found, vacios, repetidos, seen = 0, set(), 0, 0, set()
            for row in extraer(path, unidades, fuente):
                key = normalizar(row['Unidad'])
                found.add(key)
                identity = (key, row['Fecha'], row['Hora'])
                repetidos += identity in seen
                seen.add(identity)
                vacios += row['Desempeno'] is None or row['Factor_desempeno'] is None
                for writer in generales + writers_unidad(key):
                    writer.writerow(row)
                count += 1
            encontrados.update(found)
            reporte['meses'].append({**asdict(fuente), 'registros': count,
                'registros_con_metricas_vacias': vacios, 'horas_repetidas_conservadas': repetidos,
                'unidades_sin_datos': [v for k, v in unidades.items() if k not in found],
                'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
        reporte['unidades_sin_datos'] = [v for k, v in unidades.items() if k not in encontrados]
        with (stage / 'unidades_sin_datos.csv').open('w', encoding='utf-8-sig', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Unidad'])
            writer.writerows([unidad] for unidad in reporte['unidades_sin_datos'])
        reporte['registros'] = sum(m['registros'] for m in reporte['meses'])
        if reporte['registros'] == 0:
            raise ValueError('Ninguna unidad coincide con CPF Horario. Revisa los nombres del archivo de entrada.')
        for file in handles:
            file.close()
        for pares in abiertos.values():
            for file, _ in pares:
                file.close()
        (stage / 'reporte.json').write_text(json.dumps(reporte, ensure_ascii=False, indent=2), encoding='utf-8')
        stage.rename(output)
    finally:
        for file in handles:
            file.close()
        for pares in abiertos.values():
            for file, _ in pares:
                file.close()
        if stage.exists():
            shutil.rmtree(stage)
    LOG.info('%s registros; %s unidades sin coincidencia. Salida: %s',
             reporte['registros'], len(reporte['unidades_sin_datos']), output)
    return reporte


def main(argv=None):
    today = datetime.now(ZoneInfo('America/Santiago'))
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--unidades', type=Path, required=True, help='Excel o CSV con columna Unidad')
    parser.add_argument('--hoja-unidades', default='CPF', help='Hoja del Excel de entrada (CPF por defecto)')
    parser.add_argument('--anio', type=int, required=True, help='Año inicial, incluido')
    parser.add_argument('--hasta', type=int, default=today.year, help='Año final, incluido (actual por defecto)')
    parser.add_argument('--mes', type=int, nargs='+', choices=range(1, 13), help='Filtro opcional de meses')
    parser.add_argument('--salida', type=Path, default=Path('resultados') / today.strftime('cpf_%Y%m%d_%H%M%S'))
    parser.add_argument('--cache', type=Path, default=Path('.cache/cpf'))
    parser.add_argument('--actualizar', action='store_true', help='Volver a descargar los Excel almacenados')
    args = parser.parse_args(argv)
    if not 2000 <= args.anio <= args.hasta <= today.year:
        parser.error('Los años deben estar entre 2000 y el año actual, con anio <= hasta.')
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    try:
        ejecutar(args)
    except (OSError, ValueError, KeyError, BadZipFile) as exc:
        LOG.error('%s', exc)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
