from flask import Flask, flash, render_template, request, redirect, url_for, send_file, session, jsonify
import psycopg2
from psycopg2 import OperationalError, DatabaseError
from datetime import datetime, date, time
import openpyxl
import re
import os
import csv
import io as _stdio_io
import time as _time
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "pco_linea2_secret_key_dev")

CV_MAP = {
    'E20': ['2001', '2002'],
    'E21': ['2101', '2102'],
    'E22': ['2203', '2204'],
    'E23': ['2305', '2304'],
    'E24': ['2405', '2402'],
    'PV19': ['PV19'],
    'PV20': ['PV20'],
    'PV21': ['PV21'],
    'PV22': ['PV22'],
    'PV23': ['PV23'],
    'PV24': ['PV24'],
    'E20->E21': ['2001', '2003', '2005', '2101', '2103', '2003a', '2003b', '2101a', '2101b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2102'],
    'E20->E22': ['2001', '2003', '2005', '2101', '2103', '2105', '2107', '2201', '2203', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2102', '2106'],
    'E20->E23': ['2001', '2003', '2005', '2101', '2103', '2105', '2107', '2201', '2203', '2205', '2301', '2303', '2305', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2302a', '2032b', '2304', '2102', '2106', '2302'],
    'E20->E24': ['2001', '2003', '2005', '2101', '2103', '2105', '2107', '2201', '2203', '2205', '2301', '2303', '2305', '2307', '2401', '2403', '2405', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2307a', '2307b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2302a', '2032b', '2304', '2306', '2308', '2402', '2102', '2106', '2302'],
    'E21->E22': ['2103', '2105', '2107', '2201', '2203', '2105a', '2105b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2106'],
    'E21->E23': ['2103', '2105', '2107', '2201', '2203', '2205', '2301', '2303', '2305', '2105a', '2105b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2302a', '2032b', '2304', '2106', '2302'],
    'E21->E24': ['2103', '2105', '2107', '2201', '2203', '2205', '2301', '2303', '2305', '2307', '2401', '2403', '2405', '2105a', '2105b', '2307a', '2307b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2302a', '2032b', '2304', '2306', '2308', '2402', '2106', '2302'],
    'E22->E23': ['2203', '2205', '2301', '2303', '2305', '2204', '2206', '2302a', '2032b', '2304', '2302'],
    'E22->E24': ['2203', '2205', '2301', '2303', '2305', '2307', '2401', '2403', '2405', '2307a', '2307b', '2204', '2206', '2302a', '2032b', '2304', '2306', '2308', '2402', '2302'],
    'E23->E24': ['2305', '2307', '2401', '2403', '2405', '2307a', '2307b', '2304', '2306', '2308', '2402'],
    'E20->PV19': ['2007', '2001', '2010', '2002', 'PV19 BIS'],
    'E20->PV20': ['2001', '2003', '2003a', '2003b', '2002', '2004', '2006', 'PV20'],
    'E20->PV21': ['2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2102', '2106', '2001', '2003', '2005', '2101', '2103', '2105', '2107', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', 'PV21'],
    'E20->PV22': ['2001', '2003', '2005', '2101', '2103', '2105', '2107', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2102', '2106', 'PV21', '2001', '2003', '2005', '2101', '2103', '2105', '2107', '2201', '2203', '2205', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2102', '2106', 'PV22'],
    'E20->PV23': ['2001', '2003', '2005', '2101', '2103', '2105', '2107', '2201', '2203', '2205', '2301', '2303', '2305', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2302a', '2032b', '2304', '2102', '2106', '2302', 'PV23'],
    'E20->PV24': ['2001', '2003', '2005', '2101', '2103', '2105', '2107', '2201', '2203', '2205', '2301', '2303', '2305', '2307', '2401', '2403', '2405', '2407', '2409', '2411', '2003a', '2003b', '2101a', '2101b', '2105a', '2105b', '2307a', '2307b', '2002', '2004', '2006', '2008', '2102a', '2102b', '2104', '2106a', '2106b', '2108', '2002', '2204', '2206', '2302a', '2032b', '2304', '2306', '2308', '2402', '2404', '2406', '2102', '2106', '2302', 'PV24'],
    'ZM': ['ZM', '1012'],
    'ZA': ['1011', '1010', '1008b', '1008a', '1006', '1004b', '1004a', '1002', '1008', '1004', '1009', '1007', '1005', '1003', '1001', '1007b', '1007a', '1003b', '1003a', '1009', '1009', '1013', '1013', '1013', '1027', '1027', '1028', '1029', '1029', '1051', 'TK3', '1013', '126', '1030', '1036', '1013', '1013', '1033', '1034', '1035', '1043', 'TK TEST', '1042', '1010', '1031', '1032b', '1032a', '1032'],
    'TK1': ['1012', '1011', '1010', '1008b', '1008a', '1006', '1004b', '1004a', '1002', '1002', '1008', '1004', '1002'],
    'TK2': ['1001', '1009', '1007', '1005', '1003', '1001', '1001', '1007b', '1007a', '1003b', '1003a', '1009', '1009'],
    'TK3': ['1013', '1013', '1027', '1030', '1036'],
    'TK4': ['1013', '1013', '1013', '1033', '1034'],
    'TK5': ['1028'],
    'TK_TEST': ['1032', '1032a', '1032b', '1035', '1031'],
    'TESTTRACK': ['1032', '1032a', '1032b', '1035', '1031'],
}

ZONE_POSITIONS = {
    'D1': [{'top': 88, 'left': 1257, 'width': 132, 'height': 469, 'label': 'D1'}],
    'D2': [{'top': 22, 'left': 1404, 'width': 114, 'height': 547, 'label': 'D2'}],
    'E20': [{'top': 540, 'left': 162, 'width': 31, 'height': 96, 'label': 'E20'}],
    'E20->E21': [
        {'top': 542, 'left': 167, 'width': 403, 'height': 95, 'label': 'E20->E21'},
        {'top': 542, 'left': 562, 'width': 269, 'height': 95, 'label': 'E20->E21'},
    ],
    'E20->E22': [{'top': 542, 'left': 178, 'width': 660, 'height': 93, 'label': 'E20->E22'}],
    'E20->E23': [{'top': 543, 'left': 178, 'width': 909, 'height': 93, 'label': 'E20->E23'}],
    'E20->E24': [{'top': 543, 'left': 174, 'width': 1204, 'height': 93, 'label': 'E20->E24'}],
    'E20->PV19': [
        {'top': 542, 'left': 34, 'width': 141, 'height': 95, 'label': 'E20->PV19'},
        {'top': 505, 'left': 34, 'width': 41, 'height': 11, 'label': 'E20->PV19'},
    ],
    'E20->PV20': [
        {'top': 542, 'left': 165, 'width': 209, 'height': 95, 'label': 'E20->PV20'},
        {'top': 505, 'left': 344, 'width': 31, 'height': 13, 'label': 'E20->PV20'},
        {'top': 543, 'left': 371, 'width': 459, 'height': 95, 'label': 'E20->PV20'},
    ],
    'E20->PV21': [{'top': 542, 'left': 178, 'width': 569, 'height': 95, 'label': 'E20->PV21'}],
    'E20->PV22': [{'top': 542, 'left': 176, 'width': 732, 'height': 94, 'label': 'E20->PV22'}],
    'E20->PV23': [{'top': 542, 'left': 172, 'width': 955, 'height': 94, 'label': 'E20->PV23'}],
    'E20->PV24': [{'top': 542, 'left': 176, 'width': 1420, 'height': 95, 'label': 'E20->PV24'}],
    'E21': [{'top': 539, 'left': 552, 'width': 32, 'height': 99, 'label': 'E21'}],
    'E21->E20': [{'top': 544, 'left': 174, 'width': 393, 'height': 90, 'label': 'E21->E20'}],
    'E21->E22': [{'top': 542, 'left': 566, 'width': 271, 'height': 95, 'label': 'E21->E22'}],
    'E21->E23': [{'top': 543, 'left': 563, 'width': 523, 'height': 95, 'label': 'E21->E23'}],
    'E21->E24': [{'top': 541, 'left': 562, 'width': 816, 'height': 96, 'label': 'E21->E24'}],
    'E21->PV19': [{'top': 542, 'left': 32, 'width': 537, 'height': 95, 'label': 'E21->PV19'}],
    'E21->PV20': [{'top': 543, 'left': 371, 'width': 199, 'height': 95, 'label': 'E21->PV20'}],
    'E21->PV21': [{'top': 542, 'left': 565, 'width': 184, 'height': 95, 'label': 'E21->PV21'}],
    'E21->PV22': [{'top': 543, 'left': 563, 'width': 344, 'height': 95, 'label': 'E21->PV22'}],
    'E21->PV23': [{'top': 541, 'left': 564, 'width': 562, 'height': 97, 'label': 'E21->PV23'}],
    'E21->PV24': [{'top': 541, 'left': 563, 'width': 1031, 'height': 97, 'label': 'E21->PV24'}],
    'E22': [{'top': 540, 'left': 816, 'width': 30, 'height': 95, 'label': 'E22'}],
    'E22->E20': [{'top': 543, 'left': 173, 'width': 657, 'height': 94, 'label': 'E22->E20'}],
    'E22->E23': [{'top': 541, 'left': 829, 'width': 256, 'height': 96, 'label': 'E22->E23'}],
    'E22->E24': [{'top': 543, 'left': 828, 'width': 551, 'height': 92, 'label': 'E22->E24'}],
    'E22->PV19': [{'top': 540, 'left': 32, 'width': 799, 'height': 97, 'label': 'E22->PV19'}],
    'E22->PV21': [{'top': 542, 'left': 745, 'width': 85, 'height': 97, 'label': 'E22->PV21'}],
    'E22->PV22': [{'top': 543, 'left': 829, 'width': 79, 'height': 93, 'label': 'E22->PV22'}],
    'E22->PV24': [{'top': 543, 'left': 827, 'width': 768, 'height': 93, 'label': 'E22->PV24'}],
    'E22-PV23': [{'top': 542, 'left': 831, 'width': 295, 'height': 95, 'label': 'E22-PV23'}],
    'E23': [{'top': 540, 'left': 1065, 'width': 30, 'height': 98, 'label': 'E23'}],
    'E23->E20': [{'top': 542, 'left': 173, 'width': 907, 'height': 96, 'label': 'E23->E20'}],
    'E23->E21': [{'top': 543, 'left': 563, 'width': 517, 'height': 97, 'label': 'E23->E21'}],
    'E23->E22': [{'top': 542, 'left': 826, 'width': 255, 'height': 97, 'label': 'E23->E22'}],
    'E23->E24': [{'top': 541, 'left': 1079, 'width': 299, 'height': 95, 'label': 'E23->E24'}],
    'E23->PV19': [{'top': 542, 'left': 33, 'width': 1047, 'height': 95, 'label': 'E23->PV19'}],
    'E23->PV20': [{'top': 542, 'left': 371, 'width': 708, 'height': 95, 'label': 'E23->PV20'}],
    'E23->PV21': [{'top': 542, 'left': 747, 'width': 334, 'height': 94, 'label': 'E23->PV21'}],
    'E23->PV22': [{'top': 543, 'left': 904, 'width': 177, 'height': 95, 'label': 'E23->PV22'}],
    'E23->PV23': [{'top': 541, 'left': 1078, 'width': 48, 'height': 94, 'label': 'E23->PV23'}],
    'E23->PV24': [{'top': 542, 'left': 1078, 'width': 518, 'height': 95, 'label': 'E23->PV24'}],
    'E24': [{'top': 540, 'left': 1350, 'width': 45, 'height': 98, 'label': 'E24'}],
    'E24->E20': [{'top': 542, 'left': 173, 'width': 1199, 'height': 95, 'label': 'E24->E20'}],
    'E24->E21': [{'top': 541, 'left': 564, 'width': 812, 'height': 96, 'label': 'E24->E21'}],
    'E24->E22': [{'top': 542, 'left': 826, 'width': 546, 'height': 96, 'label': 'E24->E22'}],
    'E24->E23': [{'top': 543, 'left': 1076, 'width': 301, 'height': 93, 'label': 'E24->E23'}],
    'E24->PV19': [{'top': 542, 'left': 31, 'width': 1343, 'height': 96, 'label': 'E24->PV19'}],
    'E24->PV20': [{'top': 542, 'left': 370, 'width': 1005, 'height': 96, 'label': 'E24->PV20'}],
    'E24->PV21': [{'top': 542, 'left': 745, 'width': 630, 'height': 95, 'label': 'E24->PV21'}],
    'E24->PV22': [{'top': 543, 'left': 906, 'width': 464, 'height': 96, 'label': 'E24->PV22'}],
    'E24->PV23': [{'top': 542, 'left': 1123, 'width': 254, 'height': 93, 'label': 'E24->PV23'}],
    'E24->PV24': [{'top': 543, 'left': 1367, 'width': 228, 'height': 94, 'label': 'E24->PV24'}],
    'PV19': [{'top': 502, 'left': 33, 'width': 44, 'height': 16, 'label': 'PV19'}],
    'PV20': [{'top': 503, 'left': 343, 'width': 31, 'height': 16, 'label': 'PV20'}],
    'PV21': [{'top': 501, 'left': 718, 'width': 32, 'height': 17, 'label': 'PV21'}],
    'PV22': [{'top': 501, 'left': 877, 'width': 29, 'height': 18, 'label': 'PV22'}],
    'PV23': [{'top': 500, 'left': 1092, 'width': 32, 'height': 20, 'label': 'PV23'}],
    'PV24': [{'top': 503, 'left': 1570, 'width': 26, 'height': 16, 'label': 'PV24'}],
    'TK_TEST': [{'top': 442, 'left': 188, 'width': 911, 'height': 40, 'label': 'TK TEST'}],
    'TK1': [
        {'top': 89, 'left': 291, 'width': 1098, 'height': 36, 'label': 'TK1'},
        {'top': 95, 'left': 293, 'width': 47, 'height': 166, 'label': 'TK1'},
        {'top': 87, 'left': 1324, 'width': 62, 'height': 351, 'label': 'TK1'},
    ],
    'TK2': [
        {'top': 22, 'left': 160, 'width': 74, 'height': 164, 'label': 'TK2'},
        {'top': 20, 'left': 162, 'width': 1324, 'height': 41, 'label': 'TK2'},
        {'top': 20, 'left': 1407, 'width': 77, 'height': 415, 'label': 'TK2'},
    ],
    'TK3': [
        {'top': 269, 'left': 233, 'width': 223, 'height': 47, 'label': 'TK3'},
        {'top': 303, 'left': 431, 'width': 54, 'height': 68, 'label': 'TK3'},
        {'top': 333, 'left': 487, 'width': 612, 'height': 36, 'label': 'TK3'},
    ],
    'TK4': [{'top': 377, 'left': 277, 'width': 827, 'height': 50, 'label': 'TK4'}],
    'TK5': [{'top': 282, 'left': 456, 'width': 390, 'height': 35, 'label': 'TK5'}],
    'ZM': [
        {'top': 225, 'left': 452, 'width': 297, 'height': 37, 'label': 'ZM'},
        {'top': 198, 'left': 547, 'width': 71, 'height': 36, 'label': 'ZM'},
    ],
    'TK7': [{'top': 442, 'left': 188, 'width': 911, 'height': 40, 'label': 'TK TEST'}],
    'PTSA': [
        {'top': 89, 'left': 291, 'width': 1098, 'height': 36, 'label': 'PTSA'},
        {'top': 22, 'left': 160, 'width': 74, 'height': 164, 'label': 'PTSA'},
        {'top': 20, 'left': 162, 'width': 1324, 'height': 41, 'label': 'PTSA'},
    ],
    'TKTEST': [{'top': 442, 'left': 188, 'width': 911, 'height': 40, 'label': 'TK TEST'}],
    'TESTTRACK': [{'top': 442, 'left': 188, 'width': 911, 'height': 40, 'label': 'TK TEST'}],
}


def get_rects(zona):
    entry = ZONE_POSITIONS.get(zona)
    if entry is None:
        return []
    if isinstance(entry, list):
        return entry
    return [entry]


def get_cvs_for_zonas(ubicacion_zona_str):
    partes = normalizar_zonas(ubicacion_zona_str)
    cvs_encontrados = set()
    for parte in partes:
        if parte in CV_MAP:
            for cv in CV_MAP[parte]:
                cvs_encontrados.add(cv)
    return sorted(cvs_encontrados)


def get_zonas_con_info(ubicacion_zona_str):
    partes = normalizar_zonas(ubicacion_zona_str)
    resultado = []
    for parte in partes:
        rects = get_rects(parte)
        cv_list = CV_MAP.get(parte, [])
        resultado.append({
            'nombre': parte,
            'rects': rects,
            'cvs': cv_list
        })
    return resultado

def get_db_connection(max_retries=3, delay=1):
    database_url = os.environ.get("DATABASE_URL")
    last_err = None
    for i in range(max_retries):
        try:
            if database_url:
                return psycopg2.connect(database_url, connect_timeout=10)
            return psycopg2.connect(dbname="pco_db", user="ardalax", host="localhost", port="5432")
        except OperationalError as e:
            last_err = e
            if i < max_retries - 1:
                _time.sleep(delay * (i + 1))
            continue
    raise last_err


def normalizar_zonas(zonas):
    if not zonas:
        return set()

    if isinstance(zonas, str):
        partes = zonas.replace(',', '+').split('+')
    else:
        partes = zonas

    zonas_normalizadas = set()
    for parte in partes:
        zona = parte.strip().replace(" ", "").upper()
        if zona:
            zonas_normalizadas.add(zona)
    return zonas_normalizadas


def evaluar_safety_ingreso(ubicacion_zona, trabajos_activos, usa_vehiculo, tipo_vehiculo, codigo_vehiculo, empresa):
    partes_nuevas = normalizar_zonas(ubicacion_zona)

    for trab in trabajos_activos:
        empresa_activa = trab[0]
        zona_activa_cadena = trab[1]
        en_via_usa_vehiculo = bool(trab[2])
        en_via_tipo_vehiculo = trab[3] or "Vehículo Auxiliar"
        en_via_codigo_vehiculo = trab[4] or "sin código"

        partes_activas = normalizar_zonas(zona_activa_cadena)
        coincidencias = partes_nuevas.intersection(partes_activas)

        if not coincidencias:
            continue

        if en_via_usa_vehiculo or usa_vehiculo:
            if en_via_usa_vehiculo and usa_vehiculo:
                mensaje = f"❌ RECHAZADO POR SAFETY: La zona '{', '.join(sorted(coincidencias))}' ya tiene movimiento de {en_via_tipo_vehiculo} ({en_via_codigo_vehiculo}) por la empresa '{empresa_activa}' y no se permite otro vehículo en la misma zona."
            elif en_via_usa_vehiculo:
                mensaje = f"❌ RECHAZADO POR SAFETY: La zona '{', '.join(sorted(coincidencias))}' ya tiene movimiento de {en_via_tipo_vehiculo} ({en_via_codigo_vehiculo}) por la empresa '{empresa_activa}'."
            else:
                vehiculo_label = tipo_vehiculo or "vehículo auxiliar"
                codigo_label = codigo_vehiculo or "sin código"
                mensaje = f"❌ RECHAZADO POR SAFETY: No puedes ingresar el {vehiculo_label} ({codigo_label}) a la zona '{', '.join(sorted(coincidencias))}' porque actualmente hay una cuadrilla peatonal de '{empresa_activa}' trabajando ahí."
            return True, mensaje

    return False, ""

@app.route('/')
def index():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, fecha, turno, empresa, orden_trabajo, responsable, ubicacion_zona, 
               num_personas, vaf_tren, conductor, tetra, hora_inicio, hora_fin, 
               operador_turno, spco_turno, estado, 
               usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo
        FROM seguimiento_vias 
        WHERE archivado = FALSE AND (fecha = CURRENT_DATE OR estado = 'En Vía')
        ORDER BY (hora_fin IS NOT NULL) ASC, hora_inicio DESC;
    """)
    registros = cur.fetchall()
    cur.close()
    conn.close()
    
    turno_actual = session.get('turno', 'Turno 1')
    operador_actual = session.get('operador_turno', '')
    spco_actual = session.get('spco_turno', '')
    
    # LA MAGIA PARA QUE NO SE BORRE: Leemos y limpiamos la memoria al mismo tiempo
    form_previo = session.pop('form_previo', {})
    
    importados_pendientes = session.get('importados_pendientes', [])
    
    return render_template('index.html', 
                           registros=registros, 
                           turno_actual=turno_actual, 
                           operador_actual=operador_actual, 
                           spco_actual=spco_actual,
                           form_previo=form_previo,
                           importados_pendientes=importados_pendientes)

@app.route('/mapa')
def mapa():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT ubicacion_zona, empresa, usa_vehiculo, tipo_vehiculo, codigo_vehiculo, responsable, num_personas, tetra FROM seguimiento_vias WHERE hora_fin IS NULL;")
    registros_activos = cur.fetchall()
    cur.close()
    conn.close()

    zonas_ocupadas = []
    zonas_data = []
    trabajos_resumen = []
    for reg in registros_activos:
        cadena_zona = reg[0]
        empresa = reg[1] or ''
        usa_vehiculo = bool(reg[2])
        tipo_vehiculo = reg[3] or ''
        codigo_vehiculo = reg[4] or ''
        responsable = reg[5] or ''
        num_personas = reg[6] or 0
        tetra = reg[7] or ''

        trabajos_resumen.append({
            'zona': cadena_zona,
            'empresa': empresa,
            'responsable': responsable,
            'num_personas': num_personas,
            'tetra': tetra,
            'usa_vehiculo': usa_vehiculo,
            'tipo_vehiculo': tipo_vehiculo,
            'codigo_vehiculo': codigo_vehiculo,
        })

        partes = cadena_zona.replace(',', '+').split('+')
        for parte in partes:
            zona_limpia = parte.strip().replace(" ", "").upper()
            if zona_limpia and zona_limpia not in zonas_ocupadas:
                zonas_ocupadas.append(zona_limpia)
            if zona_limpia:
                rects = get_rects(zona_limpia)
                if rects:
                    for rect in rects:
                        zonas_data.append({
                            'nombre': zona_limpia,
                            'top': rect.get('top', 0),
                            'left': rect.get('left', 0),
                            'width': rect.get('width', 100),
                            'height': rect.get('height', 50),
                            'label': rect.get('label', zona_limpia),
                            'empresa': empresa,
                            'usa_vehiculo': usa_vehiculo,
                            'tipo_vehiculo': tipo_vehiculo,
                            'codigo_vehiculo': codigo_vehiculo,
                            'responsable': responsable,
                            'num_personas': num_personas,
                            'tetra': tetra,
                        })
                else:
                    zonas_data.append({
                        'nombre': zona_limpia,
                        'top': 0, 'left': 0, 'width': 100, 'height': 50,
                        'label': zona_limpia,
                        'empresa': empresa,
                        'usa_vehiculo': usa_vehiculo,
                        'tipo_vehiculo': tipo_vehiculo,
                        'codigo_vehiculo': codigo_vehiculo,
                        'responsable': responsable,
                        'num_personas': num_personas,
                        'tetra': tetra,
                    })

    return render_template('mapa.html', zonas_ocupadas=zonas_ocupadas, zonas_data=zonas_data, trabajos_resumen=trabajos_resumen)

@app.route('/coordenadas')
def coordenadas():
    return render_template('coordinate_finder.html')

@app.route('/configurar_turno', methods=['POST'])
def configurar_turno():
    session['turno'] = request.form['turno']
    session['operador_turno'] = request.form['operador_turno']
    session['spco_turno'] = request.form['spco_turno']
    return redirect(url_for('index'))

@app.route('/verificar_safety', methods=['POST'])
def verificar_safety():
    data = request.get_json() or {}
    ubicacion_zona = data.get('ubicacion_zona', '')
    usa_vehiculo = True if data.get('usa_vehiculo') == 'si' or data.get('usa_vehiculo') is True else False
    tipo_vehiculo = data.get('tipo_vehiculo', '')
    codigo_vehiculo = data.get('codigo_vehiculo', '')
    empresa = data.get('empresa', '')

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT empresa, ubicacion_zona, usa_vehiculo, tipo_vehiculo, codigo_vehiculo 
        FROM seguimiento_vias 
        WHERE hora_fin IS NULL;
    """)
    trabajos_activos = cur.fetchall()
    cur.close()
    conn.close()

    bloqueo, mensaje = evaluar_safety_ingreso(
        ubicacion_zona,
        trabajos_activos,
        usa_vehiculo,
        tipo_vehiculo,
        codigo_vehiculo,
        empresa
    )

    return jsonify({
        "bloqueo": bloqueo,
        "mensaje": mensaje
    })

@app.route('/ingresar', methods=['POST'])
def ingresar():
    if request.method == 'POST':
        turno = session.get('turno', 'Turno 1')
        operador = session.get('operador_turno', 'No Registrado')
        spco = session.get('spco_turno', 'No Registrado')
        
        empresa = request.form['empresa']
        orden_trabajo = request.form.get('orden_trabajo', '-')
        responsable = request.form['responsable']
        ubicacion_zona = request.form['ubicacion_zona']
        num_personas = request.form['num_personas'] or 0
        tetra = request.form['tetra']

        usa_vehiculo = True if request.form.get('usa_vehiculo') == 'si' else False
        tipo_vehiculo = request.form.get('tipo_vehiculo', '')
        codigo_vehiculo = request.form.get('codigo_vehiculo', '')
        conductor_vehiculo = request.form.get('conductor_vehiculo', '')

        # GUARDAMOS TODO EN LA SESIÓN (Por si hay rechazo de Safety)
        session['form_previo'] = request.form.to_dict()

        partes_nuevas = [p.strip().replace(" ", "").upper() for p in ubicacion_zona.replace(',', '+').split('+') if p.strip()]

        conn = get_db_connection()
        cur = conn.cursor()
        
        cur.execute("""
            SELECT empresa, ubicacion_zona, usa_vehiculo, tipo_vehiculo, codigo_vehiculo 
            FROM seguimiento_vias 
            WHERE hora_fin IS NULL;
        """)
        trabajos_activos = cur.fetchall()
        
        # --- MOTOR DE SAFETY ---
        bloqueo_seguridad, mensaje_alerta = evaluar_safety_ingreso(
            ubicacion_zona,
            trabajos_activos,
            usa_vehiculo,
            tipo_vehiculo,
            codigo_vehiculo,
            empresa
        )

        # --- SI HAY BLOQUEO, REDIRIGIMOS PARA MOSTRAR ALERTA ---
        if bloqueo_seguridad:
            cur.close()
            conn.close()
            flash(mensaje_alerta, "danger")
            return redirect(url_for('index'))

        # --- SI NO HAY PELIGRO, SE INSERTA NORMAL ---
        cur.execute("""
            INSERT INTO seguimiento_vias 
            (turno, operador_turno, spco_turno, empresa, orden_trabajo, responsable, ubicacion_zona, num_personas, tetra, estado, usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, 'En Vía', %s, %s, %s, %s);
        """, (turno, operador, spco, empresa, orden_trabajo, responsable, ubicacion_zona, num_personas, tetra, usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo))
        conn.commit()
        cur.close()
        conn.close()
        
        # Como fue exitoso, borramos la memoria para dejar el formulario limpio
        session.pop('form_previo', None)
        return redirect(url_for('index'))
    
@app.route('/liberar/<int:id>')
def liberar(id):
    conn = None
    cur = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("UPDATE seguimiento_vias SET hora_fin = CURRENT_TIME, estado = 'Liberado' WHERE id = %s;", (id,))
        conn.commit()
        flash('Registro liberado correctamente.', 'success')
    except (OperationalError, DatabaseError) as e:
        flash('Error de base de datos al liberar. Intenta de nuevo.', 'danger')
        app.logger.error(f"liberar DB error: {e}")
        if conn:
            conn.rollback()
    finally:
        if cur:
            cur.close()
        if conn:
            conn.close()
    return redirect(url_for('index'))

@app.route('/revertir/<int:id>')
def revertir(id):
    conn = None
    cur = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("UPDATE seguimiento_vias SET hora_fin = NULL, estado = 'En Vía' WHERE id = %s;", (id,))
        conn.commit()
        flash('Registro revertido correctamente.', 'success')
    except (OperationalError, DatabaseError) as e:
        flash('Error de base de datos al revertir. Intenta de nuevo.', 'danger')
        app.logger.error(f"revertir DB error: {e}")
        if conn:
            conn.rollback()
    finally:
        if cur:
            cur.close()
        if conn:
            conn.close()
    return redirect(url_for('index'))

@app.route('/eliminar/<int:id>')
def eliminar(id):
    conn = None
    cur = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM seguimiento_vias WHERE id = %s;", (id,))
        conn.commit()
        flash('Registro eliminado correctamente.', 'success')
    except (OperationalError, DatabaseError) as e:
        flash('Error de base de datos al eliminar. Intenta de nuevo.', 'danger')
        app.logger.error(f"eliminar DB error: {e}")
        if conn:
            conn.rollback()
    finally:
        if cur:
            cur.close()
        if conn:
            conn.close()
    return redirect(url_for('index'))

@app.route('/editar/<int:id>', methods=['POST'])
def editar(id):
    if request.method == 'POST':
        empresa = request.form['empresa']
        orden_trabajo = request.form.get('orden_trabajo', '-')
        responsable = request.form['responsable']
        ubicacion_zona = request.form['ubicacion_zona']
        num_personas = request.form['num_personas'] or 0
        tetra = request.form['tetra']
        
        usa_vehiculo = True if request.form.get('usa_vehiculo') == 'si' else False
        tipo_vehiculo = request.form.get('tipo_vehiculo', '') if usa_vehiculo else ''
        codigo_vehiculo = request.form.get('codigo_vehiculo', '') if usa_vehiculo else ''
        conductor_vehiculo = request.form.get('conductor_vehiculo', '') if usa_vehiculo else ''
        
        h_inicio_str = request.form['hora_inicio']
        h_fin_str = request.form['hora_fin']
        hora_inicio = datetime.strptime(h_inicio_str, "%H:%M:%S").time() if h_inicio_str else None
        hora_fin = datetime.strptime(h_fin_str, "%H:%M:%S").time() if h_fin_str else None
        estado = 'Liberado' if hora_fin else 'En Vía'

        conn = None
        cur = None
        try:
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute("""
                UPDATE seguimiento_vias 
                SET empresa = %s, orden_trabajo = %s, responsable = %s, ubicacion_zona = %s, num_personas = %s, 
                    tetra = %s, hora_inicio = %s, hora_fin = %s, estado = %s,
                    usa_vehiculo = %s, tipo_vehiculo = %s, codigo_vehiculo = %s, conductor_vehiculo = %s
                WHERE id = %s;
            """, (empresa, orden_trabajo, responsable, ubicacion_zona, num_personas, tetra, hora_inicio, hora_fin, estado, usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo, id))
            conn.commit()
            flash('Registro actualizado correctamente.', 'success')
        except (OperationalError, DatabaseError) as e:
            flash('Error de base de datos al actualizar. Intenta de nuevo.', 'danger')
            app.logger.error(f"editar DB error: {e}")
            if conn:
                conn.rollback()
        finally:
            if cur:
                cur.close()
            if conn:
                conn.close()
        return redirect(url_for('index'))

@app.route('/exportar')
def exportar():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT fecha, turno, operador_turno, spco_turno, empresa, responsable, ubicacion_zona, 
               num_personas, tetra, hora_inicio, hora_fin, estado,
               usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo 
        FROM seguimiento_vias 
        WHERE fecha = CURRENT_DATE 
        ORDER BY hora_inicio ASC;
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Seguimiento PCO"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells('A1:P1')
    ws['A1'] = "REPORTE DIARIO DE SEGUIMIENTO DE TRABAJOS EN VÍA"
    ws['A1'].font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    ws['A1'].fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    headers = ["Fecha", "Turno", "Operador ATS", "Supervisor SPCO", "Empresa", "Responsable", "Zona Ocupada", 
               "N° Pers.", "TETRA", "Hora Inicio", "Hora Fin", "Estado", "¿Usa Vehículo?", "Tipo Vehículo", "Código Vehículo", "Conductor"]
    ws.append([]) 
    ws.append(headers)

    header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    header_font = Font(name='Arial', size=10, bold=True, color='000000')
    thin_border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'), top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[3].height = 25

    for r_idx, row in enumerate(rows, 4):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if isinstance(value, (datetime, date)):
                cell.value = value.strftime('%Y-%m-%d')
            elif hasattr(value, 'strftime'): 
                cell.value = value.strftime('%H:%M:%S')
            elif isinstance(value, bool):
                cell.value = "Sí" if value else "No"
            else:
                cell.value = value
            cell.font = Font(name='Arial', size=10)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r_idx].height = 20

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    fecha_str = datetime.now().strftime('%Y%m%d')
    return send_file(excel_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'Seguimiento_Via_{fecha_str}.xlsx')

@app.route('/historial')
def historial():
    fecha_inicio = request.args.get('fecha_inicio', '')
    fecha_fin = request.args.get('fecha_fin', '')
    empresa = request.args.get('empresa', '')
    buscar_texto = request.args.get('buscar_texto', '')

    query = """
        SELECT id, fecha, turno, empresa, orden_trabajo, responsable, ubicacion_zona, 
               num_personas, tetra, hora_inicio, hora_fin, estado, 
               usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo,
               operador_turno, spco_turno
        FROM seguimiento_vias 
        WHERE 1=1
    """
    params = []

    if fecha_inicio:
        query += " AND fecha >= %s"
        params.append(fecha_inicio)
    if fecha_fin:
        query += " AND fecha <= %s"
        params.append(fecha_fin)
    if empresa:
        query += " AND empresa = %s"
        params.append(empresa)
    if buscar_texto:
        query += " AND (responsable ILIKE %s OR ubicacion_zona ILIKE %s OR codigo_vehiculo ILIKE %s OR conductor_vehiculo ILIKE %s)"
        term = f"%{buscar_texto}%"
        params.extend([term, term, term, term])

    query += " ORDER BY fecha DESC, hora_inicio DESC;"

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(query, tuple(params))
    registros = cur.fetchall()

    cur.execute("SELECT DISTINCT empresa FROM seguimiento_vias WHERE empresa IS NOT NULL ORDER BY empresa;")
    lista_empresas = [row[0] for row in cur.fetchall()]
    
    cur.close()
    conn.close()

    return render_template('historial.html', 
                           registros=registros, 
                           lista_empresas=lista_empresas,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           empresa_seleccionada=empresa,
                           buscar_texto=buscar_texto)

@app.route('/exportar_historial')
def exportar_historial():
    fecha_inicio = request.args.get('fecha_inicio', '')
    fecha_fin = request.args.get('fecha_fin', '')
    empresa = request.args.get('empresa', '')
    buscar_texto = request.args.get('buscar_texto', '')

    query = """
        SELECT fecha, turno, operador_turno, spco_turno, empresa, responsable, ubicacion_zona, 
               num_personas, tetra, hora_inicio, hora_fin, estado,
               usa_vehiculo, tipo_vehiculo, codigo_vehiculo, conductor_vehiculo 
        FROM seguimiento_vias 
        WHERE 1=1
    """
    params = []

    if fecha_inicio:
        query += " AND fecha >= %s"
        params.append(fecha_inicio)
    if fecha_fin:
        query += " AND fecha <= %s"
        params.append(fecha_fin)
    if empresa:
        query += " AND empresa = %s"
        params.append(empresa)
    if buscar_texto:
        query += " AND (responsable ILIKE %s OR ubicacion_zona ILIKE %s OR codigo_vehiculo ILIKE %s OR conductor_vehiculo ILIKE %s)"
        term = f"%{buscar_texto}%"
        params.extend([term, term, term, term])

    query += " ORDER BY fecha ASC, hora_inicio ASC;"

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(query, tuple(params))
    rows = cur.fetchall()
    cur.close()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historial Filtrado PCO"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells('A1:P1')
    ws['A1'] = "REPORTE HISTÓRICO DE SEGUIMIENTO DE TRABAJOS EN VÍA"
    ws['A1'].font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    ws['A1'].fill = PatternFill(start_color='333f48', end_color='333f48', fill_type='solid')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    headers = ["Fecha", "Turno", "Operador ATS", "Supervisor SPCO", "Empresa", "Responsable", "Zona Ocupada", 
               "N° Pers.", "TETRA", "Hora Inicio", "Hora Fin", "Estado", "¿Usa Vehículo?", "Tipo Vehículo", "Código Vehículo", "Conductor"]
    ws.append([]) 
    ws.append(headers)

    header_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
    header_font = Font(name='Arial', size=10, bold=True, color='000000')
    thin_border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'), top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[3].height = 25

    for r_idx, row in enumerate(rows, 4):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if isinstance(value, (datetime, date)):
                cell.value = value.strftime('%Y-%m-%d')
            elif hasattr(value, 'strftime'): 
                cell.value = value.strftime('%H:%M:%S')
            elif isinstance(value, bool):
                cell.value = "Sí" if value else "No"
            else:
                cell.value = value
            cell.font = Font(name='Arial', size=10)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r_idx].height = 20

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    return send_file(excel_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Historial_Reporte_Filtrado.xlsx')

@app.route('/archivar_turno', methods=['POST'])
def archivar_turno():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("UPDATE seguimiento_vias SET archivado = TRUE WHERE estado = 'Liberado' AND archivado = FALSE;")
    conn.commit()
    cur.close()
    conn.close()
    
    flash("✅ Relevo Preparado: Los trabajos liberados han sido archivados. Los trabajos 'En Vía' se mantienen en pantalla.", "success")
    return redirect(url_for('index'))

@app.route('/restaurar/<int:id>')
def restaurar(id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("UPDATE seguimiento_vias SET archivado = FALSE WHERE id = %s;", (id,))
    conn.commit()
    cur.close()
    conn.close()
    
    flash("✅ Registro restaurado exitosamente. El trabajo está de vuelta en tu panel de Monitoreo para que puedas corregirlo.", "success")
    return redirect(url_for('index'))

def extract_tetra(texto):
    if not texto:
        return ''
    texto = str(texto).replace(' ', '').replace('\n', ' ')
    matches = re.findall(r'(?<!\d)22\d{3}(?!\d)', texto)
    return ', '.join(matches) if matches else ''


@app.route('/importar_excel', methods=['POST'])
def importar_excel():
    if 'excel_file' not in request.files:
        return jsonify({'error': 'No se subió archivo'}), 400
    file = request.files['excel_file']
    if file.filename == '':
        return jsonify({'error': 'Nombre de archivo vacío'}), 400

    wb = openpyxl.load_workbook(file)
    ws = wb.active

    datos = []
    for row in range(6, ws.max_row + 1):
        estado = ws.cell(row=row, column=2).value or ''
        acceso = ws.cell(row=row, column=17).value or ''

        if str(estado).strip().upper() != 'CONFIRMADA':
            continue
        if str(acceso).strip().upper() != 'SI':
            continue

        spco = ws.cell(row=row, column=3).value or ''
        operador = ws.cell(row=row, column=5).value or ''
        desde = ws.cell(row=row, column=21).value or ''
        hasta = ws.cell(row=row, column=22).value or ''
        responsable = ws.cell(row=row, column=26).value or ''
        celular_tetra = ws.cell(row=row, column=27).value or ''
        orden_trabajo = ws.cell(row=row, column=10).value or ''
        empresa = ws.cell(row=row, column=15).value or ''

        tetra = extract_tetra(celular_tetra)

        datos.append({
            'spco': str(spco).strip(),
            'operador': str(operador).strip(),
            'desde': str(desde).strip(),
            'hasta': str(hasta).strip(),
            'responsable': str(responsable).strip(),
            'tetra': tetra,
            'orden_trabajo': str(orden_trabajo).strip(),
            'empresa': str(empresa).strip(),
        })

    wb.close()

    # Guardar en sesión para persistir entre recargas
    prev = session.get('importados_pendientes', [])
    prev.extend(datos)
    session['importados_pendientes'] = prev

    return jsonify({'datos': datos})


@app.route('/importar_texto', methods=['POST'])
def importar_texto():
    """Importa trabajos pegados como texto TSV (mismo formato que el Excel original).
    Detecta las columnas por su CABECERA (mas robusto al orden cambiante).
    Filtra igual que el Excel: solo filas CONFIRMADA + ACCESO A VÍA=SI."""
    data = request.get_json(silent=True) or {}
    texto = (data.get('texto') or '').strip()
    if not texto:
        return jsonify({'error': 'No se recibió texto'}), 400

    # Normalizar separadores: tabs -> tab único; detectar si viene del Excel (tab)
    # Si no hay tabs, asumimos que son múltiples espacios y los convertimos a tabs
    if '\t' not in texto and '  ' in texto:
        texto = re.sub(r' {2,}', '\t', texto)

    lineas = [ln for ln in texto.split('\n') if ln.strip()]
    if not lineas:
        return jsonify({'error': 'Texto vacío'}), 400

    # Leer con csv dialect excel-tab (maneja comillas correctamente)
    reader = csv.reader(_stdio_io.StringIO(texto), dialect='excel-tab', skipinitialspace=False)

    # Mapeo de cabeceras reconocidas -> campo interno
    # (claves normalizadas en MAYUS sin tildes ni espacios dobles)
    MAPA_CABECERAS = {
        'ESTADO': 'estado',
        'SPCO QUE VALIDA': 'spco',
        'SEGUIMIENTO (NOMBRE OPCO)': 'operador',
        'SEGUIMIENTO (NOMBRE OPCO)': 'operador',
        'ORDEN DE TRABAJO': 'orden_trabajo',
        'ORDEN DE TRABAJO': 'orden_trabajo',
        'EMPRESA CONTRATISTA': 'empresa',
        'EMPRESA CONTRATISTA': 'empresa',
        'ACCESO A VÍA': 'acceso',
        'ACCESO A VIA': 'acceso',
        'DESDE': 'desde',
        'HASTA': 'hasta',
        'RESPONSABLE DE TRABAJOS': 'responsable',
        'CELULAR Y TETRA EXT': 'celular_tetra',
    }

    def norm(s):
        # Normaliza: MAYUS, sin tildes, espacios colapsados
        s = (s or '').strip().upper()
        s = s.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')
        s = re.sub(r'\s+', ' ', s)
        return s

    # Primera fila NO vacia = cabecera. Buscarla.
    fila_iter = iter(reader)
    cabecera_idx = {}
    cab_encontrada = False
    for fila in fila_iter:
        if not fila or not any(c.strip() for c in fila):
            continue
        norm_fila = [norm(c) for c in fila]
        # Comprobar si contiene al menos 'ESTADO' y 'ACCESO A VIA' -> es cabecera
        if 'ESTADO' in norm_fila or 'ESTADO' in [n.split('(')[0].strip() for n in norm_fila]:
            for i, c in enumerate(norm_fila):
                if c in MAPA_CABECERAS:
                    cabecera_idx[MAPA_CABECERAS[c]] = i
            cab_encontrada = True
            break
        # Puede que la primera fila ya sea datos (sin cabecera) -> caer al fallback fijo
        else:
            # Reinsertar esta fila como primera fila de datos usando mapeo fijo
            datos_fijos = [_parsear_fila_fija(fila)]
            cab_encontrada = False
            break

    if not cab_encontrada and not cabecera_idx:
        # Fallback: usar mapeo posicional fijo (texto sin cabecera, igual que antes)
        # Releer todo desde el principio
        reader = csv.reader(_stdio_io.StringIO(texto), dialect='excel-tab', skipinitialspace=False)
        datos = []
        for fila in reader:
            parsed = _parsear_fila_fija(fila)
            if parsed:
                datos.append(parsed)
    else:
        # Mapeo por cabecera: leer resto de filas
        campos_requeridos = ('estado', 'acceso')
        if not all(c in cabecera_idx for c in campos_requeridos):
            return jsonify({'error': 'Faltan cabeceras obligatorias (ESTADO, ACCESO A VÍA). Cabeceras detectadas: ' + ', '.join(cabecera_idx.keys())}), 400

        datos = []
        for fila in fila_iter:
            if not fila or not any(c.strip() for c in fila):
                continue
            estado_val = fila[cabecera_idx['estado']].strip().upper() if 'estado' in cabecera_idx and len(fila) > cabecera_idx['estado'] else ''
            acceso_val = fila[cabecera_idx['acceso']].strip().upper() if 'acceso' in cabecera_idx and len(fila) > cabecera_idx['acceso'] else ''

            if estado_val != 'CONFIRMADA':
                continue
            if acceso_val != 'SI':
                continue

            def getcampo(campo):
                idx = cabecera_idx.get(campo)
                if idx is None or idx >= len(fila):
                    return ''
                return str(fila[idx]).strip()

            celular = getcampo('celular_tetra')
            tetra = extract_tetra(celular)

            datos.append({
                'spco': getcampo('spco'),
                'operador': getcampo('operador'),
                'desde': getcampo('desde'),
                'hasta': getcampo('hasta'),
                'responsable': getcampo('responsable'),
                'tetra': tetra,
                'orden_trabajo': getcampo('orden_trabajo'),
                'empresa': getcampo('empresa'),
            })

    if not datos:
        return jsonify({'error': 'No se encontraron filas CONFIRMADAS con ACCESO=SI en el texto pegado'}), 400

    # Guardar en sesión (igual que importar_excel)
    prev = session.get('importados_pendientes', [])
    prev.extend(datos)
    session['importados_pendientes'] = prev

    return jsonify({'datos': datos, 'total': len(datos)})


def _parsear_fila_fija(fila):
    """Fallback: mapeo posicional fijo cuando NO hay cabecera reconocida.
    Mantiene compatibilidad con el mapeo anterior."""
    if not fila or len(fila) < 5:
        return None
    col0 = (fila[0] or '').strip().upper()
    if col0 in ('N°', 'NRO', 'N.', 'N', 'ITEM', '#'):
        return None
    estado = (fila[1] if len(fila) > 1 else '') or ''
    acceso = (fila[13] if len(fila) > 13 else '') or ''
    if str(estado).strip().upper() != 'CONFIRMADA':
        return None
    if str(acceso).strip().upper() != 'SI':
        return None
    spco = (fila[2] if len(fila) > 2 else '') or ''
    operador = (fila[4] if len(fila) > 4 else '') or ''
    orden_trabajo = (fila[7] if len(fila) > 7 else '') or ''
    empresa = (fila[10] if len(fila) > 10 else '') or ''
    desde = (fila[16] if len(fila) > 16 else '') or ''
    hasta = (fila[17] if len(fila) > 17 else '') or ''
    responsable = (fila[19] if len(fila) > 19 else '') or ''
    celular_tetra = (fila[20] if len(fila) > 20 else '') or ''
    tetra = extract_tetra(celular_tetra)
    return {
        'spco': str(spco).strip(),
        'operador': str(operador).strip(),
        'desde': str(desde).strip(),
        'hasta': str(hasta).strip(),
        'responsable': str(responsable).strip(),
        'tetra': tetra,
        'orden_trabajo': str(orden_trabajo).strip(),
        'empresa': str(empresa).strip(),
    }


@app.route('/confirmar_importado', methods=['POST'])
def confirmar_importado():
    data = request.get_json() or {}
    idx = data.get('_idx', -1)

    # Tomar datos completos desde la sesión
    pendientes = session.get('importados_pendientes', [])
    item = pendientes[idx] if 0 <= idx < len(pendientes) else data

    turno = session.get('turno', 'Turno 1')
    operador = item.get('operador', session.get('operador_turno', 'Importado'))
    spco = item.get('spco', session.get('spco_turno', 'Importado'))
    empresa = item.get('empresa', '')
    orden_trabajo = item.get('orden_trabajo', '')
    responsable = item.get('responsable', '')
    desde = item.get('desde', '')
    hasta = item.get('hasta', '')
    zona = f"{desde} -> {hasta}" if desde and hasta else (desde or hasta)
    tetra = item.get('tetra', '')

    hora_inicio_str = data.get('hora_inicio', '')
    hora_inicio = datetime.strptime(hora_inicio_str, "%H:%M:%S").time() if hora_inicio_str else None

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO seguimiento_vias 
        (turno, operador_turno, spco_turno, empresa, orden_trabajo, responsable, ubicacion_zona, num_personas, tetra, hora_inicio, estado)
        VALUES (%s, %s, %s, %s, %s, %s, %s, 0, %s, %s, 'En Vía')
        RETURNING id;
    """, (turno, operador, spco, empresa, orden_trabajo, responsable, zona, tetra, hora_inicio))
    new_id = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()

    # Remover de la sesión
    pendientes = session.get('importados_pendientes', [])
    if 0 <= idx < len(pendientes):
        pendientes.pop(idx)
        session['importados_pendientes'] = pendientes

    return jsonify({'success': True, 'id': new_id, 'empresa': empresa})


@app.route('/revertir_importado/<int:id>')
def revertir_importado(id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT operador_turno, spco_turno, empresa, orden_trabajo, responsable, ubicacion_zona, tetra
        FROM seguimiento_vias WHERE id = %s;
    """, (id,))
    row = cur.fetchone()
    if not row:
        cur.close()
        conn.close()
        return jsonify({'error': 'No encontrado'}), 404

    operador, spco, empresa, orden_trabajo, responsable, zona, tetra = row

    cur.execute("DELETE FROM seguimiento_vias WHERE id = %s;", (id,))
    conn.commit()
    cur.close()
    conn.close()

    # Separar zona en desde -> hasta
    desde, hasta = (zona.split(' -> ') + ['', ''])[:2] if ' -> ' in zona else (zona, '')

    item = {
        'spco': spco or '',
        'operador': operador or '',
        'desde': desde,
        'hasta': hasta,
        'responsable': responsable or '',
        'tetra': tetra or '',
        'orden_trabajo': orden_trabajo or '',
        'empresa': empresa or '',
    }

    pendientes = session.get('importados_pendientes', [])
    pendientes.append(item)
    session['importados_pendientes'] = pendientes

    return jsonify({'success': True, 'item': item})


if __name__ == '__main__':
    app.run(debug=True, port=5000)