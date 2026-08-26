import os
import re
import io
import tempfile
import shutil
import subprocess
import math
from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.PyQt.QtGui import QIcon
import os
from qgis.core import (
    QgsFeature, QgsGeometry, QgsPointXY, QgsWkbTypes, QgsVectorLayer, QgsRectangle,
    QgsMessageLog
)
from qgis.utils import iface

ODA_CONVERTER = r"C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe"

DWG_VERSIONS = {
    b'AC1015': 'ACAD2000', b'AC1018': 'ACAD2004', b'AC1021': 'ACAD2007',
    b'AC1024': 'ACAD2010', b'AC1027': 'ACAD2013', b'AC1032': 'ACAD2018',
}


def _read_clipboard():
    import win32clipboard
    try:
        win32clipboard.OpenClipboard()
    except Exception:
        return None, None
    try:
        embed_id = win32clipboard.RegisterClipboardFormat("Embed Source")
        fmts = []
        f = 0
        while True:
            f = win32clipboard.EnumClipboardFormats(f)
            if f == 0:
                break
            fmts.append(f)
        if embed_id in fmts:
            try:
                data = win32clipboard.GetClipboardData(embed_id)
                if data and len(data) > 100 and data[:4] == b'\xd0\xcf\x11\xe0':
                    return 'ole', data
            except Exception as e:
                QgsMessageLog.logMessage("PasteFromCAD: OLE clipboard read: {}".format(e), "PasteFromCAD", 1)
        for fmt_id in fmts:
            try:
                data = win32clipboard.GetClipboardData(fmt_id)
                if not data or len(data) < 100:
                    continue
                if data[:4] == b'\xd0\xcf\x11\xe0':
                    return 'ole', data
                if data[:6] in DWG_VERSIONS:
                    return 'dwg', data
            except Exception as e:
                QgsMessageLog.logMessage("PasteFromCAD: format {} skip: {}".format(fmt_id, e), "PasteFromCAD", 0)
                continue
        try:
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            if text and len(text) > 10:
                return 'text', text
        except Exception as e:
            QgsMessageLog.logMessage("PasteFromCAD: text clipboard read: {}".format(e), "PasteFromCAD", 1)
        return None, None
    finally:
        win32clipboard.CloseClipboard()


def _parse_dxf_ascii(dxf_text):
    """Parser DXF ASCII simple - extrae entidades geometricas."""
    entities = []
    lines = dxf_text.split('\n')
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if line == '0':
            i += 1
            if i < n:
                etype = lines[i].strip()
                if etype in ('LWPOLYLINE', 'POLYLINE', 'LINE', 'POINT',
                             'CIRCLE', 'ARC', 'ELLIPSE', 'TEXT', 'MTEXT',
                             '3DFACE', 'SOLID', 'INSERT'):
                    tags = {}
                    while i < n - 1:
                        i += 1
                        gc = lines[i].strip()
                        if gc == '0':
                            break
                        i += 1
                        if i < n:
                            tags[gc] = lines[i].strip()
                    entities.append((etype, tags))
                    continue
        i += 1
    return entities


def _dxf_tags_to_points(tags):
    """Extrae puntos de una entidad DXF usando group codes."""
    pts = []
    # LWPOLYLINE: 10=X, 20=Y (repeat)
    gc_x = [k for k in tags if k == '10']
    gc_y = [k for k in tags if k == '20']
    if gc_x:
        for kx in sorted([k for k in tags if re.match(r'^10$', k)]):
            idx = kx
            x = float(tags[idx])
            y_key = '20'
            y = float(tags.get(y_key, '0'))
            pts.append((x, y))
    # Better approach: scan raw DXF groups for 10/20 pairs
    return pts


def _parse_dxf_file(dxf_path):
    """Parsea un archivo DXF y devuelve lista de (tipo, puntos)."""
    # Try reading as text first, fallback to bytes
    content = None
    for enc in ('utf-8', 'latin-1', 'cp1252'):
        try:
            with open(dxf_path, 'r', encoding=enc) as f:
                content = f.read()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    if content is None:
        with open(dxf_path, 'rb') as f:
            raw = f.read()
        content = raw.decode('latin-1', errors='replace')

    results = []
    lines = content.split('\n')
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i].strip()
        if line == '0' and i + 1 < n:
            etype = lines[i + 1].strip()
            if etype in ('LWPOLYLINE', 'LINE', 'POINT', 'CIRCLE', 'ARC',
                         'ELLIPSE', 'TEXT', 'MTEXT', '3DFACE', 'SOLID',
                         'POLYLINE', 'INSERT'):
                tags = []
                i += 2
                while i < n:
                    gc_str = lines[i].strip()
                    if gc_str == '0':
                        break
                    try:
                        gc = int(gc_str)
                    except ValueError:
                        i += 1
                        continue
                    i += 1
                    if i < n:
                        val = lines[i].strip()
                        tags.append((gc, val))
                    i += 1

                pts = _extract_points_from_tags(etype, tags)
                if pts:
                    results.append((etype, pts, tags))
                continue
        i += 1
    return results


def _extract_points_from_tags(etype, tags):
    """Extrae puntos de los group codes DXF."""
    pts = []

    if etype == 'LWPOLYLINE':
        xs = [v for gc, v in tags if gc == 10]
        ys = [v for gc, v in tags if gc == 20]
        for j in range(min(len(xs), len(ys))):
            try:
                pts.append((float(xs[j]), float(ys[j])))
            except ValueError as e:
                QgsMessageLog.logMessage("PasteFromCAD: LWPOLYLINE point parse: {}".format(e), "PasteFromCAD", 0)

    elif etype == 'POLYLINE':
        xs = [v for gc, v in tags if gc == 10]
        ys = [v for gc, v in tags if gc == 20]
        for j in range(min(len(xs), len(ys))):
            try:
                pts.append((float(xs[j]), float(ys[j])))
            except ValueError as e:
                QgsMessageLog.logMessage("PasteFromCAD: POLYLINE point parse: {}".format(e), "PasteFromCAD", 0)

    elif etype == 'LINE':
        x1 = next((v for gc, v in tags if gc == 10), None)
        y1 = next((v for gc, v in tags if gc == 20), None)
        x2 = next((v for gc, v in tags if gc == 11), None)
        y2 = next((v for gc, v in tags if gc == 21), None)
        if all([x1, y1, x2, y2]):
            pts = [(float(x1), float(y1)), (float(x2), float(y2))]

    elif etype == 'POINT':
        x = next((v for gc, v in tags if gc == 10), None)
        y = next((v for gc, v in tags if gc == 20), None)
        if x and y:
            pts = [(float(x), float(y))]

    elif etype == 'CIRCLE':
        cx = next((v for gc, v in tags if gc == 10), None)
        cy = next((v for gc, v in tags if gc == 20), None)
        r = next((v for gc, v in tags if gc == 40), None)
        if cx and cy and r:
            cx, cy, r = float(cx), float(cy), float(r)
            for a in range(0, 361, 5):
                pts.append((cx + r * math.cos(math.radians(a)),
                            cy + r * math.sin(math.radians(a))))

    elif etype == 'ARC':
        cx = next((v for gc, v in tags if gc == 10), None)
        cy = next((v for gc, v in tags if gc == 20), None)
        r = next((v for gc, v in tags if gc == 40), None)
        sa = next((v for gc, v in tags if gc == 50), None)
        ea = next((v for gc, v in tags if gc == 51), None)
        if cx and cy and r and sa and ea:
            cx, cy, r = float(cx), float(cy), float(r)
            start, end = float(sa), float(ea)
            if end < start:
                end += 360
            step = max(1, int((end - start) / 60))
            for a in range(int(start), int(end) + 1, step):
                pts.append((cx + r * math.cos(math.radians(a)),
                            cy + r * math.sin(math.radians(a))))

    elif etype == 'ELLIPSE':
        cx = next((v for gc, v in tags if gc == 10), None)
        cy = next((v for gc, v in tags if gc == 20), None)
        mx = next((v for gc, v in tags if gc == 11), None)
        my = next((v for gc, v in tags if gc == 21), None)
        ratio = next((v for gc, v in tags if gc == 40), None)
        if cx and cy and mx and my and ratio:
            cx, cy = float(cx), float(cy)
            mx, my = float(mx), float(my)
            ratio = float(ratio)
            major_len = math.sqrt(mx**2 + my**2)
            minor_len = major_len * ratio
            angle = math.atan2(my, mx)
            for a in range(0, 361, 5):
                rad = math.radians(a)
                x = major_len * math.cos(rad)
                y = minor_len * math.sin(rad)
                rx = x * math.cos(angle) - y * math.sin(angle) + cx
                ry = x * math.sin(angle) + y * math.cos(angle) + cy
                pts.append((rx, ry))

    elif etype in ('TEXT', 'MTEXT'):
        x = next((v for gc, v in tags if gc == 10), None)
        y = next((v for gc, v in tags if gc == 20), None)
        if x and y:
            pts = [(float(x), float(y))]

    return pts


class PasteFromCAD:
    def __init__(self, iface):
        self.iface = iface
        self.action = None

    def initGui(self):
        icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        self.action = QAction(QIcon(icon_path), 'PasteFromCAD', self.iface.mainWindow())
        self.action.setToolTip('PasteFromCAD')
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu('PasteFromCad', self.action)

    def unload(self):
        self.iface.removePluginMenu('&PasteFromCad', self.action)
        self.iface.removeToolBarIcon(self.action)

    def run(self):
        layer = self.iface.activeLayer()
        if not layer or not isinstance(layer, QgsVectorLayer):
            QMessageBox.warning(self.iface.mainWindow(), "Error", "Selecciona una capa vectorial.")
            return
        if not layer.isEditable():
            QMessageBox.warning(self.iface.mainWindow(), "Error", "La capa debe estar en modo edicion.")
            return

        kind, raw = _read_clipboard()
        if not raw:
            QMessageBox.warning(
                self.iface.mainWindow(), "Sin datos",
                "No se encontro datos de AutoCAD en el portapapeles.\n\n"
                "1. Selecciona objetos en AutoCAD\n"
                "2. Ctrl+C en AutoCAD\n"
                "3. En QGIS: capa en edicion -> Paste from CAD"
            )
            return

        layer_type = layer.geometryType()
        features = []

        if kind == 'ole':
            entities = self._ole_to_entities(raw)
            features = self._entities_to_features(entities, layer_type, layer)
        elif kind == 'dwg':
            entities = self._dwg_to_entities(raw)
            features = self._entities_to_features(entities, layer_type, layer)
        elif kind == 'text':
            features = self._text_to_features(raw, layer_type, layer)

        if features:
            layer.startEditing()
            layer.addFeatures(features)
            ok = layer.commitChanges()
            if not ok:
                errors = layer.commitErrors()
                QMessageBox.warning(
                    self.iface.mainWindow(), "Error al guardar",
                    "Error guardando en la capa:\n" + "\n".join(errors[:5])
                )
                return
            QMessageBox.information(
                self.iface.mainWindow(), "Exito",
                "Se pegaron {} entidades desde AutoCAD.".format(len(features))
            )
            canvas = iface.mapCanvas()
            geom = QgsGeometry()
            for f in features:
                geom = geom.combine(f.geometry())
            canvas.setExtent(geom.boundingBox())
            canvas.refresh()
        else:
            QMessageBox.warning(
                self.iface.mainWindow(), "Sin entidades",
                "No se pudieron convertir las entidades del portapapeles."
            )

    def _dwg_to_dxf_file(self, dwg_data, tmp_dir):
        """Convierte DWG bytes a archivo DXF, retorna ruta del DXF o None."""
        if not os.path.exists(ODA_CONVERTER):
            return None
        version = DWG_VERSIONS.get(dwg_data[:6], 'ACAD2013')
        if version not in DWG_VERSIONS.values():
            version = 'ACAD2013'
        dwg_path = os.path.join(tmp_dir, "input.dwg")
        dxf_out = os.path.join(tmp_dir, "dxf_out")
        os.makedirs(dxf_out, exist_ok=True)
        with open(dwg_path, 'wb') as f:
            f.write(dwg_data)
        abs_dwg_dir = os.path.abspath(tmp_dir)
        abs_dxf = os.path.abspath(dxf_out)
        abs_converter = os.path.abspath(ODA_CONVERTER)
        try:
            subprocess.run(  # nosec B603
                [abs_converter, abs_dwg_dir, abs_dxf, version, "DXF", "0", "0"],
                capture_output=True, timeout=30,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            )
        except Exception as e:
            QgsMessageLog.logMessage("PasteFromCAD: ODA converter: {}".format(e), "PasteFromCAD", 1)
            return None
        dxf_files = [f for f in os.listdir(dxf_out) if f.lower().endswith('.dxf')]
        if not dxf_files:
            return None
        return os.path.join(dxf_out, dxf_files[0])

    def _ole_to_entities(self, ole_data):
        import olefile
        try:
            ole = olefile.OleFileIO(io.BytesIO(ole_data))
            content = ole.openstream('Contents').read()
            ole.close()
        except Exception as e:
            QgsMessageLog.logMessage("PasteFromCAD: OLE parse: {}".format(e), "PasteFromCAD", 1)
            return []
        if content[:6] not in DWG_VERSIONS:
            return []
        tmp_dir = tempfile.mkdtemp(prefix="pastecad_")
        try:
            dxf_path = self._dwg_to_dxf_file(content, tmp_dir)
            if not dxf_path:
                return []
            return _parse_dxf_file(dxf_path)
        finally:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception as e:
                QgsMessageLog.logMessage("PasteFromCAD: temp cleanup: {}".format(e), "PasteFromCAD", 0)

    def _dwg_to_entities(self, dwg_data):
        if dwg_data[:6] not in DWG_VERSIONS:
            return []
        tmp_dir = tempfile.mkdtemp(prefix="pastecad_")
        try:
            dxf_path = self._dwg_to_dxf_file(dwg_data, tmp_dir)
            if not dxf_path:
                return []
            return _parse_dxf_file(dxf_path)
        finally:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception as e:
                QgsMessageLog.logMessage("PasteFromCAD: temp cleanup: {}".format(e), "PasteFromCAD", 0)

    def _entities_to_features(self, entities, layer_type, layer):
        features = []
        for etype, pts_2d, tags in entities:
            geom = self._make_geometry(etype, pts_2d, tags, layer_type)
            if geom:
                feat = QgsFeature(layer.fields())
                feat.setGeometry(geom)
                features.append(feat)
        return features

    def _make_geometry(self, etype, pts_2d, tags, layer_type):
        if not pts_2d:
            return None
        qpts = [QgsPointXY(p[0], p[1]) for p in pts_2d]

        if len(qpts) < 2:
            if layer_type == QgsWkbTypes.GeometryType.PointGeometry and len(qpts) == 1:
                return QgsGeometry.fromPointXY(qpts[0])
            return None

        if layer_type == QgsWkbTypes.GeometryType.PolygonGeometry:
            if qpts[0] != qpts[-1]:
                qpts.append(qpts[0])
            if len(qpts) < 4:
                return None
            return QgsGeometry.fromPolygonXY([qpts])

        elif layer_type == QgsWkbTypes.GeometryType.LineGeometry:
            return QgsGeometry.fromPolylineXY(qpts)

        elif layer_type == QgsWkbTypes.GeometryType.PointGeometry:
            return QgsGeometry.fromPointXY(qpts[0])

        return None

    def _text_to_features(self, text, layer_type, layer):
        features = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            upper = line.upper()
            if upper.startswith('_PLINE') or upper.startswith('PLINE'):
                geom = self._parse_pline(line, layer_type)
                if geom:
                    feat = QgsFeature(layer.fields())
                    feat.setGeometry(geom)
                    features.append(feat)
            elif upper.startswith('_POINT') or upper.startswith('POINT'):
                geom = self._parse_point(line)
                if geom:
                    feat = QgsFeature(layer.fields())
                    feat.setGeometry(geom)
                    features.append(feat)
        return features

    def _parse_pline(self, line, layer_type):
        coords_str = re.split(r'\s+', line, maxsplit=1)[-1] if ' ' in line else line[6:]
        points = []
        for m in re.finditer(r'([\d\.\-]+)\s*[,;]\s*([\d\.\-]+)', coords_str):
            points.append(QgsPointXY(float(m.group(1)), float(m.group(2))))
        if len(points) < 2:
            return None
        if layer_type == QgsWkbTypes.GeometryType.PolygonGeometry:
            if points[0] != points[-1]:
                points.append(points[0])
            if len(points) < 4:
                return None
            return QgsGeometry.fromPolygonXY([points])
        elif layer_type == QgsWkbTypes.GeometryType.PointGeometry:
            return QgsGeometry.fromPointXY(points[0])
        return QgsGeometry.fromPolylineXY(points)

    def _parse_point(self, line):
        m = re.search(r'([\d\.\-]+)\s*[,;]\s*([\d\.\-]+)', line[6:])
        if m:
            return QgsGeometry.fromPointXY(QgsPointXY(float(m.group(1)), float(m.group(2))))
        return None
