import os

from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.PyQt.QtGui import QIcon
from qgis.core import (
    QgsFeature, QgsGeometry, QgsPointXY, QgsWkbTypes, QgsVectorLayer
)


class CopyFromQgis:
    def __init__(self, iface):
        self.iface = iface
        self.action = None

    def initGui(self):
        icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        self.action = QAction(QIcon(icon_path), 'CopyFromQgis', self.iface.mainWindow())
        self.action.setToolTip('CopyFromQgis')
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu('CopyFromQgis', self.action)

    def unload(self):
        if self.action:
            self.iface.removeToolBarIcon(self.action)
            self.iface.removePluginMenu("Copy from QGIS", self.action)

    def run(self):
        layer = self.iface.activeLayer()
        if not layer or not isinstance(layer, QgsVectorLayer):
            QMessageBox.warning(
                self.iface.mainWindow(), "Error",
                "Selecciona una capa vectorial."
            )
            return

        selected = list(layer.selectedFeatures())
        if not selected:
            QMessageBox.warning(
                self.iface.mainWindow(), "Sin seleccion",
                "Selecciona al menos una entidad."
            )
            return

        try:
            ok, msg = self._copy_via_autocad(selected)
            if ok:
                QMessageBox.information(
                    self.iface.mainWindow(), "Exito",
                    "{} entidades en el portapapeles.\n\n"
                    "En tu dibujo de AutoCAD:\n"
                    "Ctrl+Shift+V para pegar.".format(
                        len(selected)
                    )
                )
            else:
                QMessageBox.warning(
                    self.iface.mainWindow(), "Error", msg
                )
        except Exception as e:
            QMessageBox.warning(
                self.iface.mainWindow(), "Error", str(e)
            )

    def _copy_via_autocad(self, features):
        import win32com.client
        import pythoncom
        import time

        acad = None
        for attempt in range(5):
            try:
                acad = win32com.client.GetActiveObject(
                    "AutoCAD.Application"
                )
                break
            except Exception:
                pass
            try:
                acad = win32com.client.Dispatch(
                    "AutoCAD.Application"
                )
                break
            except Exception:
                pass
            time.sleep(1)

        if acad is None:
            return (
                False,
                "Abre AutoCAD primero.",
            )

        acad.Visible = True
        doc = acad.ActiveDocument
        ms = doc.ModelSpace

        tmp = "_QGIS_CLIP"
        try:
            try:
                doc.Layers.Item(tmp)
            except Exception:
                doc.Layers.Add(tmp)
        except Exception:
            return False, "No se pudo crear capa temporal."

        created = []
        for feat in features:
            geom = feat.geometry()
            if geom.isNull() or geom.isEmpty():
                continue

            if geom.type() == QgsWkbTypes.PolygonGeometry:
                polys = geom.asMultiPolygon() if geom.isMultipart() else [geom.asPolygon()]
                for poly in polys:
                    for ring in poly:
                        if len(ring) < 3:
                            continue
                        flat = []
                        for pt in ring:
                            flat.extend([pt.x(), pt.y()])
                        v = win32com.client.VARIANT(
                            pythoncom.VT_ARRAY | pythoncom.VT_R8, flat
                        )
                        p = ms.AddLightWeightPolyline(v)
                        p.Closed = True
                        p.Layer = tmp
                        created.append(p)

            elif geom.type() == QgsWkbTypes.LineGeometry:
                lines_list = geom.asMultiPolyline() if geom.isMultipart() else [geom.asPolyline()]
                for ln in lines_list:
                    if len(ln) < 2:
                        continue
                    flat = []
                    for pt in ln:
                        flat.extend([pt.x(), pt.y()])
                    v = win32com.client.VARIANT(
                        pythoncom.VT_ARRAY | pythoncom.VT_R8, flat
                    )
                    p = ms.AddLightWeightPolyline(v)
                    p.Closed = False
                    p.Layer = tmp
                    created.append(p)

            elif geom.type() == QgsWkbTypes.PointGeometry:
                pts = geom.asMultiPoint() if geom.isMultipart() else [geom.asPoint()]
                for pt in pts:
                    v = win32com.client.VARIANT(
                        pythoncom.VT_ARRAY | pythoncom.VT_R8,
                        [pt.x(), pt.y(), 0.0]
                    )
                    p = ms.AddPoint(v)
                    p.Layer = tmp
                    created.append(p)

        if not created:
            return False, "No se crearon entidades."

        doc.Regen(1)
        time.sleep(1)

        doc.SendCommand(
            "(progn "
            "(command \"_.COPYBASE\" \"0,0,0\" "
            "(ssget \"X\" '((8 . \"{}\")))) "
            "(command \"_.ERASE\" "
            "(ssget \"X\" '((8 . \"{}\")))) "
            ")\n".format(tmp, tmp)
        )
        time.sleep(4)

        try:
            doc.Layers.Item(tmp).Delete()
        except Exception:
            pass

        try:
            doc.Layers.Item(tmp).Delete()
        except Exception:
            pass

        return True, ""
