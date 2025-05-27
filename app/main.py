import sys
import os
from pathlib import Path
# Import QHBoxLayout et Qt pour les états de la checkbox
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, \
    QCheckBox
from PyQt6.QtCore import Qt  # Pour Qt.CheckState
import xlwings as xw
import xml.etree.ElementTree as ET
import traceback  # Pour un meilleur débogage

# Déterminer le chemin racine du projet (Trading2)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

EXCEL_FILE_NAME = "TradingData.xlsm"
EXCEL_WINDOW_IDENTIFIER = "TradingData"

EXCEL_FILE_PATH = PROJECT_ROOT / "Data" / "Excel" / EXCEL_FILE_NAME

CONFIG_DIR = PROJECT_ROOT / "Data" / "config"
CONFIG_FILE_NAME = "window_settings.xml"
CONFIG_FILE_PATH = CONFIG_DIR / CONFIG_FILE_NAME

DEFAULT_EXCEL_POS = {
    "left": 3355,
    "top": 0,
    "height": 1049,
    "width": 488
}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Trading System Controller")
        self.setGeometry(100, 100, 320, 100)

        self.excel_wb = None
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)

        main_layout = QVBoxLayout()
        controls_container = QWidget()
        h_layout = QHBoxLayout(controls_container)
        h_layout.setContentsMargins(0, 0, 0, 0)
        h_layout.setSpacing(5)

        self.btn_open_excel = QPushButton(f"Ouvrir {EXCEL_FILE_NAME}")
        self.btn_open_excel.clicked.connect(self.toggle_excel_visibility)
        h_layout.addWidget(self.btn_open_excel)

        self.chk_mode_test = QCheckBox("Mode TEST")  # Renommée
        self.chk_mode_test.setChecked(False)  # Par défaut, Mode TEST est OFF
        # Connecter le changement d'état à une fonction
        self.chk_mode_test.stateChanged.connect(self.on_mode_test_changed)
        h_layout.addWidget(self.chk_mode_test)

        main_layout.addWidget(controls_container)
        container_widget = QWidget()
        container_widget.setLayout(main_layout)
        self.setCentralWidget(container_widget)

    def on_mode_test_changed(self, state):
        """Appelé lorsque l'état de la case 'Mode TEST' change."""
        if state == Qt.CheckState.Unchecked.value:  # Si on décoche "Mode TEST"
            print("Mode TEST désactivé.")
            if self.excel_wb and self.excel_wb.app.visible:
                print("Excel est ouvert. Application de la configuration sauvegardée.")
                try:
                    pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER)
                    self.set_excel_position(self.excel_wb.app.api, pos_config)
                    QMessageBox.information(self, "Mode Normal",
                                            "Mode TEST désactivé. La position/dimension sauvegardée a été appliquée à Excel.")
                except Exception as e:
                    print(f"Erreur en appliquant la config après désactivation Mode TEST: {e}")
                    QMessageBox.warning(self, "Erreur", f"Impossible d'appliquer la configuration à Excel: {e}")
        else:  # Si on coche "Mode TEST"
            print("Mode TEST activé. Les positions/dimensions Excel ne seront ni appliquées ni sauvegardées.")
            QMessageBox.information(self, "Mode TEST",
                                    "Mode TEST activé. Les modifications de position/dimension d'Excel ne seront pas sauvegardées, "
                                    "et la configuration sauvegardée ne sera pas appliquée à l'ouverture.")

    def load_excel_config(self, identifier):
        # Cette fonction charge toujours, que le Mode TEST soit actif ou non.
        # La décision d'utiliser ou non la config chargée se fait ailleurs.
        try:
            if CONFIG_FILE_PATH.exists():
                tree = ET.parse(CONFIG_FILE_PATH)
                root_config = tree.getroot()
                excel_windows_node = root_config.find("excel_window")
                if excel_windows_node is not None:
                    specific_window_node = excel_windows_node.find(identifier)
                    if specific_window_node is not None:
                        pos = {
                            "left": int(float(specific_window_node.find("left").text)),
                            "top": int(float(specific_window_node.find("top").text)),
                            "width": int(float(specific_window_node.find("width").text)),
                            "height": int(float(specific_window_node.find("height").text))
                        }
                        print(f"Config chargée pour '{identifier}' depuis {CONFIG_FILE_PATH}: {pos}")
                        return pos
        except Exception as e:
            print(f"Erreur chargement config XML pour '{identifier}' ({CONFIG_FILE_PATH}): {e}")
            QMessageBox.warning(self, "Erreur Config",
                                f"Impossible de charger config pour '{identifier}': {e}\nUtilisation des défauts.")
        print(f"Config non trouvée pour '{identifier}'. Utilisation des défauts.")
        return DEFAULT_EXCEL_POS.copy()

    def save_excel_config(self, identifier, pos):
        if self.chk_mode_test.isChecked():  # Vérifie si le Mode TEST est actif
            print(f"Mode TEST actif. Sauvegarde désactivée pour '{identifier}'.")
            return  # Ne rien faire si Mode TEST est coché

        # ... (le reste du code de sauvegarde est identique)
        try:
            if CONFIG_FILE_PATH.exists():
                try:
                    tree = ET.parse(CONFIG_FILE_PATH)
                    root_config = tree.getroot()
                except ET.ParseError:
                    print(f"Avertissement: {CONFIG_FILE_PATH} corrompu. Recréation.")
                    root_config = ET.Element("config")
            else:
                root_config = ET.Element("config")

            excel_windows_node = root_config.find("excel_window")
            if excel_windows_node is None:
                excel_windows_node = ET.SubElement(root_config, "excel_window")

            specific_window_node = excel_windows_node.find(identifier)
            if specific_window_node is None:
                specific_window_node = ET.SubElement(excel_windows_node, identifier)
            else:
                for child in list(specific_window_node):
                    specific_window_node.remove(child)

            ET.SubElement(specific_window_node, "left").text = str(int(pos["left"]))
            ET.SubElement(specific_window_node, "top").text = str(int(pos["top"]))
            ET.SubElement(specific_window_node, "width").text = str(int(pos["width"]))
            ET.SubElement(specific_window_node, "height").text = str(int(pos["height"]))

            new_tree = ET.ElementTree(root_config)
            ET.indent(new_tree, space="  ", level=0)
            new_tree.write(CONFIG_FILE_PATH, encoding="utf-8", xml_declaration=True)
            print(f"Config sauvegardée pour '{identifier}' dans {CONFIG_FILE_PATH}: {pos}")

        except Exception as e:
            print(f"Erreur sauvegarde config XML pour '{identifier}' ({CONFIG_FILE_PATH}): {e}")
            QMessageBox.critical(self, "Erreur Sauvegarde Config",
                                 f"Impossible de sauvegarder config pour '{identifier}': {e}")

    def get_current_excel_position(self):
        if self.excel_wb and hasattr(self.excel_wb.app, 'api') and self.excel_wb.app.visible:
            try:
                _ = self.excel_wb.name
                app_api = self.excel_wb.app.api
                return {
                    "left": int(app_api.Left),
                    "top": int(app_api.Top),
                    "width": int(app_api.Width),
                    "height": int(app_api.Height)
                }
            except Exception as e:
                print(f"Impossible de récupérer la position d'Excel (peut-être fermé): {e}")
        return None

    def set_excel_position(self, app_api, pos):
        if pos is None:  # Sécurité supplémentaire
            print("set_excel_position a reçu 'None', impossible d'appliquer.")
            QMessageBox.warning(self, "Erreur Interne", "Tentative d'appliquer une position 'None' à Excel.")
            return
        try:
            # Utilisation de xw.constants pour la portabilité
            if app_api.WindowState == xw.constants.WindowState.xlMaximized:
                app_api.WindowState = xw.constants.WindowState.xlNormal

            app_api.Left = pos["left"]
            app_api.Top = pos["top"]
            app_api.Width = pos["width"]
            app_api.Height = pos["height"]
            print(f"Position Excel appliquée: {pos}")
        except Exception as e:
            print(f"Erreur lors de l'application de la position à Excel: {e}")
            QMessageBox.warning(self, "Erreur Position Excel", f"Impossible de positionner la fenêtre Excel: {e}")

    def toggle_excel_visibility(self):
        if not EXCEL_FILE_PATH.exists():
            QMessageBox.critical(self, "Erreur Fichier", f"Le fichier Excel '{EXCEL_FILE_PATH}' n'a pas été trouvé.")
            return

        try:
            mode_test_actif = self.chk_mode_test.isChecked()

            if self.excel_wb:  # Si une référence existe
                try:
                    _ = self.excel_wb.name  # Test de validité
                    if self.excel_wb.app.visible:  # Excel est visible -> on le masque
                        print("Excel visible, tentative de masquage.")
                        if not mode_test_actif:  # Uniquement si Mode TEST est OFF
                            current_pos = self.get_current_excel_position()
                            if current_pos:
                                self.save_excel_config(EXCEL_WINDOW_IDENTIFIER, current_pos)
                        else:
                            print("Mode TEST actif, position non sauvegardée au masquage.")
                        self.excel_wb.app.visible = False
                        self.btn_open_excel.setText(f"Ouvrir {EXCEL_FILE_NAME}")
                    else:  # Excel est masqué -> on le réaffiche
                        print("Excel non visible, réaffichage.")
                        self.excel_wb.app.visible = True
                        self.excel_wb.activate(steal_focus=True)
                        if not mode_test_actif:  # Uniquement si Mode TEST est OFF
                            pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER)
                            self.set_excel_position(self.excel_wb.app.api, pos_config)
                        else:
                            print("Mode TEST actif, position sauvegardée non appliquée au réaffichage.")
                        self.btn_open_excel.setText(f"Masquer {EXCEL_FILE_NAME}")
                    return
                except Exception as e:
                    print(f"Référence Excel invalide ou Excel fermé ({e}). Réouverture nécessaire.")
                    self.excel_wb = None  # Réinitialiser pour forcer la réouverture ci-dessous

            # Si self.excel_wb est None (ou vient d'être mis à None) -> Ouverture d'Excel
            print(f"Tentative d'ouverture de {EXCEL_FILE_PATH}")
            try:
                # Tenter de se connecter à une instance existante
                found_wb = None
                for app_instance in xw.apps:
                    for wb_in_app in app_instance.books:
                        if wb_in_app.fullname.lower() == str(EXCEL_FILE_PATH).lower():  # Comparer les chemins complets
                            found_wb = wb_in_app
                            print(f"Fichier Excel '{EXCEL_FILE_NAME}' trouvé dans une instance existante.")
                            break
                    if found_wb:
                        break
                if found_wb:
                    self.excel_wb = found_wb
                else:
                    self.excel_wb = xw.Book(EXCEL_FILE_PATH)
            except Exception as e_open:
                print(f"Erreur spécifique lors de la tentative d'ouverture/connexion à Excel: {e_open}")
                self.excel_wb = xw.Book(EXCEL_FILE_PATH)  # Fallback sur ouverture simple

            self.excel_wb.app.visible = True
            self.excel_wb.activate(steal_focus=True)

            if not mode_test_actif:  # Uniquement si Mode TEST est OFF
                pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER)
                self.set_excel_position(self.excel_wb.app.api, pos_config)
            else:
                print("Mode TEST actif, position sauvegardée non appliquée à la première ouverture.")

            self.btn_open_excel.setText(f"Masquer {EXCEL_FILE_NAME}")
            print(f"{EXCEL_FILE_NAME} ouvert.")
            if not mode_test_actif: print("Position/dimension appliquées.")


        except Exception as e:
            self.excel_wb = None
            print(f"Erreur gestion Excel globale: {e}\n{traceback.format_exc()}")
            QMessageBox.critical(self, "Erreur Excel", f"Erreur avec Excel: {e}")
            self.btn_open_excel.setText(f"Ouvrir {EXCEL_FILE_NAME}")

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Confirmation',
                                     "Voulez-vous vraiment quitter ?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            print("Fermeture de l'application PyQt.")
            if self.excel_wb:
                try:
                    _ = self.excel_wb.name  # Vérifie si wb est valide
                    # La sauvegarde (ou non) est gérée par save_excel_config en fonction de Mode TEST
                    if self.excel_wb.app.visible and not self.chk_mode_test.isChecked():
                        current_pos = self.get_current_excel_position()
                        if current_pos:
                            self.save_excel_config(EXCEL_WINDOW_IDENTIFIER, current_pos)

                    print(f"Fermeture du fichier Excel: {self.excel_wb.name}")
                    self.excel_wb.close()  # Ferme le classeur
                    self.excel_wb = None
                except Exception as e:
                    print(f"Erreur lors de la gestion d'Excel à la fermeture de PyQt: {e}\n{traceback.format_exc()}")
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())