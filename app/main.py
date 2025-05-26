import sys
import os
from pathlib import Path
# Import QHBoxLayout
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, \
    QCheckBox
import xlwings as xw
import xml.etree.ElementTree as ET

# Déterminer le chemin racine du projet (Trading2)
PROJECT_ROOT = Path(__file__).resolve().parent.parent  # Correct si main.py est dans un sous-dossier comme 'app'

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
        # Ajustez la largeur si nécessaire, mais QHBoxLayout essaiera de s'adapter
        self.setGeometry(100, 100, 320, 100)  # Légèrement ajusté la largeur initiale

        self.excel_wb = None
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)

        # Layout principal vertical
        main_layout = QVBoxLayout()

        # --- Conteneur pour le bouton et la case à cocher (layout horizontal) ---
        controls_container = QWidget()
        h_layout = QHBoxLayout(controls_container)  # Appliquer QHBoxLayout à ce conteneur
        h_layout.setContentsMargins(0, 0, 0, 0)  # Optionnel: réduire les marges internes
        h_layout.setSpacing(5)  # Optionnel: espacement entre bouton et case

        self.btn_open_excel = QPushButton(f"Ouvrir {EXCEL_FILE_NAME}")
        self.btn_open_excel.clicked.connect(self.toggle_excel_visibility)
        h_layout.addWidget(self.btn_open_excel)  # Ajout au layout horizontal

        self.chk_un_save = QCheckBox("unSave")  # Utiliser le texte "unSave" comme sur l'image
        self.chk_un_save.setChecked(False)
        h_layout.addWidget(self.chk_un_save)  # Ajout au layout horizontal
        # -----------------------------------------------------------------------

        main_layout.addWidget(controls_container)  # Ajout du conteneur horizontal au layout vertical

        # Widget central et application du layout principal
        container_widget = QWidget()
        container_widget.setLayout(main_layout)
        self.setCentralWidget(container_widget)

    # ... (le reste de votre code load_excel_config, save_excel_config, etc. reste identique) ...
    def load_excel_config(self, identifier):
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
        if self.chk_un_save.isChecked():
            print(f"Sauvegarde désactivée pour '{identifier}' via la case à cocher 'unSave'.")
            return

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
        try:
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
            if self.excel_wb:
                try:
                    _ = self.excel_wb.name
                    if self.excel_wb.app.visible:
                        print("Excel visible, tentative de masquage.")
                        current_pos = self.get_current_excel_position()
                        if current_pos:
                            self.save_excel_config(EXCEL_WINDOW_IDENTIFIER, current_pos)
                        self.excel_wb.app.visible = False
                        self.btn_open_excel.setText(f"Ouvrir {EXCEL_FILE_NAME}")
                    else:
                        print("Excel non visible, réaffichage avec position.")
                        pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER)
                        self.excel_wb.app.visible = True
                        self.excel_wb.activate(steal_focus=True)
                        self.set_excel_position(self.excel_wb.app.api, pos_config)
                        self.btn_open_excel.setText(f"Masquer {EXCEL_FILE_NAME}")
                    return
                except Exception as e:
                    print(f"Référence Excel invalide ou Excel fermé ({e}). Réouverture.")
                    self.excel_wb = None

            print(f"Tentative d'ouverture de {EXCEL_FILE_PATH}")
            try:
                found_wb = None
                for app_instance in xw.apps:  # Renommé app en app_instance pour éviter conflit de nom
                    for wb_in_app in app_instance.books:
                        if wb_in_app.name == EXCEL_FILE_NAME:
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
                self.excel_wb = xw.Book(EXCEL_FILE_PATH)

            pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER)
            self.excel_wb.app.visible = True
            self.excel_wb.activate(steal_focus=True)
            self.set_excel_position(self.excel_wb.app.api, pos_config)
            self.btn_open_excel.setText(f"Masquer {EXCEL_FILE_NAME}")
            print(f"{EXCEL_FILE_NAME} ouvert et positionné.")

        except Exception as e:
            self.excel_wb = None
            import traceback  # Import local si non utilisé ailleurs globalement
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
                    _ = self.excel_wb.name
                    if self.excel_wb.app.visible:
                        current_pos = self.get_current_excel_position()
                        if current_pos:
                            self.save_excel_config(EXCEL_WINDOW_IDENTIFIER, current_pos)
                    print(f"Fermeture du fichier Excel: {self.excel_wb.name}")
                    self.excel_wb.close()
                    self.excel_wb = None
                except Exception as e:
                    print(f"Erreur lors de la gestion d'Excel à la fermeture de PyQt: {e}")
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    # import traceback # Déjà importé localement dans toggle_excel_visibility si besoin
    app = QApplication(sys.argv)  # 'app' est le nom standard pour QApplication
    window = MainWindow()
    window.show()
    sys.exit(app.exec())