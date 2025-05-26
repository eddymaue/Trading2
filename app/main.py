import sys
import os
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QMessageBox
import xlwings as xw
import xml.etree.ElementTree as ET

# Déterminer le chemin racine du projet (Trading2)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

EXCEL_FILE_NAME = "TradingData.xlsm"
EXCEL_WINDOW_IDENTIFIER = "TradingData" # <- Nouvel identifiant pour le XML

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
        self.setGeometry(100, 100, 300, 100)

        self.excel_wb = None
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)

        layout = QVBoxLayout()
        self.btn_open_excel = QPushButton(f"Ouvrir {EXCEL_FILE_NAME}")
        self.btn_open_excel.clicked.connect(self.toggle_excel_visibility)
        layout.addWidget(self.btn_open_excel)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_excel_config(self, identifier): # <- Prend un identifiant
        """Charge la configuration pour une fenêtre Excel spécifique."""
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
            QMessageBox.warning(self, "Erreur Config", f"Impossible de charger config pour '{identifier}': {e}\nUtilisation des défauts.")
        print(f"Config non trouvée pour '{identifier}'. Utilisation des défauts.")
        return DEFAULT_EXCEL_POS.copy()

    def save_excel_config(self, identifier, pos): # <- Prend un identifiant et la position
        """Sauvegarde la configuration pour une fenêtre Excel spécifique."""
        try:
            if CONFIG_FILE_PATH.exists():
                try:
                    tree = ET.parse(CONFIG_FILE_PATH)
                    root_config = tree.getroot()
                except ET.ParseError: # Si le XML est corrompu, on repart de zéro
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
                # Vider les anciens éléments enfants pour mettre à jour
                for child in list(specific_window_node):
                    specific_window_node.remove(child)

            ET.SubElement(specific_window_node, "left").text = str(int(pos["left"]))
            ET.SubElement(specific_window_node, "top").text = str(int(pos["top"]))
            ET.SubElement(specific_window_node, "width").text = str(int(pos["width"]))
            ET.SubElement(specific_window_node, "height").text = str(int(pos["height"]))

            # Réécrire l'arbre XML entier
            new_tree = ET.ElementTree(root_config)
            ET.indent(new_tree, space="  ", level=0)
            new_tree.write(CONFIG_FILE_PATH, encoding="utf-8", xml_declaration=True)
            print(f"Config sauvegardée pour '{identifier}' dans {CONFIG_FILE_PATH}: {pos}")

        except Exception as e:
            print(f"Erreur sauvegarde config XML pour '{identifier}' ({CONFIG_FILE_PATH}): {e}")
            QMessageBox.critical(self, "Erreur Sauvegarde Config", f"Impossible de sauvegarder config pour '{identifier}': {e}")


    def get_current_excel_position(self):
        if self.excel_wb and self.excel_wb.app.visible:
            try:
                app_api = self.excel_wb.app.api
                return {
                    "left": int(app_api.Left),
                    "top": int(app_api.Top),
                    "width": int(app_api.Width),
                    "height": int(app_api.Height)
                }
            except Exception as e:
                print(f"Impossible de récupérer la position d'Excel: {e}")
        return None

    def set_excel_position(self, app_api, pos):
        try:
            # Gérer le cas où la fenêtre est maximisée
            if app_api.WindowState == -4137: # xlMaximized (valeur pour xw.constants.WindowState.xlMaximized)
                 app_api.WindowState = -4143 # xlNormal (valeur pour xw.constants.WindowState.xlNormal)
                 # Il faut parfois un petit délai pour que le changement d'état soit pris en compte
                 # ou réappliquer les dimensions après. Pour l'instant, on essaie directement.

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
                    _ = self.excel_wb.name # Test de validité de la connexion
                    if self.excel_wb.app.visible:
                        print("Excel visible, sauvegarde position et masquage.")
                        current_pos = self.get_current_excel_position()
                        if current_pos:
                            self.save_excel_config(EXCEL_WINDOW_IDENTIFIER, current_pos) # Utilise l'identifiant
                        self.excel_wb.app.visible = False
                        self.btn_open_excel.setText(f"Ouvrir {EXCEL_FILE_NAME}")
                    else:
                        print("Excel non visible, réaffichage avec position.")
                        pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER) # Utilise l'identifiant
                        self.excel_wb.app.visible = True
                        self.excel_wb.activate()
                        self.set_excel_position(self.excel_wb.app.api, pos_config)
                        self.btn_open_excel.setText(f"Masquer {EXCEL_FILE_NAME}")
                    return
                except Exception:
                    print("Référence Excel invalide. Réouverture.")
                    self.excel_wb = None

            print(f"Tentative d'ouverture de {EXCEL_FILE_PATH}")
            self.excel_wb = xw.Book(EXCEL_FILE_PATH)
            pos_config = self.load_excel_config(EXCEL_WINDOW_IDENTIFIER) # Utilise l'identifiant
            self.excel_wb.app.visible = True
            self.excel_wb.activate()
            self.set_excel_position(self.excel_wb.app.api, pos_config)
            self.btn_open_excel.setText(f"Masquer {EXCEL_FILE_NAME}")
            print(f"{EXCEL_FILE_NAME} ouvert et positionné.")

        except Exception as e:
            self.excel_wb = None
            print(f"Erreur gestion Excel: {e}")
            QMessageBox.critical(self, "Erreur Excel", f"Erreur avec Excel: {e}")
            self.btn_open_excel.setText(f"Ouvrir {EXCEL_FILE_NAME}")

    def closeEvent(self, event):
        print("Fermeture de l'application PyQt.")
        if self.excel_wb and self.excel_wb.app.visible: # Vérifier aussi la visibilité
            current_pos = self.get_current_excel_position()
            if current_pos:
                self.save_excel_config(EXCEL_WINDOW_IDENTIFIER, current_pos) # Utilise l'identifiant
                print("Position Excel sauvegardée à la fermeture de PyQt.")
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())