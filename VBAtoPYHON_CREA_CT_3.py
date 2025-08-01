import os
import sys
from openpyxl import load_workbook
from docxtpl import DocxTemplate
from datetime import datetime
import msoffcrypto
import io
from getpass import getpass

# --- Configuración de rutas de archivos ---
# La ruta de la plantilla de Word.
word_template_path = r"Y:\COSTES\PLANTILLAS\CARACTERISTICAS TECNICAS.docx"

def get_value_or_default(sheet, cell_address, default_value=""):
    """
    Función auxiliar para obtener un valor de una celda,
    manejando celdas vacías o con errores.
    """
    value = sheet[cell_address].value
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return default_value
    return value

def create_ct_document(excel_file_path, word_template_path, output_directory):
    """
    Función principal que lee datos de Excel, procesa la lógica y
    genera el documento de Word.
    """
    try:
        # --- BLOQUE DE CÓDIGO: DESENCRIPTADO Y CARGA DEL ARCHIVO EXCEL ---
        print(f"Abriendo el archivo: {os.path.basename(excel_file_path)}")
        password = getpass("Entra la contraseña para el archivo de Excel: ")
        
        decrypted_workbook_stream = io.BytesIO()
        
        try:
            with open(excel_file_path, 'rb') as file:
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=str(password))
                office_file.decrypt(decrypted_workbook_stream)
        except msoffcrypto.exceptions.DecryptionError:
            print("Error: La contraseña es incorrecta. Por favor, inténtalo de nuevo.")
            return
        except FileNotFoundError:
            print(f"Error: No se encontró el archivo de Excel en la ruta: {excel_file_path}")
            return
        
        print("Archivo desencriptado con éxito.")
        
        workbook = load_workbook(filename=decrypted_workbook_stream, data_only=True)
        sheet = workbook["CT"]

        # --- Lógica Condicional: Asignación de valores de Excel ---
        ManometroAnalogico = ""
        ManometroElectronico = ""
        longitudMM = ""
        diamCol = ""
        diamEle = ""
        AnguloTlv = ""
        materialJaula = ""
        mater_CAL = ""
        mater_CAS = ""
        mater_CHV = ""
        mater_TLV = ""
        espTolvaValue = ""
        
        # Lectura de los valores base de las celdas
        b1_value = get_value_or_default(sheet, "B1")
        b3_value = get_value_or_default(sheet, "B3")
        b4_value = get_value_or_default(sheet, "B4")
        b5_value = get_value_or_default(sheet, "B5")
        b6_value = get_value_or_default(sheet, "B6")
        b7_value = get_value_or_default(sheet, "B7")
        b8_value = get_value_or_default(sheet, "B8")
        b9_value = get_value_or_default(sheet, "B9")
        b10_value = get_value_or_default(sheet, "B10")
        b11_value = get_value_or_default(sheet, "B11")
        b12_value = get_value_or_default(sheet, "B12")
        b13_value = get_value_or_default(sheet, "B13")
        b14_value = get_value_or_default(sheet, "B14")
        b15_value = get_value_or_default(sheet, "B15")
        b16_value = get_value_or_default(sheet, "B16")
        b17_value = get_value_or_default(sheet, "B17")
        b18_value = get_value_or_default(sheet, "B18")
        b19_value = get_value_or_default(sheet, "B19")
        b20_value = get_value_or_default(sheet, "B20")
        b21_value = get_value_or_default(sheet, "B21")
        b22_value = get_value_or_default(sheet, "B22")
        b23_value = get_value_or_default(sheet, "B23")
        b24_value = get_value_or_default(sheet, "B24")
        b25_value = get_value_or_default(sheet, "B25")
        b26_value = get_value_or_default(sheet, "B26")
        b27_value = get_value_or_default(sheet, "B27")
        b28_value = get_value_or_default(sheet, "B28")
        b30_value = get_value_or_default(sheet, "B30")
        b31_value = get_value_or_default(sheet, "B31")
        b32_value = get_value_or_default(sheet, "B32")

        # --- Traducción de la lógica condicional de VBA a Python ---
        # Manómetros
        if b30_value in [0, 1, 2, 3, 4, 5, 6]:
            ManometroAnalogico = "Esfera 0-300 mmca"
            ManometroElectronico = "Timer-manómetro. 220 Vac / IP56 / 50 Hz "
        else:
            ManometroAnalogico = "-"
            ManometroElectronico = "Timer-manómetro. 220 Vac / IP56 / 50 Hz / signal 4-20 mA"

        # Ángulo tolva 60 / 70 º
        if b22_value == "2":
            AnguloTlv = "70"
        else:
            AnguloTlv = "60"

        # Longitud de las mangas
        if b25_value == 4: longitudMM = "1.261"
        elif b25_value == 6: longitudMM = "1.870"
        elif b25_value == 8: longitudMM = "2.479"
        elif b25_value == 10: longitudMM = "3.088"
        elif b25_value == 12: longitudMM = "3.697"
        else: longitudMM = "0"
        
        # Diámetros
        if b27_value in [2, 3, 8, 9]:
            diamCol = 6
            diamEle = 1
        elif b27_value in [4, 5, 10, 11]:
            diamCol = 8
            diamEle = "1 1/2"
        elif b27_value in [6, 7, 12, 13]:
            diamCol = 8
            diamEle = 2
        else:
            diamCol = ""
            diamEle = ""

        # Material de la jaula
        materialJaula = b26_value
        if materialJaula == "Pintadas":
            materialJaula = "Acero pintado"

        # Mater_CAL
        if b18_value == 1: mater_CAL = "S235JR"
        elif b18_value == 2: mater_CAL = "AISI-304"
        elif b18_value == 3: mater_CAL = "AISI-316"
        else: mater_CAL = ""

        # Mater_CAS
        if b20_value == 1: mater_CAS = "S235JR"
        elif b20_value == 2: mater_CAS = "AISI-304"
        elif b20_value == 3: mater_CAS = "AISI-316"
        else: mater_CAS = ""

        # Mater_TLV y espTolvaValue
        if b31_value in ["A", "AE", "PL"]:
            espTolvaValue = "-"
            mater_TLV = "-"
        else:
            espTolvaValue = b17_value
            if b21_value == 1: mater_TLV = "S235JR"
            elif b21_value == 2: mater_TLV = "AISI-304"
            elif b21_value == 3: mater_TLV = "AISI-316"
            else: mater_TLV = ""

        # Mater_CHV
        if b19_value == 1: mater_CHV = "S235JR"
        elif b19_value == 2: mater_CHV = "AISI-304"
        elif b19_value == 3: mater_CHV = "AISI-316"
        else: mater_CHV = ""

        # Modificaciones para campos específicos (BAR y CONSUMO_AIRE)
        if b32_value is None:
            bar_value = "XXXXXXX"
        else:
            bar_value = b32_value

        if isinstance(b11_value, (int, float)):
            consumo_aire_value = f"{int(b11_value)}"
        else:
            consumo_aire_value = "XX"
        
        # --- Preparar el contexto para la plantilla de Word ---
        context = {
            "NOF": str(b1_value),
            "CAUDAL": f"{int(b3_value):,}" if isinstance(b3_value, (int, float)) else b3_value,
            "PRODUCT": str(b5_value),
            "CONC": str(b6_value) if not b6_value in [None, 0] else "20÷30",
            "DENS": str(b7_value) if not b7_value in [None, 0] else "800",
            "TEMPERATURA": str(b4_value) if not b4_value in [None] else "20",
            "FILTRO": str(b8_value),
            "SUP_FILTRANTE": str(b9_value),
            "RATIO_F": f"{float(b10_value):.2f}" if isinstance(b10_value, (int, float)) else b10_value,
            "P_T": str(b12_value),
            "P_D": str(b13_value),
            "ESP_CAL": str(b14_value),
            "ESP_VENT": str(b15_value),
            "ESP_CAS": str(b16_value),
            "ESP_TOLVA": str(espTolvaValue),
            "ANG": AnguloTlv,
            "MAT_MANGA": str(b23_value),
            "NUM_MANGAS": str(b24_value),
            "LONG_MANGA": str(longitudMM),
            "MAT_JAULA": str(materialJaula),
            "NUM_VALV": str(b28_value),
            "TIMER_MAN_CT": str(ManometroElectronico),
            "MAN_CT": str(ManometroAnalogico),
            "DIAM_COL": str(diamCol),
            "DIAM_ELE": str(diamEle),
            "MATER_CAL": str(mater_CAL),
            "MATER_CAS": str(mater_CAS),
            "MATER_TLV": str(mater_TLV),
            "MATER_CHV": str(mater_CHV),
            "BAR": str(bar_value),
            "CONSUMO_AIRE": consumo_aire_value
        }
        
        # --- Generar el documento de Word ---
        doc = DocxTemplate(word_template_path)
        doc.render(context)

        # Crear el nombre del archivo de salida y guardarlo
        output_filename = f"CT-{b1_value}.docx"
        output_path = os.path.join(output_directory, output_filename)
        doc.save(output_path)

        print(f"Documento completado guardado en:\n{output_path}")

    except Exception as e:
        print(f"Ocurrió un error inesperado al procesar los archivos.")
        print(f"Detalles del error: {e}")

# --- Ejecución del script ---
if __name__ == "__main__":
    # Comprueba si se ha pasado un argumento (la ruta del archivo de Excel).
    if len(sys.argv) > 1:
        excel_file_path = sys.argv[1]
        
        # Obtiene la ruta de la carpeta donde se encuentra el archivo de Excel.
        output_directory = os.path.dirname(excel_file_path)
        
        create_ct_document(excel_file_path, word_template_path, output_directory)
    else:
        print("Uso: Arrastra y suelta un archivo de Excel sobre este programa o úsalo con la opción 'Enviar a'.")
        # El programa esperará a que el usuario presione una tecla antes de cerrarse
        # para que pueda leer el mensaje de error.
        input("Presiona Enter para salir...")
