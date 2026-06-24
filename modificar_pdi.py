import os
import shutil
import openpyxl
import sys

def check_and_elevate():
    """Eleva los privilegios del script a Administrador en Windows si es necesario."""
    if os.name == 'nt':
        import ctypes
        try:
            if not ctypes.windll.shell32.IsUserAnAdmin():
                print("Elevando privilegios a Administrador para configurar la impresora a doble cara...")
                params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
                if getattr(sys, 'frozen', False):
                    # Si está compilado con PyInstaller (.exe)
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
                else:
                    # Si se ejecuta como script .py
                    script_path = os.path.abspath(sys.argv[0])
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script_path}" {params}', None, 1)
                sys.exit(0)
        except Exception as e:
            print(f"Error al intentar elevar privilegios automáticamente: {e}")
            print("Por favor, ejecuta este programa haciendo clic derecho y seleccionando 'Ejecutar como administrador'.")
            input("\nPresiona Enter para continuar...")

def is_jimny_sheet_match(wb, sheet_name, jimny_doors):
    if not jimny_doors:
        return True
    try:
        sheet = wb[sheet_name]
        d9_val = str(sheet['D9'].value or '').upper()
        is_5_door = ('5' in d9_val or '5PTAS' in d9_val or '5 PTS' in d9_val)
        if jimny_doors == "5" and is_5_door:
            return True
        if jimny_doors == "3" and not is_5_door:
            return True
    except Exception:
        pass
    return False

def process_car():
    print("=" * 60)
    print("  SUZUKI - CONFIGURADOR Y CONFIGURACIÓN RÁPIDA DE PDI")
    print("=" * 60)
    
    # 1. Resolver la ruta del archivo de Excel
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(script_dir, "FORMATO PDI.xlsx")
    
    if not os.path.exists(excel_file):
        excel_file = "FORMATO PDI.xlsx"
        if not os.path.exists(excel_file):
            print("Error: No se encontró el archivo 'FORMATO PDI.xlsx'.")
            print("Por favor, asegúrate de colocar este script en la misma carpeta que el archivo Excel.")
            return False

    # 2. Cargar el libro de trabajo (conservando fórmulas)
    print("Cargando formato Excel...")
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception as e:
        print(f"Error al abrir el archivo Excel: {e}")
        return False

    sheet_names = wb.sheetnames

    # 3. Mostrar únicamente los modelos autorizados
    models_list = ["BALENO", "SWIFT", "DZIRE", "ERTIGA", "FRONX", "JIMNY"]
    
    print("\nModelos disponibles:")
    for idx, model_opt in enumerate(models_list, 1):
        print(f"  {idx}. {model_opt}")
        
    while True:
        try:
            sel = int(input(f"Selecciona el modelo (1-{len(models_list)}): "))
            if 1 <= sel <= len(models_list):
                selected_model = models_list[sel - 1]
                break
        except ValueError:
            pass
        print("Opción inválida. Intenta de nuevo.")

    # Pregunta especial para JIMNY: ¿3 o 5 puertas?
    jimny_doors = ""
    if selected_model == "JIMNY":
        print("\n¿Cuántas puertas tiene el Jimny?")
        print("  1. 3 PTAS")
        print("  2. 5 PTAS")
        while True:
            try:
                sel_doors = int(input("Selecciona (1-2): "))
                if sel_doors == 1:
                    jimny_doors = "3"
                    break
                elif sel_doors == 2:
                    jimny_doors = "5"
                    break
            except ValueError:
                pass
            print("Opción inválida. Intenta de nuevo.")

    # 4. Extraer líneas disponibles para el modelo seleccionado (aplicando filtro de puertas si es Jimny)
    lines_set = set()
    for name in sheet_names:
        name_clean = name.strip()
        name_upper = name_clean.upper()
        if selected_model == "VITARA" and "GRAND VITARA" in name_upper:
            continue
        if selected_model in name_upper:
            if selected_model == "JIMNY" and not is_jimny_sheet_match(wb, name_clean, jimny_doors):
                continue
                
            cleaned = name_upper.replace("OK", "").replace("(2)", "").strip()
            words = cleaned.split()
            try:
                model_len = len(selected_model.split())
                trans_idx = -1
                for idx, w in enumerate(words):
                    if w in ["TA", "TM", "CVT"]:
                        trans_idx = idx
                        break
                if trans_idx != -1:
                    line_words = words[model_len:trans_idx]
                    line_str = " ".join(line_words)
                    if line_str:
                        lines_set.add(line_str)
            except Exception:
                pass
                
    lines_list = sorted(list(lines_set))
    
    # Seleccionar o escribir la línea
    selected_line = ""
    if lines_list:
        print(f"\nLíneas disponibles para {selected_model}:")
        for idx, line_opt in enumerate(lines_list, 1):
            print(f"  {idx}. {line_opt}")
        print(f"  {len(lines_list) + 1}. Escribir otra línea manualmente")
        
        while True:
            try:
                sel = int(input(f"Selecciona la línea (1-{len(lines_list) + 1}): "))
                if 1 <= sel <= len(lines_list):
                    selected_line = lines_list[sel - 1]
                    break
                elif sel == len(lines_list) + 1:
                    selected_line = input("Escribe la línea (ej. GLS, GLX, SPORT): ").strip().upper()
                    break
            except ValueError:
                pass
            print("Opción inválida. Intenta de nuevo.")
    else:
        selected_line = input("\n¿Qué línea es? (ej. GLS, GLX, SPORT): ").strip().upper()

    # Preguntar por la transmisión (TA, TM o CVT)
    print("\nTransmisiones:")
    print("  1. TA")
    print("  2. TM")
    print("  3. CVT")
    while True:
        try:
            sel_trans = int(input("Selecciona la transmisión (1-3): "))
            if sel_trans == 1:
                selected_trans = "TA"
                break
            elif sel_trans == 2:
                selected_trans = "TM"
                break
            elif sel_trans == 3:
                selected_trans = "CVT"
                break
        except ValueError:
            pass
        print("Opción inválida. Intenta de nuevo.")

    # 5. Lógica de búsqueda de la hoja con fallback
    trans_search_order = [selected_trans]
    if selected_trans == "CVT":
        trans_search_order.append("TA")
    elif selected_trans == "TA":
        trans_search_order.append("CVT")

    matched_sheet_name = None
    
    # Buscar por coincidencia exacta de modelo + línea + transmisión (y filtro Jimny si aplica)
    for t_option in trans_search_order:
        for sheet_name in sheet_names:
            name_up = sheet_name.upper()
            if selected_model == "VITARA" and "GRAND VITARA" in name_up:
                continue
            if selected_model == "JIMNY" and not is_jimny_sheet_match(wb, sheet_name, jimny_doors):
                continue
            if (selected_model in name_up) and (selected_line in name_up) and (t_option in name_up):
                matched_sheet_name = sheet_name
                break
        if matched_sheet_name:
            break

    # Si no se encuentra, buscar solo por modelo + transmisión (ignorando la línea)
    if not matched_sheet_name:
        for t_option in trans_search_order:
            for sheet_name in sheet_names:
                name_up = sheet_name.upper()
                if selected_model == "VITARA" and "GRAND VITARA" in name_up:
                    continue
                if selected_model == "JIMNY" and not is_jimny_sheet_match(wb, sheet_name, jimny_doors):
                    continue
                if (selected_model in name_up) and (t_option in name_up):
                    matched_sheet_name = sheet_name
                    break
            if matched_sheet_name:
                break

    # Si aún no se encuentra, usar la primera hoja que coincida con el modelo (y filtro Jimny si aplica)
    if not matched_sheet_name:
        for sheet_name in sheet_names:
            name_up = sheet_name.upper()
            if selected_model == "VITARA" and "GRAND VITARA" in name_up:
                continue
            if selected_model == "JIMNY" and not is_jimny_sheet_match(wb, sheet_name, jimny_doors):
                continue
            if selected_model in name_up:
                matched_sheet_name = sheet_name
                break

    if not matched_sheet_name:
        print(f"\nError: No se pudo encontrar ninguna plantilla para {selected_model}.")
        return False

    print(f"\n--> Plantilla seleccionada: '{matched_sheet_name}'")
    sheet = wb[matched_sheet_name]

    # Ajustar la página en openpyxl también por seguridad
    try:
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 2
    except Exception:
        pass

    # 6. Preguntar por los demás datos
    color = input("\n¿Qué color es? (ej. AZUL PASCAL, PLATA SIBERIA): ").strip().upper()
    key_num = input("¿Qué número de llave es? (opcional, presiona Enter para mantener la del formato): ").strip()

    # Obtener el VIN actual de la plantilla
    vin_template = sheet['D10'].value
    if not vin_template:
        vin_template = "MA3ZFEFS5VA376378" # Valor base fallback
    else:
        vin_template = str(vin_template).strip()

    # Preguntar por el dígito verificador después de la 'S'
    s_idx = vin_template.upper().find('S', 3)
    if s_idx != -1 and s_idx < len(vin_template) - 1:
        char_after_s_default = vin_template[s_idx + 1]
        print(f"\nVIN base: {vin_template}")
        print(f"La letra 'S' está en la posición {s_idx + 1}.")
        val_after_s = input(f"¿Qué número o letra tiene después de la S (actualmente '{char_after_s_default}')? ").strip().upper()
        if not val_after_s:
            val_after_s = char_after_s_default
        check_digit_idx = s_idx + 1
    else:
        char_after_s_default = vin_template[8] if len(vin_template) > 8 else "X"
        print(f"\nVIN base: {vin_template}")
        val_after_s = input(f"¿Qué número o letra tiene en la 9ª posición del VIN (actualmente '{char_after_s_default}')? ").strip().upper()
        if not val_after_s:
            val_after_s = char_after_s_default
        check_digit_idx = 8

    # Preguntar por la terminación del VIN
    while True:
        vin_end = input("\nIntroduce los últimos números del VIN (terminación, ej. 375215): ").strip()
        if vin_end.isdigit():
            break
        print("La terminación debe contener solo números. Intenta de nuevo.")

    # 7. Construir el nuevo VIN
    vin_len = len(vin_template)
    suffix_len = len(vin_end)
    
    part1 = vin_template[:check_digit_idx]
    part2 = val_after_s
    part3 = vin_template[check_digit_idx + 1 : vin_len - suffix_len]
    part4 = vin_end
    
    new_vin = part1 + part2 + part3 + part4
    print(f"\n--> Nuevo VIN generado: {new_vin}")

    # 8. Modificar la hoja de Excel
    clean_sheet_title = matched_sheet_name.replace("OK", "").replace("(2)", "").strip()
    sheet['D9'].value = f"{clean_sheet_title} {color}".strip()
    sheet['D10'].value = new_vin
    
    if key_num:
        sheet['D11'].value = key_num

    # Establecer la hoja modificada como la pestaña activa
    wb.active = wb.sheetnames.index(matched_sheet_name)

    # 9. Hacer respaldo y guardar archivo
    backup_file = os.path.join(script_dir, "FORMATO PDI_backup.xlsx")
    print(f"\nCreando respaldo en: {backup_file}")
    try:
        shutil.copyfile(excel_file, backup_file)
    except Exception as e:
        print(f"Advertencia: No se pudo crear el archivo de respaldo: {e}")

    print("Guardando cambios en Excel...")
    try:
        wb.save(excel_file)
        wb.close()
        print("¡Cambios guardados con éxito!")
    except Exception as e:
        print(f"Error al guardar los cambios: {e}")
        print("Asegúrate de que el archivo Excel no esté abierto en otro programa.")
        return False

    # 10. Comando de impresión rápida para Windows
    print("\n" + "=" * 40)
    print("              IMPRESIÓN")
    print("=" * 40)
    
    abs_excel_path = os.path.abspath(excel_file)
    
    try:
        import win32com.client
        import win32print
        
        # 10a. Configurar la impresora para impresión a doble cara (Duplex = 2)
        printer_name, original_duplex = None, None
        try:
            printer_name = win32print.GetDefaultPrinter()
            # Solicitar acceso de administración de impresora (requiere privilegios elevados, los cuales ya garantizamos con check_and_elevate)
            PRINTER_ALL_ACCESS = 0xF0000 | 0x0004 | 0x0008 | 0x0001 | 0x0002
            handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ALL_ACCESS})
            info = win32print.GetPrinter(handle, 2)
            devmode = info['pDevMode']
            original_duplex = getattr(devmode, 'Duplex', 1)
            
            # Duplex = 2 (doble cara, voltear por el borde largo)
            devmode.Duplex = 2
            win32print.SetPrinter(handle, 2, info, 0)
            win32print.ClosePrinter(handle)
            print(f"Impresora '{printer_name}' configurada a DOBLE CARA con éxito.")
        except Exception as e:
            print(f"Error: No se pudo configurar la impresora a doble cara automáticamente: {e}")
            print("Por favor, asegúrate de haber otorgado los permisos de administrador al abrir la aplicación.")

        # 10b. Imprimir la hoja usando Excel COM
        print("Iniciando instancia aislada de Excel en segundo plano...")
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        
        try:
            workbook = excel_app.Workbooks.Open(abs_excel_path)
            if workbook is None:
                workbook = excel_app.ActiveWorkbook
                
            if workbook is not None:
                active_sheet = workbook.ActiveSheet
                if active_sheet is not None:
                    # Configurar para que la hoja activa se ajuste exactamente a 1 página de ancho por 2 de alto
                    try:
                        active_sheet.PageSetup.Zoom = False
                        active_sheet.PageSetup.FitToPagesWide = 1
                        active_sheet.PageSetup.FitToPagesTall = 2
                        print("Ajuste de página configurado: Ancho = 1 pág., Alto = 2 págs. (1 hoja física doble cara).")
                    except Exception as pe:
                        print(f"Advertencia: No se pudo configurar el ajuste de tamaño de página: {pe}")
                    
                    print(f"Mandando a imprimir la hoja activa: '{active_sheet.Name}'...")
                    active_sheet.PrintOut()
                else:
                    print("Error: No se pudo acceder a la hoja activa.")
                workbook.Close(False)
                print("¡Impresión enviada correctamente!")
            else:
                print("Error: Excel no pudo abrir el libro de trabajo.")
        except Exception as e:
            print(f"Error durante el proceso de impresión con Excel: {e}")
        finally:
            excel_app.Quit()

        # 10c. Restaurar la configuración original de duplex de la impresora
        if printer_name and original_duplex is not None:
            try:
                PRINTER_ALL_ACCESS = 0xF0000 | 0x0004 | 0x0008 | 0x0001 | 0x0002
                handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ALL_ACCESS})
                info = win32print.GetPrinter(handle, 2)
                devmode = info['pDevMode']
                devmode.Duplex = original_duplex
                win32print.SetPrinter(handle, 2, info, 0)
                win32print.ClosePrinter(handle)
                print("Configuración original de la impresora restaurada.")
            except Exception as e:
                print(f"Advertencia: No se pudo restaurar el duplex original de la impresora: {e}")
            
    except ImportError:
        print("Los módulos 'pywin32' o 'win32print' no están instalados en este sistema Python.")
        if hasattr(os, 'startfile'):
            print("Intentando enviar la impresión usando el comando estándar de Windows...")
            try:
                os.startfile(abs_excel_path, "print")
                print("¡Comando de impresión de Windows ejecutado!")
            except Exception as e:
                print(f"No se pudo imprimir automáticamente: {e}")
        else:
            print("El comando de impresión rápida de Windows (os.startfile) no está disponible en este sistema operativo.")
            
    return True

def main():
    # Elevar privilegios al inicio del programa en Windows para garantizar acceso a la impresora
    check_and_elevate()

    while True:
        process_car()
        
        print("\n" + "-" * 60)
        opcion = input("¿Deseas configurar otro automóvil? (S/N): ").strip().upper()
        if opcion not in ["S", "SI"]:
            print("\n¡Gracias por utilizar la herramienta Suzuki PDI Configurator!")
            print("Proceso finalizado.")
            input("Presiona Enter para salir...")
            break
        print("\n" * 2)

if __name__ == "__main__":
    main()
