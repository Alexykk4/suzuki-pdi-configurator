import os
import shutil
import openpyxl
import sys
import subprocess
import zipfile
import re
import tempfile

def log_message(msg):
    """Muestra un mensaje en consola y lo guarda en un archivo de depuración local."""
    print(msg)
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_file = os.path.join(script_dir, "pdi_debug_log.txt")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass

def strip_printer_settings_from_zip(xlsx_path):
    """Desempaqueta el archivo Excel (.xlsx), elimina las configuraciones de impresora incrustadas (.bin) y limpia sus relaciones XML."""
    log_message("Limpiando la caché de configuración de impresora del archivo Excel...")
    temp_dir = tempfile.mkdtemp()
    temp_zip_path = os.path.join(temp_dir, "temp.zip")
    
    try:
        shutil.copyfile(xlsx_path, temp_zip_path)
        
        with zipfile.ZipFile(temp_zip_path, 'r') as zin:
            with zipfile.ZipFile(xlsx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    filename = item.filename
                    
                    if "printerSettings" in filename:
                        continue
                        
                    data = zin.read(filename)
                    
                    if filename.startswith("xl/worksheets/_rels/") and filename.endswith(".xml.rels"):
                        data_str = data.decode('utf-8', errors='ignore')
                        data_str = re.sub(
                            r'<Relationship[^>]*relationships/printerSettings[^>]*/>', 
                            '', 
                            data_str
                        )
                        data = data_str.encode('utf-8')
                        
                    elif filename.startswith("xl/worksheets/sheet") and filename.endswith(".xml"):
                        data_str = data.decode('utf-8', errors='ignore')
                        
                        def clean_pagesetup(match):
                            tag = match.group(0)
                            tag_cleaned = re.sub(r'\s+r:id="[^"]+"', '', tag)
                            return tag_cleaned
                        
                        data_str = re.sub(r'<pageSetup[^>]+/>', clean_pagesetup, data_str)
                        data = data_str.encode('utf-8')
                        
                    zout.writestr(item, data)
        log_message("--> Caché de impresora limpiada exitosamente del Excel.")
    except Exception as e:
        log_message(f"--> Advertencia al limpiar caché del Excel: {e}")
    finally:
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass

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

def get_excel_printer_string(printer_name):
    """Encuentra el puerto (ej. 'on Ne01:') en el registro de Windows para construir la cadena ActivePrinter de Excel."""
    if os.name != 'nt':
        return printer_name
    try:
        import winreg
        for key_path in [
            r"Software\Microsoft\Windows NT\CurrentVersion\Devices",
            r"Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
        ]:
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                        if name.lower().strip() == printer_name.lower().strip():
                            parts = value.split(",")
                            if len(parts) >= 2:
                                port = parts[1]
                                winreg.CloseKey(key)
                                return f"{name} on {port}"
                        i += 1
                    except OSError:
                        break
                winreg.CloseKey(key)
            except Exception:
                continue
    except Exception as e:
        log_message(f"Advertencia al leer el puerto de la impresora del registro: {e}")
        
    return printer_name

def set_excel_active_printer(excel_app, printer_name):
    """Prueba dinámicamente diferentes puertos de red (Ne00 a Ne99) en inglés y español para configurar ActivePrinter."""
    # 1. Intentar configurar directamente
    try:
        excel_app.ActivePrinter = printer_name
        log_message(f"Impresora activa configurada directamente en Excel: '{printer_name}'")
        return True
    except Exception:
        pass

    # 2. Intentar puertos en Inglés (e.g. "Nombre on Ne01:")
    for i in range(100):
        try:
            port_str = f"{printer_name} on Ne{i:02d}:"
            excel_app.ActivePrinter = port_str
            log_message(f"Impresora activa configurada en Excel: '{port_str}'")
            return True
        except Exception:
            pass

    # 3. Intentar puertos en Español (e.g. "Nombre en Ne01:")
    for i in range(100):
        try:
            port_str = f"{printer_name} en Ne{i:02d}:"
            excel_app.ActivePrinter = port_str
            log_message(f"Impresora activa configurada en Excel: '{port_str}' (Windows en Español)")
            return True
        except Exception:
            pass

    # 4. Fallback leyendo el registro de Windows
    excel_printer_str = get_excel_printer_string(printer_name)
    if excel_printer_str != printer_name:
        try:
            excel_app.ActivePrinter = excel_printer_str
            log_message(f"Impresora activa configurada desde Registro: '{excel_printer_str}'")
            return True
        except Exception:
            pass

        # Probar la versión traducida al español del registro
        if " on " in excel_printer_str:
            spanish_str = excel_printer_str.replace(" on ", " en ")
            try:
                excel_app.ActivePrinter = spanish_str
                log_message(f"Impresora activa configurada desde Registro (Español): '{spanish_str}'")
                return True
            except Exception:
                pass

    raise Exception("No se pudo asociar la impresora con ningún puerto de red de Excel (NeXX:).")

def get_current_duplex_state(printer_name):
    """Obtiene la configuración original de duplex mediante PowerShell o win32print."""
    # 1. Intentar con PowerShell
    try:
        cmd = f"(Get-PrintConfiguration -PrinterName '{printer_name}').DuplexingMode"
        res = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
        if res.returncode == 0 and res.stdout.strip():
            return "powershell", res.stdout.strip()
    except Exception:
        pass

    # 2. Fallback a win32print
    try:
        import win32print
        PRINTER_ACCESS_USE = 0x00000008
        handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})
        info = win32print.GetPrinter(handle, 2)
        devmode = info['pDevMode']
        original_duplex = getattr(devmode, 'Duplex', 1)
        win32print.ClosePrinter(handle)
        return "win32print", original_duplex
    except Exception:
        return None, None

def set_duplex_state(printer_name, duplex=True):
    """Establece la impresión a doble cara (True) o normal (False)."""
    mode = "TwoSidedLongEdge" if duplex else "OneSided"
    
    # 1. Intentar con PowerShell
    try:
        cmd = f"Set-PrintConfiguration -PrinterName '{printer_name}' -DuplexingMode {mode}"
        res = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
        if res.returncode == 0:
            log_message(f"Impresora configurada a DOBLE CARA ({mode}) vía PowerShell.")
            return True
        else:
            log_message(f"PowerShell falló con código {res.returncode}: {res.stderr}")
    except Exception as e:
        log_message(f"Error al ejecutar PowerShell: {e}")

    # 2. Fallback a win32print
    try:
        import win32print
        PRINTER_ACCESS_USE = 0x00000008
        handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})
        info = win32print.GetPrinter(handle, 2)
        devmode = info['pDevMode']
        devmode.Duplex = 2 if duplex else 1
        win32print.SetPrinter(handle, 2, info, 0)
        win32print.ClosePrinter(handle)
        log_message("Impresora configurada a DOBLE CARA vía win32print.")
        return True
    except Exception as e:
        log_message(f"No se pudo forzar la doble cara vía win32print: {e}")
        return False

def restore_duplex_state(printer_name, method, original_val):
    """Restaura la configuración de impresión a su estado original."""
    if not method or original_val is None:
        return
    if method == "powershell":
        try:
            cmd = f"Set-PrintConfiguration -PrinterName '{printer_name}' -DuplexingMode {original_val}"
            subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
            log_message("Configuración original de la impresora restaurada vía PowerShell.")
        except Exception:
            pass
    elif method == "win32print":
        try:
            import win32print
            PRINTER_ACCESS_USE = 0x00000008
            handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})
            info = win32print.GetPrinter(handle, 2)
            devmode = info['pDevMode']
            devmode.Duplex = original_val
            win32print.SetPrinter(handle, 2, info, 0)
            win32print.ClosePrinter(handle)
            log_message("Configuración original de la impresora restaurada vía win32print.")
        except Exception:
            pass

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
            log_message("Error: No se encontró el archivo 'FORMATO PDI.xlsx'.")
            return False

    # 2. Cargar el libro de trabajo (conservando fórmulas)
    log_message("Cargando formato Excel...")
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception as e:
        log_message(f"Error al abrir el archivo Excel: {e}")
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

    # 4. Líneas predefinidas para cada modelo
    model_lines = {
        "BALENO": ["GLS", "GLX"],
        "SWIFT": ["GLS", "GLX", "SPORT"],
        "DZIRE": ["GLS", "GLX"],
        "ERTIGA": ["GLS", "GLX", "XL7"],
        "FRONX": ["GLS", "GLX"],
        "JIMNY": ["GLS", "GLX"]
    }
    
    lines_list = model_lines.get(selected_model, ["GLS", "GLX"])
    
    print(f"\nLíneas disponibles para {selected_model}:")
    for idx, line_opt in enumerate(lines_list, 1):
        print(f"  {idx}. {line_opt}")
        
    while True:
        try:
            sel = int(input(f"Selecciona la línea (1-{len(lines_list)}): "))
            if 1 <= sel <= len(lines_list):
                selected_line = lines_list[sel - 1]
                break
        except ValueError:
            pass
        print("Opción inválida. Intenta de nuevo.")

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
        log_message(f"Error: No se pudo encontrar ninguna plantilla para {selected_model}.")
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
    log_message(f"--> Nuevo VIN generado: {new_vin}")

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
    log_message(f"Creando respaldo en: {backup_file}")
    try:
        shutil.copyfile(excel_file, backup_file)
    except Exception as e:
        log_message(f"Advertencia: No se pudo crear el archivo de respaldo: {e}")

    log_message("Guardando cambios en Excel...")
    try:
        wb.save(excel_file)
        wb.close()
        log_message("¡Cambios guardados con éxito!")
    except Exception as e:
        log_message(f"Error al guardar los cambios: {e}")
        log_message("Asegúrate de que el archivo Excel no esté abierto en otro programa.")
        return False

    # 9.5 LIMPIAR CACHÉ DE IMPRESORA DEL ZIP (.xlsx)
    strip_printer_settings_from_zip(excel_file)

    # 10. Comando de impresión rápida para Windows
    log_message("\n" + "=" * 40)
    log_message("              IMPRESIÓN")
    log_message("=" * 40)
    
    abs_excel_path = os.path.abspath(excel_file)
    
    try:
        import win32com.client
        import win32print
        
        # 10a. Configurar la impresora para impresión a doble cara
        printer_name = win32print.GetDefaultPrinter()
        log_message(f"Impresora predeterminada detectada: '{printer_name}'")
        method, original_val = get_current_duplex_state(printer_name)
        
        # Cambiar el duplex
        set_duplex_state(printer_name, duplex=True)

        # 10b. Imprimir la hoja usando Excel COM
        log_message("Iniciando instancia aislada de Excel en segundo plano...")
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        
        try:
            workbook = excel_app.Workbooks.Open(abs_excel_path)
            
            # Forzar ActivePrinter después de abrir el libro para obligar a Excel a recargar el driver
            try:
                set_excel_active_printer(excel_app, printer_name)
            except Exception as pe:
                log_message(f"Advertencia: No se pudo asignar ActivePrinter en Excel: {pe}")

            if workbook is not None:
                active_sheet = workbook.ActiveSheet
                if active_sheet is not None:
                    # Configurar para que la hoja activa se ajuste exactamente a 1 página de ancho por 2 de alto
                    try:
                        active_sheet.PageSetup.Zoom = False
                        active_sheet.PageSetup.FitToPagesWide = 1
                        active_sheet.PageSetup.FitToPagesTall = 2
                        log_message("Ajuste de página configurado: Ancho = 1 pág., Alto = 2 págs. (1 hoja física doble cara).")
                    except Exception as pe:
                        log_message(f"Advertencia: No se pudo configurar el ajuste de tamaño de página: {pe}")
                    
                    log_message(f"Mandando a imprimir la hoja activa: '{active_sheet.Name}'...")
                    active_sheet.PrintOut()
                else:
                    log_message("Error: No se pudo acceder a la hoja activa.")
                workbook.Close(False)
                log_message("¡Impresión enviada correctamente!")
            else:
                log_message("Error: Excel no pudo abrir el libro de trabajo.")
        except Exception as e:
            log_message(f"Error durante el proceso de impresión con Excel: {e}")
        finally:
            excel_app.Quit()

        # 10c. Restaurar la configuración original de duplex de la impresora
        restore_duplex_state(printer_name, method, original_val)
            
    except ImportError:
        log_message("Los módulos 'pywin32' o 'win32print' no están instalados en este sistema Python.")
        if hasattr(os, 'startfile'):
            log_message("Intentando enviar la impresión usando el comando estándar de Windows...")
            try:
                os.startfile(abs_excel_path, "print")
                log_message("¡Comando de impresión de Windows ejecutado!")
            except Exception as e:
                log_message(f"No se pudo imprimir automáticamente: {e}")
        else:
            log_message("El comando de impresión rápida de Windows (os.startfile) no está disponible en este sistema operativo.")
            
    return True

def main():
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_file = os.path.join(script_dir, "pdi_debug_log.txt")
        if os.path.exists(log_file):
            os.remove(log_file)
    except Exception:
        pass

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
