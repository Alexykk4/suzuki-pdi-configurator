import os
import shutil
import openpyxl

def main():
    print("=" * 60)
    print("  SUZUKI - CONFIGURADOR Y CONFIGURACIÓN RÁPIDA DE PDI")
    print("=" * 60)
    
    # 1. Resolver la ruta del archivo de Excel
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(script_dir, "FORMATO PDI.xlsx")
    
    if not os.path.exists(excel_file):
        # Intentar en el directorio de trabajo actual
        excel_file = "FORMATO PDI.xlsx"
        if not os.path.exists(excel_file):
            print(f"Error: No se encontró el archivo 'FORMATO PDI.xlsx'.")
            print("Por favor, asegúrate de colocar este script en la misma carpeta que el archivo Excel.")
            input("\nPresiona Enter para salir...")
            return

    # 2. Cargar el libro de trabajo (conservando fórmulas)
    print("Cargando formato Excel...")
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception as e:
        print(f"Error al abrir el archivo Excel: {e}")
        input("\nPresiona Enter para salir...")
        return

    # 3. Extraer modelos disponibles dinámicamente
    sheet_names = wb.sheetnames
    models_set = set()
    for name in sheet_names:
        name_clean = name.strip()
        if name_clean == "BLANCO PARA TECNICO":
            continue
        name_upper = name_clean.upper()
        if name_upper.startswith("GRAND VITARA"):
            models_set.add("GRAND VITARA")
        else:
            words = name_clean.split()
            if words:
                models_set.add(words[0].upper())
                
    models_list = sorted(list(models_set))
    
    # Preguntar: ¿Qué modelo es?
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

    # 4. Extraer líneas disponibles para el modelo seleccionado
    lines_set = set()
    for name in sheet_names:
        name_clean = name.strip()
        name_upper = name_clean.upper()
        if selected_model == "VITARA" and "GRAND VITARA" in name_upper:
            continue
        if selected_model in name_upper:
            # Limpiar nombre de hojas
            cleaned = name_upper.replace("OK", "").replace("(2)", "").strip()
            words = cleaned.split()
            try:
                # Determinar posición de la transmisión (TA, TM, CVT)
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
    
    # Preguntar: ¿Qué línea es?
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

    # Preguntar: ¿Qué transmisión es? (TA, TM o CVT)
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
    # Si responde CVT o TA, busca la que exista
    trans_search_order = [selected_trans]
    if selected_trans == "CVT":
        trans_search_order.append("TA")
    elif selected_trans == "TA":
        trans_search_order.append("CVT")

    matched_sheet_name = None
    
    # Buscar por coincidencia exacta de modelo + línea + transmisión
    for t_option in trans_search_order:
        for sheet_name in sheet_names:
            name_up = sheet_name.upper()
            # Excluir Grand Vitara si seleccionó Vitara
            if selected_model == "VITARA" and "GRAND VITARA" in name_up:
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
                if (selected_model in name_up) and (t_option in name_up):
                    matched_sheet_name = sheet_name
                    break
            if matched_sheet_name:
                break

    # Si aún no se encuentra, usar la primera hoja que coincida con el modelo
    if not matched_sheet_name:
        for sheet_name in sheet_names:
            name_up = sheet_name.upper()
            if selected_model == "VITARA" and "GRAND VITARA" in name_up:
                continue
            if selected_model in name_up:
                matched_sheet_name = sheet_name
                break

    if not matched_sheet_name:
        print(f"\nError: No se pudo encontrar ninguna plantilla para {selected_model}.")
        input("\nPresiona Enter para salir...")
        return

    print(f"\n--> Plantilla seleccionada: '{matched_sheet_name}'")
    sheet = wb[matched_sheet_name]

    # 6. Preguntar por los demás datos
    color = input("\n¿Qué color es? (ej. AZUL PASCAL, PLATA SIBERIA): ").strip().upper()
    key_num = input("¿Qué número de llave es? (opcional, presiona Enter para mantener la del formato): ").strip()

    # Obtener el VIN actual de la plantilla
    vin_template = sheet['D10'].value
    if not vin_template:
        vin_template = "MA3ZFEFS5VA376378" # Valor base por si la celda está vacía
    else:
        vin_template = str(vin_template).strip()

    # Preguntar por el dígito verificador después de la 'S'
    # Buscar la S (después de los caracteres de fabricante, ej. index >= 3)
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
        # Fallback a la posición estándar del check digit (posición 9, index 8)
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
    # El VIN tiene 17 caracteres. Reemplazamos:
    # - Posición del check digit (check_digit_idx)
    # - Posición final (los últimos len(vin_end) caracteres)
    vin_len = len(vin_template)
    suffix_len = len(vin_end)
    
    part1 = vin_template[:check_digit_idx]
    part2 = val_after_s
    part3 = vin_template[check_digit_idx + 1 : vin_len - suffix_len]
    part4 = vin_end
    
    new_vin = part1 + part2 + part3 + part4
    print(f"\n--> Nuevo VIN generado: {new_vin}")

    # 8. Modificar la hoja de Excel
    # Modificar D9 (Modelo + Línea + Transmisión + Color)
    # Limpiamos el nombre de la pestaña de sufijos como OK o (2)
    clean_sheet_title = matched_sheet_name.replace("OK", "").replace("(2)", "").strip()
    sheet['D9'].value = f"{clean_sheet_title} {color}".strip()
    
    # Modificar D10 (VIN)
    sheet['D10'].value = new_vin
    
    # Modificar D11 (Número de llave) si se ingresó algo
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
        input("\nPresiona Enter para salir...")
        return

    # 10. Comando de impresión rápida para Windows
    print("\n" + "=" * 40)
    print("              IMPRESIÓN")
    print("=" * 40)
    
    abs_excel_path = os.path.abspath(excel_file)
    
    try:
        import win32com.client
        print("Iniciando Excel en segundo plano para imprimir...")
        
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        
        try:
            workbook = excel_app.Workbooks.Open(abs_excel_path)
            active_sheet = workbook.ActiveSheet
            print(f"Mandando a imprimir la hoja: '{active_sheet.Title}'...")
            # Imprime únicamente la hoja activa
            active_sheet.PrintOut()
            workbook.Close(False)
            print("¡Impresión enviada correctamente a tu impresora predeterminada!")
        except Exception as e:
            print(f"Error durante el proceso de impresión con Excel: {e}")
        finally:
            excel_app.Quit()
            
    except ImportError:
        print("El módulo 'pywin32' no está instalado en este sistema Python.")
        if hasattr(os, 'startfile'):
            print("Intentando enviar la impresión usando el comando estándar de Windows...")
            try:
                os.startfile(abs_excel_path, "print")
                print("¡Comando de impresión de Windows ejecutado!")
                print("Nota: Esto abrirá Excel brevemente para mandar el documento a la impresora.")
            except Exception as e:
                print(f"No se pudo imprimir automáticamente: {e}")
        else:
            print("El comando de impresión rápida de Windows (os.startfile) no está disponible en este sistema operativo.")
        print("\nPara habilitar la impresión silenciosa y precisa en Windows, instala pywin32:")
        print("  pip install pywin32")

    print("\nProceso finalizado.")
    input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()
