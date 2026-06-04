import win32com.client
import os
import sys
from dotenv import load_dotenv

# Cargar variables de entorno del archivo .env local
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
else:
    load_dotenv()

CUSTOM_DATA_FILE_PATH = os.getenv("CUSTOM_DATA_FILE_PATH", "")

def get_outlook_root_folder(outlook):
    """
    Obtiene la carpeta raíz de Outlook a procesar.
    Si se configura CUSTOM_DATA_FILE_PATH en .env, intenta montar el archivo (si es PST)
    o muestra una advertencia explicativa con instrucciones de conversión (si es OST).
    Retorna (root_folder, mounted_store_root_to_remove)
    """
    custom_path = CUSTOM_DATA_FILE_PATH.strip()
    if not custom_path:
        if outlook.Folders.Count < 1:
            raise RuntimeError("No se encontraron cuentas de correo configuradas en Outlook.")
        root = outlook.Folders.Item(1)
        print(f"[INFO] Conectado con éxito a la cuenta principal activa: '{root.Name}'")
        return root, None

    # Normalizar ruta
    abs_path = os.path.abspath(custom_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"El archivo de datos especificado en CUSTOM_DATA_FILE_PATH no existe: {abs_path}")

    ext = os.path.splitext(abs_path)[1].lower()

    if ext == '.pst':
        print(f"[INFO] Cargando archivo de datos PST externo de forma dinámica: {abs_path}")
        try:
            # Montar el archivo PST
            outlook.AddStore(abs_path)
        except Exception as e:
            raise RuntimeError(f"Error al montar el archivo PST en Outlook: {e}")

        # Buscar la carpeta raíz del PST montado
        store_root = None
        for store in outlook.Stores:
            try:
                if store.FilePath and os.path.normpath(store.FilePath).lower() == os.path.normpath(abs_path).lower():
                    store_root = store.GetRootFolder()
                    break
            except Exception:
                pass

        if not store_root:
            for folder in outlook.Folders:
                try:
                    if hasattr(folder, 'Store') and folder.Store.FilePath:
                        if os.path.normpath(folder.Store.FilePath).lower() == os.path.normpath(abs_path).lower():
                            store_root = folder
                            break
                except Exception:
                    pass

        if not store_root:
            raise RuntimeError(f"El archivo PST se agregó a la sesión, pero no se pudo encontrar su carpeta raíz para la ruta: {abs_path}")

        print(f"[OK] Conectado exitosamente al archivo PST: '{store_root.Name}'")
        return store_root, store_root

    elif ext == '.ost':
        print("\n" + "="*80)
        print("[ERROR] NO SE PUEDE MONTAR UN ARCHIVO .OST DIRECTAMENTE EN OUTLOOK VIA API")
        print("="*80)
        print("La API de Microsoft Outlook (MAPI) no permite cargar archivos .ost")
        print("adicionales de forma dinámica en una sesión activa (solo admite archivos .pst).")
        print("\n>>> ¿CÓMO HACER EL CAMBIO DE OST A PST? <<<")
        print("--------------------------------------------------------------------------------")
        print("MÉTODO A: EXPORTAR DESDE OUTLOOK (Recomendado si su cuenta aún está activa)")
        print("  1. Abra Outlook normalmente.")
        print("  2. Vaya al menú: Archivo > Abrir y exportar > Importar o exportar.")
        print("  3. Seleccione 'Exportar a un archivo' y haga clic en Siguiente.")
        print("  4. Seleccione 'Archivo de datos de Outlook (.pst)' y haga clic en Siguiente.")
        print("  5. Seleccione la carpeta principal de su cuenta (marque 'Incluir subcarpetas').")
        print("  6. Elija la ruta donde guardar el archivo .pst y haga clic en Finalizar.")
        print("  7. Configure la ruta de este nuevo archivo .pst en su archivo .env:")
        print("     CUSTOM_DATA_FILE_PATH=C:\\Ruta\\A\\Su\\archivo.pst")
        print("\nMÉTODO B: MÉTODO MANUAL DE INTERCAMBIO TEMPORAL (Si el OST es un backup huérfano)")
        print("  1. Cierre Outlook por completo.")
        print("  2. Vaya a C:\\Users\\WPOSS\\AppData\\Local\\Microsoft\\Outlook")
        print("  3. Renombre su archivo .ost activo agregándole '.ACTIVO' al final.")
        print("  4. Copie el archivo .ost de respaldo y péguelo en esa carpeta con el nombre")
        print("     exacto del archivo activo original.")
        print("  5. DESCONECTE el internet de su equipo y abra la aplicación de Outlook de")
        print("     escritorio (se abrirá sin conexión y mostrará el contenido del backup).")
        print("  6. Realice la exportación a .pst siguiendo los pasos del MÉTODO A (puntos 2-6).")
        print("  7. Cierre Outlook, elimine el archivo .ost de respaldo temporal, devuelva el")
        print("     nombre original al activo, reconecte internet y configure la ruta del .pst.")
        print("\nMÉTODO C: CONVERSOR DE TERCEROS")
        print("  Use un software de conversión de OST a PST (como Stellar Converter para OST,")
        print("  Kernel para OST, u otro) para convertir directamente el archivo .ost a .pst.")
        print("="*80 + "\n")
        raise RuntimeError("Carga directa de .ost no soportada por la API de Outlook.")
    else:
        raise ValueError(f"Extensión de archivo no soportada: '{ext}'. Debe ser un archivo .pst.")

def analyze_outlook_storage(report_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root, mounted_store = get_outlook_root_folder(outlook)
    try:
        analyze_outlook_storage_internal(root, report_path)
    finally:
        if mounted_store is not None:
            try:
                outlook.RemoveStore(mounted_store)
                print("[INFO] Archivo PST externo desmontado correctamente de Outlook.")
            except Exception as remove_err:
                print(f"[WARNING] No se pudo desmontar el archivo PST externo de Outlook: {remove_err}")

def analyze_outlook_storage_internal(root, report_path):
    print(f"\n[OK] Conectado exitosamente al archivo de datos en Outlook: '{root.Name}'")
    print("Analizando carpetas y tamaños de adjuntos (esto puede demorar unos minutos)...")
    
    folder_stats = []
    
    def traverse(folder, current_path):
        try:
            items = folder.Items
            count = items.Count
        except Exception:
            return
            
        total_size_bytes = 0
        total_attachments_size_bytes = 0
        attachments_count = 0
        mail_count = 0
        other_count = 0
        
        for idx in range(1, count + 1):
            try:
                item = items.Item(idx)
                size = getattr(item, "Size", 0)
                total_size_bytes += size
                
                try:
                    msg_class = item.Class
                except Exception:
                    msg_class = 0
                    
                if msg_class == 43: # MailItem
                    mail_count += 1
                    try:
                        atts = item.Attachments
                        att_count = atts.Count
                        if att_count > 0:
                            attachments_count += att_count
                            for att_idx in range(1, att_count + 1):
                                try:
                                    att = atts.Item(att_idx)
                                    total_attachments_size_bytes += getattr(att, "Size", 0)
                                except Exception:
                                    pass
                    except Exception:
                        pass
                else:
                    other_count += 1
            except Exception:
                pass
                
        if count > 0 or total_size_bytes > 0:
            folder_stats.append({
                "path": current_path,
                "count": count,
                "mails": mail_count,
                "others": other_count,
                "total_size_mb": total_size_bytes / (1024 * 1024),
                "attachments_size_mb": total_attachments_size_bytes / (1024 * 1024),
                "attachments_count": attachments_count
            })
            
        try:
            subfolders = folder.Folders
            for idx in range(1, subfolders.Count + 1):
                sub = subfolders.Item(idx)
                sub_path = f"{current_path}\\{sub.Name}" if current_path else sub.Name
                traverse(sub, sub_path)
        except Exception:
            pass

    traverse(root, "")
    
    # Sort folders by size descending
    folder_stats.sort(key=lambda x: x["total_size_mb"], reverse=True)
    
    report_lines = []
    report_lines.append("========================================================================================")
    report_lines.append("REPORT DE ANÁLISIS DE TAMAÑO DE CARPETAS Y ADJUNTOS EN EL OST DE BACKUP (7.59 GB)")
    report_lines.append("========================================================================================")
    report_lines.append(f"{'Ruta de la Carpeta':<40} | {'Items':<6} | {'Mails':<6} | {'Tamaño Total':<12} | {'Tam. Adjuntos':<12} | {'Adjs':<5}")
    report_lines.append("-" * 96)
    
    grand_total_size = 0
    grand_total_atts_size = 0
    grand_total_items = 0
    grand_total_mails = 0
    grand_total_atts = 0
    
    for f in folder_stats:
        report_lines.append(f"{f['path'][:40]:<40} | {f['count']:<6} | {f['mails']:<6} | {f['total_size_mb']:10.2f} MB | {f['attachments_size_mb']:10.2f} MB | {f['attachments_count']:<5}")
        grand_total_size += f['total_size_mb']
        grand_total_atts_size += f['attachments_size_mb']
        grand_total_items += f['count']
        grand_total_mails += f['mails']
        grand_total_atts += f['attachments_count']
        
    report_lines.append("-" * 96)
    report_lines.append(f"{'TOTAL EN CARPETAS':<40} | {grand_total_items:<6} | {grand_total_mails:<6} | {grand_total_size:10.2f} MB | {grand_total_atts_size:10.2f} MB | {grand_total_atts:<5}")
    report_lines.append("========================================================================================\n")
    
    try:
        store = root.Store
        if hasattr(store, 'FilePath'):
            fp = store.FilePath
            if os.path.exists(fp):
                file_size_mb = os.path.getsize(fp) / (1024 * 1024)
                report_lines.append(f"Ruta del archivo OST analizado: {fp}")
                report_lines.append(f"Tamaño del archivo en Disco: {file_size_mb:.2f} MB")
                overhead = file_size_mb - grand_total_size
                report_lines.append(f"Espacio libre / Fragmentación (Sobrecarga de Base de Datos): {overhead:.2f} MB ({overhead/file_size_mb*100:.1f}%)")
    except Exception as e:
        report_lines.append(f"No se pudo calcular la fragmentación: {e}")
        
    report_text = "\n".join(report_lines)
    print(report_text)
    
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_text)
    print(f"\n[OK] Reporte completo guardado en: {report_path}")

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    report_path = os.path.join(script_dir, "analisis_backup_7GB.txt")
    
    print("=========================================================")
    print("  ASISTENTE DE ANÁLISIS Y EXTRACCIÓN DE BACKUP OST 7.59 GB")
    print("=========================================================")
    print("Este asistente analizará el archivo OST/cuenta activa en Outlook.")
    print("Asegúrese de que la aplicación Outlook de escritorio esté abierta.")
    print("=========================================================\n")
    
    custom_path = CUSTOM_DATA_FILE_PATH.strip()
    if custom_path:
        abs_path = os.path.abspath(custom_path)
        if not os.path.exists(abs_path):
            print(f"[ERROR] El archivo de datos especificado no existe: {abs_path}")
            sys.exit(1)
        ext = os.path.splitext(abs_path)[1].lower()
        if ext == '.ost':
            try:
                # Mostrar instrucciones de conversión de OST a PST y terminar inmediatamente
                get_outlook_root_folder(None)
            except Exception:
                sys.exit(1)
                
        print("PREPARACIÓN:")
        print("---------------------------------------------------------")
        print("1. Abra la aplicación de Outlook de escritorio.")
        print(f"2. Se analizará el archivo configurado: {custom_path}")
        print("---------------------------------------------------------")
    else:
        print("PREPARACIÓN:")
        print("---------------------------------------------------------")
        print("1. Abra la aplicación de Outlook de escritorio.")
        print("2. Confirme que la cuenta o el archivo OST a analizar esté activo.")
        print("---------------------------------------------------------")
    
    input("\nUna vez que Outlook esté listo, presiona ENTER aquí...")
    
    # Run analysis
    try:
        analyze_outlook_storage(report_path)
        
        # Ask if they want to run the extraction too
        print("\n=========================================================")
        ans = input("¿Desea iniciar la extracción/importación de correos de este Backup a Obsidian ahora? (S/N): ")
        if ans.strip().lower() in ['s', 'y', 'si', 'yes']:
            print("\nIniciando la extracción de correos hacia Obsidian...")
            try:
                import extractor
                extractor.extract_emails_to_obsidian()
                print("[OK] Extracción finalizada.")
            except Exception as ext_err:
                print(f"[ERROR] Error al extraer correos: {ext_err}")
        else:
            print("Extracción omitida. Solo se realizó el análisis de almacenamiento.")
            
    except Exception as e:
        print(f"[ERROR] Falló el análisis o procesamiento de la base de datos: {e}")
    
    print("\n=========================================================")
    input("Presiona ENTER para finalizar el asistente...")
    print("\n[ASISTENTE FINALIZADO]")

if __name__ == "__main__":
    main()
