import win32com.client
import os
import sys

def analyze_outlook_storage(report_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root = outlook.Folders.Item(1)
    
    print(f"\n[OK] Conectado exitosamente al archivo OST temporal en Outlook: '{root.Name}'")
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
