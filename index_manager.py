import os
import json
import re
from datetime import datetime

INDEX_FILE_NAME = ".extractor_index.json"

def parse_frontmatter(file_content):
    """
    Parsea de manera robusta el YAML Frontmatter de un archivo Markdown de correo
    sin dependencias externas de librerías de YAML.
    """
    metadata = {}
    lines = file_content.splitlines()
    if not lines or lines[0].strip() != "---":
        return metadata
    
    current_key = None
    in_frontmatter = False
    
    for i, line in enumerate(lines):
        if i == 0:
            in_frontmatter = True
            continue
        if line.strip() == "---":
            in_frontmatter = False
            break
        if not in_frontmatter:
            break
            
        # Analizar líneas de listas indentadas
        if line.startswith("  - "):
            if current_key and isinstance(metadata.get(current_key), list):
                val = line[4:].strip()
                # Quitar comillas si existen
                if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                    val = val[1:-1]
                metadata[current_key].append(val)
        else:
            parts = line.split(":", 1)
            if len(parts) == 2:
                key = parts[0].strip()
                val = parts[1].strip()
                if val == "[]":
                    metadata[key] = []
                    current_key = key
                elif not val: # Inicio de una lista
                    metadata[key] = []
                    current_key = key
                else:
                    # Quitar comillas si existen
                    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                        val = val[1:-1]
                    metadata[key] = val
                    current_key = key
    return metadata

def rebuild_index_from_vault(vault_path):
    """
    Escanea la bóveda físicamente y reconstruye la base de datos de indexación.
    Permite auto-recuperación (self-healing) total si el archivo .json es eliminado.
    """
    print("[INFO] Reconstruyendo base de datos de indexación desde archivos Markdown locales...")
    index_data = {
        "metadata": {
            "last_update": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "active_filters": {}
        },
        "emails": {}
    }
    
    if not os.path.exists(vault_path):
        return index_data
        
    for root, dirs, files in os.walk(vault_path):
        # Omitir la carpeta de adjuntos y la carpeta temporal
        parts = root.split(os.sep)
        if "attachments" in parts or "temp" in parts:
            continue
            
        for file in files:
            if file.endswith(".md") and file.lower() not in ["indice.md", "índice.md", "_correos.md"] and not file.startswith("_"):
                filepath = os.path.join(root, file)
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    metadata = parse_frontmatter(content)
                    entry_id = metadata.get("id_entrada")
                    if entry_id:
                        rel_path = os.path.relpath(filepath, vault_path)
                        # Reemplazar diagonales inversas para portabilidad
                        rel_path_portable = rel_path.replace("\\", "/")
                        
                        # Extraer tags sin los prefijos por defecto de correo y remitente
                        tags = metadata.get("tags", [])
                        categories = tags[2:] if len(tags) > 2 else []
                        
                        # Determinar asunto desde el título principal (# Asunto)
                        subject = file[:-3]
                        subj_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
                        if subj_match:
                            subject = subj_match.group(1).strip()
                            
                        index_data["emails"][entry_id] = {
                            "filepath": rel_path_portable,
                            "subject": subject,
                            "sender_name": metadata.get("remitente", "Desconocido"),
                            "sender_email": metadata.get("remitente_correo", ""),
                            "received_time": metadata.get("fecha_recepcion", ""),
                            "sent_time": metadata.get("fecha_envio", ""),
                            "outlook_folder": metadata.get("carpeta_outlook", ""),
                            "categories": categories,
                            "importance": metadata.get("importancia", "Normal"),
                            "attachments": metadata.get("adjuntos", []),
                            "status": "valid"
                        }
                except Exception as e:
                    print(f"[WARNING] Error leyendo frontmatter de {filepath} durante reconstrucción: {e}")
                    
    print(f"[INFO] Reconstrucción completada. Se indexaron {len(index_data['emails'])} correos existentes.")
    return index_data

def load_index(vault_path):
    """
    Carga el índice programático oculto desde la bóveda de salida.
    Si no existe, invoca la reconstrucción desde los archivos locales.
    """
    index_path = os.path.join(vault_path, INDEX_FILE_NAME)
    if os.path.exists(index_path):
        try:
            with open(index_path, 'r', encoding='utf-8') as f:
                index_data = json.load(f)
            # Asegurar la estructura básica
            if "metadata" not in index_data:
                index_data["metadata"] = {}
            if "emails" not in index_data:
                index_data["emails"] = {}
            return index_data
        except Exception as e:
            print(f"[WARNING] No se pudo leer {INDEX_FILE_NAME}: {e}. Se intentará reconstruir.")
            
    # Si no existe o falló la carga, se reconstruye a partir del vault
    index_data = rebuild_index_from_vault(vault_path)
    save_index(vault_path, index_data)
    return index_data

def save_index(vault_path, index_data):
    """
    Guarda el estado actual del índice en el archivo oculto.
    """
    index_path = os.path.join(vault_path, INDEX_FILE_NAME)
    index_data["metadata"]["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(index_path, 'w', encoding='utf-8') as f:
            json.dump(index_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"[ERROR] No se pudo guardar el índice programático en {index_path}: {e}")

def revalidate_filters(vault_path, index_data, folder_filter_manager, email_filter_manager, skip_senders, skip_subjects):
    """
    Revalida dinámicamente cada correo indexado contra los filtros actuales en el .env.
    Elimina físicamente los archivos .md que ya no cumplan las condiciones.
    Mantiene la trazabilidad de los descartados cambiando su estado a "filtered".
    """
    print("[INFO] Iniciando revalidación de filtros sobre información histórica...")
    emails = index_data.get("emails", {})
    deleted_ids = []
    
    for entry_id, email_info in list(emails.items()):
        # Solo revalidar los que estén actualmente activos/válidos
        if email_info.get("status") != "valid":
            continue
            
        filepath = os.path.join(vault_path, email_info["filepath"].replace("/", os.sep))
        
        # 1. Comprobar si el archivo físico fue eliminado manualmente por el usuario
        if not os.path.exists(filepath):
            email_info["status"] = "deleted"
            email_info["skip_reason"] = "Archivo físico no encontrado en disco"
            continue
            
        # 2. Evaluar filtros
        subject = email_info["subject"]
        sender_email = email_info["sender_email"]
        sender_name = email_info["sender_name"]
        outlook_folder = email_info["outlook_folder"]
        
        # Filtros de exclusión (SKIP_SENDERS, SKIP_SUBJECTS)
        subject_lower = (subject or "").lower()
        email_lower = (sender_email or "").lower()
        name_lower = (sender_name or "").lower()
        
        should_skip = False
        skip_reason = ""
        for pattern in skip_subjects:
            if pattern in subject_lower:
                should_skip = True
                skip_reason = f"asunto contiene '{pattern}'"
                break
        
        if not should_skip:
            for pattern in skip_senders:
                if pattern in email_lower or pattern in name_lower:
                    should_skip = True
                    skip_reason = f"remitente coincide con '{pattern}'"
                    break
                    
        # Filtros de carpetas (ALLOWED_FOLDERS, EXCLUDED_FOLDERS)
        folder_name = os.path.basename(outlook_folder) if outlook_folder else ""
        should_proc_folder, folder_reason = folder_filter_manager.should_process_folder(folder_name, outlook_folder)
        
        # Filtro de contenido del cuerpo (FILTER_EMAIL_CONTAINS)
        body = ""
        if email_filter_manager.filter_contains:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                parts = content.split("## Contenido\n", 1)
                if len(parts) == 2:
                    body = parts[1]
                else:
                    body = content
            except Exception as e:
                print(f"[WARNING] No se pudo leer {filepath} para validar contenido: {e}")
                
        email_ok, email_reason = email_filter_manager.should_process_email(sender_email, sender_name, subject, body)
        
        # Determinar motivo de exclusión si aplica
        reason = None
        if should_skip:
            reason = f"Filtro de exclusión ({skip_reason})"
        elif not should_proc_folder:
            reason = f"Filtro de carpetas ({folder_reason})"
        elif not email_ok:
            reason = f"Filtro de correo ({email_reason})"
            
        if reason:
            print(f"[FILTRO REVALIDADO - ELIMINANDO] '{subject}' -> {email_info['filepath']} (Motivo: {reason})")
            # Eliminar archivo físico
            try:
                os.remove(filepath)
            except Exception as e:
                print(f"[ERROR] No se pudo borrar el archivo físico {filepath}: {e}")
                
            # Actualizar estado de consistencia en el índice
            email_info["status"] = "filtered"
            email_info["skip_reason"] = reason
            deleted_ids.append(entry_id)
            
    print(f"[INFO] Revalidación de filtros completada. Se eliminaron {len(deleted_ids)} correos no conformes.")
    return deleted_ids

def cleanup_orphan_attachments(vault_path, index_data, cache_path=None):
    """
    Escanea la carpeta de adjuntos y elimina de forma segura aquellos que no estén
    referenciados por ningún correo válido activo. También actualiza attachments_cache.json.
    """
    print("[INFO] Iniciando limpieza de archivos adjuntos huérfanos...")
    emails = index_data.get("emails", {})
    referenced_attachments = set()
    
    # Recopilar todos los adjuntos en uso por correos válidos
    for entry_id, email_info in emails.items():
        if email_info.get("status") == "valid":
            for att in email_info.get("attachments", []):
                referenced_attachments.add(att)
                
    attachments_dir = os.path.join(vault_path, "attachments")
    if not os.path.exists(attachments_dir):
        print("[INFO] No existe carpeta de adjuntos. Omitiendo limpieza.")
        return
        
    deleted_count = 0
    # Listar archivos y purgar los huérfanos
    for entry in os.scandir(attachments_dir):
        if entry.is_file():
            # Ignorar el archivo de caché y archivos ocultos/del sistema
            if entry.name in ["attachments_cache.json", INDEX_FILE_NAME] or entry.name.startswith("."):
                continue
                
            if entry.name not in referenced_attachments:
                print(f"[LIMPIEZA - HUÉRFANO] Eliminando adjunto no referenciado: {entry.name}")
                try:
                    os.remove(entry.path)
                    deleted_count += 1
                except Exception as e:
                    print(f"[WARNING] No se pudo eliminar el archivo huérfano {entry.path}: {e}")
                    
    print(f"[INFO] Limpieza completada. Se eliminaron {deleted_count} archivos adjuntos huérfanos.")
    
    # Depurar y auto-recuperar attachments_cache.json
    if cache_path and os.path.exists(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
                
            original_len = len(cache_data)
            # Conservar solo los que realmente existen en el disco
            existing_files = set(os.listdir(attachments_dir))
            clean_cache = {h: fname for h, fname in cache_data.items() if fname in existing_files}
            
            if len(clean_cache) != original_len:
                with open(cache_path, 'w', encoding='utf-8') as f:
                    json.dump(clean_cache, f, ensure_ascii=False, indent=4)
                print(f"[INFO] Caché de adjuntos depurado: se eliminaron {original_len - len(clean_cache)} entradas obsoletas.")
        except Exception as e:
            print(f"[WARNING] No se pudo depurar el caché de adjuntos: {e}")

def generate_obsidian_index(vault_path, index_data, active_filters):
    """
    Genera un índice centralizado _Correos.md (MOC - Map of Content) en la raíz del vault de correos
    siguiendo las mejores prácticas modernas de Obsidian, usando enlaces bidireccionales y tablas legibles.
    """
    print("[INFO] Generando índice visual _Correos.md...")
    emails = index_data.get("emails", {})
    valid_emails = {eid: info for eid, info in emails.items() if info.get("status") == "valid"}
    
    total_emails = len(valid_emails)
    
    # Agrupaciones
    folders = set()
    unique_attachments = set()
    emails_by_folder = {}
    emails_by_year = {}
    
    for entry_id, info in valid_emails.items():
        folder = info.get("outlook_folder") or "Raíz/Desconocido"
        folders.add(folder)
        if folder not in emails_by_folder:
            emails_by_folder[folder] = []
        emails_by_folder[folder].append(info)
        
        # Agrupar por año
        rec_time = info.get("received_time")
        year = "Sin Año"
        if rec_time:
            try:
                year = rec_time.split("-")[0]
            except Exception:
                pass
        if year not in emails_by_year:
            emails_by_year[year] = []
        emails_by_year[year].append(info)
        
        # Adjuntos
        for att in info.get("attachments", []):
            unique_attachments.add(att)
            
    # Contenido del Markdown
    lines = []
    lines.append("---")
    lines.append("tags:")
    lines.append("  - categoria/correo")
    lines.append("---")
    lines.append("")
    lines.append("# 🗂️ Índice Central de Correos")
    lines.append("")
    lines.append("Este archivo sirve como el MOC (Map of Content) central para la exploración y trazabilidad de todos los correos extraídos.")
    lines.append("")
    lines.append("## 📊 Resumen Estadístico")
    lines.append(f"- **Total Correos Activos:** {total_emails}")
    lines.append(f"- **Carpetas de Outlook:** {len(folders)}")
    lines.append(f"- **Adjuntos Únicos:** {len(unique_attachments)}")
    lines.append(f"- **Última Actualización:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")
    
    # Filtros Activos
    lines.append("## ⚙️ Filtros Aplicados en este Análisis")
    has_active = False
    for name, val in active_filters.items():
        if val:
            lines.append(f"- **{name}:** `{val}`")
            has_active = True
    if not has_active:
        lines.append("- *No hay filtros activos aplicados.*")
    lines.append("")
    
    # Agrupamiento por Año
    lines.append("## 📅 Histórico por Año de Recepción")
    for year in sorted(emails_by_year.keys(), reverse=True):
        year_emails = emails_by_year[year]
        year_emails.sort(key=lambda x: x.get("received_time", ""), reverse=True)
        lines.append(f"### 📂 {year} ({len(year_emails)} correos)")
        for email in year_emails:
            # Enlace compatible con Obsidian (reemplazando barras invertidas por barras normales)
            rel_path_no_ext, _ = os.path.splitext(email["filepath"])
            link_target = rel_path_no_ext.replace("\\", "/")
            lines.append(f"- {email['received_time']} | [[{link_target}|{email['subject']}]] (De: *{email['sender_name']}*)")
        lines.append("")
        
    # Agrupamiento por Carpeta
    lines.append("## 📁 Estructura por Carpeta de Outlook")
    for folder in sorted(emails_by_folder.keys()):
        folder_emails = emails_by_folder[folder]
        folder_emails.sort(key=lambda x: x.get("received_time", ""), reverse=True)
        lines.append(f"### 📂 {folder} ({len(folder_emails)} correos)")
        for email in folder_emails:
            rel_path_no_ext, _ = os.path.splitext(email["filepath"])
            link_target = rel_path_no_ext.replace("\\", "/")
            lines.append(f"- {email['received_time']} | [[{link_target}|{email['subject']}]]")
        lines.append("")
        
    # Relación bidireccional de adjuntos
    lines.append("## 📎 Índice Relacional de Archivos Adjuntos")
    if unique_attachments:
        lines.append("| Archivo Adjunto | Correos de Procedencia |")
        lines.append("| :--- | :--- |")
        for att in sorted(list(unique_attachments)):
            # Buscar correos válidos que referencian este adjunto
            ref_links = []
            for entry_id, info in valid_emails.items():
                if att in info.get("attachments", []):
                    rel_path_no_ext, _ = os.path.splitext(info["filepath"])
                    link_target = rel_path_no_ext.replace("\\", "/")
                    ref_links.append(f"[[{link_target}|{info['subject']}]]")
            
            # Enlace al adjunto (escapando la barra vertical para que no rompa la tabla de Markdown)
            att_link = f"[[attachments/{att}\\|{att}]]"
            lines.append(f"| {att_link} | {', '.join(ref_links)} |")
    else:
        lines.append("*No se detectaron archivos adjuntos en los correos válidos.*")
        
    lines.append("")
    
    # Escribir el archivo
    index_md_path = os.path.join(vault_path, "_Correos.md")
    try:
        with open(index_md_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))
        print(f"[INFO] Índice visual de Obsidian (_Correos.md) regenerado correctamente.")
    except Exception as e:
        print(f"[ERROR] No se pudo escribir el archivo _Correos.md: {e}")

    # Eliminar archivo antiguo Indice.md si existe para evitar duplicación
    old_index_path = os.path.join(vault_path, "Indice.md")
    if os.path.exists(old_index_path):
        try:
            os.remove(old_index_path)
            print(f"[INFO] Archivo obsoleto {old_index_path} eliminado con éxito.")
        except Exception as e:
            print(f"[WARNING] No se pudo eliminar el archivo obsoleto {old_index_path}: {e}")
