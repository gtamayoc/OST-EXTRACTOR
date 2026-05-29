import win32com.client
import os
import re
import json
import hashlib
import uuid
import shutil
from datetime import datetime
from dotenv import load_dotenv
import builtins
import sys

# Redefinir print para evitar caídas por UnicodeEncodeError en consolas Windows
_original_print = builtins.print

def safe_print(*args, **kwargs):
    try:
        _original_print(*args, **kwargs)
    except UnicodeEncodeError:
        file = kwargs.get('file', sys.stdout)
        sep = kwargs.get('sep', ' ')
        end = kwargs.get('end', '\n')
        message = sep.join(str(arg) for arg in args)
        encoding = getattr(file, 'encoding', None) or 'utf-8'
        try:
            clean_message = message.encode(encoding, errors='replace').decode(encoding)
            _original_print(clean_message, **kwargs)
        except Exception:
            try:
                ascii_message = message.encode('ascii', errors='backslashreplace').decode('ascii')
                _original_print(ascii_message, **kwargs)
            except Exception:
                pass

builtins.print = safe_print

# Caché global de adjuntos para evitar descargas duplicadas
ATTACHMENT_CACHE = {}
ATTACHMENT_CACHE_INITIALIZED = False
CACHE_FILE_NAME = "attachments_cache.json"

def get_file_hash(filepath):
    """Calcula el hash SHA-256 de un archivo."""
    hasher = hashlib.sha256()
    try:
        with open(filepath, 'rb') as f:
            while chunk := f.read(8192):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception as e:
        print(f"Error calculando hash para {filepath}: {e}")
        return None

def load_attachments_cache(output_base_path):
    """Carga o inicializa el caché de adjuntos a partir de los archivos existentes y del archivo JSON."""
    global ATTACHMENT_CACHE, ATTACHMENT_CACHE_INITIALIZED
    if ATTACHMENT_CACHE_INITIALIZED:
        return
    
    # Determinar la ruta del archivo JSON en la carpeta del script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_path = os.path.join(script_dir, CACHE_FILE_NAME)
    
    # Intentar cargar desde el archivo JSON si existe
    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                ATTACHMENT_CACHE = json.load(f)
            ATTACHMENT_CACHE_INITIALIZED = True
            print(f"[INFO] Caché de adjuntos cargado ({len(ATTACHMENT_CACHE)} elementos indexados).")
            return
        except Exception as e:
            print(f"[WARNING] No se pudo leer el archivo de caché {cache_path}: {e}. Se recreará.")
    
    # Si no existe o falló la lectura, indexar los archivos físicos existentes en la carpeta attachments
    attachments_dir = os.path.join(output_base_path, "attachments")
    if os.path.exists(attachments_dir):
        print("[INFO] Indexando adjuntos existentes en la carpeta física para inicializar caché...")
        for entry in os.scandir(attachments_dir):
            # Ignorar carpetas (como temp) y archivos que no sean adjuntos reales
            if entry.is_file() and entry.name != CACHE_FILE_NAME:
                h = get_file_hash(entry.path)
                if h:
                    ATTACHMENT_CACHE[h] = entry.name
    
    ATTACHMENT_CACHE_INITIALIZED = True
    print(f"[INFO] Inicialización de caché de adjuntos completada con {len(ATTACHMENT_CACHE)} elementos.")
    # Guardar el estado inicial
    save_attachments_cache()

def save_attachments_cache():
    """Guarda el estado actual del caché de adjuntos en el archivo JSON."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_path = os.path.join(script_dir, CACHE_FILE_NAME)
    try:
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(ATTACHMENT_CACHE, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"[ERROR] No se pudo guardar el archivo de caché {cache_path}: {e}")

# Cargar variables de entorno del archivo .env local
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
else:
    load_dotenv()

class EnvConfigLoader:
    @staticmethod
    def get_list(var_name):
        val = os.getenv(var_name, "")
        if not val:
            return []
        return [item.strip() for item in val.split(",") if item.strip()]

    @staticmethod
    def get_string(var_name):
        val = os.getenv(var_name, "")
        return val.strip() if val else None

# Configuración de salida para tu bóveda de Obsidian
OBSIDIAN_VAULT_PATH = os.getenv("OBSIDIAN_VAULT_PATH", os.path.dirname(__file__))
LIMIT_EMAILS = os.getenv("LIMIT_EMAILS", "true").lower() in ("true", "1", "yes")
MAX_EMAILS = int(os.getenv("MAX_EMAILS", "10"))
ORGANIZE_BY_YEAR = os.getenv("ORGANIZE_BY_YEAR", "true").lower() in ("true", "1", "yes")

# Filtros de exclusión: si la variable está vacía la lista queda vacía y no filtra nada
_skip_senders_raw = os.getenv("SKIP_SENDERS", "")
_skip_subjects_raw = os.getenv("SKIP_SUBJECTS", "")
SKIP_SENDERS = [s.strip().lower() for s in _skip_senders_raw.split(",") if s.strip()]
SKIP_SUBJECTS = [s.strip().lower() for s in _skip_subjects_raw.split(",") if s.strip()]

# Nuevos Filtros configurables de Carpetas y Correos
ALLOWED_FOLDERS = EnvConfigLoader.get_list("ALLOWED_FOLDERS")
EXCLUDED_FOLDERS = EnvConfigLoader.get_list("EXCLUDED_FOLDERS")
FILTER_EMAIL_FROM = EnvConfigLoader.get_string("FILTER_EMAIL_FROM")
FILTER_EMAIL_SUBJECT = EnvConfigLoader.get_string("FILTER_EMAIL_SUBJECT")
FILTER_EMAIL_CONTAINS = EnvConfigLoader.get_string("FILTER_EMAIL_CONTAINS")

# Mostrar logs de configuración detectada
if ALLOWED_FOLDERS:
    print(f"[CONFIG] Carpetas permitidas detectadas (ALLOWED_FOLDERS): {', '.join(ALLOWED_FOLDERS)}")
if EXCLUDED_FOLDERS:
    print(f"[CONFIG] Carpetas excluidas detectadas (EXCLUDED_FOLDERS): {', '.join(EXCLUDED_FOLDERS)}")
if FILTER_EMAIL_FROM:
    print(f"[CONFIG] Filtro de remitente (FILTER_EMAIL_FROM): '{FILTER_EMAIL_FROM}'")
if FILTER_EMAIL_SUBJECT:
    print(f"[CONFIG] Filtro de asunto (FILTER_EMAIL_SUBJECT): '{FILTER_EMAIL_SUBJECT}'")
if FILTER_EMAIL_CONTAINS:
    print(f"[CONFIG] Filtro de contenido del cuerpo (FILTER_EMAIL_CONTAINS): '{FILTER_EMAIL_CONTAINS}'")

class FolderFilterManager:
    def __init__(self, allowed, excluded):
        self.allowed = allowed
        self.excluded = excluded
        
        # Priority rule: If allowed folders are set, ignore excluded folders.
        if self.allowed:
            self.active_allowed = True
            self.active_excluded = False
        else:
            self.active_allowed = False
            self.active_excluded = bool(self.excluded)

    @staticmethod
    def _normalize_path(path):
        if not path:
            return ""
        path = path.replace("/", "\\")
        parts = [p.strip() for p in path.split("\\") if p.strip()]
        return "\\".join(parts).lower()

    def _matches_filter(self, folder_name, relative_path, filter_item):
        norm_filter = self._normalize_path(filter_item)
        norm_rel = self._normalize_path(relative_path)
        norm_name = folder_name.strip().lower()
        
        if norm_rel == norm_filter:
            return True
        if norm_name == norm_filter:
            return True
        if norm_rel.startswith(norm_filter + "\\"):
            return True
        return False

    def _is_parent_of_any_allowed(self, folder_name, relative_path):
        norm_rel = self._normalize_path(relative_path)
        if not norm_rel:
            return True
            
        for allowed_item in self.allowed:
            norm_allowed = self._normalize_path(allowed_item)
            if norm_allowed.startswith(norm_rel + "\\"):
                return True
        return False

    def should_process_folder(self, folder_name, relative_path):
        if not self.active_allowed and not self.active_excluded:
            return True, None
            
        if self.active_allowed:
            for allowed_item in self.allowed:
                if self._matches_filter(folder_name, relative_path, allowed_item):
                    return True, None
            return False, "no está en la lista de permitidas"
            
        if self.active_excluded:
            for excluded_item in self.excluded:
                if self._matches_filter(folder_name, relative_path, excluded_item):
                    return False, "excluida por filtro"
            return True, None

    def should_traverse_folder(self, folder_name, relative_path):
        if not self.active_allowed and not self.active_excluded:
            return True
            
        if self.active_allowed:
            for allowed_item in self.allowed:
                if self._matches_filter(folder_name, relative_path, allowed_item):
                    return True
            if self._is_parent_of_any_allowed(folder_name, relative_path):
                return True
            return False
            
        if self.active_excluded:
            for excluded_item in self.excluded:
                if self._matches_filter(folder_name, relative_path, excluded_item):
                    return False
            return True

    def get_exclude_reason(self, folder_name, relative_path):
        should_proc, reason = self.should_process_folder(folder_name, relative_path)
        if not should_proc:
            return reason
        if not self.should_traverse_folder(folder_name, relative_path):
            if self.active_excluded:
                return "excluida por filtro"
            return "no está en la lista de permitidas"
        return "desconocido"

class EmailFilterManager:
    def __init__(self, filter_from, filter_subject, filter_contains):
        self.filter_from = filter_from.strip().lower() if filter_from else None
        self.filter_subject = filter_subject.strip().lower() if filter_subject else None
        self.filter_contains = filter_contains.strip().lower() if filter_contains else None
        self.has_filters = bool(self.filter_from or self.filter_subject or self.filter_contains)

    def should_process_email(self, sender_email, sender_name, subject, body):
        if not self.has_filters:
            return True, None

        if self.filter_from:
            email_lower = (sender_email or "").lower()
            name_lower = (sender_name or "").lower()
            if self.filter_from not in email_lower and self.filter_from not in name_lower:
                return False, f"remitente no coincide con FILTER_EMAIL_FROM ('{self.filter_from}')"

        if self.filter_subject:
            subject_lower = (subject or "").lower()
            if self.filter_subject not in subject_lower:
                return False, f"asunto no contiene FILTER_EMAIL_SUBJECT ('{self.filter_subject}')"

        if self.filter_contains:
            body_lower = (body or "").lower()
            if self.filter_contains not in body_lower:
                return False, f"cuerpo no contiene FILTER_EMAIL_CONTAINS ('{self.filter_contains}')"

        return True, None

folder_filter_manager = FolderFilterManager(ALLOWED_FOLDERS, EXCLUDED_FOLDERS)
email_filter_manager = EmailFilterManager(FILTER_EMAIL_FROM, FILTER_EMAIL_SUBJECT, FILTER_EMAIL_CONTAINS)

def clean_filename(title):
    """Limpia el asunto para que sea un nombre de archivo válido en Windows."""
    if not title:
        return "Sin_Asunto"
    # Eliminar caracteres prohibidos en nombres de archivo de Windows o rutas
    clean = re.sub(r'[\\/*?:"<>|]', "", str(title))
    # Quitar saltos de línea y espacios múltiples
    clean = " ".join(clean.split())
    # Limitar longitud para evitar rutas demasiado largas
    return clean[:100].strip()

def clean_tag(name):
    """Limpia un nombre para que sea un tag válido en Obsidian (sin caracteres especiales ni espacios)."""
    if not name:
        return "desconocido"
    # Eliminar acentos y caracteres especiales
    clean = re.sub(r'[^a-zA-Z0-9\s_]', '', str(name))
    # Convertir espacios a guiones bajos y a minúsculas
    clean = re.sub(r'\s+', '_', clean).lower()
    return clean.strip('_')

def should_skip_email(subject, sender_email, sender_name):
    """
    Determina si un correo debe omitirse según los filtros configurados en el .env.
    Retorna (True, motivo) si debe omitirse, o (False, None) si se debe procesar.
    Si SKIP_SUBJECTS y SKIP_SENDERS están vacíos, esta función nunca omite nada.
    """
    subject_lower = (subject or "").lower()
    email_lower = (sender_email or "").lower()
    name_lower = (sender_name or "").lower()

    for pattern in SKIP_SUBJECTS:
        if pattern in subject_lower:
            return True, f"asunto contiene '{pattern}'"

    for pattern in SKIP_SENDERS:
        if pattern in email_lower or pattern in name_lower:
            return True, f"remitente coincide con '{pattern}'"

    return False, None

def format_attachment_link(filename):
    """Genera un enlace compatible con Obsidian para los adjuntos."""
    ext = os.path.splitext(filename)[1].lower()
    # Si es imagen, se incrusta con !
    if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp']:
        return f"![[attachments/{filename}]]"
    else:
        return f"[[attachments/{filename}|{filename}]]"

def get_received_time(message):
    """Obtiene la fecha de recepción de un mensaje de forma segura."""
    try:
        return message.ReceivedTime
    except Exception:
        try:
            return message.SentOn
        except Exception:
            return None

def process_and_save_email(message, relative_folder_path, output_base_path):
    """Extrae la información de un mensaje de correo y lo guarda en formato Markdown."""
    try:
        # Extraer campos principales
        subject = getattr(message, "Subject", "Sin Asunto")
        sender_name = getattr(message, "SenderName", "Desconocido")
        
        try:
            sender_email = getattr(message, "SenderEmailAddress", "")
        except Exception:
            sender_email = ""

        # ── Verificar filtros de exclusión antes de continuar ────────────────
        skip, skip_reason = should_skip_email(subject, sender_email, sender_name)
        if skip:
            print(f"[OMITIDO] '{subject}' ({sender_name} <{sender_email}>) — {skip_reason}")
            return False, "__SKIPPED__"

        try:
            to_recipients = getattr(message, "To", "")
        except Exception:
            to_recipients = ""

        try:
            cc_recipients = getattr(message, "CC", "")
        except Exception:
            cc_recipients = ""

        # Obtener importancia
        importance_val = getattr(message, "Importance", 1)
        importance_map = {0: "Baja", 1: "Normal", 2: "Alta"}
        importance_str = importance_map.get(importance_val, "Normal")

        # Obtener categorías
        categories_raw = getattr(message, "Categories", "")
        categories = [c.strip() for c in categories_raw.split(",")] if categories_raw else []

        # Obtener IDs
        conversation_id = getattr(message, "ConversationID", "")
        entry_id = getattr(message, "EntryID", "")

        # Fechas
        received_time = get_received_time(message)
        if received_time:
            try:
                received_str = received_time.strftime("%Y-%m-%d %H:%M:%S")
                received_date_only = received_time.strftime("%Y-%m-%d")
            except Exception:
                received_str = str(received_time)
                received_date_only = datetime.now().strftime("%Y-%m-%d")
        else:
            received_str = ""
            received_date_only = datetime.now().strftime("%Y-%m-%d")

        try:
            sent_time = getattr(message, "SentOn", None)
            sent_str = sent_time.strftime("%Y-%m-%d %H:%M:%S") if sent_time else ""
        except Exception:
            sent_str = ""

        # Cuerpo del correo
        body = getattr(message, "Body", "")

        # Verificar nuevos filtros positivos de correo
        email_ok, email_reason = email_filter_manager.should_process_email(sender_email, sender_name, subject, body)
        if not email_ok:
            print(f"[OMITIDO - FILTRO] '{subject}' ({sender_name} <{sender_email}>) — {email_reason}")
            return False, "__SKIPPED_EMAIL_FILTER__"

        # Crear carpeta de destino: plano por año si ORGANIZE_BY_YEAR está activo.
        # La carpeta de Outlook (relative_folder_path) se ignora intencionalmente para
        # mantener coherencia entre años donde no se aplicó la misma estructura.
        if ORGANIZE_BY_YEAR:
            year_str = received_time.strftime("%Y") if received_time else datetime.now().strftime("%Y")
            folder_output_dir = os.path.join(output_base_path, year_str)
        else:
            folder_output_dir = output_base_path
        
        if not os.path.exists(folder_output_dir):
            os.makedirs(folder_output_dir)

        # Generar nombre del archivo y verificar si ya existe (evita procesar adjuntos en duplicados)
        safe_subject = clean_filename(subject)
        filename = f"{received_date_only}_{safe_subject}.md"
        filepath = os.path.join(folder_output_dir, filename)

        if os.path.exists(filepath):
            print(f"[OMITIDO - DUPLICADO] '{subject}' ya existe en el destino ({filename}).")
            return False, "__SKIPPED_DUPLICATE__"

        # Procesar Adjuntos (siempre se guardan en la carpeta central 'attachments' en el root de la salida)
        attachments_list = []
        attachments_links_md = []
        
        try:
            attachments = message.Attachments
            attachments_count = attachments.Count
        except Exception:
            attachments_count = 0

        if attachments_count > 0:
            attachments_dir = os.path.join(output_base_path, "attachments")
            if not os.path.exists(attachments_dir):
                os.makedirs(attachments_dir)
            
            # Asegurar que el caché esté inicializado para este vault
            load_attachments_cache(output_base_path)
            
            # Crear directorio temporal local para descargar los adjuntos inicialmente y calcular su hash
            temp_dir = os.path.join(attachments_dir, "temp")
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
            
            for att_idx in range(1, attachments_count + 1):
                try:
                    attachment = attachments.Item(att_idx)
                    att_name = attachment.FileName
                    if att_name:
                        # Generar un nombre de archivo temporal único para evitar bloqueos por procesos concurrentes en Windows
                        temp_filename = f"temp_{uuid.uuid4().hex}"
                        temp_path = os.path.join(temp_dir, temp_filename)
                        
                        try:
                            # Guardar el adjunto temporalmente
                            attachment.SaveAsFile(temp_path)
                            
                            # Calcular su hash SHA-256
                            att_hash = get_file_hash(temp_path)
                            
                            if not att_hash:
                                raise RuntimeError("No se pudo calcular el hash del archivo temporal.")
                            
                            # Si el hash ya está registrado, reutilizamos ese archivo y borramos el temporal
                            if att_hash in ATTACHMENT_CACHE:
                                clean_att_name = ATTACHMENT_CACHE[att_hash]
                                if os.path.exists(temp_path):
                                    os.remove(temp_path)
                            else:
                                # Si no está registrado, buscar un nombre limpio y resolver colisiones de nombres físicos
                                base, ext = os.path.splitext(att_name)
                                clean_base = clean_filename(base)
                                clean_att_name = clean_base + ext.lower()
                                
                                final_path = os.path.join(attachments_dir, clean_att_name)
                                
                                # Si colisiona con un archivo existente en el disco, buscamos un nombre único agregando un número
                                counter = 1
                                while os.path.exists(final_path):
                                    clean_att_name = f"{clean_base}_{counter}{ext.lower()}"
                                    final_path = os.path.join(attachments_dir, clean_att_name)
                                    counter += 1
                                
                                # Mover el archivo temporal al destino definitivo
                                shutil.move(temp_path, final_path)
                                
                                # Registrar en caché
                                ATTACHMENT_CACHE[att_hash] = clean_att_name
                            
                            if clean_att_name not in attachments_list:
                                attachments_list.append(clean_att_name)
                                attachments_links_md.append(format_attachment_link(clean_att_name))
                            
                        except Exception as att_file_err:
                            # Eliminar archivo temporal si existiera en caso de error
                            if os.path.exists(temp_path):
                                try:
                                    os.remove(temp_path)
                                except Exception:
                                    pass
                            raise att_file_err
                except Exception as att_err:
                    print(f"Error procesando adjunto #{att_idx} de '{subject}': {att_err}")

        # Formatear datos JSON para YAML seguro (evita errores de parseo en Obsidian)
        sender_tag = clean_tag(sender_name)
        
        # Construir el contenido Markdown con YAML Frontmatter
        yaml_data = {
            "aliases": [],
            "tags": ["correo", sender_tag] + [clean_tag(cat) for cat in categories],
            "fecha_recepcion": received_str,
            "fecha_envio": sent_str,
            "remitente": sender_name,
            "remitente_correo": sender_email,
            "destinatarios": to_recipients,
            "copia": cc_recipients,
            "importancia": importance_str,
            "carpeta_outlook": relative_folder_path,
            "id_conversacion": conversation_id,
            "id_entrada": entry_id,
            "adjuntos": attachments_list
        }

        # Generar bloques frontmatter
        yaml_lines = ["---"]
        for key, val in yaml_data.items():
            if isinstance(val, list):
                if not val:
                    yaml_lines.append(f"{key}: []")
                else:
                    yaml_lines.append(f"{key}:")
                    for item in val:
                        yaml_lines.append(f"  - {json.dumps(str(item), ensure_ascii=False)}")
            else:
                yaml_lines.append(f"{key}: {json.dumps(str(val), ensure_ascii=False)}")
        yaml_lines.append("---")
        frontmatter = "\n".join(yaml_lines)

        # Construir cuerpo del Markdown
        md_body = []
        md_body.append(f"# {subject}")
        md_body.append("")
        md_body.append(f"**De:** {sender_name} `<{sender_email}>`" if sender_email else f"**De:** {sender_name}")
        md_body.append(f"**Para:** {to_recipients}" if to_recipients else "")
        if cc_recipients:
            md_body.append(f"**CC:** {cc_recipients}")
        md_body.append(f"**Fecha:** {received_str}")
        md_body.append(f"**Carpeta Outlook:** {relative_folder_path}")
        md_body.append(f"**Importancia:** {importance_str}")
        
        if categories:
            md_body.append(f"**Categorías:** {', '.join(categories)}")
        
        md_body.append("\n---")
        
        # Agregar enlaces de adjuntos si existen
        if attachments_links_md:
            md_body.append("## Archivos Adjuntos")
            for link in attachments_links_md:
                md_body.append(f"- {link}")
            md_body.append("\n---")

        md_body.append("## Contenido\n")
        md_body.append(body if body else "*El correo no tiene contenido de texto.*")
        
        md_content = frontmatter + "\n\n" + "\n".join(md_body)

        # Escribir el archivo final
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(md_content)

        return True, filename
    except Exception as e:
        print(f"Error procesando un correo específico: {e}")
        return False, None

def collect_emails_recursively(folder, current_rel_path, candidate_list, limit_active, max_per_folder):
    """
    Recorre de forma recursiva las carpetas de Outlook y recolecta referencias a los correos.
    Si limit_active es True, se ordenan los correos de cada carpeta y se toman los más recientes
    para evitar saturación de memoria.
    """
    # Determinar si debemos procesar los correos de esta carpeta
    should_proc, _ = folder_filter_manager.should_process_folder(folder.Name, current_rel_path)
    if should_proc:
        try:
            items = folder.Items
            # Intentar verificar si hay elementos
            count = items.Count
            if count > 0:
                # Intentar ordenar para procesar los más recientes primero
                try:
                    items.Sort("[ReceivedTime]", True)
                except Exception:
                    pass

                local_added = 0
                for idx in range(1, count + 1):
                    try:
                        item = items.Item(idx)
                    except Exception:
                        continue

                    # Filtrar solo MailItem (Class 43)
                    try:
                        msg_class = item.Class
                    except Exception:
                        msg_class = 0

                    if msg_class == 43:
                        rec_time = get_received_time(item) or datetime.min
                        candidate_list.append({
                            "item": item,
                            "folder_path": current_rel_path,
                            "time": rec_time
                        })
                        local_added += 1
                        
                        # Si el límite está activo, no recolectamos más de lo necesario de una sola carpeta
                        if limit_active and local_added >= max_per_folder:
                            break
        except Exception as e:
            print(f"No se pudo acceder a los correos de la carpeta '{current_rel_path}': {e}")

    # Procesar subcarpetas recursivamente
    try:
        subfolders = folder.Folders
        for idx in range(1, subfolders.Count + 1):
            try:
                sub = subfolders.Item(idx)
                sub_rel_path = os.path.join(current_rel_path, sub.Name) if current_rel_path else sub.Name
                
                if folder_filter_manager.should_traverse_folder(sub.Name, sub_rel_path):
                    collect_emails_recursively(sub, sub_rel_path, candidate_list, limit_active, max_per_folder)
                else:
                    reason = folder_filter_manager.get_exclude_reason(sub.Name, sub_rel_path)
                    print(f"[OMITIDA - CARPETA] '{sub_rel_path}' (Motivo: {reason})")
            except Exception:
                continue
    except Exception:
        pass

def process_all_emails_recursively(folder, current_rel_path, output_base_path, stats):
    """
    Recorre recursivamente las carpetas y procesa/guarda cada correo directamente (Streaming).
    Ideal para migración completa ya que no mantiene referencias masivas en memoria.
    """
    should_proc, _ = folder_filter_manager.should_process_folder(folder.Name, current_rel_path)
    if should_proc:
        try:
            items = folder.Items
            count = items.Count
            if count > 0:
                for idx in range(1, count + 1):
                    try:
                        item = items.Item(idx)
                    except Exception:
                        continue

                    try:
                        msg_class = item.Class
                    except Exception:
                        msg_class = 0

                    if msg_class != 43:
                        stats["skipped_class"] += 1
                        continue

                    success, filename = process_and_save_email(item, current_rel_path, output_base_path)
                    if success:
                        stats["success"] += 1
                        print(f"[{stats['success']}] Extraído: '{getattr(item, 'Subject', 'Sin Asunto')}' -> {filename}")
                    elif filename == "__SKIPPED__":
                        stats["skipped_filter"] += 1
                    elif filename == "__SKIPPED_EMAIL_FILTER__":
                        stats["skipped_email_filter"] += 1
                    elif filename == "__SKIPPED_DUPLICATE__":
                        stats["skipped_duplicate"] += 1
                    else:
                        stats["error"] += 1
        except Exception as e:
            print(f"Error procesando correos en carpeta '{current_rel_path}': {e}")

    # Procesar subcarpetas
    try:
        subfolders = folder.Folders
        for idx in range(1, subfolders.Count + 1):
            try:
                sub = subfolders.Item(idx)
                sub_rel_path = os.path.join(current_rel_path, sub.Name) if current_rel_path else sub.Name
                
                if folder_filter_manager.should_traverse_folder(sub.Name, sub_rel_path):
                    process_all_emails_recursively(sub, sub_rel_path, output_base_path, stats)
                else:
                    reason = folder_filter_manager.get_exclude_reason(sub.Name, sub_rel_path)
                    print(f"[OMITIDA - CARPETA] '{sub_rel_path}' (Motivo: {reason})")
            except Exception:
                continue
    except Exception:
        pass

def extract_emails_to_obsidian():
    # Asegurar que exista la carpeta de salida
    if not os.path.exists(OBSIDIAN_VAULT_PATH):
        os.makedirs(OBSIDIAN_VAULT_PATH)

    print("=======================================================")
    print(f"Ruta de salida en Obsidian: {OBSIDIAN_VAULT_PATH}")
    print("Conectando a Microsoft Outlook...")
    print("Nota: La conexión no requiere saber la ruta física del archivo .ost/.pst,")
    print("ya que Outlook gestiona la base de datos automáticamente en segundo plano.")
    print("=======================================================")
    
    try:
        # Inicializar cliente COM de Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Verificar si hay carpetas/cuentas disponibles
        if outlook.Folders.Count < 1:
            print("[ERROR] No se encontraron cuentas de correo ni archivos de datos configurados en Outlook.")
            print("Por favor, asegúrate de tener Outlook instalado y configurado con tu cuenta.")
            return

        # Seleccionar la primera cuenta/almacén de datos disponible (ID 1)
        root_folder = outlook.Folders.Item(1)
        print(f"[INFO] Conectado con éxito a la cuenta principal: '{root_folder.Name}'")
        
        # Cargar/inicializar el caché de adjuntos
        load_attachments_cache(OBSIDIAN_VAULT_PATH)
        
        if LIMIT_EMAILS:
            print(f"[INFO] Modo Limitado ACTIVO: exportando los {MAX_EMAILS} correos más recientes aceptados por los filtros...")
            candidates = []
            
            # Recolectar un buffer amplio de candidatos para compensar los que sean filtrados.
            # Se usa un múltiplo de MAX_EMAILS como margen; si los filtros son muy agresivos
            # y no hay suficientes candidatos, se exportarán los que estén disponibles.
            buffer_multiplier = 10
            collect_emails_recursively(root_folder, "", candidates, limit_active=True, max_per_folder=MAX_EMAILS * buffer_multiplier)
            
            # Ordenar todos los candidatos por fecha descendente (más recientes primero)
            candidates.sort(key=lambda x: x["time"], reverse=True)
            
            print(f"[INFO] Se encontraron {len(candidates)} correos candidatos. Procesando hasta obtener {MAX_EMAILS} exportados...")
            
            success_count = 0
            skipped_filter_count = 0
            skipped_email_filter_count = 0
            skipped_duplicate_count = 0
            error_count = 0

            for candidate in candidates:
                # Parar en cuanto se alcance el objetivo de exportaciones
                if success_count >= MAX_EMAILS:
                    break

                success, filename = process_and_save_email(candidate["item"], candidate["folder_path"], OBSIDIAN_VAULT_PATH)
                if success:
                    success_count += 1
                    print(f"[{success_count}/{MAX_EMAILS}] Extraído: '{getattr(candidate['item'], 'Subject', 'Sin Asunto')}' -> {filename}")
                elif filename == "__SKIPPED__":
                    skipped_filter_count += 1
                elif filename == "__SKIPPED_EMAIL_FILTER__":
                    skipped_email_filter_count += 1
                elif filename == "__SKIPPED_DUPLICATE__":
                    skipped_duplicate_count += 1
                else:
                    error_count += 1

            print("\n=== RESUMEN DE EJECUCIÓN (MODO LIMITADO) ===")
            print(f"Correos exportados correctamente:  {success_count}")
            print(f"Correos omitidos por exclusión:    {skipped_filter_count}")
            print(f"Correos omitidos por filtro:       {skipped_email_filter_count}")
            print(f"Correos omitidos por duplicado:    {skipped_duplicate_count}")
            print(f"Errores en procesamiento:          {error_count}")
            if success_count < MAX_EMAILS:
                print(f"[AVISO] Solo se encontraron {success_count} correos válidos (se solicitaron {MAX_EMAILS}).")
            print("=============================================")

            
        else:
            print("[INFO] Modo Migración Completa ACTIVO: Procesando TODOS los correos de todas las carpetas...")
            stats = {
                "success": 0,
                "skipped_class": 0,
                "skipped_filter": 0,
                "skipped_email_filter": 0,
                "skipped_duplicate": 0,
                "error": 0
            }
            
            # Procesar en flujo recursivo directo
            process_all_emails_recursively(root_folder, "", OBSIDIAN_VAULT_PATH, stats)
            
            print("\n=== RESUMEN DE EJECUCIÓN (MIGRACIÓN COMPLETA) ===")
            print(f"Correos exportados correctamente:  {stats['success']}")
            print(f"Elementos omitidos (no correos):   {stats['skipped_class']}")
            print(f"Correos omitidos por exclusión:    {stats['skipped_filter']}")
            print(f"Correos omitidos por filtro:       {stats['skipped_email_filter']}")
            print(f"Correos omitidos por duplicado:    {stats['skipped_duplicate']}")
            print(f"Errores en procesamiento:          {stats['error']}")
            print("==================================================")

    except Exception as e:
        print(f"[ERROR] Error crítico durante la extracción: {e}")
        print("Asegúrate de que la aplicación Outlook de escritorio esté abierta y configurada correctamente.")
    finally:
        # Asegurar que se guarde el caché de adjuntos
        save_attachments_cache()
        # Limpiar la carpeta temporal de descargas de adjuntos
        temp_dir = os.path.join(OBSIDIAN_VAULT_PATH, "attachments", "temp")
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                print("[INFO] Carpeta temporal de descargas de adjuntos limpia con éxito.")
            except Exception as clean_err:
                print(f"[WARNING] No se pudo eliminar la carpeta temporal {temp_dir}: {clean_err}")

if __name__ == "__main__":
    extract_emails_to_obsidian()
