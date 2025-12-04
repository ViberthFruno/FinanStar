# Archivo: email_manager.py
# Ubicación: raíz del proyecto
# Descripción: Gestiona las operaciones de correo electrónico (SMTP e IMAP) con sistema modular de casos

import smtplib
import imaplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header, decode_header
import email
from datetime import datetime, date
import email.utils
from typing import Any, Dict, Optional, List, Set
from case_handler import CaseHandler


class EmailManager:
    def __init__(self):
        """Inicializa el gestor de correo electrónico"""
        # Definir configuraciones predeterminadas para proveedores comunes
        self.provider_configs = {
            'Gmail': {
                'smtp_server': 'smtp.gmail.com',
                'smtp_port': 587,
                'imap_server': 'imap.gmail.com',
                'imap_port': 993
            },
            'Outlook': {
                'smtp_server': 'smtp-mail.outlook.com',
                'smtp_port': 587,
                'imap_server': 'outlook.office365.com',
                'imap_port': 993
            },
            'Yahoo': {
                'smtp_server': 'smtp.mail.yahoo.com',
                'smtp_port': 587,
                'imap_server': 'imap.mail.yahoo.com',
                'imap_port': 993
            },
            'Otro': {
                'smtp_server': '',
                'smtp_port': 587,
                'imap_server': '',
                'imap_port': 993
            }
        }

        # Inicializar el manejador de casos
        self.case_handler = CaseHandler()

    def get_provider_config(self, provider):
        """Obtiene la configuración para un proveedor específico"""
        return self.provider_configs.get(provider, self.provider_configs['Otro'])

    def test_smtp_connection(self, provider, email_addr, password):
        """Prueba la conexión SMTP con los parámetros proporcionados"""
        try:
            # Obtener configuración del proveedor
            config = self.get_provider_config(provider)
            server = config['smtp_server']
            port = config['smtp_port']

            # Asegurarse de que las credenciales sean strings y eliminar caracteres problemáticos
            email_addr = self._sanitize_string(email_addr)
            password = self._sanitize_string(password)

            # Crear un contexto SSL
            context = ssl.create_default_context()

            # Conectar al servidor SMTP
            smtp = smtplib.SMTP(server, port)
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.ehlo()

            # Iniciar sesión con las credenciales
            smtp.login(email_addr, password)

            # Cerrar la conexión
            smtp.quit()

            return True

        except Exception as e:
            print(f"Error en la conexión SMTP: {str(e)}")
            return False

    def test_imap_connection(self, provider, email_addr, password):
        """Prueba la conexión IMAP con los parámetros proporcionados"""
        try:
            # Obtener configuración del proveedor
            config = self.get_provider_config(provider)
            server = config['imap_server']
            port = config['imap_port']

            # Asegurarse de que las credenciales sean strings y eliminar caracteres problemáticos
            email_addr = self._sanitize_string(email_addr)
            password = self._sanitize_string(password)

            # Crear un contexto SSL
            context = ssl.create_default_context()

            # Conectar al servidor IMAP
            imap = imaplib.IMAP4_SSL(server, port, ssl_context=context)

            # Iniciar sesión con las credenciales
            imap.login(email_addr, password)

            # Cerrar la conexión
            imap.logout()

            return True

        except Exception as e:
            print(f"Error en la conexión IMAP: {str(e)}")
            return False

    def send_email(
        self, provider, email_addr, password, to, subject, body, cc_list=None, attachments=None
    ):
        """Envía un correo electrónico a través de SMTP"""
        try:
            # Obtener configuración del proveedor
            config = self.get_provider_config(provider)
            server = config['smtp_server']
            port = config['smtp_port']

            # Sanitizar credenciales
            email_addr = self._sanitize_string(email_addr)
            password = self._sanitize_string(password)

            # Crear un mensaje MIME
            msg = MIMEMultipart()
            msg['From'] = email_addr
            msg['To'] = to
            msg['Subject'] = Header(subject, 'utf-8')

            # Añadir cabecera CC si la lista existe
            if cc_list:
                msg['Cc'] = ", ".join(cc_list)

            # Adjuntar el cuerpo del mensaje con codificación UTF-8
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            if attachments:
                for attachment in attachments:
                    if not isinstance(attachment, dict):
                        continue
                    filename = attachment.get('filename') or 'adjunto'
                    mime_type = attachment.get('mime_type') or 'application/octet-stream'
                    content = attachment.get('content')

                    if content is None:
                        continue
                    if isinstance(content, str):
                        content = content.encode('utf-8')

                    maintype, _, subtype = mime_type.partition('/')
                    maintype = maintype or 'application'
                    subtype = subtype or 'octet-stream'

                    try:
                        if maintype == 'text':
                            decoded = content.decode('utf-8', errors='ignore')
                            part = MIMEText(decoded, _subtype=subtype, _charset='utf-8')
                        elif maintype == 'application':
                            part = MIMEApplication(content, _subtype=subtype)
                        else:
                            part = MIMEBase(maintype, subtype)
                            part.set_payload(content)
                            encoders.encode_base64(part)
                    except Exception:
                        part = MIMEApplication(content, _subtype='octet-stream')

                    part.add_header('Content-Disposition', 'attachment', filename=filename)
                    msg.attach(part)

            # Crear un contexto SSL
            context = ssl.create_default_context()

            # Conectar al servidor SMTP
            with smtplib.SMTP(server, port) as smtp:
                smtp.ehlo()
                smtp.starttls(context=context)
                smtp.ehlo()

                # Iniciar sesión y enviar el correo
                smtp.login(email_addr, password)
                smtp.send_message(msg)

            return True

        except Exception as e:
            print(f"Error al enviar correo: {str(e)}")
            return False

    def read_emails(self, provider, email_addr, password, mailbox='INBOX', limit=10):
        """Lee correos de un buzón IMAP específico"""
        try:
            # Obtener configuración del proveedor
            config = self.get_provider_config(provider)
            server = config['imap_server']
            port = config['imap_port']

            # Sanitizar credenciales
            email_addr = self._sanitize_string(email_addr)
            password = self._sanitize_string(password)

            # Crear un contexto SSL
            context = ssl.create_default_context()

            # Conectar al servidor IMAP
            with imaplib.IMAP4_SSL(server, port, ssl_context=context) as imap:
                # Iniciar sesión
                imap.login(email_addr, password)

                # Seleccionar el buzón de correo
                imap.select(mailbox)

                # Buscar todos los correos no leídos
                status, messages = imap.search(None, 'UNSEEN')

                # Obtener la lista de IDs de mensajes
                message_ids = messages[0].split()

                # Limitar la cantidad de correos a procesar
                if limit > 0:
                    message_ids = message_ids[:limit]

                emails = []

                # Leer cada correo con codificación UTF-8
                for msg_id in message_ids:
                    status, data = imap.fetch(msg_id, '(RFC822)')
                    raw_email = data[0][1]
                    # Usar UTF-8 para decodificar el correo
                    email_message = email.message_from_bytes(raw_email, policy=email.policy.default)
                    emails.append(email_message)

                return emails

        except Exception as e:
            print(f"Error al leer correos: {str(e)}")
            return []

    def check_and_process_emails(self, provider, email_addr, password, search_titles, logger, cc_list=None):
        """Función principal que revisa emails y procesa los que coinciden usando el sistema modular"""
        try:
            # Obtener configuración del proveedor
            config = self.get_provider_config(provider)
            server = config['imap_server']
            port = config['imap_port']

            # Sanitizar credenciales
            email_addr = self._sanitize_string(email_addr)
            password = self._sanitize_string(password)

            # Crear un contexto SSL
            context = ssl.create_default_context()

            # Conectar al servidor IMAP
            with imaplib.IMAP4_SSL(server, port, ssl_context=context) as imap:
                # Iniciar sesión
                imap.login(email_addr, password)

                # Seleccionar el buzón de correo
                imap.select('INBOX')

                # --- LÓGICA DE BÚSQUEDA CENTRALIZADA ---
                today = date.today().strftime("%d-%b-%Y")

                filters_context = self._collect_case_filters(search_titles)
                allowed_cases: Set[str] = filters_context['allowed_cases']
                active_keywords: List[str] = filters_context['active_keywords']
                missing_keywords: List[str] = filters_context['missing_keywords']
                cases_without_keywords: List[str] = filters_context['cases_without_keywords']
                case_keywords: Dict[str, List[str]] = filters_context['case_keywords']
                allowed_keyword_norms: Set[str] = filters_context['allowed_keyword_norms']
                requested_keywords = filters_context['requested_keywords']

                if missing_keywords:
                    logger.log(
                        "Las siguientes palabras clave configuradas no corresponden a ningún caso activo: "
                        + ", ".join(sorted(missing_keywords)),
                        level="WARNING",
                    )

                if cases_without_keywords:
                    logger.log(
                        "Los siguientes casos no tienen palabras clave configuradas y no participarán en la búsqueda: "
                        + ", ".join(sorted(cases_without_keywords)),
                        level="INFO",
                    )

                if not allowed_cases:
                    logger.log(
                        "No hay casos habilitados con palabras clave configuradas para ejecutar la búsqueda de correos.",
                        level="WARNING",
                    )
                    return

                if requested_keywords and not active_keywords:
                    logger.log(
                        "Las palabras clave proporcionadas no están asociadas a ningún caso con filtros válidos. Se omite la revisión.",
                        level="WARNING",
                    )
                    return

                if active_keywords:
                    logger.log(
                        "Casos habilitados para revisión: "
                        + ", ".join(sorted(allowed_cases))
                        + " | Palabras clave activas: "
                        + ", ".join(sorted(set(active_keywords))),
                        level="INFO",
                    )

                base_tokens = ['UNSEEN']
                if today:
                    base_tokens.extend(['SINCE', today])

                candidate_case_map: Dict[bytes, Set[str]] = {}

                # Ejecutar búsqueda específica por palabra clave para cada caso habilitado
                # IMPORTANTE: Buscar en orden inverso (case12 primero) para evitar que
                # "Caso 1" haga match con "Caso 12" en la búsqueda IMAP (que usa substring)
                logger.log("=== INICIANDO BÚSQUEDA IMAP POR KEYWORDS ===", level="INFO")

                # Ordenar casos en orden inverso: case12, case11, ..., case1
                def case_sort_key(case_name):
                    # Extraer el número del case (case1 -> 1, case12 -> 12)
                    try:
                        return -int(case_name.replace('case', ''))
                    except:
                        return 0

                sorted_cases = sorted(allowed_cases, key=case_sort_key)
                logger.log(f"Orden de búsqueda IMAP: {sorted_cases}", level="INFO")

                for case_name in sorted_cases:
                    keywords = case_keywords.get(case_name, [])
                    if not keywords:
                        logger.log(f"⚠ {case_name}: NO tiene keywords configuradas", level="WARNING")
                        continue

                    logger.log(f"Buscando con {case_name}: {keywords}", level="INFO")

                    for keyword in keywords:
                        prepared_keyword = self._prepare_keyword_for_search(keyword)
                        if not prepared_keyword:
                            logger.log(f"⚠ Keyword vacía después de preparar: '{keyword}'", level="WARNING")
                            continue

                        search_tokens = self._build_keyword_search_tokens(
                            base_tokens, prepared_keyword
                        )
                        logger.log(f"  Tokens IMAP: {search_tokens}", level="DEBUG")

                        status, data = imap.search(None, *search_tokens)
                        if status != 'OK' or not data:
                            logger.log(f"  IMAP search falló: status={status}", level="DEBUG")
                            continue
                        ids = data[0].split()
                        if not ids:
                            logger.log(f"  Sin resultados para '{keyword}'", level="DEBUG")
                            continue
                        logger.log(
                            f"✓ Encontrados {len(ids)} correos para '{keyword}' del {case_name}",
                            level="INFO",
                        )
                        for mid in ids:
                            # Solo agregar el caso si el email NO está ya asignado
                            # Como iteramos en orden inverso (case12 primero), el primero que encuentra gana
                            if mid not in candidate_case_map:
                                candidate_case_map[mid] = set()
                                candidate_case_map[mid].add(case_name)
                                logger.log(f"  Email {mid.decode()[:20]}... asignado a {case_name}", level="DEBUG")
                            else:
                                logger.log(f"  Email {mid.decode()[:20]}... YA asignado a {candidate_case_map[mid]}, ignorando {case_name}", level="DEBUG")

                # Si no hubo resultados por palabras clave, realizar una búsqueda general como respaldo
                if not candidate_case_map:
                    status, data = imap.search(None, *base_tokens)
                    if status != 'OK' or not data:
                        logger.log(
                            f"Búsqueda IMAP sin resultados válidos (status={status}).",
                            level="WARNING",
                        )
                        return
                    fallback_ids = data[0].split()
                    if not fallback_ids:
                        logger.log("No se encontraron correos nuevos que coincidan con los criterios.", level="INFO")
                        return
                    logger.log(
                        "No se encontraron coincidencias directas por palabras clave; se revisarán todos los correos no leídos recientes.",
                        level="INFO",
                    )
                    for mid in fallback_ids:
                        candidate_case_map.setdefault(mid, set())

                message_ids = list(candidate_case_map.keys())
                # --- FIN DE LÓGICA DE BÚSQUEDA CENTRALIZADA ---

                if not message_ids:
                    logger.log("No se encontraron correos nuevos que coincidan con los criterios.", level="INFO")
                    return

                logger.log(f"Encontrados {len(message_ids)} emails que coinciden con la búsqueda", level="INFO")

                # Procesar cada email
                for msg_id in message_ids:
                    try:
                        # Obtener solo las cabeceras del email SIN marcarlo como leído
                        status, header_data = imap.fetch(msg_id, '(BODY.PEEK[HEADER])')

                        if status != 'OK' or not header_data:
                            logger.log(f"No se pudieron obtener las cabeceras del email {msg_id}", level="WARNING")
                            continue

                        # Parsear solo las cabeceras
                        raw_headers = header_data[0][1]
                        headers = email.message_from_bytes(raw_headers, policy=email.policy.default)

                        # Obtener y decodificar el subject del email
                        subject = self._decode_header_value(headers.get('Subject', ''))
                        sender = headers.get('From', '')

                        logger.log(f"Revisando email: '{subject}' de {sender}", level="INFO")

                        candidate_cases = candidate_case_map.get(msg_id, set())
                        if candidate_cases:
                            allowed_for_message = candidate_cases
                        else:
                            allowed_for_message = allowed_cases

                        # Buscar caso coincidente usando el sistema modular
                        matching_case = self.case_handler.find_matching_case(
                            subject,
                            logger,
                            allowed_cases=allowed_for_message,
                        )

                        if matching_case:
                            logger.log(f"Email encontrado para caso: {matching_case}", level="INFO")

                            # Marcar como leído
                            status, result = self._mark_as_read(imap, msg_id)
                            if status:
                                logger.log(f"Email marcado como leído: {result}", level="INFO")

                                # Obtener el email completo para extraer adjuntos y contenido
                                status, full_data = imap.fetch(msg_id, '(RFC822)')
                                if status != 'OK' or not full_data or not full_data[0]:
                                    logger.log(
                                        f"No se pudo obtener el contenido completo del email {msg_id}",
                                        level="ERROR",
                                    )
                                    continue

                                full_message = email.message_from_bytes(
                                    full_data[0][1], policy=email.policy.default
                                )
                                attachments_info = self._extract_attachments(full_message, logger)
                                body_text = self._extract_body(full_message)

                                if attachments_info:
                                    logger.log(
                                        f"Se detectaron {len(attachments_info)} adjunto(s) en el correo.",
                                        level="INFO",
                                    )

                                # Preparar datos del email para el caso
                                email_data = {
                                    'sender': sender,
                                    'subject': subject,
                                    'msg_id': msg_id.decode() if isinstance(msg_id, bytes) else str(msg_id),
                                    'attachments': attachments_info,
                                    'body': body_text,
                                    'email_message': full_message,
                                }

                                # Ejecutar el caso correspondiente
                                response_data = self.case_handler.execute_case(matching_case, email_data, logger)

                                if response_data:
                                    # Enviar respuesta automática (con CC si está configurado)
                                    if self._send_case_reply(provider, email_addr, password, response_data, logger,
                                                             cc_list):
                                        logger.log(f"Respuesta automática enviada usando {matching_case}", level="INFO")
                                    else:
                                        logger.log(f"Error al enviar respuesta automática", level="ERROR")
                                else:
                                    logger.log(f"Error al procesar {matching_case}", level="ERROR")
                            else:
                                logger.log(f"Error al marcar email como leído: {result}", level="ERROR")
                        else:
                            # Este log ahora es menos probable, ya que el servidor ya filtró por asunto
                            if allowed_keyword_norms:
                                normalized_subject = self._normalize_keyword(subject)
                                if not any(norm in normalized_subject for norm in allowed_keyword_norms):
                                    logger.log(
                                        f"Email omitido por no contener palabras clave configuradas: '{subject}'",
                                        level="INFO",
                                    )
                                else:
                                    logger.log(
                                        f"Email no coincide con ningún caso específico de respuesta: '{subject}'",
                                        level="INFO",
                                    )
                            else:
                                logger.log(
                                    f"Email no coincide con ningún caso específico de respuesta: '{subject}'",
                                    level="INFO",
                                )

                    except Exception as e:
                        logger.log(f"Error al procesar email individual: {str(e)}", level="ERROR")

        except Exception as e:
            logger.log(f"Error en check_and_process_emails: {str(e)}", level="ERROR")

    def _send_case_reply(self, provider, email_addr, password, response_data, logger, cc_list=None):
        """Envía una respuesta automática usando los datos del caso"""
        try:
            recipient = response_data.get('recipient', '')
            subject = response_data.get('subject', '')
            body = response_data.get('body', '')

            # Extraer solo la dirección de email del remitente
            if '<' in recipient and '>' in recipient:
                recipient = recipient.split('<')[1].split('>')[0].strip()

            # Enviar la respuesta
            attachments = response_data.get('attachments') if isinstance(response_data, dict) else None
            return self.send_email(
                provider,
                email_addr,
                password,
                recipient,
                subject,
                body,
                cc_list=cc_list,
                attachments=attachments,
            )

        except Exception as e:
            logger.log(f"Error al enviar respuesta del caso: {str(e)}", level="ERROR")
            return False

    def _collect_case_filters(self, search_titles):
        """Construye el contexto de filtros activos por caso a partir de sus palabras clave."""
        case_keywords = self.case_handler.get_case_keywords()

        case_keyword_norms: Dict[str, Set[str]] = {}
        all_normalized_keywords: Set[str] = set()

        for case_name, keywords in case_keywords.items():
            normalized_set: Set[str] = set()
            for keyword in keywords:
                normalized = self._normalize_keyword(keyword)
                if normalized:
                    normalized_set.add(normalized)
                    all_normalized_keywords.add(normalized)
            case_keyword_norms[case_name] = normalized_set

        requested_keywords: Dict[str, str] = {}
        if search_titles:
            for item in search_titles:
                if not isinstance(item, str):
                    continue
                normalized_item = self._normalize_keyword(item)
                if normalized_item:
                    requested_keywords[normalized_item] = item.strip()

        requested_norms = set(requested_keywords.keys())

        if requested_norms:
            allowed_cases = {
                case_name
                for case_name, norms in case_keyword_norms.items()
                if norms.intersection(requested_norms)
            }
        else:
            allowed_cases = {
                case_name
                for case_name, norms in case_keyword_norms.items()
                if norms
            }

        allowed_keyword_norms = {
            norm
            for case_name in allowed_cases
            for norm in case_keyword_norms.get(case_name, set())
        }

        missing_keywords = [
            original
            for norm, original in requested_keywords.items()
            if norm not in all_normalized_keywords
        ]

        cases_without_keywords = [
            case_name for case_name, norms in case_keyword_norms.items() if not norms
        ]

        active_keywords: List[str] = [
            keyword
            for case_name in allowed_cases
            for keyword in case_keywords.get(case_name, [])
        ]

        return {
            'case_keywords': case_keywords,
            'case_keyword_norms': case_keyword_norms,
            'allowed_cases': allowed_cases,
            'allowed_keyword_norms': allowed_keyword_norms,
            'requested_keywords': requested_keywords,
            'missing_keywords': missing_keywords,
            'cases_without_keywords': cases_without_keywords,
            'active_keywords': active_keywords,
        }

    def _prepare_keyword_for_search(self, keyword: str) -> str:
        """Normaliza y valida una palabra clave para comandos de búsqueda IMAP."""
        if not isinstance(keyword, str):
            return ""

        sanitized = self._sanitize_string(keyword).strip()
        return sanitized

    def _build_keyword_search_tokens(self, base_tokens: List[str], keyword: str) -> List[str]:
        """Construye los tokens de búsqueda incluyendo una palabra clave segura."""
        tokens = list(base_tokens)

        if keyword:
            tokens.extend(['SUBJECT', self._quote_for_imap(keyword)])

        return tokens

    def _quote_for_imap(self, value: str) -> str:
        """Escapa y cita un valor para usarlo en comandos IMAP."""
        if not isinstance(value, str):
            value = str(value) if value is not None else ""

        escaped = value.replace('\\', '\\\\').replace('"', '\\"')
        return f'"{escaped}"'

    def _normalize_keyword(self, keyword: str) -> str:
        """Normaliza una palabra clave para comparaciones insensibles a espacios y mayúsculas."""
        if not isinstance(keyword, str):
            return ""
        collapsed = " ".join(keyword.split())
        return collapsed.casefold()

    def _extract_attachments(self, email_message, logger):
        """Extrae los adjuntos del email en una lista de diccionarios"""
        attachments = []
        try:
            for attachment in email_message.iter_attachments():
                filename = attachment.get_filename()
                if not filename:
                    continue

                try:
                    content = attachment.get_payload(decode=True)
                except Exception:
                    content = None

                if content is None:
                    try:
                        payload = attachment.get_content()
                        if isinstance(payload, str):
                            charset = attachment.get_content_charset() or 'utf-8'
                            content = payload.encode(charset, errors='ignore')
                        else:
                            content = payload
                    except Exception:
                        content = None

                if content is None:
                    logger.log(f"No se pudo decodificar el adjunto '{filename}'", level="WARNING")
                    continue

                attachments.append({
                    'filename': filename,
                    'content': content,
                    'mime_type': attachment.get_content_type(),
                })
        except Exception as exc:
            logger.log(f"Error al extraer adjuntos: {exc}", level="ERROR")

        return attachments

    def _extract_body(self, email_message):
        """Obtiene el cuerpo del correo en texto plano"""
        if email_message.is_multipart():
            parts = []
            for part in email_message.walk():
                content_type = part.get_content_type()
                disposition = part.get_content_disposition()
                if content_type == 'text/plain' and disposition in (None, 'inline'):
                    try:
                        parts.append(part.get_content())
                    except Exception:
                        payload = part.get_payload(decode=True)
                        if payload:
                            charset = part.get_content_charset() or 'utf-8'
                            parts.append(payload.decode(charset, errors='ignore'))
            return '\n'.join(parts).strip()

        try:
            content = email_message.get_content()
            if isinstance(content, str):
                return content.strip()
        except Exception:
            payload = email_message.get_payload(decode=True)
            if payload:
                charset = email_message.get_content_charset() or 'utf-8'
                return payload.decode(charset, errors='ignore').strip()

        return ''

    def _decode_header_value(self, header_value):
        """Decodifica un valor de cabecera que puede estar codificado"""
        if not header_value:
            return ""

        try:
            decoded_parts = decode_header(header_value)
            decoded_text = ""

            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        decoded_text += part.decode(encoding)
                    else:
                        decoded_text += part.decode('utf-8', errors='ignore')
                else:
                    decoded_text += part

            return decoded_text
        except Exception as e:
            print(f"Error al decodificar cabecera: {str(e)}")
            return str(header_value)

    def _mark_as_read(self, imap_connection, msg_id):
        """Marca un email específico como leído"""
        try:
            # Añadir la flag \Seen al mensaje para marcarlo como leído
            status, result = imap_connection.store(msg_id, '+FLAGS', '\\Seen')
            if status != 'OK':
                return False, f"Estado no OK: {status}"

            # Verificar el resultado
            if not result or not result[0]:
                return False, "Resultado vacío"

            return True, "Email marcado correctamente como leído"
        except Exception as e:
            return False, f"Error al marcar email como leído: {str(e)}"

    def _is_today(self, email_date_str):
        """Verifica si un email es del día de hoy"""
        try:
            # Parsear la fecha del email
            email_date = email.utils.parsedate_to_datetime(email_date_str)
            today = datetime.now().date()

            # Comparar solo la fecha (sin hora)
            return email_date.date() == today

        except Exception as e:
            print(f"Error al verificar fecha del email: {str(e)}")
            return False

    def _sanitize_string(self, text):
        """Sanitiza un string para evitar problemas de codificación"""
        if not isinstance(text, str):
            return str(text)

        # Eliminar caracteres no imprimibles y espacios no separables
        text = ''.join(c for c in text if c.isprintable() and ord(c) != 0xA0)
        return text

    def reload_cases(self):
        """Recarga todos los casos disponibles"""
        self.case_handler.reload_cases()

    def get_available_cases(self):
        """Obtiene los casos disponibles"""
        return self.case_handler.get_available_cases()

    def get_case_info(self, case_name):
        """Obtiene información de un caso específico"""
        return self.case_handler.get_case_info(case_name)