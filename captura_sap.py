import win32com.client
import csv
import json
import sys
import time
import os
import win32clipboard

def set_clipboard_text(text):
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()

def get_sap_session():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        return session
    except Exception as e:
        print("================================================================")
        print("ERROR AL CONECTAR CON SAP")
        print("Por favor, asegúrate de tener SAP GUI abierto y haber iniciado sesión.")
        print("También verifica que el Scripting esté habilitado.")
        print("================================================================")
        sys.exit(1)

def leer_csv(ruta):
    datos = []
    try:
        with open(ruta, mode='r', encoding='utf-8-sig') as f:
            lector = csv.DictReader(f)
            for fila in lector:
                datos.append(fila)
    except FileNotFoundError:
        print(f"No se encontró el archivo: {ruta}")
    return datos

def leer_filtro_productos(ruta):
    productos = []
    if not os.path.isfile(ruta):
        return productos
    try:
        with open(ruta, mode='r', encoding='utf-8-sig') as f:
            lector = csv.DictReader(f)
            for fila in lector:
                val = (fila.get('Material') or '').strip()
                if val:
                    productos.append(val)
    except Exception as e:
        print(f"No se pudo leer el archivo de filtros: {e}")
    return productos


def clave_material(matnr):
    """Clave de comparacion de materiales: quita ceros a la izquierda en valores
    numericos para que '1093510' y '0000001093510' equivalgan. Solo se usa para
    comparar; nunca se escribe en SAP (SAP aplica su propia conversion)."""
    m = (matnr or '').strip()
    if m.isdigit():
        return m.lstrip('0') or '0'
    return m


def leer_plantas_asignadas(ruta):
    if not os.path.isfile(ruta):
        return {}
    try:
        with open(ruta, mode='r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"No se pudo leer el archivo de plantas asignadas: {e}")
        return {}

def esperar_control(session, id_control, timeout=8.0, intervalo=0.25):
    """Espera activa hasta que el control exista (en vez de sleeps fijos).
    Retorna el control o None si expira el timeout."""
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            return session.findById(id_control)
        except Exception:
            time.sleep(intervalo)
    return None

def cerrar_modales(session, max_pasadas=4):
    """Cierra popups/modales en cadena prefiriendo 'No'/Cancelar: nunca el
    boton por defecto, para no arrastrar datos de la consulta anterior."""
    for _ in range(max_pasadas):
        try:
            session.findById("wnd[1]")
        except Exception:
            return  # No hay modal
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()  # Boton "No"
        except Exception:
            try:
                session.findById("wnd[1]/tbar[0]/btn[12]").press()  # Cancelar
            except Exception:
                try:
                    session.findById("wnd[1]").sendVKey(12)
                except Exception:
                    pass
        time.sleep(0.3)

def avanzar_pagina(session, table_id):
    """Avanza una pagina de la tabla usando el mismo mecanismo verificado que
    funciona al insertar: foco en el CONTROL de tabla (no en una celda) +
    Page Down + verificacion de que la vista realmente avanzo.
    Retorna True si avanzo, False si estaba al final (o no se pudo verificar)."""
    try:
        session.findById(table_id).setFocus()
    except Exception:
        pass
    try:
        pos_antes = session.findById(table_id).FirstVisibleRow
    except Exception:
        pos_antes = None
    try:
        session.findById("wnd[0]").sendVKey(82)  # Page Down
        time.sleep(0.1)
    except Exception:
        return False
    try:
        pos_despues = session.findById(table_id).FirstVisibleRow
        if pos_antes is not None and pos_despues == pos_antes:
            return False  # no avanzo: fin de la tabla
        return True
    except Exception:
        return True  # no se pudo verificar; el dedup de paginas evita bucles


def scroll_al_inicio(session, table_id=None):
    """Scroll al inicio de la tabla. El setFocus garantiza que el scroll
    llegue a la tabla (sin foco, los Page Up no la afectan)."""
    if table_id:
        try:
            session.findById(f"{table_id}/ctxtMV10A-MATNR[0,0]").setFocus()
        except Exception:
            pass
    pos_previa = None
    for _ in range(6):
        try:
            session.findById("wnd[0]").sendVKey(81)  # Page Up
            time.sleep(0.05)
            if table_id:
                pos = session.findById(table_id).FirstVisibleRow
                if pos == pos_previa:
                    return
                pos_previa = pos
        except Exception:
            return

def scroll_al_final(session, table_id=None):
    """Scroll al final de la tabla con Page Down verificado (con foco en ella)."""
    if table_id:
        try:
            session.findById(table_id).setFocus()
        except Exception:
            pass
    pos_previa = None
    for _ in range(6):
        try:
            session.findById("wnd[0]").sendVKey(82)  # Page Down
            time.sleep(0.05)
            if table_id:
                pos = session.findById(table_id).FirstVisibleRow
                if pos == pos_previa:
                    return
                pos_previa = pos
        except Exception:
            return

def main():
    session = get_sap_session()
    
    # Leer datos de los archivos CSV
    clientes = leer_csv('clientes.csv')
    productos_dict = leer_csv('productos.csv')
    plantas = leer_plantas_asignadas('plantas_asignadas.json')
    
    if not clientes or not productos_dict:
        print("Asegúrate de que 'clientes.csv' y 'productos.csv' tengan datos y estén en la misma carpeta.")
        return
        
    productos = [p['Material'].strip() for p in productos_dict if p['Material'].strip()]
    
    productos_filtro = leer_filtro_productos('productos_filtros.csv')
    if productos_filtro:
        print(f"Se encontraron {len(productos_filtro)} productos para aplicar como filtro en la consulta.")
        # Cuando hay filtro, SOLO se gestionan esos productos en toda la corrida
        productos_objetivo = productos_filtro
    else:
        print("No hay productos de filtro o el archivo productos_filtros.csv esta vacio/ausente. Se omite el filtro.")
        productos_objetivo = productos
    
    print(f"Se encontraron {len(clientes)} clientes y {len(productos)} productos.")
    print("Iniciando procesamiento...")
    print("-" * 50)

    t_inicio_corrida = time.time()

    for idx_cliente, cliente in enumerate(clientes, 1):
        kunnr = cliente['Cliente'].strip()
        vkorg = cliente['OrgVentas'].strip()
        vtweg = cliente['CanalDist'].strip()
        planta_manual = (cliente.get('Planta') or cliente.get('Plant') or '').strip()
        plant = planta_manual or (plantas.get(kunnr, '') or '').strip()
        
        if not plant:
            print(f"[{idx_cliente}/{len(clientes)}] Cliente: {kunnr} -> Saltado: Planta vacia. Ejecuta asignar_plantas.py primero.")
            continue
        
        print(f"[{idx_cliente}/{len(clientes)}] Procesando Cliente: {kunnr} | Planta: {plant}")
        t_cliente = time.time()

        try:
            # OPTIMIZACION: si ya estamos en la pantalla de seleccion de VD52
            # (el guardado del cliente anterior regresa aqui), se reutiliza la
            # pantalla y se omiten los F12 y el /nVD52: menos flashazos y menos
            # tiempo muerto por cliente.
            cerrar_modales(session)  # por si quedo un popup del guardado anterior
            en_vd52 = False
            try:
                if session.info.Transaction == "VD52":
                    en_vd52 = esperar_control(session, "wnd[0]/usr/ctxtKUNNR", timeout=2.0, intervalo=0.2) is not None
            except Exception:
                en_vd52 = False

            if not en_vd52:
                # Cancelar cualquier transaccion/pantalla previa que haya quedado
                # abierta. F12 repetido cancela primero los modales pendientes
                # ("¿Guardar cambios?") y luego la pantalla activa.
                for _ in range(4):
                    try:
                        session.findById("wnd[0]").sendVKey(12)  # F12 = Cancelar
                        time.sleep(0.3)
                    except Exception:
                        break

                # Abrir VD52 reintentando la NAVEGACION completa.
                max_nav = 4
                for nav_intent in range(1, max_nav + 1):
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVD52"
                    session.findById("wnd[0]").sendVKey(0)

                    # Espera activa a la pantalla de seleccion (no sleep fijo)
                    if esperar_control(session, "wnd[0]/usr/ctxtKUNNR", timeout=6.0, intervalo=0.25):
                        break

                    cerrar_modales(session)
                    txn = ""
                    try:
                        txn = session.info.Transaction
                    except Exception:
                        pass
                    print(f"   -> Pantalla de seleccion no disponible (intento {nav_intent}/{max_nav}, txn={txn}). Reintentando...")
                    time.sleep(1.0)

            # Verificar que SI estamos en la pantalla de seleccion de VD52
            # y llenar los datos del cliente solo cuando el control exista
            if esperar_control(session, "wnd[0]/usr/ctxtKUNNR", timeout=6.0, intervalo=0.25) is None:
                txn = ""
                try:
                    txn = session.info.Transaction
                except Exception:
                    pass
                raise RuntimeError(
                    f"No se pudo abrir la pantalla de seleccion de VD52 (txn={txn}). "
                    "Verifica que SAP este logueado y sin ventanas pendientes."
                )
            session.findById("wnd[0]/usr/ctxtKUNNR").text = kunnr
            session.findById("wnd[0]/usr/ctxtVKORG").text = vkorg
            session.findById("wnd[0]/usr/ctxtVTWEG").text = vtweg
            
            print(f"   -> Ingresando Cliente: {kunnr}")
            print(f"   -> Ingresando Org Ventas: {vkorg}")
            print(f"   -> Ingresando Canal Dist: {vtweg}")
            
            # Limpiar el campo Material (rango) que queda de la consulta anterior.
            # Via rapida: IDs directos. Fallback: enumeracion completa de campos.
            print("   -> Limpiando campos extra de la pantalla de selección...")
            limpiado_directo = False
            for cid in ("wnd[0]/usr/ctxtMATNR_R-LOW", "wnd[0]/usr/ctxtMATNR_R-HIGH"):
                try:
                    campo = session.findById(cid)
                    limpiado_directo = True
                    if campo.text and campo.text.strip():
                        campo.text = ""
                        print(f"      Campo limpiado: {cid}")
                except Exception:
                    pass
            
            if not limpiado_directo:
                # Fallback: limpiar TODOS los campos de texto excepto los tres
                # requeridos, sin importar el ID exacto ni si es rango o valor.
                # Usamos los SUFIJOS de los IDs porque SAP puede devolver el path
                # completo (/app/con[0]/ses[0]/wnd[0]/usr/ctxtKUNNR).
                sufijos_requeridos = ("ctxtKUNNR", "ctxtVKORG", "ctxtVTWEG")
                try:
                    usr_container = session.findById("wnd[0]/usr")
                    for i in range(usr_container.Children.Count):
                        try:
                            child = usr_container.Children(i)
                            child_id = child.Id
                            es_requerido = any(child_id.endswith(s) for s in sufijos_requeridos)
                            if not es_requerido:
                                try:
                                    if child.text and child.text.strip():
                                        child.text = ""
                                        print(f"      Campo limpiado: {child_id}")
                                except Exception:
                                    pass
                        except Exception:
                            pass
                except Exception as enum_err:
                    print(f"   -> No se pudo enumerar campos (ignorando): {enum_err}")
            
            # --- APLICAR FILTRO DE PRODUCTOS (si existe) ---
            if productos_filtro:
                print(f"   -> Aplicando filtro de {len(productos_filtro)} productos a la consulta...")
                try:
                    # Abrir el dialogo de seleccion multiple del campo Material
                    session.findById("wnd[0]/usr/btn%_MATNR_R_%_APP_%-VALU_PUSH").press()
                    esperar_control(session, "wnd[1]/tbar[0]/btn[24]", timeout=4.0, intervalo=0.2)

                    # Eliminar cualquier lista previa que haya quedado en la "memoria" del autocompletar de SAP
                    try:
                        session.findById("wnd[1]/tbar[0]/btn[16]").press()
                        time.sleep(0.2)
                    except Exception:
                        pass

                    # Cargar los materiales en el portapapeles de Windows (uno por linea)
                    clipboard_text = "\r\n".join(productos_filtro)
                    set_clipboard_text(clipboard_text)

                    # Pegar desde el portapapeles (boton Paste del dialogo de rangos)
                    session.findById("wnd[1]/tbar[0]/btn[24]").press()
                    time.sleep(0.5)

                    # Aceptar el dialogo de rangos (Copy/Continue)
                    session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    time.sleep(0.2)
                except Exception as filtro_err:
                    print(f"   -> ERROR al aplicar filtro de productos: {filtro_err}")
                    # Intentar cerrar el dialogo de rangos si quedo abierto
                    try:
                        session.findById("wnd[1]").sendVKey(12)
                    except Exception:
                        pass
            
            table_id = "wnd[0]/usr/tblSAPMV10ATC_CU_MA"
            
            # Ir directamente a Ejecutar (F8) SIN presionar Enter primero.
            # Presionar Enter activa el historial de campo de SAP y puede restaurar
            # el último material buscado como filtro no deseado.
            print("   -> Ejecutando búsqueda (F8) sin pasar por Enter...")
            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                # Si F8 no está disponible directamente, usar Enter como fallback
                print("   -> F8 no disponible, usando Enter como fallback...")
                session.findById("wnd[0]").sendVKey(0)
                try:
                    session.findById("wnd[0]/tbar[1]/btn[8]").press()
                except Exception:
                    pass
            
            # Espera activa a los resultados: tabla presente O error de consulta
            fin_espera = time.time() + 10.0
            while time.time() < fin_espera:
                try:
                    session.findById(f"{table_id}/ctxtMV10A-MATNR[0,0]")
                    break
                except Exception:
                    try:
                        if session.findById("wnd[0]/sbar").MessageType == "E":
                            break  # error de consulta: lo detectara la Fase 1
                    except Exception:
                        pass
                    time.sleep(0.25)
            
            # Leer el número de filas visibles una sola vez al inicio para evitar errores COM tras modificar celdas
            tabla = session.findById(table_id)
            try:
                vis_rows = tabla.visibleRowCount
            except Exception:
                vis_rows = 15  # Fallback seguro

            # Detectar si la tabla tiene columna de Planta (WERKS): permite escribir
            # la planta directamente en la celda en vez de abrir el detalle (F2) por
            # producto. Si no se detecta o la celda no acepta el valor, se usa el
            # metodo de detalle de siempre (fallback seguro).
            celda_planta_prefijo = None
            try:
                columnas = tabla.ColumnOrder
                try:
                    columnas = list(columnas)
                except Exception:
                    columnas = [columnas(i) for i in range(columnas.Count)]
                for col in columnas:
                    if str(col).upper().endswith("WERKS"):
                        celda_planta_prefijo = f"{table_id}/ctxtMV10A-{col}"
                        break
            except Exception:
                celda_planta_prefijo = None
            if celda_planta_prefijo:
                print(f"   -> Columna de planta detectada en tabla: escritura directa ({celda_planta_prefijo.split('/')[-1]})")
            else:
                print("   -> Columna de planta no detectada: se usara la via detalle (F2) por producto.")
            
            # --- FASE 1: LEER PRODUCTOS EXISTENTES EN SAP (OPTIMIZADO) ---
            # OPTIMIZACION: como la consulta ya viene filtrada por Material
            # (productos_filtros.csv), una tabla vacia implica que NINGUN producto
            # del filtro existe para este cliente => se omite el escaneo completo
            # y se insertan todos. Si hay registros, se escanea solo lo necesario:
            # salida temprana al encontrar todos los objetivo y deteccion de fin
            # de tabla con una sola lectura COM por pagina.
            existing_materials = set()
            existing_claves = set()

            # Proteccion para corridas largas: si SAP rechazo la consulta (ej.
            # combinacion org/canal invalida), abortar el cliente con un mensaje
            # claro en vez de insertar sobre una pantalla rota.
            try:
                sbar = session.findById("wnd[0]/sbar")
                if sbar.MessageType == "E":
                    raise RuntimeError(f"SAP rechazo la consulta: {sbar.Text}")
            except RuntimeError:
                raise
            except Exception:
                pass

            # Leer solo la primera fila: si esta vacia, no hay nada que validar
            try:
                primera_fila = session.findById(f"{table_id}/ctxtMV10A-MATNR[0,0]").text.strip()
            except Exception:
                primera_fila = ""

            if not primera_fila:
                print("   -> Tabla vacia: ningun producto del filtro existe para este cliente (escaneo omitido).")
            else:
                print("   -> Leyendo productos existentes en SAP...")
                target_claves = {clave_material(p) for p in productos_objetivo}
                paginas_vistas = set()
                while True:
                    current_page = []
                    for row_in_screen in range(vis_rows):
                        try:
                            val = session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").text.strip()
                            current_page.append(val)
                        except Exception:
                            current_page.append("")

                    # Red de seguridad contra bucles: no repetir paginas
                    firma = tuple(current_page)
                    if firma in paginas_vistas:
                        break
                    paginas_vistas.add(firma)

                    for val in current_page:
                        if val:
                            existing_materials.add(val)
                            existing_claves.add(clave_material(val))

                    # Salida temprana: ya encontramos todos los productos objetivo
                    if target_claves and target_claves.issubset(existing_claves):
                        break

                    # Avanzar pagina con el mecanismo verificado (foco en la tabla)
                    if not avanzar_pagina(session, table_id):
                        break  # no avanzo: fin de la tabla

                print(f"   -> Se encontraron {len(existing_materials)} productos existentes.")
            
            # --- FASE 2: INSERTAR PRODUCTOS QUE NO EXISTEN ---
            # Usa la lista objetivo (filtro si existe, si no todos los productos).
            # La comparacion se hace por clave (sin ceros a la izquierda) para que
            # el formato del CSV no importe frente a como SAP muestra el material.
            nuevos_productos = [p for p in productos_objetivo if clave_material(p) not in existing_claves]
            
            inserted_any = False
            if nuevos_productos:
                print(f"   -> Insertando {len(nuevos_productos)} nuevos productos...")

                # Ir al final de la tabla (solo si hay registros previos; si la
                # tabla esta vacia la primera fila libre esta al inicio)
                if existing_materials:
                    scroll_al_final(session, table_id)
                
                def buscar_fila_vacia():
                    """Busca la primera fila vacía en la pantalla actual.
                    Retorna el índice en pantalla o -1 si no hay filas vacías."""
                    for r in range(vis_rows):
                        try:
                            val = session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{r}]").text.strip()
                            if not val:
                                return r
                        except Exception:
                            return r
                    return -1
                
                # Buscar la primera fila vacía de forma segura
                row_in_screen = buscar_fila_vacia()
                while row_in_screen == -1:
                    print("   -> Pantalla sin filas vacías. Validando y bajando más...")
                    session.findById("wnd[0]").sendVKey(0)   # Enter para validar
                    cerrar_modales(session)
                    try:
                        session.findById(f"{table_id}/ctxtMV10A-MATNR[0,0]").setFocus()
                    except Exception:
                        pass
                    session.findById("wnd[0]").sendVKey(82)  # Page Down
                    time.sleep(0.3)
                    scroll_al_final(session, table_id)
                    row_in_screen = buscar_fila_vacia()
                
                # Insertar los nuevos productos
                for matnr in nuevos_productos:
                    # Si llenamos la página actual, buscar la siguiente fila vacía
                    if row_in_screen >= vis_rows or row_in_screen == -1:
                        print("   -> Página llena. Validando y avanzando de página...")
                        session.findById("wnd[0]").sendVKey(0)   # Enter para validar
                        cerrar_modales(session)  # por si SAP abre popup al validar
                        try:
                            session.findById(f"{table_id}/ctxtMV10A-MATNR[0,0]").setFocus()
                        except Exception:
                            pass
                        session.findById("wnd[0]").sendVKey(82)  # Page Down
                        time.sleep(0.3)
                        
                        # Ir al final nuevamente
                        scroll_al_final(session, table_id)
                        
                        # Buscar fila vacía de forma segura
                        row_in_screen = buscar_fila_vacia()
                        while row_in_screen == -1:
                            print("   -> Pantalla sin filas vacías. Validando y bajando más...")
                            session.findById("wnd[0]").sendVKey(0)   # Enter
                            cerrar_modales(session)
                            try:
                                session.findById(f"{table_id}/ctxtMV10A-MATNR[0,0]").setFocus()
                            except Exception:
                                pass
                            session.findById("wnd[0]").sendVKey(82)  # Page Down
                            time.sleep(0.3)
                            scroll_al_final(session, table_id)
                            row_in_screen = buscar_fila_vacia()
                    
                    print(f"   -> Escribiendo nuevo producto {matnr} en fila visible {row_in_screen}")
                    session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").text = matnr
                    row_in_screen += 1
                    inserted_any = True
            
            if inserted_any:
                print("   -> Validando nuevos productos insertados...")
                session.findById("wnd[0]").sendVKey(0)  # Enter
                time.sleep(1)
            
            # --- FASE 3: ASIGNAR/ACTUALIZAR PLANTA A LOS PRODUCTOS DEL LISTADO ---
            # Conjunto exacto de materiales que hay que marcar con la planta:
            # los que ya existian (solo los objetivo) mas los recien insertados.
            # Se trabaja con claves (sin ceros a la izquierda) para que la
            # comparacion no dependa del formato con que SAP muestra el material.
            claves_por_marcar = ({clave_material(v) for v in existing_materials} & {clave_material(p) for p in productos_objetivo}) | {clave_material(p) for p in nuevos_productos}
            updated_products = set()
            filas_sin_planta = []

            if claves_por_marcar:
                print(f"   -> Asignando/actualizando plantas para {len(claves_por_marcar)} productos...")

                # Volver al inicio de la tabla
                scroll_al_inicio(session, table_id)

                paginas_vistas_f3 = set()
                while True:
                    current_page = []
                    for row_in_screen in range(vis_rows):
                        try:
                            val = session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").text.strip()
                            current_page.append(val)
                        except Exception:
                            current_page.append("")

                    # Red de seguridad contra bucles: no repetir paginas
                    firma = tuple(current_page)
                    if firma in paginas_vistas_f3:
                        break
                    paginas_vistas_f3.add(firma)

                    for row_in_screen, matnr in enumerate(current_page):
                        clave_fila = clave_material(matnr)
                        if not matnr or clave_fila not in claves_por_marcar or clave_fila in updated_products:
                            continue

                        print(f"   -> Asignando planta a {matnr} (fila {row_in_screen})...")

                        # Intento 1: escritura directa en la celda de planta (rapido).
                        # Se lee ANTES de escribir: si ya tiene la planta objetivo no
                        # se toca (y evita una "verificacion" vacua que siempre pasa).
                        asignada = False
                        if celda_planta_prefijo:
                            try:
                                celda = session.findById(f"{celda_planta_prefijo}[0,{row_in_screen}]")
                                if celda.text.strip() == plant:
                                    asignada = True
                                else:
                                    celda.text = plant
                                    if celda.text.strip() == plant:
                                        asignada = True
                                    else:
                                        print(f"      Celda no acepto el valor; usando via detalle...")
                            except Exception:
                                asignada = False

                        # Intento 2 (fallback): pantalla de detalle (F2) como siempre
                        if not asignada:
                            try:
                                # Re-obtener la tabla para evitar COM obsoletos
                                tabla = session.findById(table_id)

                                # Foco en el material e ir a detalle (F2)
                                session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").setFocus()
                                session.findById("wnd[0]").sendVKey(2)  # F2
                                time.sleep(0.3)

                                # Asignar la Planta
                                session.findById("wnd[0]/usr/ctxtMV10A-WERKS").text = plant
                                session.findById("wnd[0]/tbar[0]/btn[3]").press()  # F3 - Volver
                                time.sleep(0.3)
                                asignada = True
                            except Exception as e:
                                print(f"      No se pudo asignar planta a {matnr}: {e}")

                        if asignada:
                            updated_products.add(clave_fila)
                        else:
                            filas_sin_planta.append(matnr)

                    # Salida temprana: ya se marco todo lo necesario
                    if claves_por_marcar.issubset(updated_products):
                        break

                    # Avanzar pagina con el mecanismo verificado (foco en la tabla)
                    if not avanzar_pagina(session, table_id):
                        break  # no avanzo: fin de la tabla

                if filas_sin_planta:
                    print(f"   -> AVISO: no se pudo asignar planta a: {', '.join(filas_sin_planta)}")

                # Si se escribieron plantas directo en celdas, validar antes de guardar
                if updated_products and celda_planta_prefijo:
                    session.findById("wnd[0]").sendVKey(0)  # Enter para validar ediciones
                    cerrar_modales(session)
                    time.sleep(0.3)

                # Alertar si quedaron productos sin pasar por la asignacion
                faltantes = claves_por_marcar - updated_products
                if faltantes:
                    muestra = ", ".join(sorted(faltantes)[:5])
                    print(f"   -> AVISO: {len(faltantes)} productos quedaron sin asignar planta: {muestra}")

            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            cerrar_modales(session)  # por si el guardado abre popup de confirmacion
            
            # Verificacion post-guardado: SAP avisa si no hubo cambios pendientes
            try:
                sbar_txt = session.findById("wnd[0]/sbar").Text or ""
                if "no info to be saved" in sbar_txt.lower():
                    print(f"   -> AVISO: SAP reporto '{sbar_txt}' (no habia cambios pendientes que guardar).")
            except Exception:
                pass
            
            print(f" -> ¡Cliente {kunnr} guardado con éxito!")

            # Estadisticas para corridas largas
            t_cliente = time.time() - t_cliente
            promedio = (time.time() - t_inicio_corrida) / idx_cliente
            restante_min = promedio * (len(clientes) - idx_cliente) / 60
            print(f" -> Tiempo cliente: {t_cliente:.1f}s | Promedio: {promedio:.1f}s/cliente | Restante estimado: {restante_min:.0f} min")
            
        except Exception as e:
            import traceback
            print(f" -> ERROR procesando cliente {kunnr}:")
            traceback.print_exc()
            
            # Si hubo un error, cancelamos la pantalla actual (F12) varias veces para volver al inicio
            try:
                session.findById("wnd[0]").sendVKey(12) # Cancelar
                session.findById("wnd[0]").sendVKey(12) # Cancelar de nuevo por si acaso
            except:
                pass

    print("-" * 50)
    print("Proceso finalizado.")

if __name__ == "__main__":
    main()
