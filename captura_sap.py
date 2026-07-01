import win32com.client
import csv
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

def scroll_al_inicio(session):
    for _ in range(5):
        try:
            session.findById("wnd[0]").sendVKey(81)  # Page Up
            time.sleep(0.1)
        except Exception:
            break

def scroll_al_final(session):
    for _ in range(5):
        try:
            session.findById("wnd[0]").sendVKey(82)  # Page Down
            time.sleep(0.1)
        except Exception:
            break

def main():
    session = get_sap_session()
    
    # Leer datos de los archivos CSV
    clientes = leer_csv('clientes.csv')
    productos_dict = leer_csv('productos.csv')
    
    if not clientes or not productos_dict:
        print("Asegúrate de que 'clientes.csv' y 'productos.csv' tengan datos y estén en la misma carpeta.")
        return
        
    productos = [p['Material'].strip() for p in productos_dict if p['Material'].strip()]
    
    productos_filtro = leer_filtro_productos('productos_filtros.csv')
    if productos_filtro:
        print(f"Se encontraron {len(productos_filtro)} productos para aplicar como filtro en la consulta.")
    else:
        print("No hay productos de filtro o el archivo productos_filtros.csv esta vacio/ausente. Se omite el filtro.")
    
    print(f"Se encontraron {len(clientes)} clientes y {len(productos)} productos.")
    print("Iniciando procesamiento...")
    print("-" * 50)
    
    for idx_cliente, cliente in enumerate(clientes, 1):
        kunnr = cliente['Cliente'].strip()
        vkorg = cliente['OrgVentas'].strip()
        vtweg = cliente['CanalDist'].strip()
        plant = cliente['Planta'].strip()
        
        print(f"[{idx_cliente}/{len(clientes)}] Procesando Cliente: {kunnr} | Planta: {plant}")
        
        try:
            # Forzar cierre de cualquier transacción previa y abrir VD52 limpiamente
            # Usamos /nVD52 en el campo okcd para garantizar que se cierra lo anterior
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVD52"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.5)  # Esperar a que SAP cargue la pantalla de VD52
            
            # Cerrar cualquier popup/modal que aparezca (ej. "¿Guardar cambios?")
            for _ in range(3):
                try:
                    modal = session.findById("wnd[1]")
                    # Intentar presionar "No" o el botón de la izquierda para descartar
                    try:
                        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()  # Botón "No"
                    except:
                        try:
                            session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Botón genérico
                        except:
                            modal.sendVKey(12)  # Cancelar modal
                    time.sleep(0.5)
                except:
                    break  # No hay modal, continuamos
            
            # Pantalla inicial de Selección - con reintentos por si la pantalla tarda en cargar
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # Verificar que los campos requeridos son accesibles
                    session.findById("wnd[0]/usr/ctxtKUNNR").text = kunnr
                    session.findById("wnd[0]/usr/ctxtVKORG").text = vkorg
                    session.findById("wnd[0]/usr/ctxtVTWEG").text = vtweg
                    break
                except Exception as field_err:
                    if attempt < max_retries - 1:
                        print(f"   -> Reintentando acceso a pantalla (intento {attempt + 2}/{max_retries})...")
                        time.sleep(1.5)
                    else:
                        raise field_err
            
            print(f"   -> Ingresando Cliente: {kunnr}")
            print(f"   -> Ingresando Org Ventas: {vkorg}")
            print(f"   -> Ingresando Canal Dist: {vtweg}")
            
            # Limpiar TODOS los campos de texto de la pantalla excepto los tres requeridos.
            # Esto elimina cualquier filtro de "memoria de campo" de SAP (ej. Material)
            # sin importar cuál sea el ID exacto del campo ni si es rango o valor simple.
            print("   -> Limpiando campos extra de la pantalla de selección...")
            # Usamos los SUFIJOS de los IDs porque SAP puede devolver el path completo
            # (/app/con[0]/ses[0]/wnd[0]/usr/ctxtKUNNR) en lugar del path corto (wnd[0]/usr/ctxtKUNNR)
            sufijos_requeridos = ("ctxtKUNNR", "ctxtVKORG", "ctxtVTWEG")
            try:
                usr_container = session.findById("wnd[0]/usr")
                for i in range(usr_container.Children.Count):
                    try:
                        child = usr_container.Children(i)
                        child_id = child.Id
                        # Verificar si el campo es uno de los requeridos por sufijo
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
            
            # Re-confirmar los valores requeridos por si acaso
            session.findById("wnd[0]/usr/ctxtKUNNR").text = kunnr
            session.findById("wnd[0]/usr/ctxtVKORG").text = vkorg
            session.findById("wnd[0]/usr/ctxtVTWEG").text = vtweg
            
            # --- APLICAR FILTRO DE PRODUCTOS (si existe) ---
            if productos_filtro:
                print(f"   -> Aplicando filtro de {len(productos_filtro)} productos a la consulta...")
                try:
                    # Abrir el dialogo de seleccion multiple del campo Material
                    session.findById("wnd[0]/usr/btn%_MATNR_R_%_APP_%-VALU_PUSH").press()
                    time.sleep(0.5)

                    # Eliminar cualquier lista previa que haya quedado en la "memoria" del autocompletar de SAP
                    try:
                        session.findById("wnd[1]/tbar[0]/btn[16]").press()
                        time.sleep(0.3)
                    except Exception:
                        pass

                    # Cargar los materiales en el portapapeles de Windows (uno por linea)
                    clipboard_text = "\r\n".join(productos_filtro)
                    set_clipboard_text(clipboard_text)

                    # Pegar desde el portapapeles (boton Paste del dialogo de rangos)
                    session.findById("wnd[1]/tbar[0]/btn[24]").press()
                    time.sleep(0.8)

                    # Aceptar el dialogo de rangos (Copy/Continue)
                    session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    time.sleep(0.3)
                except Exception as filtro_err:
                    print(f"   -> ERROR al aplicar filtro de productos: {filtro_err}")
                    # Intentar cerrar el dialogo de rangos si quedo abierto
                    try:
                        session.findById("wnd[1]").sendVKey(12)
                    except Exception:
                        pass
            
            # Ir directamente a Ejecutar (F8) SIN presionar Enter primero.
            # Presionar Enter activa el historial de campo de SAP y puede restaurar
            # el último material buscado como filtro no deseado.
            print("   -> Ejecutando búsqueda (F8) sin pasar por Enter...")
            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                time.sleep(1.5)
            except Exception:
                # Si F8 no está disponible directamente, usar Enter como fallback
                print("   -> F8 no disponible, usando Enter como fallback...")
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.5)
                try:
                    session.findById("wnd[0]/tbar[1]/btn[8]").press()
                    time.sleep(1)
                except Exception:
                    pass
            
            table_id = "wnd[0]/usr/tblSAPMV10ATC_CU_MA"
            
            # Leer el número de filas visibles una sola vez al inicio para evitar errores COM tras modificar celdas
            tabla = session.findById(table_id)
            try:
                vis_rows = tabla.visibleRowCount
            except Exception:
                vis_rows = 15  # Fallback seguro
            
            # --- FASE 1: LEER PRODUCTOS EXISTENTES EN SAP (ESCANEO POR TECLADO) ---
            print("   -> Leyendo productos existentes en SAP...")
            scroll_al_inicio(session)
            
            existing_materials = set()
            last_page_signature = None
            
            while True:
                tabla = session.findById(table_id)
                current_page = []
                for row_in_screen in range(vis_rows):
                    try:
                        val = session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").text.strip()
                        current_page.append(val)
                    except Exception:
                        current_page.append("")
                
                if current_page == last_page_signature:
                    break
                last_page_signature = current_page
                
                read_any = False
                for val in current_page:
                    if val:
                        existing_materials.add(val)
                        read_any = True
                        
                if not read_any:
                    break
                    
                # Avanzar página
                try:
                    session.findById("wnd[0]").sendVKey(82)  # Page Down
                    time.sleep(0.3)
                except Exception:
                    break
            
            print(f"   -> Se encontraron {len(existing_materials)} productos existentes.")
            
            # --- FASE 2: INSERTAR PRODUCTOS QUE NO EXISTEN ---
            nuevos_productos = [p for p in productos if p not in existing_materials]
            
            inserted_any = False
            if nuevos_productos:
                print(f"   -> Insertando {len(nuevos_productos)} nuevos productos...")
                
                # Ir al final de la tabla
                scroll_al_final(session)
                
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
                    session.findById("wnd[0]").sendVKey(82)  # Page Down
                    time.sleep(0.5)
                    scroll_al_final(session)
                    row_in_screen = buscar_fila_vacia()
                
                # Insertar los nuevos productos
                for matnr in nuevos_productos:
                    # Si llenamos la página actual, buscar la siguiente fila vacía
                    if row_in_screen >= vis_rows or row_in_screen == -1:
                        print("   -> Página llena. Validando y avanzando de página...")
                        session.findById("wnd[0]").sendVKey(0)   # Enter para validar
                        session.findById("wnd[0]").sendVKey(82)  # Page Down
                        time.sleep(0.5)
                        
                        # Ir al final nuevamente
                        scroll_al_final(session)
                        
                        # Buscar fila vacía de forma segura
                        row_in_screen = buscar_fila_vacia()
                        while row_in_screen == -1:
                            print("   -> Pantalla sin filas vacías. Validando y bajando más...")
                            session.findById("wnd[0]").sendVKey(0)   # Enter
                            session.findById("wnd[0]").sendVKey(82)  # Page Down
                            time.sleep(0.5)
                            scroll_al_final(session)
                            row_in_screen = buscar_fila_vacia()
                    
                    print(f"   -> Escribiendo nuevo producto {matnr} en fila visible {row_in_screen}")
                    session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").text = matnr
                    row_in_screen += 1
                    inserted_any = True
            
            if inserted_any:
                print("   -> Validando nuevos productos insertados...")
                session.findById("wnd[0]").sendVKey(0)  # Enter
                time.sleep(1)
            
            # --- FASE 3: ASIGNAR/ACTUALIZAR PLANTA A LOS PRODUCTOS DE PRODUCTOS.CSV (UNIFICADO) ---
            print("   -> Asignando/actualizando plantas para los productos del listado...")
            
            # Volver al inicio de la tabla
            scroll_al_inicio(session)
            
            updated_products = set()
            last_page_signature = None
            
            while True:
                tabla = session.findById(table_id)
                current_page = []
                for row_in_screen in range(vis_rows):
                    try:
                        val = session.findById(f"{table_id}/ctxtMV10A-MATNR[0,{row_in_screen}]").text.strip()
                        current_page.append(val)
                    except Exception:
                        current_page.append("")
                        
                if current_page == last_page_signature:
                    break
                last_page_signature = current_page
                
                read_any = False
                for row_in_screen, matnr in enumerate(current_page):
                    if matnr:
                        read_any = True
                        if matnr in productos and matnr not in updated_products:
                            print(f"   -> Actualizando planta para {matnr} en fila visible {row_in_screen}")
                            
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
                            
                            updated_products.add(matnr)
                            
                if not read_any:
                    break
                    
                # Avanzar página
                try:
                    session.findById("wnd[0]").sendVKey(82)  # Page Down
                    time.sleep(0.3)
                except Exception:
                    break
                
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            print(f" -> ¡Cliente {kunnr} guardado con éxito!")
            
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
