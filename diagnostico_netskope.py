import os
import platform
import socket
import ssl
import subprocess
import sys
import urllib.parse
import urllib.request

try:
    import certifi
except ImportError:
    certifi = None

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import winreg
except ImportError:
    winreg = None

HOST = "api.mapbox.com"
PORT = 443
TOKEN = os.getenv("MAPBOX_ACCESS_TOKEN", "")


def encabezado(titulo):
    print("=" * 60)
    print(titulo)
    print("=" * 60)


def prueba_https(host, puerto, contexto, etiqueta):
    try:
        req = urllib.request.Request(f"https://{host}:{puerto}/")
        with urllib.request.urlopen(req, timeout=10, context=contexto) as resp:
            print(f"   [OK] {etiqueta} -> HTTP {resp.status}")
            return True
    except Exception as e:
        print(f"   [ERROR] {etiqueta} -> {type(e).__name__}: {e}")
        return False


def contexto_ambos():
    ctx = ssl.create_default_context()
    if certifi:
        try:
            ctx.load_verify_locations(cafile=certifi.where())
        except ssl.SSLError:
            pass
    ctx.load_default_certs(ssl.Purpose.SERVER_AUTH)
    return ctx


def contexto_sistema():
    ctx = ssl.create_default_context()
    ctx.load_default_certs(ssl.Purpose.SERVER_AUTH)
    return ctx


def contexto_certifi():
    ctx = ssl.create_default_context(cafile=certifi.where() if certifi else None)
    return ctx


def contexto_sin_verificar():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def leer_proxy_windows():
    if not winreg:
        return None
    def leer_valor(clave, nombre, defecto=""):
        try:
            valor, _ = winreg.QueryValueEx(clave, nombre)
            return valor
        except OSError:
            return defecto

    try:
        clave = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        )
        habilitado = leer_valor(clave, "ProxyEnable", 0)
        servidor = leer_valor(clave, "ProxyServer", "")
        pac = leer_valor(clave, "AutoConfigURL", "")
        winreg.CloseKey(clave)
        return {"habilitado": bool(habilitado), "servidor": servidor, "pac": pac}
    except OSError as e:
        return {"error": str(e)}


def issuer_del_servidor(host, puerto):
    try:
        import tempfile
        pem = ssl.get_server_certificate((host, puerto))
        cert = {}
        if hasattr(ssl._ssl, "_test_decode_cert"):
            ruta_tmp = os.path.join(tempfile.gettempdir(), "mapbox_cert.pem")
            with open(ruta_tmp, "w", encoding="utf-8") as f:
                f.write(pem)
            cert = ssl._ssl._test_decode_cert(ruta_tmp)
        emisor = dict(
            (k[0][0], k[0][1]) for k in cert.get("issuer", [])
        ) if cert.get("issuer") else {}
        sujeto = dict(
            (k[0][0], k[0][1]) for k in cert.get("subject", [])
        ) if cert.get("subject") else {}
        return {
            "emisor_cn": emisor.get("commonName", "?"),
            "emisor_org": emisor.get("organizationName", "?"),
            "sujeto_cn": sujeto.get("commonName", "?"),
        }
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}


def token_enmascarado(token):
    if not token:
        return "(vacio)"
    if token.startswith("pk.") and len(token) > 12:
        return token[:10] + "..." + token[-4:]
    return "(presente, formato no pk.)"


def hay_netskope_local():
    hostname = platform.node()
    netskope_en_proceso = False
    try:
        salida = subprocess.run(
            ["tasklist", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=15,
        ).stdout.lower()
        netskope_en_proceso = "netskope" in salida
    except Exception:
        pass
    return hostname, netskope_en_proceso


def main():
    encabezado("DIAGNOSTICO DE CONEXION A MAPBOX / NETSKOPE")

    print("Fecha:", __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    print("Maquina:", platform.node())
    print("Sistema:", platform.system(), platform.release(), platform.machine())
    print("Python:", platform.python_version(), sys.executable)

    hostname, proceso_netskope = hay_netskope_local()
    print("Proceso Netskope en ejecucion:", "SI" if proceso_netskope else "NO / no detectable")

    print()
    encabezado("1. CERTIFICADOS (certifi)")
    if certifi:
        ruta = certifi.where()
        print("certifi instalado:", "SI")
        print("Paquete de CAs:", ruta)
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                contenido = f.read()
            cantidad = contenido.count("BEGIN CERTIFICATE")
            print("Cantidad de CAs en el bundle:", cantidad)
        except Exception as e:
            print("No se pudo leer el bundle:", e)
    else:
        print("certifi NO instalado. Ejecuta: py -m pip install certifi")

    print()
    encabezado("2. TOKEN MAPBOX")
    print("Token configurado:", token_enmascarado(TOKEN))

    print()
    encabezado("3. PROXY (variables de entorno)")
    for var in ["HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY", "http_proxy", "https_proxy"]:
        valor = os.getenv(var, "")
        print(f"   {var} = {valor if valor else '(vacio)'}")

    print()
    encabezado("4. PROXY (configuracion de Windows / Netskope)")
    config = leer_proxy_windows()
    if config:
        if "error" in config:
            print("   No se pudo leer el registro:", config["error"])
        else:
            print("   Proxy habilitado:", config["habilitado"])
            print("   Servidor:", config["servidor"] or "(ninguno)")
            print("   PAC / AutoConfig:", config["pac"] or "(ninguno)")
    else:
        print("   winreg no disponible (no es Windows).")

    print()
    encabezado("5. DNS")
    try:
        ip = socket.gethostbyname(HOST)
        print(f"   [OK] {HOST} -> {ip}")
    except Exception as e:
        print(f"   [ERROR] Resolucion DNS fallo: {e}")

    print()
    encabezado("6. PUERTO TCP 443")
    try:
        with socket.create_connection((HOST, PORT), timeout=10):
            print(f"   [OK] Conexion TCP a {HOST}:{PORT} exitosa")
    except Exception as e:
        print(f"   [ERROR] No hay conexion TCP: {e}")

    print()
    encabezado("7. PRUEBAS TLS / HTTPS A MAPBOX")
    prueba_https(HOST, PORT, contexto_sin_verificar(), "Conexion sin verificar (control)")
    prueba_https(HOST, PORT, contexto_sistema(), "Confianza SOLO almacen Windows")
    prueba_https(HOST, PORT, contexto_certifi(), "Confianza SOLO certifi")
    prueba_https(HOST, PORT, contexto_ambos(), "Confianza Windows + certifi (fix actual)")

    print()
    encabezado("8. CERTIFICADO PRESENTADO POR EL SERVIDOR")
    info = issuer_del_servidor(HOST, PORT)
    if "error" in info:
        print("   No se pudo obtener:", info["error"])
    else:
        print("   Emisor (CN):", info["emisor_cn"])
        print("   Emisor (Org):", info["emisor_org"])
        print("   Sujeto (CN):", info["sujeto_cn"])
        print()
        print("   Si el emisor contiene 'Netskope', 'Zscaler', 'Fortinet',")
        print("   'SSL Interception' o el nombre de la empresa, confirma que")
        print("   el trafico pasa por inspeccion SSL de Netskope.")

    print()
    encabezado("9. PRUEBA DE GEOCODIFICACION REAL")
    if TOKEN and TOKEN != "tu_token_aqui":
        direccion = "8700 SW 8th St, MIAMI, FL, 33174"
        url_dir = urllib.parse.quote(direccion)
        url = f"https://{HOST}/geocoding/v5/mapbox.places/{url_dir}.json?access_token={TOKEN}&limit=1"
        try:
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req, timeout=15, context=contexto_ambos()) as resp:
                data = resp.read().decode("utf-8")
                print("   [OK] Geocodificacion con el fix actual exitosa.")
                print("   Respuesta (primeros 200 chars):", data[:200])
        except Exception as e:
            print(f"   [ERROR] Geocodificacion fallo: {type(e).__name__}: {e}")
    else:
        print("   Token no configurado, se omite la prueba.")

    print()
    encabezado("RESULTADO PARA EL TICKET DE IT")
    print("Copia todo lo anterior y adjuntalo al ticket con esta info:")
    print()
    print(" 1. Si la seccion 7 muestra que la conexion 'sin verificar' SI funciona")
    print("    pero 'certifi' falla, el problema es de confianza del certificado.")
    print(" 2. Copia el 'Emisor' de la seccion 8: es el certificado que Netskope")
    print("    esta inyectando. El departamento de IT debe poder identificarlo.")
    print(" 3. Pide a IT que verifique en Netskope que la app api.mapbox.com")
    print("    (o el host completo) este en la lista blanca / permitida sin")
    print("    inspeccion SSL, O que el CA de inspeccion este instalado en la")
    print("    maquina en 'Certificados del equipo > Entidades de certificacion")
    print("    raiz de confianza'.")
    print()
    print("Si seccion 5 o 6 fallan, es un bloqueo de red/DNS de Netskope y el")
    print("ticket debe pedir que se permita el trafico hacia api.mapbox.com:443.")


if __name__ == "__main__":
    main()
