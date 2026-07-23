import csv
import json
import os
import sys
import time
import urllib.parse
import urllib.request

try:
    from shapely.geometry import Point, Polygon
except ImportError:
    print("================================================================")
    print("ERROR: Falta la libreria 'shapely'")
    print("Instala con:  pip install shapely")
    print("================================================================")
    sys.exit(1)

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

CLIENTES_CSV = "clientes.csv"
CELULAS_JSON = "celulas_geograficas.json"
CACHE_JSON = "geocode_cache.json"
MAPBOX_TOKEN = os.getenv("MAPBOX_ACCESS_TOKEN", "")


def cargar_celulas(ruta):
    try:
        with open(ruta, mode="r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: No se encontro el archivo: {ruta}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"ERROR: JSON invalido en {ruta}: {e}")
        sys.exit(1)

    celulas = []
    for c in data.get("celulas", []):
        nombre = c.get("nombre", "Sin nombre")
        planta = c.get("planta", "")
        coords = c.get("poligono", [])
        if len(coords) < 3:
            print(f"ADVERTENCia: Celula '{nombre}' tiene menos de 3 puntos, se omite.")
            continue
        poligono = Polygon([(lng, lat) for lat, lng in coords])
        celulas.append({"nombre": nombre, "planta": planta, "poligono": poligono})
    return celulas


def cargar_cache(ruta):
    if os.path.isfile(ruta):
        try:
            with open(ruta, mode="r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def guardar_cache(ruta, cache):
    try:
        with open(ruta, mode="w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"ADVERTENCia: No se pudo guardar el cache: {e}")


def geocodificar_mapbox(direccion, cache, token):
    if direccion in cache:
        return cache[direccion]

    if not token:
        print("ERROR: No se encontro MAPBOX_ACCESS_TOKEN.")
        print("Crea un archivo .env con: MAPBOX_ACCESS_TOKEN=tu_token")
        sys.exit(1)

    url_direccion = urllib.parse.quote(direccion)
    url = f"https://api.mapbox.com/geocoding/v5/mapbox.places/{url_direccion}.json?access_token={token}&limit=1"

    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"   -> Error de geocodificacion: {e}")
        return None

    features = data.get("features", [])
    if not features:
        return None

    coords = features[0].get("center", [])
    if len(coords) != 2:
        return None

    resultado = {"lng": coords[0], "lat": coords[1]}
    cache[direccion] = resultado
    return resultado


def asignar_planta(lat, lng, celulas):
    punto = Point(lng, lat)
    for cel in celulas:
        if punto.within(cel["poligono"]):
            return cel["nombre"], cel["planta"]
    return None, None


def leer_clientes(ruta):
    clientes = []
    try:
        with open(ruta, mode="r", encoding="utf-8-sig") as f:
            lector = csv.DictReader(f)
            for fila in lector:
                clientes.append(fila)
    except FileNotFoundError:
        print(f"ERROR: No se encontro el archivo: {ruta}")
        sys.exit(1)
    return clientes


def guardar_clientes(ruta, clientes, nombres_columnas):
    try:
        with open(ruta, mode="w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=nombres_columnas)
            writer.writeheader()
            for c in clientes:
                writer.writerow(c)
    except Exception as e:
        print(f"ERROR al escribir {ruta}: {e}")
        sys.exit(1)


def main():
    if not MAPBOX_TOKEN or MAPBOX_TOKEN == "tu_token_aqui":
        print("================================================================")
        print("ERROR: Token de Mapbox no configurado.")
        print("Edita el archivo .env y coloca tu token real:")
        print("  MAPBOX_ACCESS_TOKEN=pk.eyJ1Ijoie...")
        print("================================================================")
        sys.exit(1)

    print("Cargando configuracion de celulas...")
    celulas = cargar_celulas(CELULAS_JSON)
    print(f"Se cargaron {len(celulas)} celulas:")
    for c in celulas:
        print(f"   - {c['nombre']} -> Planta: {c['planta']}")

    print("\nCargando clientes...")
    clientes = leer_clientes(CLIENTES_CSV)
    if not clientes:
        print("No hay clientes en el archivo. Nada que procesar.")
        return
    print(f"Se encontraron {len(clientes)} clientes.\n")

    print("Cargando cache de geocodificacion...")
    cache = cargar_cache(CACHE_JSON)
    if cache:
        print(f"Cache con {len(cache)} direcciones pre-codificadas.\n")

    columnas = list(clientes[0].keys())

    print("-" * 50)
    for i, cliente in enumerate(clientes, 1):
        kunnr = cliente.get("Cliente", "").strip()
        street = cliente.get("Street", "").strip()
        city = cliente.get("City", "").strip()
        state = cliente.get("State", "").strip()
        zip_code = cliente.get("Zip", "").strip()

        partes = [p for p in [street, city, state, zip_code] if p]
        direccion_completa = ", ".join(partes)

        if not direccion_completa:
            print(f"[{i}/{len(clientes)}] Cliente {kunnr}: sin direccion, saltando.")
            continue

        print(f"[{i}/{len(clientes)}] Cliente {kunnr}: {direccion_completa}")

        coords = geocodificar_mapbox(direccion_completa, cache, MAPBOX_TOKEN)
        if not coords:
            print(f"   -> No se pudo geocodificar. Planta sin asignar.")
            continue

        lat = coords["lat"]
        lng = coords["lng"]
        print(f"   -> Coordenadas: lat={lat:.6f}, lng={lng:.6f}")

        nombre_celula, planta = asignar_planta(lat, lng, celulas)
        if planta:
            cliente["Planta"] = planta
            print(f"   -> Celula: {nombre_celula} | Planta asignada: {planta}")
        else:
            print(f"   -> ADVERTENCia: No cayo en ninguna celula. Planta sin asignar.")

        time.sleep(0.2)

    print("-" * 50)

    print("\nGuardando clientes.csv actualizado...")
    guardar_clientes(CLIENTES_CSV, clientes, columnas)

    print("Guardando cache de geocodificacion...")
    guardar_cache(CACHE_JSON, cache)

    print("\nProceso finalizado.")


if __name__ == "__main__":
    main()