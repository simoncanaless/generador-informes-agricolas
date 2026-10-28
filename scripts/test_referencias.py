"""Las tres fuentes de ODEPA, juntas y consultables por especie.

Lo que se comprueba es lo que puede enganar:

- Que el nombre del CATASTRO encuentre el dato, aunque cada fuente use otro
  vocabulario: el cliente escribe «Avellano europeo», Aduanas dice «Avellana»
  y CIREN dice «Avellano».
- Que cada numero llegue con su reparo pegado. Un FOB sin su advertencia al
  lado es el error que §10.2 previene.
- Que el precio de mercado interno sea el MAYORISTA. El de supermercado es
  seis veces mas alto y no es lo que recibe el productor.
- Que lo que falta se diga en vez de fingirse.

Uso: python scripts/test_referencias.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import ajustes                              # noqa: E402
from services.referencias import (                               # noqa: E402
    ANIOS_DE_LA_DECADA, TEMPORADAS_DEL_RESUMEN, cargar_referencias,
    resumen_de_serie,
)

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def main():
    print("=" * 72)
    print("REFERENCIAS DE ODEPA (FOB + mercado interno + superficie)")
    print("=" * 72)

    biblioteca = cargar_referencias(ajustes.raiz / "config")
    fuentes = biblioteca.fuentes

    print("\n[1] Las tres fuentes estan cargadas, y se dice cual falta")
    for nombre in ("fob", "mercado_interno", "superficie"):
        comprobar(fuentes[nombre]["cargado"],
                  "«{}» disponible".format(nombre),
                  "falta su archivo en config/")
    comprobar(len(biblioteca.especies) > 50,
              "hay especies consultables",
              "{}".format(len(biblioteca.especies)))

    print("\n[2] El nombre del catastro cruza los tres vocabularios")
    # El caso que motivo todo: el catastro de Cliente A dice «Avellano
    # europeo», Aduanas «Avellana» y CIREN «Avellano».
    r = biblioteca.de("Avellano europeo")
    comprobar(r.fob is not None, "encuentra el FOB («Avellana»)")
    comprobar(r.superficie is not None, "y la superficie («Avellano»)")
    n = biblioteca.de("Nogal")
    comprobar(n.fob is not None and "Nuez" in n.fob.nombre,
              "«Nogal» encuentra «Nuez», que no comparte ni la raiz",
              n.fob.nombre if n.fob else "no lo encontro")
    comprobar(n.superficie is not None and "Nogal" in n.superficie.nombre,
              "y la superficie sigue siendo la del nogal")

    print("\n[2.bis] Los nombres REALES del catalogo de curvas")
    # La leccion mas cara de este frente: probe con los nombres que se me
    # ocurrieron -«Avellana», «Nogal», «Palto»- y daba todo verde, mientras
    # CINCO de las trece curvas reales no encontraban nada. El catalogo dice
    # «Cerezos», «Nogales» y «Limoneros», en plural, y «Arandano Secoya» con
    # la variedad pegada. Hay que probar con lo que hay, no con lo que uno
    # elige.
    for escrito in ("Cerezos", "Nogales", "Limoneros", "Arandano Secoya",
                    "Palto Clonal", "Naranjo Lane Late", "Ciruelo",
                    "Avellano Giffoni", "Manzano", "Peral", "Duraznero",
                    "Mandarino", "Almendro", "Kiwi"):
        # Variable propia: `r` es el avellano y lo siguen usando las secciones
        # de mas abajo. Pisarla hizo que la seccion [4] midiera el kiwi y
        # afirmara que el avellano tiene 7.588 ha.
        del_catalogo = biblioteca.de(escrito)
        comprobar(del_catalogo.fob is not None,
                  "«{}» encuentra su FOB".format(escrito),
                  "sin FOB no hay ningun boton que apretar")

    print("\n[2.ter] Ninguna especie se pierde entre los tres vocabularios")
    # LA comprobacion que faltaba, y que habria atrapado lo del ciruelo: cada
    # nombre que EXISTE en una fuente tiene que encontrarse desde esa fuente.
    # Probar con los nombres que uno elige es probar el emparejador contra si
    # mismo (§16.a.quadragies.quinquies).
    import json
    import io as _io
    fuentes_de = {}
    for archivo, marca, rama in (
            ("precios_fob_odepa.json", "fob", "grupos"),
            ("precios_mercado_interno.json", "interno", "productos"),
            ("catastro_ciren.json", "superficie", "especies")):
        with _io.open(str(ajustes.raiz / "config" / archivo),
                      encoding="utf-8") as fh:
            for registro in json.load(fh)[rama].values():
                fuentes_de.setdefault(registro["nombre"], set()).add(marca)

    perdidos = []
    for nombre, marcas in sorted(fuentes_de.items()):
        encontrado = biblioteca.de(nombre)
        for marca in marcas:
            if getattr(encontrado, marca) is None:
                perdidos.append("{} ({})".format(nombre, marca))
    comprobar(not perdidos,
              "los {} nombres de las tres fuentes se encuentran".format(
                  len(fuentes_de)),
              "; ".join(perdidos[:6]))

    print("\n[2.quater] El arbol y la fruta son la misma especie")
    # El caso que reporto el usuario: Aduanas nombra el ARBOL («Ciruelo») y el
    # boletin de mayoristas la FRUTA («Ciruela»). El puente iba en un solo
    # sentido y con un solo valor, asi que el ciruelo no encontraba su precio
    # de mercado interno.
    for escrito in ("Ciruelo", "Ciruela", "Ciruelos"):
        c = biblioteca.de(escrito)
        comprobar(c.fob is not None and c.interno is not None
                  and c.superficie is not None,
                  "«{}» cruza las tres fuentes".format(escrito),
                  "FOB={} interno={} CIREN={}".format(
                      c.fob and c.fob.nombre, c.interno and c.interno.nombre,
                      c.superficie and c.superficie.nombre))

    # Y no es solo el ciruelo: los dos sentidos, con nombres que no comparten
    # ni la raiz.
    for arbol, fruta in (("Nogal", "Nuez"), ("Palto", "Palta"),
                         ("Limonero", "Limón"), ("Duraznero", "Durazno")):
        uno, otro = biblioteca.de(arbol), biblioteca.de(fruta)
        comprobar(uno.fob and otro.fob and uno.fob.nombre == otro.fob.nombre,
                  "«{}» y «{}» dan el mismo FOB".format(arbol, fruta),
                  "{} vs {}".format(uno.fob and uno.fob.nombre,
                                    otro.fob and otro.fob.nombre))

    # Fruta FRESCA, tambien en el catastro: CIREN separa el duraznero de
    # consumo fresco del tipo conservero, y el agrónomo pidio el fresco. Va
    # DECLARADO, no por el orden alfabetico del archivo.
    durazno = biblioteca.de("Duraznero")
    comprobar(durazno.superficie is not None
              and "Fresco" in durazno.superficie.nombre,
              "el duraznero apunta al de CONSUMO FRESCO, no al conservero",
              durazno.superficie.nombre if durazno.superficie else "nada")

    print("\n[2.quinquies] Cuando no hay, se PROPONE")
    # Pedido de el usuario: «que le diga ¿quieres usar la especie tal de la base de
    # datos de ODEPA?». Se propone, no se decide.
    for escrito, esperado in (("kivi", "Kiwi"),          # error de tipeo
                              ("nogal", "Nuez"),         # sinonimo declarado
                              ("ciruel", "Ciruela")):    # a medio escribir
        propuestas = biblioteca.parecidas(escrito)
        comprobar(esperado in propuestas,
                  "«{}» propone «{}»".format(escrito, esperado),
                  ", ".join(propuestas) or "(nada)")
    comprobar(not biblioteca.parecidas("qwerty"),
              "y no propone nada cuando no se parece a nada")
    comprobar(len(biblioteca.parecidas("pera")) <= 6,
              "la lista es corta: es para elegir de un vistazo")

    print("\n[3] Cada numero trae su reparo")
    comprobar(r.fob is not None and not r.fob.listo_para_el_flujo,
              "el FOB no esta listo para la fila del flujo")
    comprobar(r.fob_es_mezcla and bool(r.fob.advertencia),
              "el avellano avisa que mezcla con y sin cascara")
    comprobar(len(r.fob_codigos) >= 2, "y trae el desglose para elegir",
              "{} codigos".format(len(r.fob_codigos)))

    palto = biblioteca.de("Palto")
    comprobar(palto.interno is not None, "el palto tiene precio interno")
    if palto.interno:
        comprobar(palto.interno.punto_de_venta == "Mercado Mayorista",
                  "y es el MAYORISTA, no el de supermercado",
                  palto.interno.punto_de_venta)
        retail = palto.interno.otros_puntos.get("Supermercado")
        comprobar(retail and retail > palto.interno.clp_kg,
                  "el de supermercado se guarda aparte y es mas alto",
                  "{} vs {}".format(palto.interno.clp_kg, retail))
        # El tipo de cambio es del INFORME, no del dato: por eso se convierte
        # al consultar y no al importar.
        comprobar(palto.interno.usd_kg(941) is not None
                  and palto.interno.usd_kg(None) is None,
                  "se convierte a USD con el cambio que declara el tasador")

    print("\n[4] La superficie sirve para situar al cliente")
    s = r.superficie
    comprobar(s is not None and s.hectareas > 10000,
              "el avellano tiene superficie nacional",
              "{:,.0f} ha".format(s.hectareas) if s else "")
    if s:
        en_biobio = s.en_la_region("Biobío")
        comprobar(en_biobio and en_biobio > 0,
                  "y desglosada por region", "Biobío: {}".format(en_biobio))
        # La frase que el banco entiende: 760 ha de Cliente A sobre las de
        # la region.
        peso = s.peso_del_cliente(760.0, "Biobío")
        comprobar(peso and 0.05 < peso < 0.9,
                  "dice que fraccion de la region es el cliente",
                  "{:.1%}".format(peso) if peso else "no lo calculo")
        comprobar(bool(s.variedades),
                  "y las variedades que se plantan en Chile",
                  "{} listadas".format(len(s.variedades)))

    print("\n[4.bis] El rango reciente: minimo, maximo y promedio")
    # Lo pidio el agrónomo. Un FOB suelto no dice si la temporada fue normal o
    # excepcional; contra las ultimas seis, si.
    resumen = r.resumen_fob()
    comprobar(resumen is not None, "el avellano trae su resumen de FOB")
    if resumen:
        comprobar(resumen.n == TEMPORADAS_DEL_RESUMEN,
                  "resume la ultima temporada y cinco hacia atras",
                  "{} temporadas".format(resumen.n))
        # La ventana es la del FINAL de la serie, no la del principio: el
        # error facil es recortar con [:6] y resumir 2016-2021.
        completa = sorted(r.fob_serie.items(), key=lambda kv: kv[0][-4:])
        ultimas = [v for _k, v in completa][-TEMPORADAS_DEL_RESUMEN:]
        comprobar(resumen.hasta == completa[-1][0],
                  "y termina en la temporada mas reciente",
                  "{} vs {}".format(resumen.hasta, completa[-1][0]))
        comprobar(resumen.minimo <= resumen.promedio <= resumen.maximo,
                  "minimo <= promedio <= maximo",
                  "{} / {} / {}".format(resumen.minimo, resumen.promedio,
                                        resumen.maximo))
        # El promedio se recalcula aparte: dos caminos dando el mismo numero
        # es el contraste que este proyecto usa para todo.
        aparte = round(sum(ultimas) / len(ultimas), 4)
        comprobar(abs(resumen.promedio - aparte) < 1e-9,
                  "el promedio coincide con el calculado aparte",
                  "{} vs {}".format(resumen.promedio, aparte))
        comprobar(len(resumen.puntos) == resumen.n
                  and resumen.puntos[-1]["etiqueta"] == resumen.hasta,
                  "y viajan los puntos de esa ventana, que son los que se "
                  "dibujan")
        comprobar(not resumen.parcial and resumen.de_cuantos > resumen.n,
                  "la serie completa es mas larga y se dice cuanto",
                  "{} de {}".format(resumen.n, resumen.de_cuantos))

    print("\n[4.ter] La superficie, en la ultima decada")
    decada = s.decada() if s else None
    comprobar(decada is not None, "el avellano trae su decada")
    if decada:
        comprobar(decada.n == ANIOS_DE_LA_DECADA, "son diez anios",
                  "{}".format(decada.n))
        anios = sorted(int(a) for a in s.serie_nacional_ha)
        comprobar(decada.hasta == str(anios[-1]) and int(decada.desde) ==
                  anios[-ANIOS_DE_LA_DECADA],
                  "y son los diez ULTIMOS",
                  "{} a {}".format(decada.desde, decada.hasta))
        crecimiento = round((decada.ultimo - decada.primero) / decada.primero, 4)
        comprobar(abs(decada.variacion - crecimiento) < 1e-9,
                  "la variacion es de punta a punta de la ventana",
                  "{:.0%}".format(decada.variacion))

    print("\n[4.quater] Una serie corta no miente sobre su tamano")
    corta = resumen_de_serie({"2023": 10.0, "2024": 12.0}, 6)
    comprobar(corta is not None and corta.n == 2,
              "con dos datos resume dos, no seis")
    comprobar(corta.parcial,
              "y queda marcada como parcial, para poder rotularla bien")
    comprobar(resumen_de_serie({}, 6) is None,
              "sin serie no se inventa un rango")
    uno = resumen_de_serie({"2024": 7.0}, 6)
    comprobar(uno is not None and uno.minimo == uno.maximo == uno.promedio,
              "con un solo dato el rango es ese dato, no un intervalo")

    print("\n[4.quinquies] El precio que le queda al productor")
    # Lo pidio el agrónomo: el mayorista NO es lo que recibe el agricultor.
    # Entre medio hay flete, comision y merma.
    productor = palto.interno.al_productor() if palto.interno else None
    comprobar(productor is not None, "el palto trae su precio al productor")
    if productor:
        comprobar(productor["clp_kg_bajo"] < productor["clp_kg_alto"]
                  < palto.interno.clp_kg,
                  "es un RANGO, y siempre por debajo del mayorista",
                  "{:,.0f} a {:,.0f} de {:,.0f}".format(
                      productor["clp_kg_bajo"], productor["clp_kg_alto"],
                      palto.interno.clp_kg))
        esperado = round(palto.interno.clp_kg * 0.5, 2)
        comprobar(abs(productor["clp_kg_bajo"] - esperado) < 0.01,
                  "el extremo bajo es el 50% del mayorista",
                  "{} vs {}".format(productor["clp_kg_bajo"], esperado))
        comprobar("No es un dato publicado" in productor["criterio"],
                  "y dice que es una estimacion, no un dato de ODEPA")
    # Una especie sin precio mayorista no puede tener precio al productor:
    # inventarlo seria peor que no tenerlo.
    comprobar(r.interno is None or r.interno.al_productor() is None,
              "sin precio mayorista no se inventa uno al productor")

    print("\n[4.sexies] De donde sale el precio: fruta FRESCA")
    # el agrónomo: los datos tienen que salir de la fruta fresca, no de conserva.
    # Ya era asi, pero no se veia en ninguna parte.
    for escrito, esperado in (("Cerezos", "Fresca"), ("Nogales", "Frutos secos")):
        hallado = biblioteca.de(escrito)
        comprobar(hallado.fob_procesamiento == esperado,
                  "«{}» dice que su precio es de «{}»".format(escrito, esperado),
                  str(hallado.fob_procesamiento))

    print("\n[5] Lo que no hay, se dice")
    # El avellano no se vende en mercados mayoristas: es un fruto seco que va
    # a la industria. Que no tenga precio interno es correcto, no un fallo.
    comprobar(r.interno is None or r.interno.clp_kg is None,
              "el avellano no tiene precio de mercado mayorista, y no se "
              "inventa uno")
    inventada = biblioteca.de("Arbol que no existe")
    comprobar(not inventada.hay_algo,
              "una especie inexistente no devuelve nada")

    print("\n[6] Lo que se ofrece mientras se escribe")
    # El campo Especie de las curvas es un buscador: pide 8 sugerencias y las
    # espera ordenadas como en Mercado Libre. El orden lo decide el servidor
    # (una sola copia de la regla), asi que se prueba aca.
    ocho = biblioteca.parecidas("a", 8)
    comprobar(len(ocho) == 8, "«a» ofrece 8 y no 6", str(len(ocho)))
    comprobar(len(biblioteca.parecidas("a")) == 6,
              "sin pedirlo, el «¿es alguna de estas?» de ODEPA sigue en 6")

    ch = biblioteca.parecidas("ch", 8)
    comprobar(ch[:2] == ["Chirimoya", "Chirimoyo"],
              "«ch» pone primero lo que EMPIEZA con «ch»", str(ch))
    comprobar("Pistacho" in ch and ch.index("Pistacho") > 1,
              "y despues lo que lo CONTIENE", str(ch))
    comprobar(biblioteca.parecidas("kivi", 8) == ["Kiwi"],
              "un error de tecleo («kivi») ofrece Kiwi")
    comprobar("Nuez" in biblioteca.parecidas("nogal", 8),
              "un sinonimo («nogal») ofrece Nuez aunque no comparta letras")
    comprobar(biblioteca.parecidas("xyzq", 8) == [],
              "lo que no se parece a nada no inventa sugerencias")
    comprobar(biblioteca.parecidas("CHIRIM", 8)[:1] == ["Chirimoya"]
              and biblioteca.parecidas("Níspero", 8)[:1] == ["Níspero"]
              and biblioteca.parecidas("nispero", 8)[:1] == ["Níspero"],
              "no distingue mayusculas ni tildes")

    # Y que el endpoint deje pasar el numero: sin esto el navegador pide 8 y
    # recibe 6 sin ningun error.
    import asyncio
    from api.routers.referencias import indice
    via_web = asyncio.run(indice(parecidas="a", cuantas=8))
    comprobar(len(via_web["parecidas"]) == 8,
              "el endpoint respeta `cuantas` (el navegador pide 8)",
              str(len(via_web["parecidas"])))
    comprobar(asyncio.run(indice(parecidas=None, cuantas=8))["parecidas"] == [],
              "y sin `parecidas` no devuelve nada")

    print("\n" + "=" * 72)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(main())
