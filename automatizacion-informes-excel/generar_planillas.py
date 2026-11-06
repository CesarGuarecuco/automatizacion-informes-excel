"""Genera planillas de ventas INVENTADAS y desordenadas, como las que suele
mandar cada sucursal de una pyme: columnas con nombres distintos, fechas en
varios formatos, montos como texto, filas repetidas y totales a mano.

Ningún dato corresponde a una empresa real.
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEMILLA = 7
PRODUCTOS = {
    "Cemento 25 kg": ("Construcción", 6490), "Pintura látex 1 gl": ("Pinturas", 15990),
    "Taladro percutor": ("Herramientas", 54990), "Tornillos caja 100": ("Fijaciones", 3290),
    "Cerámica m2": ("Construcción", 8990), "Brocha 3 pulgadas": ("Pinturas", 2490),
    "Huincha de medir 5 m": ("Herramientas", 4990), "Silicona transparente": ("Fijaciones", 3990),
    "Escalera aluminio": ("Herramientas", 69990), "Esmalte al agua 1/4 gl": ("Pinturas", 7990),
}
SUCURSALES = {"Sucursal Río": 1.0, "Sucursal Plaza": 0.8, "Sucursal Puerto": 0.6}
DIAS = pd.date_range("2026-10-01", "2026-10-31")


def ventas_base(sucursal, factor, rng):
    filas = []
    for dia in DIAS:
        n = rng.poisson(14 * factor * (1.3 if dia.dayofweek == 5 else 1.0))
        for _ in range(n):
            prod = rng.choice(list(PRODUCTOS))
            filas.append({"fecha": dia, "producto": prod, "cantidad": int(rng.integers(1, 6)),
                          "precio": PRODUCTOS[prod][1], "vendedor": f"V{rng.integers(1, 4)}"})
    df = pd.DataFrame(filas)
    df.insert(0, "boleta", np.arange(1001, 1001 + len(df)))
    return df


def main():
    rng = np.random.default_rng(SEMILLA)
    carpeta = Path(__file__).parent / "entrada"
    carpeta.mkdir(exist_ok=True)

    # Sucursal Río: formato "ordenado", pero con filas repetidas
    rio = ventas_base("Sucursal Río", 1.0, rng)
    rio = pd.concat([rio, rio.sample(12, random_state=1)]).sort_values("boleta")
    rio.columns = ["N° Boleta", "Fecha", "Producto", "Cantidad", "Precio Unitario", "Vendedor"]
    rio.to_excel(carpeta / "ventas_rio_octubre.xlsx", index=False)

    # Sucursal Plaza: fechas como texto dd/mm/aaaa, montos con signo $ y punto de miles
    plaza = ventas_base("Sucursal Plaza", 0.8, rng)
    plaza = pd.concat([plaza, plaza.sample(5, random_state=2)]).sort_values("boleta")
    plaza["fecha"] = plaza.fecha.dt.strftime("%d/%m/%Y")
    plaza["precio"] = plaza.precio.map(lambda x: f"${x:,.0f}".replace(",", "."))
    plaza.columns = ["Boleta", "FECHA VENTA", "Descripción", "Cant.", "Valor", "Vend."]
    plaza.to_excel(carpeta / "Plaza - ventas oct.xlsx", index=False)

    # Sucursal Puerto: CSV con punto y coma, nombres de producto en mayúsculas y
    # espacios extra, y una fila de "TOTAL" escrita a mano al final
    puerto = ventas_base("Sucursal Puerto", 0.6, rng)
    puerto["producto"] = puerto.producto.str.upper() + "  "
    puerto["fecha"] = puerto.fecha.dt.strftime("%Y-%m-%d")
    total = pd.DataFrame([{"boleta": "", "fecha": "TOTAL", "producto": "", "cantidad": puerto.cantidad.sum(),
                           "precio": "", "vendedor": ""}])
    puerto = pd.concat([puerto, total])
    puerto.columns = ["folio", "dia", "item", "unidades", "precio_unit", "vendedor"]
    puerto.to_csv(carpeta / "puerto_oct.csv", sep=";", index=False, encoding="utf-8-sig")

    print("Planillas inventadas guardadas en", carpeta)


if __name__ == "__main__":
    main()
