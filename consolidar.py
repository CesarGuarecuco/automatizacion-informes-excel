"""Une las planillas de ventas de cada sucursal, las limpia y genera el
informe mensual en Excel, listo para enviar.

Uso:
    python consolidar.py                # lee la carpeta entrada/
    python consolidar.py otra_carpeta/  # lee otra carpeta

Datos de demostración inventados (ver generar_planillas.py).
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Cada sucursal manda su planilla con su propio formato. Aquí se describe una
# sola vez cómo leer cada una; si llega una sucursal nueva, se agrega una línea.
FUENTES = {
    "ventas_rio_octubre.xlsx": {
        "sucursal": "Sucursal Río",
        "formato_fecha": "ISO8601",
        "columnas": {"N° Boleta": "boleta", "Fecha": "fecha", "Producto": "producto", "Cantidad": "cantidad",
                     "Precio Unitario": "precio", "Vendedor": "vendedor"},
    },
    "Plaza - ventas oct.xlsx": {
        "sucursal": "Sucursal Plaza",
        "formato_fecha": "%d/%m/%Y",
        "columnas": {"Boleta": "boleta", "FECHA VENTA": "fecha", "Descripción": "producto", "Cant.": "cantidad",
                     "Valor": "precio", "Vend.": "vendedor"},
    },
    "puerto_oct.csv": {
        "sucursal": "Sucursal Puerto",
        "formato_fecha": "%Y-%m-%d",
        "columnas": {"folio": "boleta", "dia": "fecha", "item": "producto", "unidades": "cantidad",
                     "precio_unit": "precio", "vendedor": "vendedor"},
    },
}
CATEGORIAS = {
    "Cemento 25 kg": "Construcción", "Cerámica m2": "Construcción",
    "Pintura látex 1 gl": "Pinturas", "Brocha 3 pulgadas": "Pinturas", "Esmalte al agua 1/4 gl": "Pinturas",
    "Taladro percutor": "Herramientas", "Huincha de medir 5 m": "Herramientas", "Escalera aluminio": "Herramientas",
    "Tornillos caja 100": "Fijaciones", "Silicona transparente": "Fijaciones",
}
VIOLETA = "2D1B8F"


def leer(ruta: Path) -> pd.DataFrame:
    if ruta.suffix == ".csv":
        return pd.read_csv(ruta, sep=";", dtype=str, encoding="utf-8-sig")
    return pd.read_excel(ruta, dtype=str)


def limpiar(df: pd.DataFrame, sucursal: str, columnas: dict, formato_fecha: str, registro: list) -> pd.DataFrame:
    df = df.rename(columns=columnas)[list(columnas.values())].copy()
    n0 = len(df)

    # Filas de totales escritas a mano
    es_total = df.fecha.str.upper().str.contains("TOTAL", na=False)
    df = df[~es_total]
    registro.append((sucursal, "2. Filas de total escritas a mano", int(es_total.sum())))

    df["fecha"] = pd.to_datetime(df.fecha, format=formato_fecha, errors="coerce")
    df["boleta"] = pd.to_numeric(df.boleta, errors="coerce")
    df["producto"] = df.producto.str.strip().str.capitalize()
    nombres = {p.lower(): p for p in CATEGORIAS}
    df["producto"] = df.producto.str.lower().map(nombres).fillna(df.producto)
    df["precio"] = pd.to_numeric(df.precio.str.replace(r"[$.\s]", "", regex=True), errors="coerce")
    df["cantidad"] = pd.to_numeric(df.cantidad, errors="coerce")

    invalidas = df[["boleta", "fecha", "precio", "cantidad"]].isna().any(axis=1)
    registro.append((sucursal, "3. Filas con boleta, fecha, precio o cantidad ilegible", int(invalidas.sum())))
    df = df[~invalidas]

    duplicadas = df.duplicated(subset="boleta")
    registro.append((sucursal, "4. Boletas repetidas (misma boleta dos veces)", int(duplicadas.sum())))
    df = df[~duplicadas]

    df["sucursal"] = sucursal
    df["categoria"] = df.producto.map(CATEGORIAS).fillna("Sin categoría")
    df["monto"] = df.cantidad * df.precio
    registro.append((sucursal, "1. Filas recibidas", n0))
    registro.append((sucursal, "5. Filas válidas en el informe", len(df)))
    return df


def dar_formato(hoja, ancho_min=12):
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=VIOLETA)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, col in enumerate(hoja.columns, start=1):
        largo = max(len(str(c.value or "")) for c in col)
        hoja.column_dimensions[get_column_letter(i)].width = max(ancho_min, min(largo + 2, 40))
    hoja.freeze_panes = "A2"


def formato_pesos(hoja, columnas):
    for col in columnas:
        for celda in hoja[col][1:]:
            celda.number_format = '"$"#,##0'


def main(carpeta: Path):
    registro, partes = [], []
    for archivo, cfg in FUENTES.items():
        ruta = carpeta / archivo
        if not ruta.exists():
            print(f"Aviso: falta {archivo}, se omite.")
            continue
        partes.append(limpiar(leer(ruta), cfg["sucursal"], cfg["columnas"], cfg["formato_fecha"], registro))
    ventas = pd.concat(partes, ignore_index=True).sort_values(["fecha", "sucursal"])
    mes = ventas.fecha.dt.strftime("%Y-%m").mode()[0]

    resumen = ventas.groupby("sucursal").agg(ventas_clp=("monto", "sum"), unidades=("cantidad", "sum"),
                                             boletas=("monto", "size")).sort_values("ventas_clp", ascending=False)
    resumen["ticket_promedio"] = (resumen.ventas_clp / resumen.boletas).round()
    resumen["participacion"] = resumen.ventas_clp / resumen.ventas_clp.sum()
    resumen.loc["Total"] = [resumen.ventas_clp.sum(), resumen.unidades.sum(), resumen.boletas.sum(),
                            round(resumen.ventas_clp.sum() / resumen.boletas.sum()), 1.0]

    ventas["semana"] = ventas.fecha.dt.to_period("W-SUN").dt.start_time
    por_semana = ventas.pivot_table(index="semana", columns="sucursal", values="monto", aggfunc="sum", fill_value=0)
    por_semana.index = por_semana.index.strftime("Semana del %d/%m")
    por_categoria = ventas.pivot_table(index="categoria", columns="sucursal", values="monto", aggfunc="sum",
                                       fill_value=0, margins=True, margins_name="Total").sort_values("Total", ascending=False)
    top = (ventas.groupby("producto").agg(unidades=("cantidad", "sum"), ventas_clp=("monto", "sum"))
           .sort_values("ventas_clp", ascending=False).head(10))
    calidad = pd.DataFrame(registro, columns=["sucursal", "revision", "filas"]).pivot(
        index="revision", columns="sucursal", values="filas")

    salida = Path(__file__).parent / "salida" / f"informe_ventas_{mes}.xlsx"
    salida.parent.mkdir(exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as xl:
        resumen.reset_index().rename(columns={
            "sucursal": "Sucursal", "ventas_clp": "Ventas (CLP)", "unidades": "Unidades", "boletas": "Boletas",
            "ticket_promedio": "Ticket promedio", "participacion": "Participación"}).to_excel(xl, sheet_name="Resumen", index=False)
        por_semana.rename_axis("Semana").reset_index().to_excel(xl, sheet_name="Por semana", index=False)
        por_categoria.rename_axis("Categoría").reset_index().to_excel(xl, sheet_name="Por categoría", index=False)
        top.reset_index().rename(columns={"producto": "Producto", "unidades": "Unidades", "ventas_clp": "Ventas (CLP)"}).to_excel(
            xl, sheet_name="Top 10 productos", index=False)
        calidad.reset_index().rename(columns={"revision": "Revisión"}).to_excel(xl, sheet_name="Revisión de datos", index=False)
        ventas.drop(columns="semana").to_excel(xl, sheet_name="Datos limpios", index=False)

        libro = xl.book
        for hoja in libro.worksheets:
            dar_formato(hoja)
        formato_pesos(libro["Resumen"], ["B", "E"])
        for celda in libro["Resumen"][libro["Resumen"].max_row]:
            celda.font = Font(bold=True)
        for celda in libro["Resumen"]["F"][1:]:
            celda.number_format = "0.0%"
        h = libro["Por semana"]
        formato_pesos(h, [get_column_letter(i) for i in range(2, h.max_column + 1)])
        h = libro["Por categoría"]
        formato_pesos(h, [get_column_letter(i) for i in range(2, h.max_column + 1)])
        formato_pesos(libro["Top 10 productos"], ["C"])
        formato_pesos(libro["Datos limpios"], ["E", "I"])
        for celda in libro["Datos limpios"]["B"][1:]:
            celda.number_format = "dd-mm-yyyy"

        g = BarChart(); g.title = "Ventas del mes por sucursal"; g.y_axis.title = "CLP"; g.legend = None
        n = len(resumen)  # sin la fila Total
        g.add_data(Reference(libro["Resumen"], min_col=2, min_row=1, max_row=n), titles_from_data=True)
        g.set_categories(Reference(libro["Resumen"], min_col=1, min_row=2, max_row=n))
        g.height, g.width = 7.5, 15
        libro["Resumen"].add_chart(g, "H2")

        h = libro["Por semana"]
        l = LineChart(); l.title = "Ventas por semana"; l.y_axis.title = "CLP"
        l.add_data(Reference(h, min_col=2, max_col=h.max_column, min_row=1, max_row=h.max_row), titles_from_data=True)
        l.set_categories(Reference(h, min_col=1, min_row=2, max_row=h.max_row))
        l.height, l.width = 7.5, 15
        h.add_chart(l, "F2")

    # Gráfico para el README
    fig, ejes = plt.subplots(1, 2, figsize=(12, 3.8))
    plt.rcParams.update({"axes.spines.top": False})
    (por_semana / 1e6).plot(ax=ejes[0], marker="o", color=["#2d1b8f", "#7a6cc9", "#e08e0b"])
    ejes[0].set(title="Ventas por semana (millones CLP)", xlabel="", ylabel="")
    ejes[0].legend(frameon=False, fontsize=8)
    (por_categoria.drop(index="Total").drop(columns="Total") / 1e6).plot.barh(
        stacked=True, ax=ejes[1], color=["#2d1b8f", "#7a6cc9", "#e08e0b"])
    ejes[1].set(title="Ventas por categoría (millones CLP)", xlabel="", ylabel="")
    ejes[1].legend(frameon=False, fontsize=8)
    for ax in ejes:
        ax.spines[["top", "right"]].set_visible(False)
        ax.tick_params(axis="x", labelsize=8)
    fig.text(0.99, 0.01, "Ejemplo · datos inventados", ha="right", fontsize=8, color="#9aa0ad")
    fig.tight_layout()
    fig.savefig(Path(__file__).parent / "images" / "informe.png", dpi=110, bbox_inches="tight")

    print(f"Informe generado: {salida}")
    print(calidad.to_string())
    print(resumen.to_string())


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "entrada")
