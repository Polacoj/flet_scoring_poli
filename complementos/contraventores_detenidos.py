from pathlib import Path

import pandas as pd


class ContraventoresDetenidosProcessor:
    """Procesa un DataFrame para extraer filas con contraventores y detenidos.

    Uso:
        proc = ContraventoresDetenidosProcessor()
        df_result = proc.process_df(df)

    También se puede usar `process_file(ruta)` para leer y procesar un archivo.
    """

    def __init__(
        self,
        cont_col="CONTRAVENTORES",
        det_col="DETENIDOS",
        cierre_col="TIPO DE CIERRE",
        barrio_col="BARRIO",
        origen_col="ORIGEN",
        gap_col="GAP",
        sae_col="SAE",
        jefe_col="JEFE DE SERVICIO",
        oficial_index=5,  # Índice de la columna que contiene el nombre del oficial
    ):
        self.cont_col = cont_col
        self.det_col = det_col
        self.cierre_col = cierre_col
        self.barrio_col = barrio_col
        self.origen_col = origen_col
        self.gap_col = gap_col
        self.sae_col = sae_col
        self.jefe_col = jefe_col
        self.oficial_index = oficial_index

    def process_df(self, df: pd.DataFrame) -> pd.DataFrame:
        filas_contraventor = []
        filas_detenido = []

        for idx, fila in df.iterrows():
            if (
                fila.get(self.cont_col, 0) >= 1
                and fila.get(self.cierre_col, "") == "FINALIZA CON IMPUTADO/S PROPIO/S"
            ):
                oficial_val = (
                    fila.iloc[self.oficial_index]
                    if len(fila) > self.oficial_index
                    else None
                )
                filas_contraventor.append(
                    {
                        "FECHA": fila.get("FECHA"),
                        "HORA": "",
                        "JEFE DE SERVICIO": fila.get(self.jefe_col),
                        "OFICIAL": oficial_val,
                        "TIPO DE EVENTO": "EVENTO CON CONTRAVENTOR",
                        "BARRIO": "ADICIONAL",
                        "ORIGEN": fila.get(self.origen_col),
                        "GAP": fila.get(self.gap_col),
                        "SAE": fila.get(self.sae_col),
                    }
                )

        for idx, fila in df.iterrows():
            if (
                fila.get(self.det_col, 0) >= 1
                and fila.get(self.cierre_col, "") == "FINALIZA CON IMPUTADO/S PROPIO/S"
            ):
                oficial_val = (
                    fila.iloc[self.oficial_index]
                    if len(fila) > self.oficial_index
                    else None
                )
                filas_detenido.append(
                    {
                        "FECHA": fila.get("FECHA"),
                        "HORA": "",
                        "JEFE DE SERVICIO": fila.get(self.jefe_col),
                        "OFICIAL": oficial_val,
                        "TIPO DE EVENTO": "EVENTO CON DETENIDOS",
                        "BARRIO": "ADICIONAL",
                        "ORIGEN": fila.get(self.origen_col),
                        "GAP": fila.get(self.gap_col),
                        "SAE": fila.get(self.sae_col),
                    }
                )

        resultado = pd.concat(
            [
                pd.DataFrame(filas_contraventor),
                pd.DataFrame(filas_detenido),
            ],
            ignore_index=True,
        )

        return resultado

    def process_file(self, ruta_archivo) -> pd.DataFrame:
        df = pd.read_excel(ruta_archivo, engine="openpyxl")
        return self.process_df(df)


def process_file(ruta_archivo, **kwargs):
    proc = ContraventoresDetenidosProcessor(**kwargs)
    return proc.process_file(ruta_archivo)


if __name__ == "__main__":
    # Comportamiento hacia atrás compatible: si se ejecuta como script, intenta procesar
    # el archivo especificado por ruta por defecto (si existe) o debe editarse por el usuario.
    default_path = Path.cwd() / "CMU - SEPTIEMBRE.xlsx"
    if default_path.exists():
        resultado = process_file(default_path)
        salida = Path.cwd() / "contraventores_detenidos.xlsx"
        resultado.to_excel(salida, index=False)
        print(f"Archivo '{salida}' generado correctamente.")
    else:
        print(
            "Ejemplo: importar ContraventoresDetenidosProcessor desde otro script y usar process_file()."
        )
