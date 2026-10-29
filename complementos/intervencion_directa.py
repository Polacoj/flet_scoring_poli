import unicodedata

import pandas as pd


class ColaboracionPropiaProcessor:
    """Procesa un DataFrame para extraer filas con eventos de colaboracion con intervencion propios.

    Uso:
        proc = ColaboracionPropiaProcessor()
        df_result = proc.process_df(df)

    También se puede usar `process_file(ruta)` para leer y procesar un archivo.
    """

    def __init__(
        self,
        js_col="JEFE DE SERVICIO",
        of_col="OFICIAL",
        org_col="ORIGEN",
        cierre_col="TIPO DE CIERRE",
        gap_col="GAP",
        sae_col="SAE",
    ):
        self.js_col = js_col
        self.of_col = of_col
        self.org_col = org_col
        self.cierre_col = cierre_col
        self.gap_col = gap_col
        self.sae_col = sae_col

    def remove_accents(self, s):
        s = str(s)
        return "".join(
            ch
            for ch in unicodedata.normalize("NFKD", s)
            if not unicodedata.combining(ch)
        )

    def process_df(self, df: pd.DataFrame) -> pd.DataFrame:
        filas_colaboracion = []

        for idx, fila in df.iterrows():
            tip_val = self.remove_accents(fila.get(self.cierre_col, "")).strip().lower()
            if "colaboracion con imputado/s positiva" in tip_val:
                filas_colaboracion.append(
                    {
                        "FECHA": fila.get("FECHA"),
                        "HORA": "",
                        "JEFE DE SERVICIO": fila.get(self.js_col),
                        "OFICIAL": fila.get(self.of_col),
                        "TIPO DE EVENTO": "COLABORACION CON IMPUTADO/S POSITIVA",
                        "BARRIO": "ADICIONAL",
                        "ORIGEN": fila.get(self.org_col),
                        "GAP": fila.get(self.gap_col),
                        "SAE": fila.get(self.sae_col),
                    }
                )

        return pd.DataFrame(filas_colaboracion)

    def process_file(self, ruta_archivo: str) -> pd.DataFrame:
        df = pd.read_excel(ruta_archivo, engine="openpyxl")
        return self.process_df(df)
