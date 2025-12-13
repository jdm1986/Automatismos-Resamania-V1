import os
import pandas as pd


def load_data_file(folder_path: str, base_name: str) -> pd.DataFrame:
    """
    Load a data file prioritizing CSV (Resamania exports) and falling back to XLSX.
    Args:
        folder_path: Directory where the exports live.
        base_name: Filename without extension, e.g. "RESUMEN CLIENTE".
    Returns:
        pandas.DataFrame with the loaded contents.
    Raises:
        FileNotFoundError if no matching file is found.
    """
    candidates = [
        (f"{base_name}.csv", "csv"),
        (f"{base_name}.xlsx", "excel"),
    ]

    for filename, kind in candidates:
        full_path = os.path.join(folder_path, filename)
        if not os.path.exists(full_path):
            continue

        if kind == "csv":
            try:
                # sep=None with engine="python" infers delimiter (Resamania exports often use ';').
                return pd.read_csv(full_path, sep=None, engine="python", encoding="utf-8-sig")
            except UnicodeDecodeError:
                # Fallback a Latin-1 si el CSV viene en ANSI/Windows-1252.
                try:
                    return pd.read_csv(full_path, sep=None, engine="python", encoding="latin-1")
                except Exception:
                    return pd.read_csv(full_path, sep=";", encoding="latin-1")
            except Exception:
                # Fallback a separador ';' en UTF-8 si la inferencia falla por delimitador.
                return pd.read_csv(full_path, sep=";", encoding="utf-8-sig")

        return pd.read_excel(full_path)

    raise FileNotFoundError(
        f"No se encontro ninguno de estos archivos en {folder_path}: "
        f"{', '.join(name for name, _ in candidates)}"
    )
