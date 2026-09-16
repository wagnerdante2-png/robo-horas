from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pandas as pd
import yaml
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


BASE_DIR = Path(__file__).resolve().parent


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="RPA local BI -> Excel -> WhatsApp")
    parser.add_argument("--config", default="config.yaml", help="Arquivo YAML de configuracao")
    parser.add_argument("--input", help="Usa um Excel ja exportado e pula a etapa do BI")
    parser.add_argument("--setup-auth", action="store_true", help="Abre BI e WhatsApp para gravar login no perfil local")
    parser.add_argument("--dry-run", action="store_true", help="Forca simulacao sem enviar WhatsApp")
    return parser.parse_args()


def load_config(path: str) -> dict[str, Any]:
    cfg_path = (BASE_DIR / path).resolve()
    if not cfg_path.exists():
        raise FileNotFoundError(
            f"Configuracao nao encontrada: {cfg_path}. Execute setup.bat ou copie config.example.yaml para config.yaml."
        )
    with cfg_path.open("r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    return cfg


def resolve_local(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path.resolve()


def ensure_dirs(cfg: dict[str, Any]) -> dict[str, Path]:
    output_dir = resolve_local(cfg.get("app", {}).get("output_dir", "output"))
    paths = {
        "output": output_dir,
        "downloads": output_dir / "downloads",
        "stores": output_dir / "lojas",
        "previews": output_dir / "previews",
        "logs": output_dir / "logs",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths


def setup_logging(paths: dict[str, Path], level: str = "INFO") -> Path:
    log_file = paths["logs"] / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )
    return log_file


def build_driver(cfg: dict[str, Any], downloads_dir: Path) -> webdriver.Chrome:
    browser_cfg = cfg.get("browser", {})
    profile_dir = resolve_local(browser_cfg.get("profile_dir", ".rpa_profile"))
    profile_dir.mkdir(parents=True, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={profile_dir}")

    if browser_cfg.get("headless", False):
        options.add_argument("--headless=new")

    prefs = {
        "download.default_directory": str(downloads_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("prefs", prefs)

    try:
        driver = webdriver.Chrome(options=options)
    except WebDriverException as exc:
        raise RuntimeError(
            "Nao foi possivel iniciar o Chrome. Feche outras instancias do Chrome que estejam usando "
            "o perfil .rpa_profile e confirme que o Google Chrome esta instalado."
        ) from exc

    timeout = int(browser_cfg.get("page_timeout_seconds", 90))
    driver.set_page_load_timeout(timeout)
    return driver


def setup_auth(driver: webdriver.Chrome, cfg: dict[str, Any]) -> None:
    bi_url = cfg.get("bi", {}).get("url", "https://app.powerbi.com/")
    wa_url = cfg.get("whatsapp", {}).get("web_url", "https://web.whatsapp.com/")

    logging.info("Abrindo BI para autenticacao...")
    driver.get(bi_url)
    input("Faca login no BI nesta janela e pressione ENTER aqui quando concluir... ")

    logging.info("Abrindo WhatsApp Web para autenticacao...")
    driver.get(wa_url)
    input("Faca login no WhatsApp Web nesta janela e pressione ENTER aqui quando concluir... ")

    logging.info("Autenticacao concluida. O perfil local foi preservado.")


BY_MAP = {
    "css": By.CSS_SELECTOR,
    "xpath": By.XPATH,
    "id": By.ID,
    "name": By.NAME,
    "class": By.CLASS_NAME,
    "tag": By.TAG_NAME,
}


def perform_export_actions(driver: webdriver.Chrome, actions: list[dict[str, Any]]) -> None:
    if not actions:
        raise RuntimeError(
            "Modo selectors ativado, mas nenhuma acao foi cadastrada em bi.export.actions."
        )

    for index, action in enumerate(actions, start=1):
        method = str(action.get("by", "xpath")).lower()
        value = str(action.get("value", "")).strip()
        if method not in BY_MAP or not value:
            raise ValueError(f"Acao {index} invalida: {action}")

        timeout = int(action.get("timeout_seconds", 30))
        wait = WebDriverWait(driver, timeout)
        logging.info("BI: executando acao %s/%s", index, len(actions))

        element = wait.until(EC.element_to_be_clickable((BY_MAP[method], value)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        try:
            element.click()
        except WebDriverException:
            driver.execute_script("arguments[0].click();", element)

        sleep_after = float(action.get("sleep_after_seconds", 1))
        if sleep_after > 0:
            time.sleep(sleep_after)


def excel_snapshot(downloads_dir: Path) -> set[Path]:
    return {
        p.resolve()
        for pattern in ("*.xlsx", "*.xls")
        for p in downloads_dir.glob(pattern)
        if not p.name.startswith("~$")
    }


def wait_for_new_excel(downloads_dir: Path, before: set[Path], timeout: int) -> Path:
    deadline = time.time() + timeout
    last_sizes: dict[Path, int] = {}

    while time.time() < deadline:
        candidates = [
            p.resolve()
            for pattern in ("*.xlsx", "*.xls")
            for p in downloads_dir.glob(pattern)
            if not p.name.startswith("~$") and p.resolve() not in before
        ]

        partial_downloads = list(downloads_dir.glob("*.crdownload"))
        if candidates and not partial_downloads:
            newest = max(candidates, key=lambda p: p.stat().st_mtime)
            size = newest.stat().st_size
            if size > 0 and last_sizes.get(newest) == size:
                logging.info("Excel detectado: %s", newest)
                return newest
            last_sizes[newest] = size

        time.sleep(1)

    raise TimeoutError(f"Nenhum novo Excel foi detectado em {downloads_dir} dentro de {timeout}s.")


def export_bi_excel(
    driver: webdriver.Chrome,
    cfg: dict[str, Any],
    downloads_dir: Path,
) -> Path:
    bi_cfg = cfg.get("bi", {})
    export_cfg = bi_cfg.get("export", {})
    url = str(bi_cfg.get("url", "")).strip()
    if not url:
        raise ValueError("Defina bi.url em config.yaml.")

    mode = str(export_cfg.get("mode", "manual")).lower()
    timeout = int(export_cfg.get("timeout_seconds", 180))
    before = excel_snapshot(downloads_dir)

    logging.info("Abrindo BI: %s", url)
    driver.get(url)

    if mode == "selectors":
        perform_export_actions(driver, export_cfg.get("actions", []))
    elif mode == "manual":
        print("\n[BI] O relatorio foi aberto.")
        print("[BI] Exporte o visual/relatorio para Excel. O robo detectara o arquivo automaticamente.\n")
    else:
        raise ValueError("bi.export.mode deve ser 'manual' ou 'selectors'.")

    return wait_for_new_excel(downloads_dir, before, timeout)


def find_sheet_with_store_column(
    excel_path: Path,
    preferred_sheet: str,
    header_row: int,
    store_column: str,
) -> tuple[str, pd.DataFrame]:
    if preferred_sheet:
        df = pd.read_excel(excel_path, sheet_name=preferred_sheet, header=header_row)
        if store_column not in df.columns:
            raise KeyError(f"Coluna '{store_column}' nao encontrada na aba '{preferred_sheet}'.")
        return preferred_sheet, df

    book = pd.ExcelFile(excel_path)
    for sheet in book.sheet_names:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
        except Exception:
            continue

        df.columns = [str(c).strip() for c in df.columns]
        if store_column in df.columns:
            return sheet, df

    raise KeyError(
        f"Nenhuma aba do Excel contem a coluna de loja '{store_column}'. "
        f"Abas encontradas: {book.sheet_names}"
    )


def normalize_store(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if re.fullmatch(r"\d+(?:\.0)?", text):
        number = int(float(text))
        return f"ML{number:02d}"
    match = re.fullmatch(r"ML\s*0*(\d+)", text)
    if match:
        return f"ML{int(match.group(1)):02d}"
    return re.sub(r"\s+", "", text)


def normalize_phone(value: Any) -> str:
    return re.sub(r"\D", "", "" if pd.isna(value) else str(value))


def is_active(value: Any) -> bool:
    text = str(value).strip().lower()
    return text in {"1", "true", "sim", "s", "yes", "y", "ativo"}


def load_store_mapping(cfg: dict[str, Any]) -> dict[str, str]:
    stores_cfg = cfg.get("stores", {})
    mapping_path = resolve_local(stores_cfg.get("mapping_file", "data/lojas.csv"))
    if not mapping_path.exists():
        raise FileNotFoundError(f"Cadastro de lojas nao encontrado: {mapping_path}")

    df = pd.read_csv(mapping_path, dtype=str, sep=None, engine="python")
    code_col = stores_cfg.get("code_column", "loja")
    phone_col = stores_cfg.get("phone_column", "telefone")
    active_col = stores_cfg.get("active_column", "ativo")

    for required in (code_col, phone_col):
        if required not in df.columns:
            raise KeyError(f"Coluna obrigatoria '{required}' ausente em {mapping_path}")

    mapping: dict[str, str] = {}
    for _, row in df.iterrows():
        if active_col in df.columns and not is_active(row.get(active_col, "")):
            continue
        store = normalize_store(row.get(code_col))
        phone = normalize_phone(row.get(phone_col))
        if store and phone:
            mapping[store] = phone

    if not mapping:
        raise ValueError("Nenhuma loja ativa com telefone valido foi encontrada no cadastro.")

    return mapping


def safe_filename(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("_") or "loja"


def format_value(value: Any) -> str:
    if pd.isna(value):
        return "-"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(value).strip()


def build_message(store: str, group: pd.DataFrame, cfg: dict[str, Any]) -> str:
    msg_cfg = cfg.get("message", {})
    title = str(msg_cfg.get("title", "📊 PARCIAL OPERACIONAL | {loja}")).format(loja=store)
    fields = list(msg_cfg.get("fields") or [])
    max_auto = int(msg_cfg.get("max_auto_fields", 10))
    store_column = cfg.get("excel", {}).get("store_column", "Loja")

    if not fields:
        fields = [str(c) for c in group.columns if str(c) != store_column][:max_auto]

    row = group.iloc[0]
    lines = [title, ""]

    if msg_cfg.get("include_record_count", True) and len(group) > 1:
        lines.append(f"Registros no relatorio: {len(group)}")
        lines.append("")

    for field in fields:
        if field in group.columns:
            lines.append(f"{field}: {format_value(row[field])}")

    footer = str(msg_cfg.get("footer", "")).strip()
    if footer:
        lines.extend(["", footer])

    return "\n".join(lines).strip()


def split_by_store(
    excel_path: Path,
    cfg: dict[str, Any],
    paths: dict[str, Path],
) -> list[tuple[str, str, str, Path]]:
    excel_cfg = cfg.get("excel", {})
    sheet_name = str(excel_cfg.get("sheet_name", "") or "").strip()
    header_row = int(excel_cfg.get("header_row", 0))
    store_column = str(excel_cfg.get("store_column", "Loja"))

    selected_sheet, df = find_sheet_with_store_column(
        excel_path=excel_path,
        preferred_sheet=sheet_name,
        header_row=header_row,
        store_column=store_column,
    )
    df.columns = [str(c).strip() for c in df.columns]

    if excel_cfg.get("drop_empty_store", True):
        df = df[df[store_column].notna()].copy()

    df["_loja_normalizada"] = df[store_column].map(normalize_store)
    df = df[df["_loja_normalizada"] != ""].copy()

    mapping = load_store_mapping(cfg)
    results: list[tuple[str, str, str, Path]] = []

    logging.info(
        "Excel carregado: %s | aba=%s | linhas=%s | lojas=%s",
        excel_path.name,
        selected_sheet,
        len(df),
        df["_loja_normalizada"].nunique(),
    )

    for store, group in df.groupby("_loja_normalizada", sort=True):
        clean_group = group.drop(columns=["_loja_normalizada"])
        out_file = paths["stores"] / f"{safe_filename(store)}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        clean_group.to_excel(out_file, index=False)

        phone = mapping.get(store, "")
        if not phone:
            logging.warning("%s: sem telefone ativo no cadastro; nao sera enviado.", store)
            continue

        message = build_message(store, clean_group, cfg)
        results.append((store, phone, message, out_file))

    return results


def save_preview(store: str, phone: str, message: str, paths: dict[str, Path]) -> Path:
    preview = paths["previews"] / f"{safe_filename(store)}_{datetime.now():%Y%m%d_%H%M%S}.txt"
    preview.write_text(f"Loja: {store}\nTelefone: {phone}\n\n{message}\n", encoding="utf-8")
    return preview


def send_whatsapp_message(
    driver: webdriver.Chrome,
    phone: str,
    message: str,
    cfg: dict[str, Any],
) -> None:
    wa_cfg = cfg.get("whatsapp", {})
    timeout = int(wa_cfg.get("chat_load_timeout_seconds", 60))

    url = f"https://web.whatsapp.com/send?phone={quote(phone)}&text={quote(message)}"
    driver.get(url)

    wait = WebDriverWait(driver, timeout)
    try:
        composer = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//footer//*[@contenteditable='true']")
            )
        )
    except TimeoutException as exc:
        raise TimeoutError(
            f"WhatsApp nao carregou a conversa do telefone {phone}. "
            "Confirme login, telefone e conexao."
        ) from exc

    composer.click()
    time.sleep(0.5)
    ActionChains(driver).send_keys(Keys.ENTER).perform()


def process_dispatches(
    driver: webdriver.Chrome | None,
    dispatches: list[tuple[str, str, str, Path]],
    cfg: dict[str, Any],
    paths: dict[str, Path],
    force_dry_run: bool,
) -> None:
    wa_cfg = cfg.get("whatsapp", {})
    dry_run = force_dry_run or bool(wa_cfg.get("dry_run", True))
    delay = float(wa_cfg.get("delay_between_messages_seconds", 4))

    if not dispatches:
        logging.warning("Nenhum disparo elegivel foi encontrado.")
        return

    logging.info("Disparos preparados: %s | dry_run=%s", len(dispatches), dry_run)

    for position, (store, phone, message, store_file) in enumerate(dispatches, start=1):
        preview = save_preview(store, phone, message, paths)
        logging.info(
            "[%s/%s] %s -> %s | excel=%s | preview=%s",
            position,
            len(dispatches),
            store,
            phone,
            store_file.name,
            preview.name,
        )

        if dry_run:
            print("\n" + "=" * 60)
            print(f"PREVIA {store} -> {phone}")
            print(message)
            print("=" * 60)
            continue

        if driver is None:
            raise RuntimeError("Driver do navegador nao foi iniciado para envio.")

        send_whatsapp_message(driver, phone, message, cfg)
        logging.info("%s: mensagem enviada.", store)

        if position < len(dispatches) and delay > 0:
            time.sleep(delay)


def main() -> int:
    args = parse_args()
    driver: webdriver.Chrome | None = None

    try:
        cfg = load_config(args.config)
        paths = ensure_dirs(cfg)
        log_file = setup_logging(paths, cfg.get("logging", {}).get("level", "INFO"))
        logging.info("Inicio do robo. Log: %s", log_file)

        if args.setup_auth:
            driver = build_driver(cfg, paths["downloads"])
            setup_auth(driver, cfg)
            return 0

        input_excel: Path | None = None
        if args.input:
            input_excel = resolve_local(args.input)
            if not input_excel.exists():
                raise FileNotFoundError(f"Excel informado nao existe: {input_excel}")
        else:
            driver = build_driver(cfg, paths["downloads"])
            input_excel = export_bi_excel(driver, cfg, paths["downloads"])

        dispatches = split_by_store(input_excel, cfg, paths)

        dry_run = args.dry_run or bool(cfg.get("whatsapp", {}).get("dry_run", True))
        if not dry_run and driver is None:
            driver = build_driver(cfg, paths["downloads"])

        process_dispatches(driver, dispatches, cfg, paths, args.dry_run)
        logging.info("Execucao concluida com sucesso.")
        return 0

    except KeyboardInterrupt:
        logging.warning("Execucao interrompida pelo usuario.")
        return 130
    except Exception as exc:
        logging.exception("Falha: %s", exc)
        print(f"\n[ERRO] {exc}")
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
