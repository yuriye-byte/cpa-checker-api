import pandas as pd
import re


def parse_summary(summary_text: str):
    rows = []

    lines = summary_text.split("\n")

    for line in lines:
        line = line.strip()
        if not line:
            continue

        try:
            country = re.search(r"Country:\s*([^;]+)", line)
            baseline = re.search(r"Baseline:\s*([\d\.]+)", line)
            wager = re.search(r"Wager:\s*([\d\.]+)", line)
            cpa = re.search(r"CPA:\s*\$?([\d\.]+)", line)
            ftd = re.search(r"FTD:\s*(\d+)", line)

            rows.append({
                "country": country.group(1) if country else None,
                "baseline": float(baseline.group(1)) if baseline else None,
                "wager": float(wager.group(1)) if wager else None,
                "cpa": float(cpa.group(1)) if cpa else None,
                "ftd": int(ftd.group(1)) if ftd else None,
                "raw": line
            })

        except Exception:
            rows.append({"raw": line})

    return pd.DataFrame(rows)


def process_file(input_path: str, summary_text: str, output_path: str):
    # читаем Excel
    df = pd.read_excel(input_path)

    # парсим summary
    rules_df = parse_summary(summary_text)

    # добавляем колонки
    df["validated"] = "OK"
    df["summary"] = summary_text

    # создаём Excel с несколькими листами
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
        rules_df.to_excel(writer, index=False, sheet_name="rules")

    return {
        "status": "ok",
        "rows": len(df),
        "rules": len(rules_df)
    }