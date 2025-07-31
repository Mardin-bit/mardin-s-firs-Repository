import pandas as pd
import math

MONTH_NAMES = [
    "فروردین",
    "اردیبهشت",
    "خرداد",
    "تیر",
    "مرداد",
    "شهریور",
    "مهر",
    "آبان",
    "آذر",
    "دی",
    "بهمن",
    "اسفند",
]

def is_jalali_leap(year: int) -> bool:
    """Return True if the given Jalali year is a leap year."""
    a = year - 474
    b = (a % 2820) + 474
    return ((b + 38) * 682) % 2816 < 682

def days_in_jalali_month(year: int, month: int) -> int:
    """Return number of days in a Jalali month for the given year."""
    if month <= 6:
        return 31
    if month <= 11:
        return 30
    return 30 if is_jalali_leap(year) else 29

def prepare_monthly_spi(input_excel: str, output_excel: str) -> None:
    """Read daily data from *input_excel* and write monthly SPI prep to *output_excel*."""
    # Read the first two columns: date and precipitation
    df = pd.read_excel(input_excel)

    # Ensure precipitation is numeric; treat errors as NaN
    df["Precipitation"] = pd.to_numeric(df["precipitation"], errors="coerce")

    # Split the Persian date into year, month and day
    df["Date"] = df["Persian Date"].astype(str).str.strip()

    df[["Year", "Month", "Day"]] = df["Date"].str.split("/", expand=True)
    df["Year"] = df["Year"].astype(int)
    df["Month"] = df["Month"].astype(int)
    df["Day"] = df["Day"].astype(int)

    years = list(range(1359, 1397))
    records = []
    for year in years:
        row = {"سال": year}
        for month in range(1, 13):
            mask = (df["Year"] == year) & (df["Month"] == month)
            series = df.loc[mask, "Precipitation"]
            total_precip = series.sum(skipna=True)
            valid_count = series.notna().sum()
            month_days = days_in_jalali_month(year, month)
            percent_valid = valid_count / month_days if month_days > 0 else math.nan
            month_name = MONTH_NAMES[month - 1]
            row[month_name] = total_precip
            row[f"درصد داده‌های {month_name}"] = percent_valid
            if percent_valid >= 0.75:
                adjusted = total_precip / percent_valid
            else:
                adjusted = math.nan
            row[f"بارش اصلاح‌شده {month_name}"] = adjusted
        records.append(row)

    result_df = pd.DataFrame(records)
    result_df.to_excel(output_excel, index=False)

if __name__ == "__main__":
    prepare_monthly_spi("gutvand.xlsx", "gutvand_monthly_spi.xlsx")
