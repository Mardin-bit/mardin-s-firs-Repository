import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage


def extract_month_name(column: str) -> str:
    """Return the Persian month name extracted from column header."""
    return column.replace("SPI", "").strip()


def create_spi_charts(input_file: str = "chart_of_spi.xlsx", output_file: str = "spi_area_charts.xlsx") -> None:
    """Create SPI area charts for each month and save them in an Excel file."""
    df = pd.read_excel(input_file)
    # Remove leading/trailing spaces from column names
    df.columns = [c.strip() for c in df.columns]

    year_col = "year"
    if year_col not in df.columns:
        raise KeyError(f"Column '{year_col}' not found in the Excel file")

    month_columns = [c for c in df.columns if c != year_col]

    images = []
    for col in month_columns:
        month_name = extract_month_name(col)
        data = df[[year_col, col]].dropna()
        years = data[year_col]
        values = data[col]

        plt.figure(figsize=(8, 4))
        plt.fill_between(years, values, 0, where=values >= 0, interpolate=True, color="blue", alpha=0.7)
        plt.fill_between(years, values, 0, where=values < 0, interpolate=True, color="red", alpha=0.7)
        plt.plot(years, values, color="black", linewidth=1)
        plt.title(f"SPI Chart for {month_name}")
        plt.xlabel("Year")
        plt.ylabel("SPI")
        plt.tight_layout()
        image_name = f"{month_name}.png"
        plt.savefig(image_name)
        plt.close()
        images.append((month_name, image_name))

    wb = Workbook()
    # Remove the default sheet created by Workbook
    wb.remove(wb.active)

    for month_name, image_name in images:
        ws = wb.create_sheet(title=month_name)
        img = XLImage(image_name)
        ws.add_image(img, "A1")

    wb.save(output_file)


if __name__ == "__main__":
    create_spi_charts()
