import streamlit as st
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import io


# ==============================================================================
# 1. МОДУЛЬ ПАРСИНГУ PDF (Ф1 та Ф2 за кодами рядків)
# ==============================================================================
def parse_pdf_report(pdf_file):
    """Витягує показники фінансової звітності з PDF за кодами рядків."""
    data = {}
    if pdf_file is None:
        return data

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for line in text.split('\n'):
                # Нормалізуємо пробіли (замінюємо нерозривні пробіли на звичайні)
                clean_line = line.replace('\xa0', ' ').strip()
                
                # Регулярне вираження шукає:
                # 1. Код рядка (точно 4 цифри)
                # 2. Перше число після нього (звітний період), яке може бути:
                #    - у дужках: (1 234,5) або (1234)
                #    - зі знаком мінус: -1 234
                #    - звичайним: 1 234,5
                match = re.search(
                    r'\b(\d{4})\b.*?\b(\(?\-?\s*\d+[\d\s]*[.,]?\d*\)?)\b', 
                    clean_line
                )
                if match:
                    code = match.group(1)
                    val_str = match.group(2).strip()
                    
                    # Перевірка на від'ємне значення у дужках: (123) -> -123
                    is_negative = False
                    if val_str.startswith('(') and val_str.endswith(')'):
                        is_negative = True
                        val_str = val_str[1:-1]
                    
                    # Видаляємо пробіли всередині числа та замінюємо кому на крапку
                    val_str = val_str.replace(" ", "").replace(",", ".")
                    
                    try:
                        val = float(val_str)
                        if is_negative:
                            val = -val
                        
                        # Записуємо код лише один раз (перше знайдене співпадіння)
                        if code not in data:
                            data[code] = val
                    except ValueError:
                        pass
                        
    return data


# ==============================================================================
# 2. МОДУЛЬ ГЕНЕРАЦІЇ EXCEL (4 ТАБЛИЦІ МІНФІНУ С ЗВ'ЯЗАНИМИ ФОРМУЛАМИ)
# ==============================================================================
def generate_excel_report(pdf_y1, pdf_y2, pdf_y3):
    d1 = parse_pdf_report(pdf_y1)
    d2 = parse_pdf_report(pdf_y2)
    d3 = parse_pdf_report(pdf_y3)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Фінансовий_Аналіз"

    # Стилі
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2B579A", end_color="2B579A", fill_type="solid")
    title_font = Font(name="Calibri", size=12, bold=True, color="002060")
    bold_font = Font(name="Calibri", size=11, bold=True)

    r = 1

    # --- ТАБЛИЦЯ 1: СТРУКТУРА І ДИНАМІКА МАЙНА ТА ДЖЕРЕЛ ---
    ws.cell(row=r, column=1, value="1. СТРУКТУРА І ДИНАМІКА МАЙНА ТА ДЖЕРЕЛ ЙОГО УТВОРЕННЯ").font = title_font
    r += 1

    h1 = ["Показники", "На 31.12.2023", "", "На 31.12.2024", "", "На 31.12.2025", "", "Відхилення 2025 р. до 2023 р.",
          "", ""]
    h2 = ["", "тис. грн.", "%", "тис. грн.", "%", "тис. грн.", "%", "+,-", "%", "п.с."]

    ws.append(h1)
    ws.append(h2)

    for row_idx in range(r, r + 2):
        for col_idx in range(1, 11):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r += 2

    t1_start_row = r
    items_t1 = [
        ("АКТИВИ", None),
        ("1. Необоротні активи:", "1095"),
        ("в т.ч. основні засоби", "1010"),
        ("2. Оборотні активи", "1195"),
        ("в т.ч. запаси", "1100"),
        ("дебіторська заборгованість", "1125"),
        ("грошові кошти", "1165"),
        ("ПАСИВИ", None),
        ("1. Власний капітал", "1495"),
        ("в т.ч. зареєстрований капітал", "1400"),
        ("нерозподілений прибуток", "1420"),
        ("2. Довгострокові зобов'язання і забезпечення", "1595"),
        ("3. Поточні зобов'язання і забезпечення", "1695"),
        ("БАЛАНС:", "1300")
    ]

    balance_row = t1_start_row + len(items_t1) - 1

    for name, code in items_t1:
        ws.cell(row=r, column=1, value=name)
        if code is None:
            ws.row_dimensions[r].font = bold_font
            ws.cell(row=r, column=1).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        else:
            ws.cell(row=r, column=2, value=d1.get(code, 0.0))
            ws.cell(row=r, column=4, value=d2.get(code, 0.0))
            ws.cell(row=r, column=6, value=d3.get(code, 0.0))

            ws.cell(row=r, column=3, value=f"=(B{r}/$B${balance_row})*100")
            ws.cell(row=r, column=5, value=f"=(D{r}/$D${balance_row})*100")
            ws.cell(row=r, column=7, value=f"=(F{r}/$F${balance_row})*100")

            ws.cell(row=r, column=8, value=f"=F{r}-B{r}")
            ws.cell(row=r, column=9, value=f"=IF(B{r}<>0, ((F{r}-B{r})/B{r})*100, 0)")
            ws.cell(row=r, column=10, value=f"=G{r}-C{r}")

            for c in [2, 4, 6, 8]: ws.cell(row=r, column=c).number_format = "#,##0.0"
            for c in [3, 5, 7, 9, 10]: ws.cell(row=r, column=c).number_format = "0.00"

            if name == "БАЛАНС:":
                ws.row_dimensions[r].font = bold_font
        r += 1

    # --- ТАБЛИЦЯ 2: МЕХАНІЗМ ФОРМУВАННЯ ФІНАНСОВИХ РЕЗУЛЬТАТІВ ---
    r += 2
    ws.cell(row=r, column=1, value="2. МЕХАНІЗМ ФОРМУВАННЯ ФІНАНСОВИХ РЕЗУЛЬТАТІВ ДІЯЛЬНОСТІ").font = title_font
    r += 1

    ws.append(["Показники", "2023", "2024", "2025", "Відхилення 2025 р. до 2023 р. (+,-)",
               "Відхилення 2025 р. до 2023 р. (%)"])
    for col_idx in range(1, 7):
        cell = ws.cell(row=r, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r += 1

    t2_start_row = r
    items_t2 = [
        ("Чистий дохід від реалізації продукції (товарів, робіт, послуг)", "2000"),
        ("Собівартість реалізованої продукції (товарів, робіт, послуг)", "2050"),
        ("Валовий: прибуток (збиток)", "2090"),
        ("Інші операційні доходи", "2120"),
        ("Адміністративні витрати", "2130"),
        ("Витрати на збут", "2150"),
        ("Інші операційні витрати", "2180"),
        ("Фінансовий результат від операційної діяльності: прибуток (збиток)", "2190"),
        ("Фінансові витрати", "2250"),
        ("Інші витрати", "2270"),
        ("Фінансовий результат до оподаткування: прибуток (збиток)", "2290"),
        ("Чистий фінансовий результат: прибуток (збиток)", "2350")
    ]

    for name, code in items_t2:
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=d1.get(code, 0.0))
        ws.cell(row=r, column=3, value=d2.get(code, 0.0))
        ws.cell(row=r, column=4, value=d3.get(code, 0.0))

        ws.cell(row=r, column=5, value=f"=D{r}-B{r}")
        ws.cell(row=r, column=6, value=f"=IF(B{r}<>0, ((D{r}-B{r})/B{r})*100, 0)")

        for c in [2, 3, 4, 5]: ws.cell(row=r, column=c).number_format = "#,##0.0"
        ws.cell(row=r, column=6).number_format = "0.00"
        r += 1

    # --- ТАБЛИЦЯ 3: ОСНОВНІ ЕКОНОМІЧНІ ПОКАЗНИКИ ---
    r += 2
    ws.cell(row=r, column=1,
            value="3. ОСНОВНІ ЕКОНОМІЧНІ ПОКАЗНИКИ ДІЯЛЬНОСТІ СУБ'ЄКТА ГОСПОДАРЮВАННЯ").font = title_font
    r += 1

    ws.append(["№", "Показник", "2023", "2024", "2025", "Відхилення (+,-)", "Відхилення (%)"])
    for col_idx in range(1, 8):
        cell = ws.cell(row=r, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r += 1

    t3_start_row = r
    indicators_t3 = [
        "Чистий дохід, тис. грн.", "Валовий прибуток, тис. грн.", "Чистий прибуток, тис. грн.",
        "Середньорічна вартість активів, тис. грн.", "Середньорічна вартість необоротних активів, тис. грн.",
        "Середньорічна вартість оборотних активів, тис. грн.", "Середньооблікова чисельність працівників, осіб",
        "Фондовіддача", "Фондомісткість", "Фондоозброєність, тис. грн /ос.",
        "Продуктивність праці, тис. грн. /ос.", "Коефіцієнт оборотності оборотних активів"
    ]

    for idx, name in enumerate(indicators_t3, 1):
        ws.cell(row=r, column=1, value=idx)
        ws.cell(row=r, column=2, value=name)

        if idx == 1:
            ws.cell(row=r, column=3, value=f"=B{t2_start_row}")
            ws.cell(row=r, column=4, value=f"=C{t2_start_row}")
            ws.cell(row=r, column=5, value=f"=D{t2_start_row}")
        elif idx == 2:
            ws.cell(row=r, column=3, value=f"=B{t2_start_row + 2}")
            ws.cell(row=r, column=4, value=f"=C{t2_start_row + 2}")
            ws.cell(row=r, column=5, value=f"=D{t2_start_row + 2}")
        elif idx == 3:
            ws.cell(row=r, column=3, value=f"=B{t2_start_row + 11}")
            ws.cell(row=r, column=4, value=f"=C{t2_start_row + 11}")
            ws.cell(row=r, column=5, value=f"=D{t2_start_row + 11}")
        elif idx == 4:
            ws.cell(row=r, column=3, value=f"=B{balance_row}")
            ws.cell(row=r, column=4, value=f"=D{balance_row}")
            ws.cell(row=r, column=5, value=f"=F{balance_row}")
        elif idx == 5:
            ws.cell(row=r, column=3, value=f"=B{t1_start_row + 1}")
            ws.cell(row=r, column=4, value=f"=D{t1_start_row + 1}")
            ws.cell(row=r, column=5, value=f"=F{t1_start_row + 1}")
        elif idx == 6:
            ws.cell(row=r, column=3, value=f"=B{t1_start_row + 3}")
            ws.cell(row=r, column=4, value=f"=D{t1_start_row + 3}")
            ws.cell(row=r, column=5, value=f"=F{t1_start_row + 3}")
        elif idx == 7:
            ws.cell(row=r, column=3, value=50)
            ws.cell(row=r, column=4, value=50)
            ws.cell(row=r, column=5, value=50)
        elif idx == 8:
            ws.cell(row=r, column=3, value=f"=C{t3_start_row}/C{t3_start_row + 4}")
            ws.cell(row=r, column=4, value=f"=D{t3_start_row}/D{t3_start_row + 4}")
            ws.cell(row=r, column=5, value=f"=E{t3_start_row}/E{t3_start_row + 4}")
        elif idx == 9:
            ws.cell(row=r, column=3, value=f"=1/C{t3_start_row + 7}")
            ws.cell(row=r, column=4, value=f"=1/D{t3_start_row + 7}")
            ws.cell(row=r, column=5, value=f"=1/E{t3_start_row + 7}")
        elif idx == 10:
            ws.cell(row=r, column=3, value=f"=C{t3_start_row + 4}/C{t3_start_row + 6}")
            ws.cell(row=r, column=4, value=f"=D{t3_start_row + 4}/D{t3_start_row + 6}")
            ws.cell(row=r, column=5, value=f"=E{t3_start_row + 4}/E{t3_start_row + 6}")
        elif idx == 11:
            ws.cell(row=r, column=3, value=f"=C{t3_start_row}/C{t3_start_row + 6}")
            ws.cell(row=r, column=4, value=f"=D{t3_start_row}/D{t3_start_row + 6}")
            ws.cell(row=r, column=5, value=f"=E{t3_start_row}/E{t3_start_row + 6}")
        elif idx == 12:
            ws.cell(row=r, column=3, value=f"=C{t3_start_row}/C{t3_start_row + 5}")
            ws.cell(row=r, column=4, value=f"=D{t3_start_row}/D{t3_start_row + 5}")
            ws.cell(row=r, column=5, value=f"=E{t3_start_row}/E{t3_start_row + 5}")

        ws.cell(row=r, column=6, value=f"=E{r}-C{r}")
        ws.cell(row=r, column=7, value=f"=IF(C{r}<>0, ((E{r}-C{r})/C{r})*100, 0)")

        for c in [3, 4, 5, 6]: ws.cell(row=r, column=c).number_format = "#,##0.00"
        ws.cell(row=r, column=7).number_format = "0.00"
        r += 1

    # --- ТАБЛИЦЯ 4: ДИНАМІКА ПОКАЗНИКІВ ФІНАНСОВОГО СТАНУ ---
    r += 2
    ws.cell(row=r, column=1,
            value="4. ДИНАМІКА ПОКАЗНИКІВ ОЦІНКИ ФІНАНСОВОГО СТАНУ СУБ'ЄКТА ГОСПОДАРЮВАННЯ").font = title_font
    r += 1

    ws.append(["Показники", "2023", "2024", "2025", "Відхилення 2025 р. до 2023 р. (+,-)"])
    for col_idx in range(1, 6):
        cell = ws.cell(row=r, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r += 1

    ratios_t4 = [
        "1. Коефіцієнт загальної ліквідності", "2. Коефіцієнт абсолютної ліквідності", "3. Чистий оборотний капітал",
        "4. Коефіцієнт автономії", "5. Коефіцієнт фінансової стабільності",
        "6. Коефіцієнт забезпеченості власними засобами",
        "7. Коефіцієнт оборотності активів", "8. Коефіцієнт оборотності власного капіталу",
        "9. Рентабельність господарської діяльності, %",
        "10. Рентабельність власного капіталу, %", "11. Рентабельність продукції, %"
    ]

    for idx, name in enumerate(ratios_t4, 1):
        ws.cell(row=r, column=1, value=name)

        if idx == 1:
            ws.cell(row=r, column=2, value=f"=B{t1_start_row + 3}/B{t1_start_row + 12}")
            ws.cell(row=r, column=3, value=f"=D{t1_start_row + 3}/D{t1_start_row + 12}")
            ws.cell(row=r, column=4, value=f"=F{t1_start_row + 3}/F{t1_start_row + 12}")
        elif idx == 2:
            ws.cell(row=r, column=2, value=f"=B{t1_start_row + 6}/B{t1_start_row + 12}")
            ws.cell(row=r, column=3, value=f"=D{t1_start_row + 6}/D{t1_start_row + 12}")
            ws.cell(row=r, column=4, value=f"=F{t1_start_row + 6}/F{t1_start_row + 12}")
        elif idx == 3:
            ws.cell(row=r, column=2, value=f"=B{t1_start_row + 3}-B{t1_start_row + 12}")
            ws.cell(row=r, column=3, value=f"=D{t1_start_row + 3}-D{t1_start_row + 12}")
            ws.cell(row=r, column=4, value=f"=F{t1_start_row + 3}-F{t1_start_row + 12}")
        elif idx == 4:
            ws.cell(row=r, column=2, value=f"=B{t1_start_row + 8}/B{balance_row}")
            ws.cell(row=r, column=3, value=f"=D{t1_start_row + 8}/D{balance_row}")
            ws.cell(row=r, column=4, value=f"=F{t1_start_row + 8}/F{balance_row}")
        elif idx == 5:
            ws.cell(row=r, column=2, value=f"=B{t1_start_row + 8}/(B{t1_start_row + 11}+B{t1_start_row + 12})")
            ws.cell(row=r, column=3, value=f"=D{t1_start_row + 8}/(D{t1_start_row + 11}+D{t1_start_row + 12})")
            ws.cell(row=r, column=4, value=f"=F{t1_start_row + 8}/(F{t1_start_row + 11}+F{t1_start_row + 12})")
        elif idx == 6:
            ws.cell(row=r, column=2, value=f"=(B{t1_start_row + 8}-B{t1_start_row + 1})/B{t1_start_row + 3}")
            ws.cell(row=r, column=3, value=f"=(D{t1_start_row + 8}-D{t1_start_row + 1})/D{t1_start_row + 3}")
            ws.cell(row=r, column=4, value=f"=(F{t1_start_row + 8}-F{t1_start_row + 1})/F{t1_start_row + 3}")
        elif idx == 7:
            ws.cell(row=r, column=2, value=f"=B{t2_start_row}/B{balance_row}")
            ws.cell(row=r, column=3, value=f"=C{t2_start_row}/D{balance_row}")
            ws.cell(row=r, column=4, value=f"=D{t2_start_row}/F{balance_row}")
        elif idx == 8:
            ws.cell(row=r, column=2, value=f"=B{t2_start_row}/B{t1_start_row + 8}")
            ws.cell(row=r, column=3, value=f"=C{t2_start_row}/D{t1_start_row + 8}")
            ws.cell(row=r, column=4, value=f"=D{t2_start_row}/F{t1_start_row + 8}")
        elif idx == 9:
            ws.cell(row=r, column=2, value=f"=(B{t2_start_row + 11}/B{t2_start_row})*100")
            ws.cell(row=r, column=3, value=f"=(C{t2_start_row + 11}/C{t2_start_row})*100")
            ws.cell(row=r, column=4, value=f"=(D{t2_start_row + 11}/D{t2_start_row})*100")
        elif idx == 10:
            ws.cell(row=r, column=2, value=f"=(B{t2_start_row + 11}/B{t1_start_row + 8})*100")
            ws.cell(row=r, column=3, value=f"=(C{t2_start_row + 11}/D{t1_start_row + 8})*100")
            ws.cell(row=r, column=4, value=f"=(D{t2_start_row + 11}/F{t1_start_row + 8})*100")
        elif idx == 11:
            ws.cell(row=r, column=2, value=f"=(B{t2_start_row + 2}/B{t2_start_row + 1})*100")
            ws.cell(row=r, column=3, value=f"=(C{t2_start_row + 2}/C{t2_start_row + 1})*100")
            ws.cell(row=r, column=4, value=f"=(D{t2_start_row + 2}/D{t2_start_row + 1})*100")

        ws.cell(row=r, column=5, value=f"=D{r}-B{r}")
        for c in range(2, 6): ws.cell(row=r, column=c).number_format = "0.00"
        r += 1

    ws.column_dimensions['A'].width = 50
    for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws.column_dimensions[col].width = 18

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# 3. СТОРІНКА STREAMLIT
# ==============================================================================
st.set_page_config(page_title="Фінансовий Аналізатор", page_icon="📈", layout="wide")

st.title("📈 Онлайн-сервіс аналізу фінансової звітності (2023–2025)")
st.markdown(
    "Завантажте 3 PDF-файли звітності підприємства за 2023, 2024 та 2025 роки. Сервіс автоматично сформує Excel-файл з усіма 4-ма аналітичними таблицями.")

col1, col2, col3 = st.columns(3)

with col1:
    pdf_2023 = st.file_uploader("Звітність за 2023 рік (PDF)", type=["pdf"], key="y2023")
with col2:
    pdf_2024 = st.file_uploader("Звітність за 2024 рік (PDF)", type=["pdf"], key="y2024")
with col3:
    pdf_2025 = st.file_uploader("Звітність за 2025 рік (PDF)", type=["pdf"], key="y2025")

if pdf_2023 and pdf_2024 and pdf_2025:
    st.success("Усі 3 PDF-файли успішно завантажено!")

    if st.button("🚀 Сформувати аналітичний Excel-звіт", type="primary"):
        with st.spinner("Зчитуємо PDF та обчислюємо коефіцієнти..."):
            excel_data = generate_excel_report(pdf_2023, pdf_2024, pdf_2025)

            st.download_button(
                label="📥 Завантажити підсумковий Excel-файл",
                data=excel_data,
                file_name="Financial_Analysis_2023_2025.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Будь ласка, завантажте PDF-файли за всі три роки для початку розрахунку.")
