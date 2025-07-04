from docxtpl import DocxTemplate
import os
import pandas as pd
from num2words import num2words
import calendar
from tkinter import filedialog

class AcceptanceTransferActCreator:
    def extract_metadata(self, wb, delivery_df, is_import):
        from datetime import datetime

        def safe_float(val):
            try:
                return float(str(val).replace(",", ".").replace(" ", ""))
            except:
                return 0.0

        def safe_int(val):
            try:
                return int(float(str(val).replace(",", ".").replace(" ", "")))
            except:
                return 0

        def fmt_date(val):
            try:
                return pd.to_datetime(val).strftime("%d.%m.%Y") if val else ""
            except:
                return ""

        contract_sheet = wb["ContractINFO"]
        sheet = wb["Deposit IMPORT"] if is_import else wb["Deposit EXPORT"]

        d5_date = pd.to_datetime(sheet["D5"].value)
        last_day = calendar.monthrange(d5_date.year, d5_date.month)[1]
        date_act = d5_date.replace(day=last_day).strftime("%d.%m.%Y")

        uk_months_locative = [
            "", "СІЧНІ", "ЛЮТОМУ", "БЕРЕЗНІ", "КВІТНІ",
            "ТРАВНІ", "ЧЕРВНІ", "ЛИПНІ", "СЕРПНІ",
            "ВЕРЕСНІ", "ЖОВТНІ", "ЛИСТОПАДІ", "ГРУДНІ"
        ]
        month_year_ua = f"{uk_months_locative[d5_date.month]} {d5_date.year}"
        month_year_en = f"{d5_date.strftime('%B').upper()} {d5_date.year}"

        def get_cell(sheet, cell):
            return str(sheet[cell].value).strip() if sheet[cell].value else ""

        direction_raw = get_cell(sheet, "E5").upper()
        direction_code = direction_raw.replace(">", "-").replace(" ", "")
        country_code = direction_code.replace("UA", "").replace("-", "") or "XX"

        direction_ua = direction_raw.replace("SK", "Словаччина").replace("UA", "Україна")
        direction_en = direction_raw.replace("SK", "Slovakia").replace("UA", "Ukraine")

        month = f"{d5_date.month:02d}"
        year = d5_date.year
        act_number = "01" if is_import else "02"
        act_num = f"{act_number}-{month}{year}-{country_code}"

        col = 26 if is_import else 54
        if delivery_df.shape[0] <= 72 or delivery_df.shape[1] <= col:
            raise ValueError("Delivery table is missing required rows or columns.")

        total = safe_int(delivery_df.iloc[33, col])
        cost = safe_float(delivery_df.iloc[71, col])
        price = safe_float(delivery_df.iloc[72, col])
        dt = delivery_df.iloc[33, col]

        if is_import:
            row = 3
        else:
            row = 9

        seller_ua = get_cell(contract_sheet, f"B{row}")
        seller_en = get_cell(contract_sheet, f"C{row}")
        sellerstate_ua = get_cell(contract_sheet, f"D{row}")
        sellerstate_en = sellerstate_ua
        sellerrep_ua = get_cell(contract_sheet, f"E{row}")
        sellerrep_en = get_cell(contract_sheet, f"F{row}")
        sellerpos_ua = get_cell(contract_sheet, f"G{row}")
        sellerpos_en = get_cell(contract_sheet, f"H{row}")

        row += 1

        buyer_ua = get_cell(contract_sheet, f"B{row}")
        buyer_en = get_cell(contract_sheet, f"C{row}")
        buyerstate_ua = get_cell(contract_sheet, f"D{row}")
        buyerstate_en = buyerstate_ua
        buyerrep_ua = get_cell(contract_sheet, f"E{row}")
        buyerrep_en = get_cell(contract_sheet, f"F{row}")
        buyerpos_ua = get_cell(contract_sheet, f"G{row}")
        buyerpos_en = get_cell(contract_sheet, f"H{row}")

        efet_date = contract_sheet[f"L{row-1 if is_import else row}"].value
        contract_no = get_cell(contract_sheet, f"K{row}")
        ic_date = contract_sheet[f"L{row}"].value
        context = {
            "EFET_DATE": fmt_date(efet_date),
            "CONTRACT_NO": contract_no,
            "IC_DATE": fmt_date(ic_date),
            "BUYER_UA": buyer_ua,
            "BUYERSTATE_UA": buyerstate_ua,
            "BUYERREP_UA": buyerrep_ua,
            "BUYERPOS_UA": buyerpos_ua,
            "BUYER_EN": buyer_en,
            "BUYERSTATE_EN": buyerstate_en,
            "BUYERREP_EN": buyerrep_en,
            "BUYERPOS_EN": buyerpos_en,
            "SELLER_UA": seller_ua,
            "SELLER_EN": seller_en,
            "SELLERSTATE_UA": sellerstate_ua,
            "SELLERSTATE_EN": sellerstate_en,
            "SELLERREP_UA": sellerrep_ua,
            "SELLERPOS_UA": sellerpos_ua,
            "SELLERREP_EN": sellerrep_en,
            "SELLERPOS_EN": sellerpos_en,
            "MONTH_YEAR_UA": month_year_ua,
            "MONTH_YEAR_EN": month_year_en,
            "DIRECTION_UA": direction_ua,
            "DIRECTION_EN": direction_en,
            "TOTAL": str(total),
            "PRICE": f"{price:,.5f}".replace(",", "\u00A0").replace(".", ","),
            "COST": f"{cost:,.2f}".replace(",", "\u00A0").replace(".", ","),
            "DT": str(dt) if pd.notna(dt) else "-",
            "COST_WORDS_UA": number_to_ua_words(int(round(cost))),
            "COST_WORDS_EN": num2words(cost, lang='en').capitalize() + " hryvnias",
            "DIRECTION_CODE": direction_code,
            "ACT_NUM": act_num,
            "DATE_ACT": date_act,
        }
        return context

    def render_docx(self, template_path, context, output_path):
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)

    def get_default_filename(self, context, is_import):
        direction_code = context.get("DIRECTION_CODE", "DIRECTION")
        month_year_en = context.get("MONTH_YEAR_EN", "MONTH YEAR")
        company = context.get("SELLER_EN" if is_import else "BUYER_EN", "COMPANY").split()[0]
        act_number = "01" if is_import else "02"

        countries = [c.strip() for c in direction_code.split('-') if c.strip()]
        country_code = next((c for c in countries if c != "UA"), "XX")

        date_act = context.get("DATE_ACT", "")
        month = date_act[3:5] if len(date_act) == 10 else "MM"
        year = date_act[6:10] if len(date_act) == 10 else "YYYY"

        code = f"{act_number}-{month}{year}-{country_code}"
        return f"{direction_code}_Acceptance Certificate_{company}_{month_year_en}, {code}.docx".replace(' ', '_')

    def create_act(self, wb, delivery_df, is_import, template_path=None, output_path=None):
        context = self.extract_metadata(wb, delivery_df, is_import)
        default_filename = self.get_default_filename(context, is_import)

        if output_path is None:
            output_path = filedialog.asksaveasfilename(
                title="Зберегти акт приймання-передачі",
                initialfile=default_filename,
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx")]
            )
            if not output_path:
                return

        if template_path is None:
            template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                         'templates', 'word', 'Acceptance Certificate_FORM.docx')

        self.render_docx(template_path, context, output_path)


def number_to_ua_words(amount):
    units = ['', 'одна', 'дві', 'три', 'чотири', 'п\'ять', 'шість', 'сім', 'вісім', 'дев\'ять']
    teens = ['десять', 'одинадцять', 'дванадцять', 'тринадцять', 'чотирнадцять',
             'п\'ятнадцять', 'шістнадцять', 'сімнадцять', 'вісімнадцять', 'дев\'ятнадцять']
    tens = ['', '', 'двадцять', 'тридцять', 'сорок', 'п\'ятдесят', 'шістдесят', 'сімдесят', 'вісімдесят', 'дев\'яносто']
    hundreds = ['', 'сто', 'двісті', 'триста', 'чотириста', 'п\'ятсот', 'шістсот', 'сімсот', 'вісімсот', 'дев\'ятсот']

    if amount == 0:
        return 'нуль'

    words = []
    n = int(amount)
    if n >= 100:
        words.append(hundreds[n // 100])
        n %= 100
    if n >= 20:
        words.append(tens[n // 10])
        n %= 10
    elif n >= 10:
        words.append(teens[n - 10])
        n = 0
    if n > 0:
        words.append(units[n])

    return ' '.join([w for w in words if w])
