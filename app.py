import streamlit as st
import pandas as pd
import io
import re


# =========================================================
# Helpers
# =========================================================

def extract_product_code(campaign_name):
    """
    استخراج كود المنتج من اسم الحملة.
    مثال:
    KW010102MKIO99 ايربودز لترجمة المحادثات 8-11
    => KW010102MKIO99
    """

    if pd.isna(campaign_name):
        return None

    text = str(campaign_name)

    # إزالة علامات الاتجاه المخفية
    text = text.replace("\u200e", "").replace("\u200f", "").strip()

    # الأكواد الموجودة في الشيتات تبدأ غالباً بـ KW
    match = re.search(r'\bKW[A-Z0-9]+\b', text, re.IGNORECASE)

    if match:
        return match.group(0).upper()

    return None


def extract_product_name(campaign_name, product_code):
    """
    استخراج اسم المنتج من اسم الحملة بعد إزالة:
    - الكلمات قبل الكود مثل Lead / Sales / CPA
    - التاريخ الموجود في نهاية الحملة
    - Copy
    - BID
    """

    if pd.isna(campaign_name) or not product_code:
        return ""

    text = str(campaign_name)

    # إزالة علامات الاتجاه المخفية
    text = text.replace("\u200e", "").replace("\u200f", "")

    # نأخذ النص بعد كود المنتج
    pattern = re.escape(product_code)

    parts = re.split(pattern, text, flags=re.IGNORECASE)

    if len(parts) > 1:
        name = parts[1].strip()
    else:
        name = text.strip()

    # حذف الكلمات الخاصة بنسخ الحملات
    name = re.sub(r'\s*-\s*Copy.*$', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s*-\s*\d+.*$', '', name)
    name = re.sub(r'\bCopy\s*\d*\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\bBID\b', '', name, flags=re.IGNORECASE)

    # حذف التاريخ الموجود غالباً في نهاية اسم الحملة
    # أمثلة: 8-10 / 8-23 / 6-15
    name = re.sub(r'\s+\d{1,2}-\d{1,2}.*$', '', name)

    # تنظيف المسافات
    name = re.sub(r'\s+', ' ', name).strip()

    return name


def get_spend_column(df):
    """
    تحديد عمود الصرف حسب نوع التقرير
    """

    possible_columns = [
        "Amount spent (USD)",
        "Spend",
        "Amount spent",
        "Cost"
    ]

    for col in possible_columns:
        if col in df.columns:
            return col

    return None


def get_orders_column(df):
    """
    تحديد عمود النتائج / الأوردرات.

    في ملفات Meta:
    Results

    في الملفات الأخرى الموجودة عندك:
    Conversions
    """

    # الأولوية للـ Results لو كانت رقمية فعلاً
    if "Results" in df.columns:

        test = pd.to_numeric(
            df["Results"],
            errors="coerce"
        )

        if test.notna().sum() > 0:
            return "Results"

    # في الشيتات اللي رفعتها الأوردرات موجودة هنا
    if "Conversions" in df.columns:
        return "Conversions"

    # احتمالات إضافية
    possible_columns = [
        "Orders",
        "Conversions",
        "Purchases"
    ]

    for col in possible_columns:
        if col in df.columns:
            return col

    return None


def process_file(file):
    """
    قراءة كل Sheets الموجودة داخل الملف
    وتحويلها إلى DataFrame موحد
    """

    all_frames = []

    xls = pd.read_excel(
        file,
        sheet_name=None,
        engine="openpyxl"
    )

    for sheet_name, df in xls.items():

        df = df.dropna(how="all")

        if df.empty:
            continue

        all_frames.append(df)

    if not all_frames:
        return pd.DataFrame()

    return pd.concat(
        all_frames,
        ignore_index=True,
        sort=False
    )


def clean_and_prepare(df, source_file):

    if df.empty:
        return pd.DataFrame()

    if "Campaign name" not in df.columns:
        return pd.DataFrame()

    spend_col = get_spend_column(df)
    orders_col = get_orders_column(df)

    if spend_col is None:
        return pd.DataFrame()

    # إنشاء DataFrame جديد
    result = pd.DataFrame()

    result["اسم الحملة"] = df["Campaign name"].astype(str)

    # استخراج كود المنتج
    result["كود المنتج"] = result["اسم الحملة"].apply(
        extract_product_code
    )

    # حذف أي صف ليس فيه كود منتج
    # وبالتالي صفوف Total of results مش هتدخل
    result = result[
        result["كود المنتج"].notna()
    ].copy()

    # استخراج اسم المنتج
    result["اسم المنتج"] = result.apply(
        lambda row: extract_product_name(
            row["اسم الحملة"],
            row["كود المنتج"]
        ),
        axis=1
    )

    # تحويل الصرف إلى رقم
    result["الصرف"] = pd.to_numeric(
        df.loc[result.index, spend_col],
        errors="coerce"
    ).fillna(0)

    # تحويل الأوردرات إلى رقم
    if orders_col:

        result["الأوردرات"] = pd.to_numeric(
            df.loc[result.index, orders_col],
            errors="coerce"
        ).fillna(0)

    else:
        result["الأوردرات"] = 0

    # تحديد العملة
    if "Currency" in df.columns:

        result["العملة"] = (
            df.loc[result.index, "Currency"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

    else:

        # لو اسم العمود Amount spent USD
        if "USD" in spend_col.upper():
            result["العملة"] = "USD"
        else:
            result["العملة"] = ""

    # لو العملة مش موجودة نحاول نحددها من اسم الملف
    result["العملة"] = result["العملة"].replace(
        {
            "NAN": "",
            "NONE": ""
        }
    )

    # إضافة مصدر الملف
    result["مصدر الملف"] = source_file

    # حذف الحملات اللي صرفها صفر
    result = result[
        result["الصرف"] > 0
    ].copy()

    return result


# =========================================================
# Streamlit App
# =========================================================

st.set_page_config(
    page_title="💎 Campaign Product Analyzer",
    layout="wide"
)

st.title("💎 Campaign Product Analyzer")

st.markdown("""
ارفع تقارير الحملات، والبرنامج هيقوم بـ:

- استخراج كود المنتج من اسم الحملة
- تجميع كل الحملات اللي فيها نفس الكود
- حذف الحملات اللي صرفها صفر
- جمع الصرف بالدولار
- جمع الصرف بالمصري
- تحويل المصري إلى دولار
- حساب إجمالي الصرف النهائي بالدولار
- جمع عدد الأوردرات لكل منتج
""")


# =========================================================
# Exchange Rate
# =========================================================

st.subheader("💱 سعر تحويل الدولار")

exchange_rate = st.number_input(
    "1 USD = كام جنيه مصري؟",
    min_value=1.0,
    value=50.0,
    step=0.5
)


# =========================================================
# Upload Files
# =========================================================

uploaded_files = st.file_uploader(
    "ارفع ملفات Excel الخاصة بالحملات",
    accept_multiple_files=True,
    type=["xlsx"]
)


if uploaded_files:

    all_data = []

    # -----------------------------------------------------
    # قراءة الملفات
    # -----------------------------------------------------

    for file in uploaded_files:

        try:

            df = process_file(file)

            if df.empty:
                continue

            cleaned_df = clean_and_prepare(
                df,
                file.name
            )

            if not cleaned_df.empty:
                all_data.append(cleaned_df)

        except Exception as e:

            st.error(
                f"حصلت مشكلة أثناء قراءة الملف: {file.name}"
            )

            st.error(str(e))


    # =====================================================
    # لو فيه بيانات
    # =====================================================

    if all_data:

        final_data = pd.concat(
            all_data,
            ignore_index=True,
            sort=False
        )


        # =================================================
        # توحيد أسماء العملات
        # =================================================

        final_data["العملة"] = (
            final_data["العملة"]
            .astype(str)
            .str.upper()
            .str.strip()
        )


        # لو فيه ملف عملته USD من اسم العمود
        final_data["العملة"] = final_data["العملة"].replace(
            {
                "US DOLLAR": "USD",
                "$": "USD",
                "EGYPTIAN POUND": "EGP",
                "LE": "EGP"
            }
        )


        # =================================================
        # فصل الصرف بالدولار والمصري
        # =================================================

        final_data["الصرف بالدولار"] = final_data.apply(
            lambda row:
            row["الصرف"]
            if row["العملة"] == "USD"
            else 0,
            axis=1
        )


        final_data["الصرف بالمصري"] = final_data.apply(
            lambda row:
            row["الصرف"]
            if row["العملة"] == "EGP"
            else 0,
            axis=1
        )


        # =================================================
        # تجميع الحملات حسب كود المنتج
        # =================================================

        summary = (
            final_data
            .groupby(
                "كود المنتج",
                as_index=False
            )
            .agg(
                **{
                    "اسم المنتج": (
                        "اسم المنتج",
                        lambda x: next(
                            (
                                str(v)
                                for v in x
                                if pd.notna(v)
                                and str(v).strip() != ""
                            ),
                            ""
                        )
                    ),

                    "إجمالي الصرف بالدولار": (
                        "الصرف بالدولار",
                        "sum"
                    ),

                    "إجمالي الصرف بالمصري": (
                        "الصرف بالمصري",
                        "sum"
                    ),

                    "إجمالي الأوردرات": (
                        "الأوردرات",
                        "sum"
                    ),

                    "عدد الحملات": (
                        "اسم الحملة",
                        "count"
                    ),

                    "الحملات": (
                        "اسم الحملة",
                        lambda x: " | ".join(
                            dict.fromkeys(
                                x.astype(str)
                            )
                        )
                    )
                }
            )
        )


        # =================================================
        # تحويل المصري إلى دولار
        # =================================================

        summary["الصرف المصري بالدولار"] = (
            summary["إجمالي الصرف بالمصري"]
            / exchange_rate
        )


        # =================================================
        # إجمالي الصرف النهائي بالدولار
        # =================================================

        summary["إجمالي الصرف بالدولار النهائي"] = (
            summary["إجمالي الصرف بالدولار"]
            + summary["الصرف المصري بالدولار"]
        )


        # =================================================
        # تكلفة الأوردر
        # =================================================

        summary["تكلفة الأوردر بالدولار"] = (
            summary["إجمالي الصرف بالدولار النهائي"]
            / summary["إجمالي الأوردرات"].replace(0, pd.NA)
        )


        # =================================================
        # ترتيب الأعمدة
        # =================================================

        summary = summary[
            [
                "كود المنتج",
                "اسم المنتج",
                "عدد الحملات",
                "إجمالي الأوردرات",
                "إجمالي الصرف بالدولار",
                "إجمالي الصرف بالمصري",
                "الصرف المصري بالدولار",
                "إجمالي الصرف بالدولار النهائي",
                "تكلفة الأوردر بالدولار",
                "الحملات"
            ]
        ]


        # =================================================
        # تقريب الأرقام
        # =================================================

        numeric_columns = [
            "إجمالي الأوردرات",
            "إجمالي الصرف بالدولار",
            "إجمالي الصرف بالمصري",
            "الصرف المصري بالدولار",
            "إجمالي الصرف بالدولار النهائي",
            "تكلفة الأوردر بالدولار"
        ]

        for col in numeric_columns:

            if col in summary.columns:

                summary[col] = (
                    pd.to_numeric(
                        summary[col],
                        errors="coerce"
                    )
                    .round(2)
                )


        # =================================================
        # ترتيب حسب أعلى صرف
        # =================================================

        summary = summary.sort_values(
            "إجمالي الصرف بالدولار النهائي",
            ascending=False
        )


        # =================================================
        # عرض النتائج
        # =================================================

        st.success(
            "✅ تم تجميع الحملات والمنتجات بنجاح"
        )


        # -------------------------------------------------
        # Metrics
        # -------------------------------------------------

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "عدد المنتجات",
            len(summary)
        )

        col2.metric(
            "إجمالي الأوردرات",
            int(
                summary["إجمالي الأوردرات"].sum()
            )
        )

        col3.metric(
            "إجمالي الصرف بالدولار",
            f"{summary['إجمالي الصرف بالدولار النهائي'].sum():,.2f} $"
        )

        col4.metric(
            "إجمالي الصرف بالمصري",
            f"{summary['إجمالي الصرف بالمصري'].sum():,.2f} EGP"
        )


        # =================================================
        # عرض الجدول
        # =================================================

        st.dataframe(
            summary,
            use_container_width=True
        )


        # =================================================
        # إنشاء ملف Excel
        # =================================================

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # Sheet الملخص
            summary.to_excel(
                writer,
                index=False,
                sheet_name="ملخص المنتجات"
            )


            # Sheet التفاصيل
            details_columns = [
                "كود المنتج",
                "اسم المنتج",
                "اسم الحملة",
                "العملة",
                "الصرف",
                "الأوردرات",
                "مصدر الملف"
            ]

            details = final_data[
                [
                    col for col in details_columns
                    if col in final_data.columns
                ]
            ]

            details.to_excel(
                writer,
                index=False,
                sheet_name="تفاصيل الحملات"
            )


            # -------------------------------------------------
            # تنسيق Excel
            # -------------------------------------------------

            workbook = writer.book

            for sheet_name in workbook.sheetnames:

                worksheet = workbook[sheet_name]

                # تثبيت الصف الأول
                worksheet.freeze_panes = "A2"

                # Auto width
                for column_cells in worksheet.columns:

                    max_length = 0
                    column_letter = column_cells[0].column_letter

                    for cell in column_cells:

                        try:
                            cell_length = len(
                                str(cell.value)
                            )

                            if cell_length > max_length:
                                max_length = cell_length

                        except:
                            pass

                    worksheet.column_dimensions[
                        column_letter
                    ].width = min(
                        max_length + 2,
                        50
                    )


        output.seek(0)


        # =================================================
        # Download Button
        # =================================================

        st.download_button(
            label="⬇️ تحميل تقرير المنتجات Excel",
            data=output.getvalue(),
            file_name="Campaign_Product_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


    else:

        st.warning(
            "⚠️ لم يتم العثور على حملات فيها صرف أكبر من صفر."
        )
