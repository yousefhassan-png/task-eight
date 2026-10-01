import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# تحديد مكان قاعدة البيانات ومجلد حفظ النتائج
مسار_القاعدة = Path("store.db")
مجلد_النتائج = Path(".")

# الاتصال بقاعدة البيانات
اتصال = sqlite3.connect(مسار_القاعدة)

# 1. الكمية المباعة شهريًا
استعلام_الكمية = """
SELECT strftime('%Y-%m', sale_date) AS الشهر,
       SUM(quantity) AS إجمالي_الكمية
FROM sales
GROUP BY strftime('%Y-%m', sale_date)
ORDER BY الشهر
"""
الكمية_الشهرية = pd.read_sql_query(استعلام_الكمية, اتصال)

# 2. الإيراد الشهري
استعلام_الإيراد = """
SELECT strftime('%Y-%m', s.sale_date) AS الشهر,
       SUM(s.quantity * p.price) AS الإيراد
FROM sales AS s
JOIN products AS p ON s.product_id = p.product_id
GROUP BY strftime('%Y-%m', s.sale_date)
ORDER BY الشهر
"""
الإيراد_الشهري = pd.read_sql_query(استعلام_الإيراد, اتصال)

# 3. الإيراد حسب الفئة
استعلام_الفئات = """
SELECT p.category AS الفئة,
       SUM(s.quantity * p.price) AS الإيراد
FROM sales AS s
JOIN products AS p ON s.product_id = p.product_id
GROUP BY p.category
ORDER BY الإيراد DESC
"""
إيراد_الفئات = pd.read_sql_query(استعلام_الفئات, اتصال)

# 4. الإيراد حسب المدينة
استعلام_المدن = """
SELECT s.city AS المدينة,
       SUM(s.quantity * p.price) AS الإيراد
FROM sales AS s
JOIN products AS p ON s.product_id = p.product_id
GROUP BY s.city
ORDER BY الإيراد DESC
"""
إيراد_المدن = pd.read_sql_query(استعلام_المدن, اتصال)

# 5. المنتج الأعلى إيرادًا في كل مدينة
استعلام_أفضل_منتج = """
WITH إيراد_المنتجات AS (
    SELECT s.city AS المدينة, p.product_name AS المنتج,
           SUM(s.quantity * p.price) AS الإيراد
    FROM sales AS s
    JOIN products AS p ON s.product_id = p.product_id
    GROUP BY s.city, p.product_id, p.product_name
),
ترتيب_المنتجات AS (
    SELECT المدينة, المنتج, الإيراد,
           ROW_NUMBER() OVER (PARTITION BY المدينة ORDER BY الإيراد DESC) AS الترتيب
    FROM إيراد_المنتجات
)
SELECT المدينة, المنتج, الإيراد
FROM ترتيب_المنتجات
WHERE الترتيب = 1
ORDER BY المدينة
"""
أفضل_منتج_لكل_مدينة = pd.read_sql_query(استعلام_أفضل_منتج, اتصال)

# 6. متوسط قيمة عملية البيع حسب الفئة
استعلام_المتوسط = """
SELECT p.category AS الفئة,
       AVG(s.quantity * p.price) AS متوسط_قيمة_العملية
FROM sales AS s
JOIN products AS p ON s.product_id = p.product_id
GROUP BY p.category
ORDER BY متوسط_قيمة_العملية DESC
"""
متوسط_البيع = pd.read_sql_query(استعلام_المتوسط, اتصال)

# إعداد مؤشرات الملخص
إجمالي_الإيراد = float(الإيراد_الشهري["الإيراد"].sum()) if not الإيراد_الشهري.empty else 0
أعلى_شهر_كمية = الكمية_الشهرية.loc[الكمية_الشهرية["إجمالي_الكمية"].idxmax()] if not الكمية_الشهرية.empty else None
أعلى_شهر_إيراد = الإيراد_الشهري.loc[الإيراد_الشهري["الإيراد"].idxmax()] if not الإيراد_الشهري.empty else None
أعلى_فئة = إيراد_الفئات.iloc[0] if not إيراد_الفئات.empty else None
أعلى_مدينة = إيراد_المدن.iloc[0] if not إيراد_المدن.empty else None

# إنشاء الرسم الأول: الكمية الشهرية
plt.figure(figsize=(10, 5))
plt.plot(الكمية_الشهرية["الشهر"], الكمية_الشهرية["إجمالي_الكمية"], marker="o")
plt.title("Monthly Sales Quantity")
plt.xlabel("Month")
plt.ylabel("Quantity")
plt.xticks(rotation=45, ha="right")
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("monthly_quantity.png", dpi=150)
plt.close()

# إنشاء الرسم الثاني: الإيراد الشهري
plt.figure(figsize=(10, 5))
plt.plot(الإيراد_الشهري["الشهر"], الإيراد_الشهري["الإيراد"], marker="o")
plt.title("Monthly Revenue")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.xticks(rotation=45, ha="right")
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("monthly_revenue.png", dpi=150)
plt.close()

# تصدير جميع النتائج إلى ملف Excel متعدد الصفحات
with pd.ExcelWriter("sales_report.xlsx", engine="openpyxl") as كاتب:
    ملخص = pd.DataFrame({
        "المؤشر": ["إجمالي الإيراد", "أعلى شهر من حيث الكمية", "أعلى شهر من حيث الإيراد",
                   "أعلى فئة إيرادًا", "أعلى مدينة إيرادًا"],
        "القيمة": [
            إجمالي_الإيراد,
            f'{أعلى_شهر_كمية["الشهر"]} ({أعلى_شهر_كمية["إجمالي_الكمية"]})' if أعلى_شهر_كمية is not None else "لا توجد بيانات",
            f'{أعلى_شهر_إيراد["الشهر"]} ({أعلى_شهر_إيراد["الإيراد"]:.2f})' if أعلى_شهر_إيراد is not None else "لا توجد بيانات",
            f'{أعلى_فئة["الفئة"]} ({أعلى_فئة["الإيراد"]:.2f})' if أعلى_فئة is not None else "لا توجد بيانات",
            f'{أعلى_مدينة["المدينة"]} ({أعلى_مدينة["الإيراد"]:.2f})' if أعلى_مدينة is not None else "لا توجد بيانات"
        ]
    })
    ملخص.to_excel(كاتب, sheet_name="Summary", index=False)
    الكمية_الشهرية.to_excel(كاتب, sheet_name="Monthly Quantity", index=False)
    الإيراد_الشهري.to_excel(كاتب, sheet_name="Monthly Revenue", index=False)
    إيراد_الفئات.to_excel(كاتب, sheet_name="Category Revenue", index=False)
    إيراد_المدن.to_excel(كاتب, sheet_name="City Revenue", index=False)
    أفضل_منتج_لكل_مدينة.to_excel(كاتب, sheet_name="Top Product by City", index=False)
    متوسط_البيع.to_excel(كاتب, sheet_name="Average Sale", index=False)

# كتابة الملخص التنفيذي في فقرتين
نص_الملخص = (
    f"بلغ إجمالي الإيرادات المسجلة في قاعدة البيانات {إجمالي_الإيراد:.2f}. "
    + (f"وسجل شهر {أعلى_شهر_إيراد['الشهر']} أعلى إيراد بقيمة {أعلى_شهر_إيراد['الإيراد']:.2f}، "
       f"بينما كان شهر {أعلى_شهر_كمية['الشهر']} الأعلى من حيث الكمية المباعة "
       f"بإجمالي {أعلى_شهر_كمية['إجمالي_الكمية']} وحدة. "
       if أعلى_شهر_إيراد is not None and أعلى_شهر_كمية is not None else "")
    + "\n\n"
    + (f"على مستوى توزيع الإيرادات، حققت فئة {أعلى_فئة['الفئة']} أعلى إيراد "
       f"({أعلى_فئة['الإيراد']:.2f})، وكانت مدينة {أعلى_مدينة['المدينة']} الأعلى إيرادًا "
       f"({أعلى_مدينة['الإيراد']:.2f}). "
       if أعلى_فئة is not None and أعلى_مدينة is not None else "")
    + "تساعد هذه النتائج على متابعة تغير المبيعات وتحديد الفئات والمناطق التي تستحق دراسة أعمق. "
      "ويجب تفسير الفروق بين الكمية والإيراد بالرجوع إلى أسعار المنتجات ومزيج المنتجات المباع، "
      "كما أن النتائج تصف البيانات المتاحة فقط ولا تثبت أسباب التغير."
)
with open("executive_summary.txt", "w", encoding="utf-8") as ملف:
    ملف.write(نص_الملخص)

اتصال.close()
print("تم إنشاء: sales_report.xlsx و monthly_quantity.png و monthly_revenue.png و executive_summary.txt")
print(نص_الملخص)
