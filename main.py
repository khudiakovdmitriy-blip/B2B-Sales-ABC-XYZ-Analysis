import os
import glob
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
excel_files = [f for f in glob.glob(os.path.join(BASE_DIR, '*.xlsx')) if not os.path.basename(f).startswith('~$') and 'Results' not in f]

if not excel_files:
    raise FileNotFoundError(f"В папке {BASE_DIR} не найдено ни одного Excel-файла! Список файлов в папке: {os.listdir(BASE_DIR)}")

FILE_PATH = excel_files[0]
print(f"Успешно найден файл: {FILE_PATH}")

fact_sales = pd.read_excel(FILE_PATH, sheet_name='Fact_Sales')
dim_products = pd.read_excel(FILE_PATH, sheet_name='Dim_Products')

print(f"Всего строк в Fact_Sales: {len(fact_sales):,}")

# 2. Очистка от возвратов (учитываем только чистые продажи)
clean_sales = fact_sales[fact_sales['Is_Returned'] == 0].copy()
clean_sales['Order_Date'] = pd.to_datetime(clean_sales['Order_Date'])
clean_sales['Year_Month'] = clean_sales['Order_Date'].dt.to_period('M')

print(f"Строк после исключения возвратов: {len(clean_sales):,}")

# 3. Агрегация показателей по каждому товару (Product_ID)
prod_metrics = (
    clean_sales.groupby('Product_ID', as_index=False)
    .agg(
        Total_Revenue=('Revenue_UAH', 'sum'),
        Total_Profit=('Gross_Profit_UAH', 'sum'),
        Total_Qty=('Quantity', 'sum'),
        Orders_Count=('Order_ID', 'nunique')
    )
)

# Присоединяем справочник товаров (чтобы учесть даже товары с нулевыми продажами, если есть)
abc_xyz = dim_products[
    ['Product_ID', 'SKU_Code', 'Product_Name', 'Category', 'Subcategory', 'Brand', 'Demand_Pattern']].merge(
    prod_metrics, on='Product_ID', how='left'
).fillna({'Total_Revenue': 0, 'Total_Profit': 0, 'Total_Qty': 0, 'Orders_Count': 0})


# 4. Многофакторный ABC-анализ (по Выручке и по Валовой прибыли)
def assign_abc(df, value_col, group_col_name):
    df = df.sort_values(by=value_col, ascending=False).reset_index(drop=True)
    total_val = df[value_col].sum()
    df[f'{value_col}_Share'] = df[value_col] / total_val
    df[f'{value_col}_CumShare'] = df[f'{value_col}_Share'].cumsum()

    conditions = [
        df[f'{value_col}_CumShare'] <= 0.80,
        df[f'{value_col}_CumShare'] <= 0.95
    ]
    df[group_col_name] = np.select(conditions, ['A', 'B'], default='C')
    return df


abc_xyz = assign_abc(abc_xyz, 'Total_Revenue', 'ABC_Revenue')
abc_xyz = assign_abc(abc_xyz, 'Total_Profit', 'ABC_Profit')
abc_xyz['ABC_Combined'] = abc_xyz['ABC_Revenue'] + abc_xyz['ABC_Profit']

# 5. XYZ-анализ (коэффициент вариации помесячного спроса в штуках)
# Создаем полную сетку "Все товары x Все месяцы", чтобы нулевые продажи в месяц не терялись
all_months = pd.period_range(clean_sales['Year_Month'].min(), clean_sales['Year_Month'].max(), freq='M')
grid = pd.MultiIndex.from_product([dim_products['Product_ID'], all_months],
                                  names=['Product_ID', 'Year_Month']).to_frame(index=False)

monthly_demand = (
    clean_sales.groupby(['Product_ID', 'Year_Month'], as_index=False)['Quantity']
    .sum()
)

monthly_full = grid.merge(monthly_demand, on=['Product_ID', 'Year_Month'], how='left').fillna({'Quantity': 0})

# Считаем среднее, стандартное отклонение и коэффициент вариации CV (%)
xyz_stats = (
    monthly_full.groupby('Product_ID', as_index=False)['Quantity']
    .agg(Mean_Monthly_Qty='mean', Std_Monthly_Qty='std')
)
xyz_stats['CV_Percent'] = np.where(
    xyz_stats['Mean_Monthly_Qty'] > 0,
    (xyz_stats['Std_Monthly_Qty'] / xyz_stats['Mean_Monthly_Qty']) * 100,
    999.0  # Для товаров без продаж
)

# Классические пороги XYZ: X <= 10%, Y <= 25%, Z > 25%
xyz_conditions = [
    xyz_stats['CV_Percent'] <= 10.0,
    xyz_stats['CV_Percent'] <= 25.0
]
xyz_stats['XYZ_Group'] = np.select(xyz_conditions, ['X', 'Y'], default='Z')

# Объединяем ABC и XYZ в единую витрину
abc_xyz = abc_xyz.merge(xyz_stats[['Product_ID', 'Mean_Monthly_Qty', 'CV_Percent', 'XYZ_Group']], on='Product_ID',
                        how='left')
abc_xyz['ABC_XYZ'] = abc_xyz['ABC_Revenue'] + abc_xyz['XYZ_Group']

# 6. Бизнес-рекомендации по управлению запасами для каждой из 9 групп
recommendations = {
    'AX': 'Высокий вклад в выручку, стабильный спрос. Автозаказ (JIT), минимальный страховой запас, жесткий контроль наличия.',
    'AY': 'Высокий вклад, сезонные колебания. Планирование закупок с учетом сезонности, повышенный страховой запас.',
    'AZ': 'Высокий вклад, непредсказуемый спрос. Поставка под заказ клиента или частое прогнозирование, контроль рисков неликвида.',
    'BX': 'Средний вклад, стабильный спрос. Стандартная система пополнения запасов (Min-Max), автоматизация закупок.',
    'BY': 'Средний вклад, умеренные колебания. Увеличенный страховой запас перед пиками сезона.',
    'BZ': 'Средний вклад, нерегулярный спрос. Закупка малыми партиями или по предзаказу.',
    'CX': 'Низкий вклад, стабильный спрос. Редкие крупные заказы для экономии на логистике (максимальный интервал поставки).',
    'CY': 'Низкий вклад, сезонный спрос. Минимальные остатки вне сезона, распродажа остатков в конце сезона.',
    'CZ': 'Низкий вклад, хаотичный спрос. Кандидаты на вывод из ассортимента (де-листинг) или работа строго под предоплату.'
}
abc_xyz['Inventory_Strategy'] = abc_xyz['ABC_XYZ'].map(recommendations)

# 7. Сводная матрица 9 групп (количество SKU, выручка, прибыль, доля выручки)
matrix_summary = (
    abc_xyz.groupby('ABC_XYZ', as_index=False)
    .agg(
        SKU_Count=('Product_ID', 'count'),
        Revenue_UAH=('Total_Revenue', 'sum'),
        Profit_UAH=('Total_Profit', 'sum'),
        Avg_CV_Percent=('CV_Percent', 'mean')
    )
)
matrix_summary['Revenue_Share_%'] = (matrix_summary['Revenue_UAH'] / matrix_summary['Revenue_UAH'].sum() * 100).round(2)
matrix_summary['Profit_Share_%'] = (matrix_summary['Profit_UAH'] / matrix_summary['Profit_UAH'].sum() * 100).round(2)

print("\n=== СВОДНАЯ МАТРИЦА ABC-XYZ (9 ГРУПП) ===")
print(matrix_summary.to_string(index=False))

# Проверка скрытых инсайтов: товары группы A по выручке, но C по прибыли (высокие скидки / низкая маржа)
hidden_margin_killers = abc_xyz[(abc_xyz['ABC_Revenue'] == 'A') & (abc_xyz['ABC_Profit'].isin(['B', 'C']))]
print(f"\nНайдено SKU группы A по выручке, но B/C по валовой прибыли: {len(hidden_margin_killers)}")

# Валидация расчетного XYZ против эталонного Demand_Pattern из справочника
print("\n=== СВЕРКА РАСЧЕТНОГО XYZ С ПРОФИЛЕМ СПРОСА В СПРАВОЧНИКЕ ===")
print(pd.crosstab(abc_xyz['Demand_Pattern'], abc_xyz['XYZ_Group']))

# 8. Экспорт итоговой таблицы для загрузки в Power BI и отчета
OUTPUT_PATH = r'C:\Users\Дмитрий\Desktop\Проект\ABC_XYZ_Results.xlsx'
with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
    abc_xyz.to_excel(writer, sheet_name='ABC_XYZ_Products', index=False)
    matrix_summary.to_excel(writer, sheet_name='Matrix_9_Groups', index=False)

print(f"\nГотово! Результаты сохранены в файл: {OUTPUT_PATH}")