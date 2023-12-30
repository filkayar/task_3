import pandas as pd

# Задание 1: Загрузка файла и вывод информации о датафрейме
df_orders = pd.read_excel('orders.xlsx')
# Преобразование столбца с датой в формат DateTime
df_orders['order_date'] = pd.to_datetime(df_orders['order_date'])
rows, columns = df_orders.shape
row_names = df_orders.index.tolist()
column_names = df_orders.columns.tolist()
data_types = df_orders.dtypes

print(f'Количество строк: {rows}')
print(f'Количество столбцов: {columns}')
print(f'Названия строк: {row_names}')
print(f'Названия столбцов: {column_names}')
print(f'Типы данных столбцов:\n{data_types}')

# Задание 2: Вывод столбца со стоимостью проданных товаров
sales_column = df_orders['sales']
average_sales = sales_column.mean()
std_deviation_sales = sales_column.std()
median_sales = sales_column.median()

print(f'Стоимость проданных товаров:\n{sales_column}')
print(f'Средняя стоимость: {average_sales}')
print(f'Стандартное отклонение: {std_deviation_sales}')
print(f'Медиана: {median_sales}')

# Задание 3: Вывод заданных столбцов и строк
two_columns = df_orders[['id', 'order_date']]
specified_rows = df_orders[df_orders['id'].isin([100678, 100706, 100762])]
specified_rows_range = df_orders[100:106]
customer_sales = specified_rows.loc[:, ['customer_id', 'sales']]

print(f'Два столбца:\n{two_columns}')
print(f'Три строки с заданными id:\n{specified_rows}')
print(f'Строки 100-105:\n{specified_rows_range}')
print(f'Данные с id покупателя и стоимостью покупки:\n{customer_sales}')

# Задание 4: Определение количества заказов стоимостью более 43000 долларов
high_cost_orders = df_orders[df_orders['sales'] > 43000]
num_high_cost_orders = len(high_cost_orders)
num_high_cost_orders_shipped_second_class = len(high_cost_orders[high_cost_orders['ship_mode'] == 'Second'])

print(f'Количество заказов стоимостью более 43000 долларов: {num_high_cost_orders}')
print(f'Количество заказов стоимостью более 43000 долларов с доставкой вторым классом: \
    {num_high_cost_orders_shipped_second_class}')

# Задание 5: Сумма заказов для каждого класса доставки
delivery_class_total_sales = df_orders.groupby('ship_mode')['sales'].sum()

print(f'Сумма заказов для каждого класса доставки:\n{delivery_class_total_sales}')

# Задание 6: Группировка данных по классу доставки и дате заказа с подсчетом суммы и количества заказов
grouped_data = df_orders.groupby(['ship_mode', 'order_date']).agg({'sales': 'sum', 'id': 'count'})

print(f'Сгруппированные данные:\n{grouped_data}')

# Задание 7: Сортировка датафрейма по выручке и определение даты с наибольшей выручкой
sales_by_date = df_orders.groupby('order_date')['sales'].sum().reset_index()
sales_by_date = sales_by_date.sort_values('sales', ascending=False)
# Нахождение даты с самой высокой выручкой
max_sales_date = sales_by_date.iloc[0]['order_date']

print(f'Дата с наибольшей выручкой: {max_sales_date}')

# Задание 8: Загрузка данных о клиентах и объединение таблиц
df_customers = pd.read_excel('customers.xlsx')
merged_df = df_orders.merge(df_customers, left_on='customer_id', right_on='id', how='inner')
merged_columns = merged_df.columns.tolist()

print(f'Названия столбцов объединенного датафрейма:\n{merged_columns}')

# Задание 9: Топ-5 городов с самой большой выручкой
revenue_by_city = df_orders[df_orders['order_date'].dt.year == 2016]\
    .merge(df_customers, left_on='customer_id', right_on='id', how='inner').groupby('city')['sales'].sum()
top_five_cities_revenue = revenue_by_city.nlargest(5)

print(f'Пять городов с наибольшей выручкой в 2016 году:\n{top_five_cities_revenue}')

# Задание 10: Количество заказов, отправленных первым классом за последние 5 лет
last_five_years_orders = df_orders[df_orders['order_date'].dt.year >= (df_orders['order_date'].dt.year.max() - 4)]
first_class_orders = len(last_five_years_orders[last_five_years_orders['ship_mode'] == 'First'])

print(f'Количество заказов, отправленных первым классом за последние 5 лет: {first_class_orders}')

# Задания 11: Количество клиентов из Калифорнии
california_customers = len(df_customers[df_customers['state'] == 'California'])

print(f'Количество клиентов из Калифорнии: {california_customers}')

# Задания 12: Количество заказов из Калифорнии
california_orders = len(df_orders[df_orders['customer_id']\
                        .isin(df_customers[df_customers['state'] == 'California']['id'])])

print(f'Количество заказов от клиентов из Калифорнии: {california_orders}')

# Задание 13: Построение сводной таблицы средних чеков по штатам и годам
pivot_table = df_orders.merge(df_customers, left_on='customer_id', right_on='id', how='inner')\
    .pivot_table(values='sales', index='state', columns=df_orders['order_date'].dt.year, aggfunc='mean')

print(f'Сводная таблица средних чеков:\n{pivot_table}')
