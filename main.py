import sys
import os
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl import load_workbook

sys.stdout.write("Example：1月30日 -> 0130\n")
date_today = input("請輸入今天的日期：")
date_begin = input("請輸入報表的起始日期：")
date_end = input("請輸入報表的結束日期：")

orders_folder = input("請輸入報表存放的資料夾：")

template_path = input("請輸入模板存放的資料夾：")

template_file_name = 'Shopee_template.xlsx'

# 路徑懶得處理 不要有中文

today_orders = f'Order.all.2022{date_begin}_2022{date_end}.xlsx'
file_path = os.path.join(orders_folder, today_orders)

# 今日訂單存放位置
save_folder_path = os.path.join(orders_folder, f'2022{date_today}')
if not os.path.exists(save_folder_path):
    os.mkdir(save_folder_path)

os.chdir(orders_folder)

# 讀取報表
wb = load_workbook(file_path)

# 指定第一個工作頁
order_sheet = wb[wb.sheetnames[0]]

# 自訂欄位順序
sort_target_cols_lst = [
    1, 39, 34, 33, 6, 12, 13, 14, 16, 18, 8, 24, 25, 29, 30, 27, 32, 11, 49, 50, 2, 40
]

consumer_lst = [[0 for x in range(len(sort_target_cols_lst))] for y in range(order_sheet.max_row)]

# 欄位標題 放進去
for col in range(len(sort_target_cols_lst)):
    consumer_lst[0][col] = order_sheet.cell(1, sort_target_cols_lst[col]).value

# 訂單資料
for row in range(order_sheet.max_row):

    # 只要 "待出貨" 訂單
    # 直接 continue 所以這行都是空的
    if order_sheet.cell(row + 1, 2).value != "待出貨":
        continue

    # 按照我要的順序塞進 list (順序 與 template 的欄位順序相關)
    for col in range(len(sort_target_cols_lst)):
        consumer_lst[row][col] = order_sheet.cell(row + 1, sort_target_cols_lst[col]).value

wb.close()

y = 1

# 用到 6 欄 A ~ F
title_lst = ['A', 'B', 'C', 'D', 'E', 'F']

# 中等粗細的框線 黑色
medium = Side(border_style="medium", color="000000")

while y < len(consumer_lst):
    # 【待出貨】以外的訂單
    # 前面塞資料時，直接略過，所以都沒有資料
    # 這邊也略過不處理，直接進到下一筆訂單
    if not consumer_lst[y][0]:
        y += 1
        continue

    # 讀取模板
    sole_order = load_workbook(os.path.join(template_path, template_file_name))

    # 指定第一個工作頁
    current_order = sole_order[sole_order.sheetnames[0]]

    # 訂單編號 收件者姓名 收件者電話 收件者地址 訂單日期 出貨分店(這欄手動添加)
    for i in range(5):
        current_order[f'{title_lst[i]}2'] = consumer_lst[y][i]
        current_order[f'{title_lst[i]}2'].alignment = Alignment(vertical="center")
        current_order[f'{title_lst[i]}2'].font = Font(name='微軟正黑體')

    current_order['D2'].alignment = Alignment(wrap_text=True,
                                              horizontal="center",
                                              vertical="center")

    current_order['F2'] = "44"
    current_order['F2'].font = Font(name='Calibri', color="FF0000", bold=True)
    current_order['F2'].alignment = Alignment(horizontal="center", vertical="center")

    # 蝦皮補貼金額 蝦幣折抵 銀行信用卡活動折抵 賣場優惠券 優惠券 買家支付運費
    for i in range(5, 11):
        current_order[f'{title_lst[i-5]}4'] = consumer_lst[y][i]
        current_order[f'{title_lst[i-5]}4'].font = Font(name='Calibri')
        current_order[f'{title_lst[i-5]}4'].alignment = Alignment(horizontal="right",
                                                                  vertical="center")

    # 商品名稱 商品選項名稱 商品選項貨號 數量 商品活動價格 蝦皮促銷組合折扣:促銷組合標籤
    same_order_row = y
    same_order_count = 0
    cur_order_num = consumer_lst[y][0]

    # 統計一次買多少件商品
    if same_order_row + 1 != len(consumer_lst):
        while consumer_lst[same_order_row + 1][0] == cur_order_num:
            if same_order_row == len(consumer_lst):
                break
            same_order_count += 1
            same_order_row += 1

    goods_index_begin = 6
    goods_index_end = 6 + same_order_count

    for g in range(goods_index_begin, goods_index_end + 1):
        for i in range(11, 17):
            current_order[f'{title_lst[i-11]}{g}'] = consumer_lst[y + g - 6][i]
            current_order[f'{title_lst[i-11]}{g}'].alignment = Alignment(vertical="center")
            current_order[f'{title_lst[i-11]}{g}'].border = Border(top=medium,
                                                                   left=medium,
                                                                   right=medium,
                                                                   bottom=medium)

    for g in range(goods_index_begin, goods_index_end + 1):
        current_order[f'A{g}'] = current_order[f'A{g}'].value.replace(" ToysRUs玩具反斗城", "")
        current_order[f'A{g}'].alignment = Alignment(wrap_text=True)

        current_order[f'C{g}'].alignment = Alignment(horizontal="center", vertical="center")

        current_order[f'D{g}'].font = Font(color="FF0000", bold=True)
        current_order[f'D{g}'].alignment = Alignment(horizontal="center", vertical="center")

        current_order[f'E{g}'].alignment = Alignment(horizontal="right", vertical="center")

        current_order[f'F{g}'].alignment = Alignment(horizontal="center", vertical="center")

    # 買家總支付金額
    current_order[f'A{goods_index_end + 1}'] = consumer_lst[0][17]
    current_order[f'A{goods_index_end + 1}'].alignment = Alignment(vertical="center")
    current_order[f'A{goods_index_end + 1}'].fill = PatternFill(start_color="FFFF00",
                                                                fill_type="solid")

    current_order[f'E{goods_index_end + 1}'] = consumer_lst[y][17]
    current_order[f'E{goods_index_end + 1}'].alignment = Alignment(horizontal="right",
                                                                   vertical="center")

    for i in range(6):
        current_order[f'{title_lst[i]}{goods_index_end + 1}'].border = Border(top=medium,
                                                                              left=medium,
                                                                              right=medium,
                                                                              bottom=medium)

    # 買家備註
    if consumer_lst[y][18]:
        current_order[f'A{goods_index_end + 3}'] = consumer_lst[0][18]
        current_order[f'A{goods_index_end + 3}'].font = Font(color="FFFFFF", bold=True)
        current_order[f'A{goods_index_end + 3}'].fill = PatternFill(start_color="C00000",
                                                                    fill_type="solid")
        current_order[f'A{goods_index_end + 3}'].border = Border(top=medium,
                                                                 left=medium,
                                                                 right=medium,
                                                                 bottom=medium)

        current_order[f'A{goods_index_end + 4}'] = consumer_lst[y][18]
        current_order[f'A{goods_index_end + 4}'].font = Font(name='微軟正黑體')
        current_order[f'A{goods_index_end + 4}'].alignment = Alignment(vertical="center")
        current_order[f'A{goods_index_end + 4}'].border = Border(top=medium,
                                                                 left=medium,
                                                                 right=medium,
                                                                 bottom=medium)

    # 備註
    if consumer_lst[y][19]:
        current_order[f'C{goods_index_end + 3}'] = consumer_lst[0][19]
        current_order[f'C{goods_index_end + 3}'].font = Font(color="FFFFFF")
        current_order[f'C{goods_index_end + 3}'].fill = PatternFill(start_color="C00000",
                                                                    fill_type="solid")
        current_order[f'C{goods_index_end + 3}'].border = Border(top=medium,
                                                                 left=medium,
                                                                 right=medium,
                                                                 bottom=medium)

        current_order[f'C{goods_index_end + 4}'] = consumer_lst[y][19]
        current_order[f'C{goods_index_end + 4}'].font = Font(name='微軟正黑體')
        current_order[f'C{goods_index_end + 4}'].alignment = Alignment(vertical="center")
        current_order[f'C{goods_index_end + 4}'].border = Border(top=medium,
                                                                 left=medium,
                                                                 right=medium,
                                                                 bottom=medium)

    # 寄出方式
    shipping_method = '全家' if consumer_lst[y][-1] == '全家' else '宅配'

    # 訂單成立時間
    order_date = str(consumer_lst[y][4]).split(" ")[0]

    file_name = f'Shopee_{order_date}_{shipping_method}_{str(consumer_lst[y][0])}_{str(consumer_lst[y][1]).replace("*","x")}.xlsx'
    save_file_path = os.path.join(save_folder_path, file_name)
    sole_order.save(save_file_path)
    sole_order.close()

    y += same_order_count + 1
