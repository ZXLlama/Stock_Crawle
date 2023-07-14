import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import Scale
import xml.etree.ElementTree as ET
import tkinter.messagebox as messagebox
import datetime
import html

def update_xml():
    # 獲取用戶界面中的數值
    start_year = start_year_combobox.get()
    start_month = start_month_combobox.get()
    end_year = end_year_combobox.get()
    end_month = end_month_combobox.get()
    stock_no = stock_no_entry.get()
    excel_name = html.unescape(excel_name_entry.get())
    delay_sec = delay_scale.get()

    # 判斷起始年月是否小於結束年月，若是則交換
    if int(start_year) > int(end_year) or (int(start_year) == int(end_year) and int(start_month) > int(end_month)):
        start_year, end_year = end_year, start_year
        start_month, end_month = end_month, start_month
        messagebox.showwarning("錯誤", "起始年月份不能大於結束年月份！已自動調整。")

    # 更新XML文件中的數值
    tree = ET.parse('data.xml')
    root = tree.getroot()
    root.find('startYear').text = start_year
    root.find('startMonth').text = start_month
    root.find('endYear').text = end_year
    root.find('endMonth').text = end_month
    root.find('stockNo').text = stock_no
    root.find('excelName').text = excel_name
    root.find('delaySec').text = str(delay_sec)
    tree.write('data.xml', encoding='UTF-8')  # 指定編碼為UTF-8
    messagebox.showinfo("更新成功", "XML文件已成功更新！")

# 創建主窗口
window = tk.Tk()
window.title("每日股票竊取器")

# 設置窗口寬度
window.geometry("400x400")

# 創建輸入字段和標簽
start_year_label = tk.Label(window, text="起始年份:")
start_year_label.pack()

current_year = datetime.datetime.now().year
start_year_values = [str(year) for year in range(2010, current_year + 1)]
start_year_combobox = Combobox(window, values=start_year_values)
start_year_combobox.pack()

start_month_label = tk.Label(window, text="起始月份:")
start_month_label.pack()
start_month_combobox = Combobox(window, values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"])
start_month_combobox.pack()

end_year_label = tk.Label(window, text="結束年份:")
end_year_label.pack()
end_year_combobox = Combobox(window, values=start_year_values)
end_year_combobox.pack()

end_month_label = tk.Label(window, text="結束月份:")
end_month_label.pack()
end_month_combobox = Combobox(window, values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"])
end_month_combobox.pack()

stock_no_label = tk.Label(window, text="股票代碼:")
stock_no_label.pack()
stock_no_entry = tk.Entry(window)
stock_no_entry.pack()

excel_name_label = tk.Label(window, text="Excel名稱:")
excel_name_label.pack()
excel_name_entry = tk.Entry(window)
excel_name_entry.pack()

# 創建滑塊用於調整delaySec參數
delay_label = tk.Label(window, text="延遲秒數:")
delay_label.pack()
delay_scale = Scale(window, from_=0.1, to=5, resolution=0.01, orient=tk.HORIZONTAL, length=300)
delay_scale.set(1.0)  # 默認值為1.0
delay_scale.pack()

# 創建更新按鈕
update_button = tk.Button(window, text="更新", command=update_xml, width=10)
update_button.pack(pady=10)

# 啟動主循環
window.mainloop()
