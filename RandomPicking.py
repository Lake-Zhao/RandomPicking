# ----------导入所需库----------
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from random import randint
import pandas as pd
import sys
import os
import shutil
import ctypes

# ----------初始化----------
print("Loading...")
if getattr(sys, 'frozen', False):
    '''
    此功能由DeepSeek辅助编写
    打包为 .exe 文件时添加 message.log 文件
    检查是否为打包后的可执行文件
    '''
    # 打包运行时将日志文件释放到exe所在目录
    exe_dir = os.path.dirname(sys.executable)
    src_log = os.path.join(sys._MEIPASS, "message.log")
    dst_log = os.path.join(exe_dir, "message.log")
    if not os.path.exists(dst_log):
        shutil.copy(src_log, exe_dir)

# 读取log内容
log = open("message.log", "r+", encoding="utf-8-sig")
content = log.read()
theme_list = content.split("\n")
changed_list = theme_list
log.close()
# 初始化DataFrame
df = pd.DataFrame()
df_changed = df  # 筛选后的DataFrame
df_history = pd.DataFrame(columns=[theme_list[2], theme_list[3], theme_list[5], theme_list[8]])  # 创建历史记录的DataFrame

# 初始化分类列表
first_class_list = ["(无)"]
second_class_list = ["(无)"]

# 创建主窗口
root = tk.Tk()
root.geometry("1000x500")
root.resizable(False, False)
root.title(theme_list[0])
root.attributes("-topmost", 1)
root.update()
root.attributes("-topmost", 0)
print("Setting...")
# 最小化终端窗口
cmd_window_handle = ctypes.windll.kernel32.GetConsoleWindow()
ctypes.windll.user32.ShowWindow(cmd_window_handle, 6)

# ----------全局变量定义----------
# 模式选择 (0=文件模式, 1=数字模式)
var_model_choose = tk.IntVar()
var_model_choose.set(0)
# Excel文件路径
var_file_path = tk.StringVar()
var_file_path.set(theme_list[1])
# 第一分类项
var_first_class = tk.StringVar()
var_first_class.set(theme_list[6])
# 第二分类项
var_second_class = tk.StringVar()
var_second_class.set(theme_list[9])
# 抽取数量
var_number = tk.StringVar()
var_number.set("1")
# 是否允许重复抽取
var_replace = tk.IntVar()
var_replace.set(0)
# 提示信息
var_tip = tk.StringVar()
var_tip.set("")
# 主显示区内容
var_show_main = tk.StringVar()
var_show_main.set("")
# 详情显示区内容
var_show_detail = tk.StringVar()
var_show_detail.set("")
# 数字模式最小值
var_min = tk.StringVar()
var_min.set("0")
# 数字模式最大值
var_max = tk.StringVar()
var_max.set("50")
# 数字模式显示
var_num_show = tk.StringVar()
var_num_show.set("")
# 分类启用的状态变量
var_on_0 = tk.IntVar()
if changed_list[4] == "T":
    var_on_0.set(1)
else:
    var_on_0.set(0)

var_on_1 = tk.IntVar()
if changed_list[7] == "T":
    var_on_1.set(1)
else:
    var_on_1.set(0)

var_on_2 = tk.IntVar()
if changed_list[10] == "T":
    var_on_2.set(1)
else:
    var_on_2.set(0)

# 配置变量
# 窗口标题
var_title = tk.StringVar()
var_title.set(changed_list[0])
# 主抽取项名称
var_def_m = tk.StringVar()
var_def_m.set(changed_list[2])
# 从抽取项名称
var_def_0 = tk.StringVar()
var_def_0.set(changed_list[3])
# 第一分类项名称
var_def_1 = tk.StringVar()
var_def_1.set(changed_list[5])
# 第二分类项名称
var_def_2 = tk.StringVar()
var_def_2.set(changed_list[8])
# 状态标记
is_readfile = False  # 标记Excel文件是否成功读取
is_changed = False  # 标记是否进行修改
old_index = -1  # 旧索引指向-1 (用于避免重复抽取)


def page_change(p):
    """此函数用于切换不同的页面布局 """
    global is_readfile
    if p == "1-2":  # 文件模式 -> 数字模式
        var_tip.set("")
        root.geometry("510x460")
        but_settings.place_forget()
        lab_choose_file.place_forget()
        ent_file_path.place_forget()
        but_choose_file.place_forget()
        but_upload.place_forget()
        lab_first_class.place_forget()
        com_first_class.place_forget()
        lab_second_class.place_forget()
        com_second_class.place_forget()
        lab_number.place_forget()
        ent_number.place_forget()
        che_replace.place_forget()
        ent_show_main.place_forget()
        ent_show_detail.place_forget()
        lis_history.place_forget()
        but_to_history.place_forget()
        but_pick.config(state=tk.NORMAL)
        scr_history.place_forget()

        lab_min.place(x=20, y=50)
        ent_min.place(x=110, y=50)
        lab_max.place(x=220, y=50)
        ent_max.place(x=310, y=50)
        lab_tip.place(x=110, y=90)
        ent_show_num.place(x=20, y=130)
        but_pick.place(x=30, y=370)
        but_quit.place(x=270, y=370)
        but_quit.config(font=("Times", 28), width=9)
        but_pick.config(state=tk.NORMAL)
        scr_history.place(x=977, y=150, height=240)
    if p == "2-1":  # 数字模式 -> 文件模式
        var_tip.set("")
        root.geometry("1000x500")
        lab_model.place(x=20, y=13)
        rad_excel_model.place(x=100, y=13)
        rad_num_model.place(x=240, y=12)
        but_settings.place(x=400, y=8)
        lab_choose_file.place(x=20, y=50)
        ent_file_path.place(x=135, y=50)
        but_choose_file.place(x=635, y=45)
        but_upload.place(x=770, y=45)
        lab_number.place(x=20, y=97)
        ent_number.place(x=90, y=97)
        che_replace.place(x=180, y=94)
        if theme_list[7] == "T":
            lab_first_class.place(x=330, y=97)
            com_first_class.place(x=420, y=97)
        else:
            lab_first_class.place_forget()
            com_first_class.place_forget()
        if theme_list[10] == "T":
            lab_second_class.place(x=510, y=97)
            com_second_class.place(x=600, y=97)
        else:
            lab_second_class.place_forget()
            com_second_class.place_forget()
        lab_tip.place(x=700, y=97)
        ent_show_main.place(x=20, y=150)
        ent_show_detail.place(x=20, y=320)
        lis_history.place(x=720, y=150)
        but_pick.place(x=220, y=400)
        but_quit.place(x=720, y=420)
        but_quit.config(font=("Times", 17), width=8)
        but_to_history.place(x=850, y=420)
        if is_readfile:
            but_pick.config(state=tk.NORMAL)
        else:
            but_pick.config(state=tk.DISABLED)

        lab_min.place_forget()
        ent_min.place_forget()
        lab_max.place_forget()
        ent_max.place_forget()
        ent_show_num.place_forget()
        upload_class_data()
    if p == "1-3":  # 主界面 -> 设置界面
        var_tip.set("")
        root.geometry("500x350")
        lab_model.place_forget()
        rad_excel_model.place_forget()
        rad_num_model.place_forget()
        but_settings.place_forget()
        lab_choose_file.place_forget()
        ent_file_path.place_forget()
        but_choose_file.place_forget()
        but_upload.place_forget()
        lab_first_class.place_forget()
        com_first_class.place_forget()
        lab_second_class.place_forget()
        com_second_class.place_forget()
        lab_number.place_forget()
        ent_number.place_forget()
        che_replace.place_forget()
        ent_show_main.place_forget()
        ent_show_detail.place_forget()
        lis_history.place_forget()
        but_pick.place_forget()
        but_to_history.place_forget()
        but_quit.place_forget()
        scr_history.place_forget()

        lab_settings.place(x=30, y=20)
        lab_title.place(x=45, y=60)
        ent_title.place(x=180, y=60)
        lab_def_m.place(x=45, y=100)
        ent_def_m.place(x=180, y=100)
        lab_def_0.place(x=45, y=140)
        ent_def_0.place(x=180, y=140)
        che_on_0.place(x=290, y=137)
        lab_def_1.place(x=45, y=180)
        ent_def_1.place(x=180, y=180)
        che_on_1.place(x=290, y=177)
        lab_def_2.place(x=45, y=220)
        ent_def_2.place(x=180, y=220)
        che_on_2.place(x=290, y=217)
        but_ok.place(x=200, y=260)
        but_cancel.place(x=330, y=260)
        lab_tip.place(x=45, y=310)
    if p == "3-1":  # 设置界面 -> 主界面
        var_tip.set("")
        root.geometry("1000x500")
        lab_model.place(x=20, y=13)
        rad_excel_model.place(x=100, y=13)
        rad_num_model.place(x=240, y=12)
        but_settings.place(x=400, y=8)
        lab_choose_file.place(x=20, y=50)
        ent_file_path.place(x=135, y=50)
        but_choose_file.place(x=635, y=45)
        but_upload.place(x=770, y=45)
        lab_number.place(x=20, y=97)
        ent_number.place(x=90, y=97)
        che_replace.place(x=180, y=94)
        if theme_list[7] == "T":
            lab_first_class.place(x=330, y=97)
            com_first_class.place(x=420, y=97)
        else:
            lab_first_class.place_forget()
            com_first_class.place_forget()
        if theme_list[10] == "T":
            lab_second_class.place(x=510, y=97)
            com_second_class.place(x=600, y=97)
        else:
            lab_second_class.place_forget()
            com_second_class.place_forget()
        lab_tip.place(x=700, y=97)
        ent_show_main.place(x=20, y=150)
        ent_show_detail.place(x=20, y=320)
        lis_history.place(x=720, y=150)
        but_pick.place(x=220, y=400)
        but_quit.place(x=720, y=420)
        but_to_history.place(x=850, y=420)
        scr_history.place(x=977, y=150, height=240)

        lab_settings.place_forget()
        lab_title.place_forget()
        ent_title.place_forget()
        lab_def_m.place_forget()
        ent_def_m.place_forget()
        lab_def_0.place_forget()
        ent_def_0.place_forget()
        che_on_0.place_forget()
        lab_def_1.place_forget()
        ent_def_1.place_forget()
        che_on_1.place_forget()
        lab_def_2.place_forget()
        ent_def_2.place_forget()
        che_on_2.place_forget()
        but_ok.place_forget()
        but_cancel.place_forget()
    if p == "able":  # 启用抽取功能组件
        ent_number.config(state=tk.NORMAL)
        che_replace.config(state=tk.NORMAL)
        com_first_class.config(state=tk.NORMAL)
        com_second_class.config(state=tk.NORMAL)
        but_pick.config(state=tk.NORMAL)
        var_first_class.set("(无)")
        var_second_class.set("(无)")
        var_show_main.set("")
        var_show_detail.set("")
    if p == "unable":  # 禁用抽取功能组件
        ent_number.config(state=tk.DISABLED)
        che_replace.config(state=tk.DISABLED)
        com_first_class.config(state=tk.DISABLED)
        com_second_class.config(state=tk.DISABLED)
        but_pick.config(state=tk.DISABLED)


def is_not_empty(s):
    """检查字符串是否非空"""
    return s.strip() != ''


def to_list(c):
    """将DataFrame的对应列列转换为去重列表，添加'(无)'选项"""
    global df
    try:
        if c in df.columns:
            c_list = ["(无)"] + df.groupby(c, as_index=False).describe()[c].values.tolist()
        else:
            c_list = ["(无)"]
        i = 0
        while i < len(c_list):
            if not is_not_empty(c_list[i]):
                c_list.pop(i)
                i -= 1
            i += 1
        return c_list
    except Exception as error:
        print(error)
        return ["(无)"]


def isint(s):
    """检查字符串是否能转换为整数"""
    s = str(s)
    try:
        int(s)
        return True
    except ValueError:
        return False


def upload_class_data(event=None):
    """根据选择的分类依据筛选数据"""
    global df, df_changed
    try:
        but_pick.config(state=tk.NORMAL)
        # 获取当前选择的分类
        c1 = var_first_class.get()
        c2 = var_second_class.get()
        class_1 = theme_list[5]
        class_2 = theme_list[8]
        # 根据选择组合筛选条件
        if c1 != "(无)" and c2 != "(无)":
            df_changed = df[(df[class_1] == c1) & (df[class_2] == c2)].reset_index(drop=True)
        else:
            if c1 != "(无)":
                df_changed = df[df[class_1] == c1].reset_index(drop=True)
            elif c2 != "(无)":
                df_changed = df[df[class_2] == c2].reset_index(drop=True)
            else:
                df_changed = df.reset_index(drop=True)
        print("抽取范围修改：")
        print(df_changed)
        theme_list[6] = c1
        theme_list[9] = c2
        if len(df_changed) <= 1:
            var_tip.set("筛选长度小于等于1！")
            but_pick.config(state=tk.DISABLED)
        else:
            var_tip.set("筛选长度为" + str(len(df_changed)))
    except Exception as err:
        print(f"筛选错误:", err)


def pick_one(i):
    """显示单个抽取结果并记录历史"""
    global df_changed, df_history
    # 构建信息
    detail_content = ""
    if is_not_empty(theme_list[3]):
        detail_content += (str(theme_list[3]) + ":" + str(df_changed.at[i, theme_list[3]])
                           + " ")
    if is_not_empty(theme_list[5]):
        detail_content += str(theme_list[5]) + ":" + str(df_changed.at[i, theme_list[5]]) + " "
    if is_not_empty(theme_list[8]):
        detail_content += str(theme_list[8]) + ":" + str(df_changed.at[i, theme_list[8]]) + " "
    # 显示内容并记录历史
    var_show_main.set(str(df_changed.at[i, theme_list[2]]))
    var_show_detail.set(detail_content)
    lis_history.insert(0, str(df_changed.at[i, theme_list[2]]) + " " + detail_content)
    df_history = pd.concat([df_history, df_changed[i:i + 1]], ignore_index=True)


def settings():
    """打开设置页面"""
    page_change("1-3")
    # 根据是否修改过显示不同内容
    if is_changed:
        var_title.set(changed_list[0])
        var_def_m.set(changed_list[2])
        var_def_0.set(changed_list[3])
        var_def_1.set(changed_list[5])
        var_def_2.set(changed_list[8])
        var_tip.set("已恢复配置信息")
    else:
        var_title.set(theme_list[0])
        var_def_m.set(theme_list[2])
        var_def_0.set(theme_list[3])
        var_def_1.set(theme_list[5])
        var_def_2.set(theme_list[8])


def get_file():
    """打开文件选择对话框并获取文件路径 """
    f_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx;*.xls")])
    var_file_path.set(f_path)


def model_change():
    """切换抽取模式"""
    model_num = var_model_choose.get()
    if model_num == 0:
        page_change("2-1")
    if model_num == 1:
        page_change("1-2")


def upload():
    """上传并解析Excel文件"""
    global is_readfile, first_class_list, second_class_list, df, df_changed, theme_list
    try:
        file_path = var_file_path.get()
        # 读取Excel文件
        df = pd.read_excel(file_path, dtype=str, keep_default_na=False)
        df = df.reindex(columns=[theme_list[2], theme_list[3], theme_list[5], theme_list[8]]).fillna("")
        print("读取到文件：")
        print(df)
        df_changed = df
        if not file_path == theme_list[1]:
            page_change("able")
        is_readfile = True
        # 更新分类下拉菜单
        first_class_list = to_list(theme_list[5])
        com_first_class["values"] = first_class_list
        second_class_list = to_list(theme_list[8])
        com_second_class["values"] = second_class_list
        theme_list[1] = file_path
    except FileNotFoundError:
        var_tip.set("文件错误，请重试")
        var_file_path.set(theme_list[1])
        is_readfile = False
    except Exception as err:
        var_tip.set("文件错误，请重试")
        var_file_path.set("")
        print("未知错误：", err)
        is_readfile = False


# 分类项启用状态切换函数
def is_on_ob0():
    fg = var_on_0.get()
    if fg == 0:
        ent_def_0.config(state=tk.DISABLED)
        var_def_0.set("")
    else:
        ent_def_0.config(state=tk.NORMAL)


def is_on_ob1():
    fg = var_on_1.get()
    if fg == 0:
        ent_def_1.config(state=tk.DISABLED)
        var_def_1.set("")
    else:
        ent_def_1.config(state=tk.NORMAL)


def is_on_ob2():
    fg = var_on_2.get()
    if fg == 0:
        ent_def_2.config(state=tk.DISABLED)
        var_def_2.set("")
    else:
        ent_def_2.config(state=tk.NORMAL)


def pick():
    """执行抽取操作"""
    global df_changed, df_history, theme_list, old_index
    but_pick.config(state=tk.DISABLED)
    var_tip.set("")
    model_num = var_model_choose.get()
    # 文件模式抽取
    if model_num == 0:
        pick_num = var_number.get()
        if isint(pick_num):
            num = int(pick_num)
            if num >= 1:
                re_flag = var_replace.get()
                if num == 1:
                    if re_flag == 0:
                        pick_index = randint(0, len(df_changed) - 1)
                        while old_index == pick_index:
                            pick_index = randint(0, len(df_changed) - 1)
                        old_index = pick_index
                    else:
                        pick_index = randint(0, len(df_changed) - 1)
                    pick_one(pick_index)
                else:
                    if re_flag == 0:
                        if num <= len(df_changed):
                            index_list = []
                            while num > 0:
                                pick_index = randint(0, len(df_changed) - 1)
                                while pick_index in index_list:
                                    pick_index = randint(0, len(df_changed) - 1)
                                index_list.append(pick_index)
                                num -= 1
                            print("抽取到条目索引：")
                            print(index_list)
                            for i in index_list:
                                root.after(800, pick_one(i))
                                root.update()
                        else:
                            var_tip.set("数量大于抽取范围！")
                            var_number.set("1")
                    else:
                        index_list = []
                        while num > 0:
                            pick_index = randint(0, len(df_changed) - 1)
                            index_list.append(pick_index)
                            num -= 1
                        print("抽取到条目索引：")
                        print(index_list)
                        for i in index_list:
                            root.after(800, pick_one(i))
                            root.update()
            else:
                var_tip.set("请输入正确的数量！")
                var_number.set("1")
        else:
            var_tip.set("请输入正确的数量！")
            var_number.set("1")
    # 数字模式抽取
    if model_num == 1:
        in_min, in_max = var_min.get(), var_max.get()
        if isint(in_min):
            in_min = int(in_min)
            if isint(in_max):
                in_max = int(in_max)
                if in_min >= in_max:
                    var_min.set("1")
                    var_max.set("100")
                    var_tip.set("最小数不能大于等于最大数！")
                else:
                    var_tip.set("")
                    var_num_show.set(str(randint(in_min, in_max)))
            else:
                var_max.set("100")
                var_tip.set("请输入正确的整数！")
        else:
            var_min.set("1")
            var_tip.set("请输入正确的整数！")
    but_pick.config(state=tk.NORMAL)


def restart():
    """
    此函数由DeepSeek辅助编写
    重启应用程序
    """
    try:
        if getattr(sys, 'frozen', False):
            # 打包后的可执行文件路径
            file_path = sys.executable
        else:
            # 未打包时的脚本路径
            file_path = sys.argv[0]
        os.startfile(file_path)
        print(file_path)
        return True
    except Exception as err:
        messagebox.showerror("重启失败", f"无法重启程序: {str(err)}")
        return False


def root_quit():
    """关闭应用程序"""
    root.quit()
    root.destroy()


def to_history():
    """导出历史记录到Excel"""
    global df_history
    print("历史记录：")
    print(df_history)
    try:
        df_history.to_excel("历史记录.xlsx")
        var_tip.set("历史记录导出成功！")
    except Exception as err:
        var_tip.set("导出失败，请重试")
        print("未知错误：", err)


def s_ok():
    """保存设置并重启应用"""
    var_tip.set("")
    global changed_list, is_changed
    d_title = var_title.get()
    d_m = var_def_m.get()
    d_0 = var_def_0.get()
    d_1 = var_def_1.get()
    d_2 = var_def_2.get()
    if var_on_0.get() == 0:
        c4 = "F"
    else:
        c4 = "T"
        if not is_not_empty(var_def_0.get()):
            var_tip.set("输入的内容不允许为空！")
            var_def_0.set("")
            return False
    if var_on_1.get() == 0:
        c7 = "F"
    else:
        c7 = "T"
        if not is_not_empty(var_def_1.get()):
            var_tip.set("输入的内容不允许为空！")
            var_def_1.set("")
            return False
    if var_on_2.get() == 0:
        c10 = "F"
    else:
        c10 = "T"
        if not is_not_empty(var_def_2.get()):
            var_tip.set("输入的内容不允许为空！")
            var_def_2.set("")
            return False
    changed_list = [d_title, "", d_m, d_0, c4, d_1, "(无)", c7, d_2, "(无)", c10]
    m = tk.messagebox.askyesnocancel(title="确认修改并重启",
                                     message="配置信息已更新\n立即重启后生效\n\n是否立即关闭该程序？")
    if m:
        is_changed = True
        root_quit()
    elif not m:
        is_changed = True
        page_change("3-1")
    else:
        page_change("3-1")


def s_cancel():
    """取消设置更改"""
    page_change("3-1")


# ----------GUI组件初始化----------
# 设置基础组件
lab_model = tk.Label(root, text="模式：", font=("Times", 17))
rad_excel_model = tk.Radiobutton(root, text="文件模式", variable=var_model_choose,
                                 value=0, command=model_change, font=("Times", 17))
rad_num_model = tk.Radiobutton(root, text="数字模式", variable=var_model_choose,
                               value=1, command=model_change, font=("Times", 17))
but_settings = tk.Button(root, text="用户配置", command=settings, font=("Times", 15), width=8)

lab_choose_file = tk.Label(root, text="选择文件：", font=("Times", 17))
ent_file_path = tk.Entry(root, textvariable=var_file_path, font=("Times", 17), width=45)
but_choose_file = tk.Button(root, text="选择", command=get_file, font=("Times", 15), width=8)
but_upload = tk.Button(root, text="上传", command=upload, font=("Times", 15), width=8)

lab_first_class = tk.Label(root, text=theme_list[5] + "：", font=("Times", 17))
com_first_class = ttk.Combobox(root, textvariable=var_first_class, font=("Times", 17), width=6)
com_first_class["values"] = first_class_list
com_first_class.config(state=tk.DISABLED)
com_first_class.bind("<<ComboboxSelected>>", upload_class_data)
lab_second_class = tk.Label(root, text=theme_list[8] + "：", font=("Times", 17))
com_second_class = ttk.Combobox(root, textvariable=var_second_class, font=("Times", 17), width=6)
com_second_class["values"] = second_class_list
com_second_class.config(state=tk.DISABLED)
com_second_class.bind("<<ComboboxSelected>>", upload_class_data)
lab_number = tk.Label(root, text="数量：", font=("Times", 17))
ent_number = tk.Entry(root, textvariable=var_number, font=("Times", 17), width=6)
ent_number.config(state=tk.DISABLED)
che_replace = tk.Checkbutton(root, text="是否重复？", variable=var_replace, onvalue=1, offvalue=0, font=("Times", 17))
che_replace.config(state=tk.DISABLED)
lab_tip = tk.Label(root, textvariable=var_tip, font=("Times", 17))

ent_show_main = tk.Entry(root, textvariable=var_show_main, font=("Times", 100, "bold"),
                         width=10, disabledforeground="black", justify='center')
ent_show_main.config(state=tk.DISABLED, fg="black", bg="white")
ent_show_detail = tk.Entry(root, textvariable=var_show_detail, font=("Times", 40, "bold"),
                           width=25, disabledforeground="black")
ent_show_detail.config(state=tk.DISABLED, fg="black", bg="white")
lis_history = tk.Listbox(root, font=("Times", 13), width=28, height=12)
scr_history = tk.Scrollbar(root, command=lis_history.yview)
lis_history.config(yscrollcommand=scr_history.set)

but_pick = tk.Button(root, text="抽取", command=pick, font=("Times", 28), width=9)
but_pick.config(state=tk.DISABLED)
but_quit = tk.Button(root, text="退出", command=root_quit, font=("Times", 17), width=8)
but_to_history = tk.Button(root, text="导出历史", command=to_history, font=("Times", 17), width=8)

# 排列组件
lab_model.place(x=20, y=13)
rad_excel_model.place(x=100, y=13)
rad_num_model.place(x=240, y=12)
but_settings.place(x=400, y=8)

lab_choose_file.place(x=20, y=50)
ent_file_path.place(x=135, y=50)
but_choose_file.place(x=635, y=45)
but_upload.place(x=770, y=45)

lab_number.place(x=20, y=97)
ent_number.place(x=90, y=97)
che_replace.place(x=180, y=94)
if theme_list[7] == "T":
    lab_first_class.place(x=330, y=97)
    com_first_class.place(x=420, y=97)
else:
    lab_first_class.place_forget()
    com_first_class.place_forget()
if theme_list[10] == "T":
    lab_second_class.place(x=510, y=97)
    com_second_class.place(x=600, y=97)
else:
    lab_second_class.place_forget()
    com_second_class.place_forget()
lab_tip.place(x=700, y=97)

ent_show_main.place(x=20, y=150)
ent_show_detail.place(x=20, y=320)
lis_history.place(x=720, y=150)

but_pick.place(x=220, y=400)
but_quit.place(x=720, y=420)
scr_history.place(x=977, y=150, height=240)
but_to_history.place(x=850, y=420)

# 模式选择的隐藏组件
# 定义
lab_min = tk.Label(root, text="最小值：", font=("Times", 17))
ent_min = tk.Entry(root, textvariable=var_min, font=("Times", 17), width=8, justify='right')
lab_max = tk.Label(root, text="最大值：", font=("Times", 17))
ent_max = tk.Entry(root, textvariable=var_max, font=("Times", 17), width=8, justify='right')
ent_show_num = tk.Entry(root, textvariable=var_num_show, font=("Times", 140, "bold"), width=5,
                        disabledforeground="black", justify='center')
ent_show_num.config(state=tk.DISABLED, fg="black", bg="white")
# 隐藏
lab_min.place_forget()
ent_min.place_forget()
lab_max.place_forget()
ent_max.place_forget()
ent_show_num.place_forget()

# 用户配置的隐藏组件
# 定义
lab_settings = tk.Label(root, text="请自定义抽取项和分类项：", font=("Times", 17))
lab_title = tk.Label(root, text="标题：", font=("Times", 17))
ent_title = tk.Entry(root, textvariable=var_title, font=("Times", 17), width=9)

lab_def_m = tk.Label(root, text="主抽取项：", font=("Times", 17))
ent_def_m = tk.Entry(root, textvariable=var_def_m, font=("Times", 17), width=9)
lab_def_0 = tk.Label(root, text="从抽取项：", font=("Times", 17))
ent_def_0 = tk.Entry(root, textvariable=var_def_0, font=("Times", 17), width=9)
if changed_list[4] == "F":
    ent_def_0.config(state=tk.DISABLED)
che_on_0 = tk.Checkbutton(root, text="是否启用此项？", variable=var_on_0, onvalue=1, offvalue=0,
                          font=("Times", 17), command=is_on_ob0)

lab_def_1 = tk.Label(root, text="第一分类项：", font=("Times", 17))
ent_def_1 = tk.Entry(root, textvariable=var_def_1, font=("Times", 17), width=9)
if changed_list[7] == "F":
    ent_def_1.config(state=tk.DISABLED)
che_on_1 = tk.Checkbutton(root, text="是否启用此项？", variable=var_on_1, onvalue=1, offvalue=0,
                          font=("Times", 17), command=is_on_ob1)
lab_def_2 = tk.Label(root, text="第二分类项：", font=("Times", 17))
ent_def_2 = tk.Entry(root, textvariable=var_def_2, font=("Times", 17), width=9)
if changed_list[10] == "F":
    ent_def_2.config(state=tk.DISABLED)
che_on_2 = tk.Checkbutton(root, text="是否启用此项？", variable=var_on_2, onvalue=1, offvalue=0,
                          font=("Times", 17), command=is_on_ob2)

but_ok = tk.Button(root, text="确定", command=s_ok, font=("Times", 17), width=8)
but_cancel = tk.Button(root, text="取消", command=s_cancel, font=("Times", 17), width=8)

# 隐藏
lab_settings.place_forget()

lab_def_m.place_forget()
ent_def_m.place_forget()
lab_def_0.place_forget()
ent_def_0.place_forget()
che_on_0.place_forget()

lab_def_1.place_forget()
ent_def_1.place_forget()
che_on_1.place_forget()
lab_def_2.place_forget()
ent_def_2.place_forget()
che_on_2.place_forget()

but_ok.place_forget()
but_cancel.place_forget()

try:
    # 尝试初始加载Excel文件
    df = pd.read_excel(theme_list[1], dtype=str, keep_default_na=False)
    df = df.reindex(columns=[theme_list[2], theme_list[3], theme_list[5], theme_list[8]]).fillna("")
    print("读取到文件：")
    print(df)
    df_changed = df
    page_change("able")
    is_readfile = True
    first_class_list = to_list(theme_list[5])
    com_first_class["values"] = first_class_list
    second_class_list = to_list(theme_list[8])
    com_second_class["values"] = second_class_list
    var_tip.set("读取记录成功！")
    upload_class_data()
except FileNotFoundError:
    var_tip.set("欢迎使用！")
    theme_list[1] = ""
    var_file_path.set(theme_list[1])
    page_change("unable")
    is_readfile = False
except Exception as e:
    var_tip.set("欢迎使用！")
    print("未知错误：", e)
    theme_list[1] = ""
    var_file_path.set(theme_list[1])
    page_change("unable")
    is_readfile = False

if theme_list[6] in first_class_list and theme_list[9] in second_class_list:
    # 设置初始分类选择
    var_first_class.set(theme_list[6])
    var_second_class.set(theme_list[9])
else:
    var_first_class.set("(无)")
    var_second_class.set("(无)")

# 启动主事件循环
root.mainloop()

# ----------程序退出处理----------
log = open("message.log", "w", encoding="utf-8-sig")
if is_changed:
    log.write('\n'.join(map(str, changed_list)))
else:
    log.write('\n'.join(map(str, theme_list)))
log.close()
if is_changed:
    restart()
