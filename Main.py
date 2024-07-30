import tkinter as tk
import os
import Functions

# 路径配置
DATABASE_PATH = os.path.join(os.getcwd(), "发票数据库")
REPEAT_DATABASE_PATH = os.path.join(os.getcwd(), '查重台账.xlsx')

if __name__ == '__main__':

    # 主窗口
    MAIN_FRAME = tk.Tk()

    # 标题栏
    MESSAGE_LABEL = tk.Label(master=MAIN_FRAME, text='请输入发票号码。需要同时查重的发票请勾选查重',
                             font=("微软雅黑", 14), pady=14)

    # 文字栏
    CHECK_INVOICE_LABEL = tk.Label(master=MAIN_FRAME, text="请输入发票号码以检查真伪:", font=("微软雅黑", 14),
                                   border=10)
    # 发票输入框
    INVOICE_NUMBER_ENTRY = tk.Entry(master=MAIN_FRAME, font=("微软雅黑", 14), width=25)

    # 是否查重的复选框
    REPEAT_CHECKBOX = tk.Checkbutton(master=MAIN_FRAME, text="是否同时查重", font=("微软雅黑", 14))
    REPEAT_CHECKBOX.select()

    # 启动按钮
    CHECK_INVOICE_BUTTON = tk.Button(master=MAIN_FRAME, text="检查发票", font=("微软雅黑", 14))

    # 结果显示框
    RESULT_LABEL = tk.Label(master=MAIN_FRAME, font=("微软雅黑", 12), anchor='w', pady=12, padx=12, justify='left')

    # 各种路径
    # 初始化用户主界面
    MAIN_FRAME.resizable(width=False, height=False)
    MAIN_FRAME.geometry('1280x720')
    MAIN_FRAME.title('发票查验和查重')
    MESSAGE_LABEL.pack(side=tk.TOP)

    # 初始化显示发票查验提示框
    CHECK_INVOICE_LABEL.pack(side=tk.TOP, fill=tk.X)

    # 初始化发票查验输入框和对应绑定的值
    INVOICE_NUMBER = tk.StringVar()
    INVOICE_NUMBER_ENTRY.pack(side=tk.TOP)
    INVOICE_NUMBER_ENTRY.config(textvariable=INVOICE_NUMBER)

    # 初始化查重复选框并绑定对应的值
    IS_CHECK_REPEAT = tk.IntVar()
    REPEAT_CHECKBOX.config(variable=IS_CHECK_REPEAT)
    REPEAT_CHECKBOX.pack(side=tk.TOP)

    # 初始化按钮 程序要到最后再绑定处理程序
    CHECK_INVOICE_BUTTON.pack(side=tk.TOP)
    RESULT_LABEL.pack(side=tk.TOP, fill=tk.X, )

    # 主程序区域 需要都把元素排布好再做主程序部分
    # 两个核心变量 INVOICE_NUMBER.get() 用来获取发票号码
    # IS_CHECK_REPEAT.get() 用来获取是否被选中

    # 程序启动时在结果框体显示初始化信息
    result_string = ''
    if not os.path.exists(DATABASE_PATH):
        result_string += '无发票数据库，程序无法工作，请检查数据库是否存在\n'

    if not os.path.exists(REPEAT_DATABASE_PATH):
        result_string += '无查重台账，程序无法工作，请检查文件是否存在\n'

    RESULT_LABEL.config(text=result_string)


    # 按钮绑定的主函数，之后将主函数绑定到按钮上
    def start_check():
        CHECK_INVOICE_BUTTON.config(state=tk.DISABLED)
        MAIN_FRAME.update()
        RESULT_LABEL.config(text='')
        # print(IS_CHECK_REPEAT.get())
        # 设置最终的显示字符串
        result = ''
        # 获取发票字符串并检查是否有效
        invoice_number_string = INVOICE_NUMBER_ENTRY.get().strip()
        if not Functions.validate_invoice_number(invoice_number_string):
            result += "发票号码: {} 无效，请输入8位或20位发票号码。".format(invoice_number_string)
            RESULT_LABEL.config(text=result)
            CHECK_INVOICE_BUTTON.config(state=tk.NORMAL)
            INVOICE_NUMBER.set('')
            return

        # 查重的逻辑
        # 如果勾选查重 先到查重表里查找相应的号码，如果没有，则进行查真，查真之后，将发票信息写入查重表
        # 如果勾选查重，找到了查重表里的号码，则提示该发票已经查重过

        # 勾选了查重的情况下
        if IS_CHECK_REPEAT.get():
            # 先到查重表里寻找编码看是否能够找到该号码,如果找到直接组装字符串返回
            repeat_result = Functions.find_repeated_invoice(invoice_number_string)
            if len(repeat_result) > 0:
                result += '发票号码：{} 已经存在于查重台账中，属于重复发票。'.format(invoice_number_string)
                RESULT_LABEL.config(text=result)
                CHECK_INVOICE_BUTTON.config(state=tk.NORMAL)
                INVOICE_NUMBER.set('')
                return
            # 未找到则要先去验真，然后将验真数据写入查重台账，之后返回结果
            else:
                result = '发票号码：{} 不存在于查重台账中，将进行验真。\n'.format(invoice_number_string)
                # 进行验真, 返回查找的结果
                found_string = Functions.find_invoice(invoice_number_string)
                # 验真结果为真, 先写入字符串
                if len(found_string) > 0:
                    result += Functions.assemble_find_invoice_result(found_string, invoice_number_string)
                    # 将结果写入到查重台账中
                    Functions.write_found_invoice_to_repeat_database(found_string)
                    result += "已将上述查找结果写入查重台账"
                    RESULT_LABEL.config(text=result)
                    CHECK_INVOICE_BUTTON.config(state=tk.NORMAL)
                    INVOICE_NUMBER.set('')
                    return
                else:
                    # 验真结果为假和未勾选的结果是一样的，直接交给组装函数
                    RESULT_LABEL.config(
                        text=Functions.assemble_find_invoice_result(found_string, invoice_number_string))
                    CHECK_INVOICE_BUTTON.config(state=tk.NORMAL)
                    INVOICE_NUMBER.set('')
                    return
        else:
            # print("未勾选查重，仅仅验真")
            RESULT_LABEL.config(
                text=Functions.assemble_find_invoice_result(Functions.find_invoice(invoice_number_string),
                                                            invoice_number_string))
            CHECK_INVOICE_BUTTON.config(state=tk.NORMAL)
            INVOICE_NUMBER.set('')
            return


    CHECK_INVOICE_BUTTON.config(command=start_check)

    MAIN_FRAME.mainloop()
