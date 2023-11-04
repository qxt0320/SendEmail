# 导入自定义的函数
from read_excel import read_excel_file
from email_content_generator import generate_email_content, create_grades_entry
from send_email import send_email
import tkinter as tk
from tkinter import filedialog


filepath = ""
root = tk.Tk()
root.geometry("400x200")
root.title("成绩单发送程序")
root.attributes('-alpha', 0.95)

def main():
    # 定义Excel文件路径，假定文件位于项目的data文件夹下
    # excel_file_path = 'data/成绩表.xlsx'

    gui()
    # 读取Excel文件，获取DataFrame
    df = read_excel_file(filepath)

    # 确保DataFrame不为空
    if df is not None:
        # 获取学生学号列表，我们假设学号列存在且每个学号是唯一的
        student_ids = df['学号'].unique()

        # 遍历每个学号，为每位学生生成一封成绩通知邮件
        for student_id in student_ids:
            # 创建该学生的成绩条目列表
            grade_entries = create_grades_entry(df, student_id)

            # 创建该学生的绩点条目列表
            gpa_entries = create_gpa_entry(df, student_id)

            # 从DataFrame中获取该学生的姓名，假设同一个学号对应的姓名是唯一的
            student_name = df[df['学号'] == student_id]['姓名'].iloc[0]

            # 生成电子邮件内容
            email_content = generate_email_content(student_id, student_name, grade_entries)

            # print(email_content)
            # 发送邮件
            send_email(student_name, student_id, email_content)


def end():
    root.destroy()
def gui():
    file_frame = tk.Frame(root)
    select_button = tk.Button(file_frame, text="选择文件", command=select_file, width=15, font=("Arial", 12, "bold"), bg="green", fg="white")
    select_button.pack(side=tk.LEFT, padx=10)

    global filepath_label
    filepath_label = tk.Label(file_frame, text="未选择文件", bg="white", bd=1, relief=tk.SOLID, width=30, fg="black", font=("Arial", 12))
    filepath_label.pack(side=tk.LEFT, padx=20)
    file_frame.pack(pady=40)

    send_button = tk.Button(root, text="发送文件", command=end, font=("Arial", 14, "bold"), bg="green", fg="white")
    send_button.pack(pady=20)

    root.configure(bg="white")
    root.mainloop()

def select_file():
    global filepath_label  # 将filepath_label声明为全局变量
    global filepath
    filepath = filedialog.askopenfilename(initialdir="/", title="选择文件", filetypes=(("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
    if filepath:
        filepath_label.config(text=filepath)


# 确保当该脚本被直接运行时，会调用main函数
if __name__ == "__main__":
    main()
