# 导入自定义的函数
from read_excel import read_excel_file
from email_content_generator import generate_email_content, create_grades_entry, create_gpa_entry
from send_email import send_email


def main():
    # 定义Excel文件路径，假定文件位于项目的data文件夹下
    excel_file_path = 'data/成绩表.xlsx'

    # 读取Excel文件，获取DataFrame
    df = read_excel_file(excel_file_path)

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
            email_content = generate_email_content(student_id, student_name, grade_entries, gpa_entries)

            # 发送邮件
            send_email(student_name, student_id, email_content)



# 确保当该脚本被直接运行时，会调用main函数
if __name__ == "__main__":
    main()
