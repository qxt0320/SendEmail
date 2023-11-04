import re
def create_grades_entry(df, student_id):
    # 根据学生ID筛选出该生的所有成绩
    student_grades = df[df['学号'] == student_id]

    # 创建成绩条目的HTML列表
    grades_entry = '<ul>'
    for _, row in student_grades.iterrows():
        # 检查成绩是否为数值类型
        if isinstance(row['百分成绩'], (int, float)):
            grade = f"{row['百分成绩']}分"
        else:
            # 非数值类型的成绩直接转换为字符串
            grade = str(row['百分成绩'])
        grades_entry += f"<li>{row['课程名称']}: {grade}</li>"
    grades_entry += '</ul>'
    return grades_entry
def create_gpa_entry(df, student_id):
    # 根据学生ID筛选出该生的所有绩点
    student_gpa = df[df['学号'] == student_id]
    # 创建绩点条目的列表
    gpa_entry = []
    for _, row in student_gpa.iterrows():
        gpa = float(row['五分成绩'])
        gpa_entry.append(gpa)
    return gpa_entry

def generate_email_content(student_id, student_name, grade_entries):
    # 将成绩条目转换为 HTML 格式的字符串


    high_grades = []
    low_grades = []

    pattern = r"<li>(.*?)</li>"
    matches = re.findall(pattern, grade_entries)
    for entry in matches:
        course,grade = entry.split(":")
        grade = grade.strip()

        if grade[0].isdigit():
            grade = int(grade.split("分")[0])
            if grade >= 90:
                high_grades.append(course)
            elif grade < 60:
                low_grades.append(course)
    high_grades_str = "，".join(high_grades)
    low_grades_str = "，".join(low_grades)

    grades_str = grade_entries
    email_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            li {{ margin-bottom: 5px; }}
        </style>
    </head>
    <body>
        <p>亲爱的<strong>{student_id}{student_name}</strong>同学:</p>
        <p>祝贺您顺利完成本学期的学习！教务处在此向您发送最新的成绩单。</p>
        <ul>
            {grades_str}
        </ul>
    """

    if high_grades_str:
        email_content += f"<p>您在学科{high_grades_str}表现很好，请继续保持。</p>"

    if low_grades_str:
        email_content += f"<p>您在学科{low_grades_str}需要加油，请继续努力。</p>"

    email_content += """
        <p>希望您能够对自己的成绩感到满意，并继续保持努力和积极的学习态度。如果您在某些科目上没有达到预期的成绩，不要灰心，这也是学习过程中的一部分。我们鼓励您与您的任课教师或辅导员进行交流，他们将很乐意为您解答任何疑问并提供帮助。请记住，学习是一个持续不断的过程，我们相信您有能力克服困难并取得更大的进步。</p>
        <p>再次恭喜您，祝您学习进步、事业成功！</p>
        <p>教务处</p>
    </body>
    </html>
    """
    return email_content





