import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time


# -------------------------- 固定XPath配置（列名按新规则映射） --------------------------
# 键：新列名（原始分=科目名，等级=首字+级）；值：对应的XPath
score_xpaths = {
    # 原始分（列名=科目名）
    "语文": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[2]",  # 语文原始分→列名“语文”
    "数学": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[3]",  # 数学原始分→列名“数学”
    "英语": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[4]",  # 英语原始分→列名“英语”
    "物理": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[5]",  # 物理原始分→列名“物理”
    "化学": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[6]",  # 化学原始分→列名“化学”
    "道德与法治": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[7]",  # 道德与法治原始分→列名“道德与法治”
    "历史": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[8]",  # 历史原始分→列名“历史”
    "地理": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[9]",  # 地理原始分→列名“地理”
    "生物": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[10]",  # 生物原始分→列名“生物”
    "体育": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[11]",  # 体育原始分→列名“体育”（无等级）

    # 等级（列名=首字+级）
    "语级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[2]",  # 语文等级→“语级”
    "数级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[3]",  # 数学等级→“数级”
    "英级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[4]",  # 英语等级→“英级”
    "物级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[5]",  # 物理等级→“物级”
    "化级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[6]",  # 化学等级→“化级”
    "道级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[7]",  # 道德与法治等级→“道级”（首字“道”）
    "历级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[8]",  # 历史等级→“历级”
    "地级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[9]",  # 地理等级→“地级”
    "生级": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[5]/td[10]",  # 生物等级→“生级”

    # 总分（列名不变）
    "总分": "//*[@id='wrapper']/div[3]/div[2]/table/tbody/tr[3]/td[12]"
}

# 科目分类（有等级/无等级）
subjects_with_grade = [
    ("语文", "语级"), 
    ("数学", "数级"), 
    ("英语", "英级"), 
    ("物理", "物级"), 
    ("化学", "化级"), 
    ("道德与法治", "道级"), 
    ("历史", "历级"), 
    ("地理", "地级"), 
    ("生物", "生级")
]  # 元组：(原始分列名, 等级列名)
subjects_no_grade = ["体育"]  # 仅原始分，列名“体育”
# --------------------------------------------------------------------------------


# 1. 读取考生信息Excel
data = pd.read_excel('考生信息.xlsx')  # 确保文件在代码同目录

# 2. 初始化结果列（按新规则定义列名）
result_columns = []
# 添加有等级的科目列（原始分+等级）
for sub, grade_col in subjects_with_grade:
    result_columns.append(sub)  # 原始分列（如“语文”）
    result_columns.append(grade_col)  # 等级列（如“语级”）
# 添加无等级的科目列（仅原始分）
result_columns.extend(subjects_no_grade)  # 如“体育”
# 添加总分列
result_columns.append("总分")

# 初始化所有结果列
for col in result_columns:
    data[col] = None


# 3. 初始化浏览器
options = webdriver.EdgeOptions()
options.add_argument("--ignore-certificate-errors")  # 忽略证书错误
driver = webdriver.Edge(
    service=Service(EdgeChromiumDriverManager().install()),
    options=options
)
query_url = 'http://27.150.22.198:9001/login.aspx'
driver.get(query_url)
wait = WebDriverWait(driver, 15)  # 延长等待时间


# 4. 逐个处理考生
for index, row in data.iterrows():
    try:
        # 提取考生信息
        ks_ksno = str(row['考生号']).strip()
        zkzh = str(row['准考证号']).strip()
        ks_xm = str(row['姓名']).strip()
        print(f'\n处理第{index+1}条：{ks_xm}（考生号：{ks_ksno}）')

        # 输入信息并查询
        考生号_input = wait.until(EC.presence_of_element_located((By.ID, 'ks_ksno')))
        考生号_input.clear()
        考生号_input.send_keys(ks_ksno)

        准考证号_input = wait.until(EC.presence_of_element_located((By.ID, 'zkzh')))
        准考证号_input.clear()
        准考证号_input.send_keys(zkzh)

        姓名_input = wait.until(EC.presence_of_element_located((By.ID, 'ks_xm')))
        姓名_input.clear()
        姓名_input.send_keys(ks_xm)

        查询按钮 = wait.until(EC.element_to_be_clickable((By.ID, 'Button1')))
        查询按钮.click()
        time.sleep(2)  # 等待结果加载


        # -------------------------- 提取分数（映射到新列名） --------------------------
        for col_name, xpath in score_xpaths.items():  # col_name是新列名（如“语文”“语级”）
            try:
                score = driver.find_element(By.XPATH, xpath).text.strip()
                data.at[index, col_name] = score  # 直接存入新列名
            except Exception as e:
                data.at[index, col_name] = f'未找到：{str(e)[:5]}'


        print(f"✅ 提取完成：{ks_xm}的成绩已记录（列名按新规则显示）")
        driver.back()  # 返回查询页
        time.sleep(1)

    except Exception as e:
        print(f"❌ 处理失败：{str(e)}")
        for col in result_columns:
            data.at[index, col] = '查询失败'
        driver.get(query_url)
        time.sleep(1)


# 5. 保存结果
try:
    data.to_excel('考生成绩结果.xlsx', index=False)
    print('\n结果已保存到“考生成绩结果.xlsx”（列名按新规则显示）')
except PermissionError:
    data.to_excel('考生成绩_备用.xlsx', index=False)
    print('\n原文件被占用，结果已保存到“考生成绩_备用.xlsx”')

driver.quit()