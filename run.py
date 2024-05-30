from DrissionPage import ChromiumPage, ChromiumOptions
from DrissionPage.common import Actions
import ddddocr
import pandas
import numpy
import re
import os

pandas.set_option('display.max_rows', None)  # 显示所有行
pandas.set_option('display.max_columns', None)  # 显示所有列
pandas.set_option('display.max_colwidth', None) # 调整列宽

# 更改工作路径(个人习惯，可删，如果遇到问题，改回工作路径即可)
os.chdir()
# 名单目录路径(格式看README.MD)
excel=''
# 账号/密码/班级
account=''
password=''
your_class=''
# 大学习官网
url='https://youth.cq.cqyl.org.cn/login'

def get_data():
    # 请改为你电脑内Chrome可执行文件路径(Edge)
    co=ChromiumOptions().set_browser_path(path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")
    page = ChromiumPage(addr_or_opts=co)

    # 打开网页
    page.get(url=url,retry=3,interval=2,timeout=10)
    page.wait.load_start()
    
    # 过验证码
    ocr=ddddocr.DdddOcr(use_gpu=True,show_ad=False)
    if page.ele('xpath://img[@alt="验证码"]'):
        img=page.ele('xpath://img[@alt="验证码"]').src()
        res=ocr.classification(img)

    # 输入信息
    if page.ele('xpath://input[@placeholder="请输入账号/手机号"]').attr('value') == account:
        if page.ele('xpath://input[@placeholder="请输入验证码"]'):
            page.ele('xpath://input[@placeholder="请输入验证码"]').input(res)
        page.ele('xpath://input[@placeholder="请输入密码"]').input(password)
    else:
        page.ele('xpath://input[@placeholder="请输入账号/手机号"]').input(account)
        if page.ele('xpath://input[@placeholder="请输入验证码"]'):
            page.ele('xpath://input[@placeholder="请输入验证码"]').input(res)
        page.ele('xpath://input[@placeholder="请输入密码"]').input(password)
    # login
    

    # 跳转 开始登录
    while page.ele('登 录').click():
        if page.wait.ele_displayed('xpath://div[@class="commonUse___rGa9d"]//div[@class="menuIcon___bpCSk no_icon___utWZq"]',timeout=None,raise_err=None):
            page.ele('xpath://div[@class="menuIcon___bpCSk no_icon___utWZq"]').click()
            # 打开输入框
            page.ele('xpath://div[@class="ant-select ant-tree-select ant-select-outlined ant-select-in-form-item css-1kllxaf ant-select-single ant-select-allow-clear ant-select-show-arrow ant-select-show-search"]//div[@class="ant-select-selector"]').click()
            
            # 选择班级
            page.ele('xpath://span[@title=\"{}\"]'.format(your_class)).click()
            break

        else:
            page.wait(1)
    
    """# 创建动作链对象
    ac = Actions(page)
    # 左键按住元素
    ac.hold('xpath://div[@class="ant-select-tree-list-scrollbar-thumb"]')
    # 向下移动鼠标300像素
    ac.down(200)
    # 释放左键
    ac.release()"""
    
    while page.ele('查 询').click():
        if page.wait.ele_displayed('xpath://div[@class="ant-card-body"]//tr[@class="ant-table-row ant-table-row-level-0 lightd-table-row"]',timeout=None,raise_err=None):
            break
        else:
            page.wait(1)
     
    # 确定页数
    s=page.ele('xpath://span[@class="lightd-table-pagination-show-total"]').text
    x=re.search(pattern='\s(\d+)\s',string=s)
    num=int(x.group(1))
    print('当前完成人数：{}{}剩余完成人数：{}'.format(num,'\n',22-num))
    page_num=num//10
    if num%10>0:
        page_num+=1

    ls=[]
    for i in range(page_num):
        for name in page.eles('xpath://tr[@class="ant-table-row ant-table-row-level-0 lightd-table-row"]'):
        # 获取姓名
            name=name.ele('xpath://td/a').text
            ls.append(name)
            
        # 下一页
        if i != page_num-1:
            page.ele('xpath://li[@title="下一页"]').click()
            page.wait(35)
        else:
            break


    print('完成人员名单:'+ls)
    return(ls)

#-----------------------------
def match_data():
    # 源数据
    df=pandas.read_excel(excel)

    return(df)

#-----------------------------
def concatenate_names(series):
    return ', '.join(series)
    
if __name__=='__main__':
    # 获取官网已经完成的大学习姓名列表
    data=get_data()
    
    # 读取excel
    excel_df=match_data()

    # 遍历 data['姓名'] 中的每个姓名
    for name in data:
        # 使用 numpy.where() 根据条件选择性地修改值
        excel_df['是否完成大学习'] = numpy.where(excel_df['姓名'] == name, '是', excel_df['是否完成大学习'])

    # pivot_table 基本结构
    # concatenate_namesm 内部调用
    pivot_df = pandas.pivot_table(excel_df, index='是否完成大学习', values=['姓名'] ,aggfunc=concatenate_names)
    
    # 打印输出
    print(pivot_df)
