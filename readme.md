# 第零步：下载chromedriver.exe
打开你的chrome浏览器，点击右上角的三个点->help->about Google Chrome，查看相应的版本
例如Version 69.0.3497.81 (Official Build) (64-bit)
去chromedriver.storage.googleapis.com/index.html中查询该版本对应的chromedriver是什么版本（当然也可以百度直接搜匹配的版本）
国内用户在npm.taobao.org/mirrors/chromedriver中下载对应版本的chromedriver的win32的压缩包
解压后把chromedriver.exe放在代码目录下
# 第一步：检查安装好了python3.7的环境 并且安装好了如下第三方库
pip install selenium -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
修改一下excel_handler.py中的wb_name和new_wb_name
# 第二步：在cmd中输入python.exe excel_handler.py A B
A 表示excel中的起始行， B表示excel中的末尾行
python.exe excel_handler.py A B 表示从A行执行到B行（包括B行）
python.exe excel_handler.py 表示从头执行到尾
python.exe excel_handler.py A 表示从A行执行到尾
# 第三步：网页自动填写好信息
# 第四步：手工输入验证码
# 第五步：点击“查询”按钮
# 第六步：切换回cmd，输入ok
# 第七步：切换回网页，跳转到第三步

如果中途想退出输入exit即可
