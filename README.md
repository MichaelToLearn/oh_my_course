# oh_my_course
 【北理工】填一个excel表格就能抢课的python脚本。
## 使用
环境：`python 3`
依赖：
- xlrd
- requrests
1. `clone`或者下载仓库
```bash
git clone https://github.com/MichaelToLearn/oh_my_course.git
```
2. 安装依赖
```bash
pip install xlrd
pip install requests
```
3. 填写自己的账户和选课信息
打开`抢课模板.xlsx`，填入下面的信息：
![在这里插入图片描述](https://img-blog.csdnimg.cn/20190115204244685.png)
4. 选课前，双击“运行一次.bat”，在选课之前提前建立好课程数据库（也可以选课时建立，提前的话省时间）
5. 双击“运行一次.bat”，打开一个窗口。双击“运行多次.bat”，多个窗口同时抢课。如下图所示:
![在这里插入图片描述](https://img-blog.csdnimg.cn/20190115204610628.gif)

## 原理
首先实现统一认证系统`Login.py`，然后为这个糟糕的选课系统写第三方`API`，省略了页面文本、图片、`css`、`js`的加载，只关注数据的传输，因此能够在浏览器进不去的情况下，脚本可以进去。
## 免责
1. 代码仓促，`bug`正常。
2. 爬虫有害健康。
