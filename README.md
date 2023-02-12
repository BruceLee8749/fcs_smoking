## 项目简介  
DCS和FCS产品接口参数多达84个，现有的测试工具如Postman和JMeter无法满足庞大的测试需要，故针对公司产品开发DCS&FCS自动化接口测试工具。
测试工具基于Python的第三方库requests和openpyxl开发，具有良好的可维护性、可拓展性和易用性，同样也在健壮性和容错性做出了优化。


## 前置条件
####环境依赖
```shell
本脚本仅限于在Windows下chrome浏览器执行
```
####需要安装的软件
```shell
pycharm 
chrome 
chromedriver 
python3.7，勾选Add Python 3.x to PATH 
驱动下载及放置位置
1.chrome版本查看方式：在浏览器中输入chrome://version/ 查看chrome版本
2.再到http://chromedriver.storage.googleapis.com/index.html中下载相应版本的chromedriver
3.放到python的安装目录下
```
### 依赖的python第三方库
```shell
selenium
pypiwin32
zipp
pillow
numpy
opencv-python
shutil
configparser
openpyxl
math
operator
functools
可直接使用setup.bat运行安装
如果安装失败，先执行下python -m pip install --upgrade pip升级pip在进行三方库安装
```


## 目录结构说明
```shell
1.FCS测试用例:存放测试用例的excel文件
2.FCS测试结果：存放测试结果的excel文件
3.conf.ini:配置文件：地址、数据管理
4.main.py：核心文件，主要用于串跑和发送请求、读取excel表，浏览器页面操作，截图及对比等方法
5.tool.py:内部调用的工具，写入log，读取ini文件，excel表读取、写入方法，系统操作方法
5.fcs_convert.py: convert_run():参数接口转换  convert_rerun():参数接口转换失败后的用例重跑
6.case_run.py:执行所有测试用例，并产生html测试报告
```


## 配置文件操作说明
```shell
conf.ini：配置文件
    - 测试接口：测试目标服务器以及具体接口，若要修改主要修改服务器地址
    - 指定文件：若为空，运行测试用例文件夹下所有xlsx文件，若为多值，请用“空格”连接，如：aaa.xlsx bbb.xlsx
    - 指定工作表：若为空，运行该表格下所有工作表，若为多值，请用“空格”连接，如：sheet1 sheet2
    - 存放文件夹：测试文件存放的文件夹名，为了统一管理，大家把测试文件放到一个文件夹内
    - 概括为空时，运行所有的，多个值请用“空格”连接
    
        | 类型 | 文件空值 | 文件单值 | 文件多值 |
        | :----: | :----: | :----: | :----: |
        | 工作表空值 | 所有 | 该文件下所有 |该些文件下所有|
        | 工作表单值 | 所有 | 该文件夹下该工作表 |该些文件下所有|
        | 工作表多值 | 所有 | 该文件夹下该些工作表 |该些文件下所有|  
```
## 测试用例编写
- 用例名：自定义用例名以区分，举个例子：isDownload_1_doc_1，isDownload参数为1，doc转换成convertType为1
- 本地文件路径：只用写测试文件夹名
- 转换类型：即convertType
- 其他参数：**请严格按照JSON格式书写，英文状态下符号**，如{"isPrint": 1,"isCopy":0, "isZip": 0}，如不需要填写参数，直接写{}
- 验证点：**请严格按照JSON格式书写，英文状态下符号**，如responseBody是嵌套的JSON格式，验证点无需嵌套，单一JSON格式即可
- 测试用例过多的话，可以分多个excel表和工作表编写测试用例，用于区分
## 结果表解读
- 返回body：记录返回的responseBody，用于确认的结果

- 测试结果：
    - 通过：标记绿色
    - 失败：标记红色，和验证点不一致
    - 阻塞：标记黄色，服务端故障导致无法进行测试
    - 格式错误：标记灰色，请检查你的参数和验证点，是否存在JSON格式错误

- 转换后文件地址：生成的URL的超链接，直接点击可查阅线上转换好的文件，DCS接口若是多个文件，会分多个显示
- 跳转页码：跳转到其他页码
- 期望截图：作为截图对比标准，保留正确截图后，不再变动
- 实际截图：每次版本更新，进行实际截图，与期望截图进行对比
- 截图对比结果：实际截图预期望截图进行对比。对比值越大，差异越大。

## 注意事项
```shell
1.脚本启动时,测试用例excel、测试结果excel文件请关闭,否则脚本运行时会报错

2.脚本启动时,Download文件夹请关闭,否则脚本运行时会报错

3.本脚本执行时间大概为xx小时左右,期间禁止操作设备,防止脚本执行出现异常

4.执行完成后FReport文件夹下进行查看测试报告用例通过在90%以上本次部署即为通过

5.本脚本有两种方式查看错误情况:测试结果excel和log日志

```

