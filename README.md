# learn-import-mark

清华大学旧版网络学堂批量导入成绩的工具。

1. 使用网络学堂帐号登录
2. 选择课程
3. 选择作业（在这一步之前，已交作业的学生才能导入；可以在网络学堂上批量代交作业）
4. 读取成绩文件（格式如下）
5. 导入成绩

成绩文件是纯文本文件，格式[示例](examples/example-0.txt)：

```
2013010001 95 每行对应于一个学生的分数
2013010001 90 这里写评语。学号、分数、评语用空格（或 Tab）隔开
2013010002 88.8 分数可以保留一位小数
2013010003 - 如果只写评语，分数为空，则分数用半角减号“-”代替
  2013010003   -   多几个空格或空行不会影响程序识别
2013010004 91 评语可以为空
2013010005 100
```

注意事项：

1. 为了避免频繁登录，cookies 保存在本机当前运行目录下。在公共场所用使用完后请清理 `cookie.txt`。
2. 您需要知道，网络学堂的漏洞很多，很有可能受到 XSS 注入等攻击，导致程序获得的是 XSS 攻击后的数据，从而出现逻辑错误。
3. 使用本软件请遵守相关规章制度，如果不被允许，请立即停止使用本软件，并按要求删除有关数据。

## Usage

Install requirements

```
$ pip install -r requirements.txt
```

Run GUI

```
$ apt install python3-tk    # if you are using python3
$ python src/main.py
```

Distribute as `.exe` (Windows only):

```
$ # install [py2exe](http://www.py2exe.org/)
$ python src/setup.py py2exe
```
