# NavReporter
A tool automatically scans emails, downloads Excel files,  parse the spreadsheet, collect the data and send them to related persons.
1. 自动扫描邮件，下载指定附件。
2. 解析Excel电子表格，提取感兴趣信息。
3. 自动填写Excel电子表格，生成报告。
4. 自动发送报告到指定电子邮件账户。

最新修改：

1. 下载的代码可以直接运行。使用测试邮箱 nav_report@163.com，请勿修改邮箱设置和密码。
2. 163.com 邮箱默认不开启pop3功能，如果自己编写程序，记得设置开启pop3收信功能。
3. 代码中Calendar.py中第32行被hardcode成 2019.9.19，在实际使用中需要修改。
4. 目前程序做成了一个实例，程序会登录nav_report@163.com，下载两个邮件附件，里面是模拟的数据表，然后读取数据，填写本地的报告表格，再把报告发送回nav_report@163.com。
5. 程序运行两种方式，假设下载代码后，放置于 NavReporter目录：
  1). 运行方法一： python_path/python.exe NavReporter ，这种运行方法会执行 NavReporter/__main__.py中的入口
  2). 运行方法二： python_path/python.exe NavReporter/NavReporter.py，这种方法会执行NavRepoerter/NavReporter.py中的__main__入口。
