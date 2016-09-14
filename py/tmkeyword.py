'''
							  .======.
							  | INRI |
							  |      |
							  |      |
					 .========'      '========.
					 |   _      xxxx      _   |
					 |  /_;-.__ / _\  _.-;_\  |
					 |     `-._`'`_/'`.-'     |
					 '========.`\   /`========'
							  | |  / |
							  |/-.(  |
							  |\_._\ |
							  | \ \`;|
							  |  > |/|
							  | / // |
							  | |//  |
							  | \(\  |
							  |  ``  |
							  |      |
							  |      |
							  |      |
							  |      |
							  .======.
    ……………………………………………………………………………………

                    　　　　　　！！！！！
                    　　　　　　 \\ - - //
                    　　　　　　 (-● ●-)
                    　　　　　　　\ (_) /
                    　　　　　　　 \ u /
                    ┏oOOo-━━━━━━━━┓
                    ┃　　　　　　　　　　 ┃
                    ┃　　　耶稣保佑！　　 ┃
                    ┃ 		   永无BUG！！！┃
                    ┃　　　　　　　　　　 ┃
                    ┗━━━━━━━━-oOOo┛

    ……………………………………………………………………………………

                                  _oo0oo_
                                 088888880
                                 88" . "88
                                 (| -_- |)
                                  0\ = /0
                               ___/'---'\___
                             .' \\\\|     |// '.
                            / \\\\|||  :  |||// \\
                           /_ ||||| -:- |||||- \\
                          |   | \\\\\\  -  /// |   |
                          | \_|  ''\---/''  |_/ |
                          \  .-\__  '-'  __/-.  /
                        ___'. .'  /--.--\  '. .'___
                     ."" '<  '.___\_<|>_/___.' >'  "".
                    | | : '-  \'.;'\ _ /';.'/ - ' : | |
                    \  \ '_.   \_ __\ /__ _/   .-' /  /
                ====='-.____'.___ \_____/___.-'____.-'=====
                                  '=---='


              ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                        佛祖保佑           永无BUG




							  ┌─┐       ┌─┐
						   ┌──┘ ┴───────┘ ┴──┐
						   │                 │
						   │       ───       │
						   │  ─┬┘       └┬─  │
						   │                 │
						   │       ─┴─       │
						   │                 │
						   └───┐         ┌───┘
							   │         │
							   │         │
 							   │         │
							   │         └──────────────┐
							   │                        │
							   │                        ├─┐
							   │                        ┌─┘
							   │                        │
							   └─┐  ┐  ┌───────┬──┐  ┌──┘
								 │ ─┤ ─┤       │ ─┤ ─┤
								 └──┴──┘       └──┴──┘
									 神兽保佑
									 代码无BUG!
'''
# !/usr/bin/python3.4
# -*- coding: utf-8 -*-

# 前排烧香
# 永无BUG

import requests
import re
import time
import random
import xlsxwriter
import datetime


def geturl(url):
    # 制作头部
    header = {
        'User-Agent': 'Mozilla/5.0 (iPad; U; CPU OS 4_3_4 like Mac OS X; ja-jp) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8K2 Safari/6533.18.5',
        'Referer': 'https://top.taobao.com/index.php?topId=TR_FS&leafId=50012010&rank=sale&type=up&s=0',
        'Host': 'top.taobao.com',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
        'Connection': 'keep-alive'
    }
    # get参数
    res = requests.get(url=url, headers=header)
    # ('UTF-8')('unicode_escape')('gbk','ignore')
    resdata = res.content
    return resdata

def timetochina(longtime, formats='{}天{}小时{}分钟{}秒'):
    day = 0
    hour = 0
    minutue = 0
    second = 0
    try:
        if longtime > 60:
            second = longtime % 60
            minutue = longtime // 60
        else:
            second = longtime
        if minutue > 60:
            hour = minutue // 60
            minutue = minutue % 60
        if hour > 24:
            day = hour // 24
            hour = hour % 24
        return formats.format(day, hour, minutue, second)
    except:
        raise Exception('时间非法')

if __name__ == '__main__':

    a = time.clock()
    while 1:
        try:
            type = int(input("查询销售上升榜请按0，销售热门排行请按1，搜索上升榜请按2，,搜索热门排行榜请按3，品牌上升榜请按4，品牌热门排行请按5："))
            if type >= 0 and type <= 5:
                break
            else:
                print("请正确输入！")
        except:
            print("请正确输入！")

    try:
        page = int(input("请输入查询页数（默认5页）："))
        if page > 5:
            page = 5 * 20
        elif page < 0:
            page = 5 * 20
        else:
            page = page * 20
    except:
        page = 5 * 20

    # 转成字符串
    # %y 两位数的年份表示（00 - 99）
    # %Y 四位数的年份表示（000 - 9999）
    # %m 月份（01 - 12）
    # %d 月内中的一天（0 - 31）
    # %H 24小时制小时数（0 - 23）
    # %I 12小时制小时数（01 - 12）
    # %M 分钟数（00 = 59）
    # %S 秒（00 - 59）
    today = time.strftime('%Y%m%d%H%M', time.localtime())
    # 创建一个Excel文件
    workbook = xlsxwriter.Workbook('../excel/' + today + '.xlsx')
    # 创建一个工作表
    worksheet = workbook.add_worksheet("keyword")

    # 总表
    itemlist = []
    # 当查询销售上升销售热门的时候用这个
    if type == 0 or type == 1:
        # 制定查询规则
        rank = "sale"
        if type == 0:
            type = "up"
        else:
            type = "hot"

        # 写入excel的行数
        num = 1
        for i in range(0, page, 20):
            url = "https://top.taobao.com/index.php?topId=TR_FS&leafId=50012010&rank=" + rank + "&type=" + str(
                type) + "&s=" + str(i)
            # 解析网页
            contents = geturl(url)

            # 筛选出字典
            dict = {}
            dict = re.compile(r'("list":\[)(.+?)(\])').search(
                contents.decode("unicode_escape", "ignore").replace("\\", "")).group(2)
            # 将字符串转化为字典
            product = eval(dict)
            print("正在抓取...")

            for item in product:
                # 取字典里面的值呀
                # 排名
                itemlist.append(item['col1']['text'])
                # 包包名
                itemlist.append(item['col2']['text'])
                # 详情页
                itemlist.append("https:" + item['col2']['url'])
                # 参考价
                itemlist.append(item['col3']['text'])
                # 成交指数
                itemlist.append(item['col4']['num'])
                # 升降位次
                itemlist.append(item['col5']['text'])
                # 判断升降，1是升
                itemlist.append(item['col5']['upOrDown'])
            # 每一次下载都暂停1-3秒
            loadtime = random.randint(1, 3)
            print("暂停" + str(loadtime) + "秒")
            time.sleep(loadtime)
        print("正在写入...")

        # 一个队itemlist的计数器
        count = 0
        # 记录第一行
        first = ['排名', '关键词', '详情页', '参考价', '成交指数', '升降位次', '0为降1为升2为空']
        while 1:
            try:
                # 写入第一行
                for m in range(0, len(first)):
                    worksheet.write(0, m, first[m])
                    # 写入数据
                    worksheet.write(num, m, itemlist[count])
                    count = count + 1
            except Exception as err:
                print(err)
                break
            num = num + 1
        print("写入完毕...")

    elif type == 2 or type == 3:
        # 制定查询规则
        rank = "search"
        if type == 2:
            type = "up"
        else:
            type = "hot"

        # 写入excel的行数
        num = 1
        for i in range(0, page, 20):
            url = "https://top.taobao.com/index.php?topId=TR_FS&leafId=50012010&rank=" + rank + "&type=" + str(
                type) + "&s=" + str(i)
            # 解析网页
            contents = geturl(url)
            # print(contents.decode("unicode_escape", "ignore"))
            # 筛选出字典
            dict = {}
            dict = re.compile(r'("list":\[)(.+?})(\])').search(
                contents.decode("unicode_escape", "ignore").replace("\\", "")).group(2)
            # print(dict)
            # 将字符串转化为字典
            product = eval(dict)
            print("正在抓取...")

            for item in product:
                # 取字典里面的值呀
                # 排名
                itemlist.append(item['col1']['text'])
                # 关键词
                itemlist.append(item['col2']['text'])
                # 详情页
                itemlist.append("https:" + item['col2']['url'])
                # 关注指数
                itemlist.append(item['col4']['num'])
                # 升降位次
                itemlist.append(item['col5']['text'])
                # 判断升降，1是升
                itemlist.append(item['col5']['upOrDown'])
                # 升降幅度
                itemlist.append(item['col6']['text'])
            # 每一次下载都暂停1-3秒
            loadtime = random.randint(1, 3)
            print("暂停" + str(loadtime) + "秒")
            time.sleep(loadtime)
        print("正在写入...")

        # 一个队itemlist的计数器
        count = 0
        # 记录第一行
        first = ['排名', '关键词','详情页', '关注指数', '升降位次', '0为降1为升2为空', '升降幅度']
        while 1:
            try:
                # 写入第一行
                for m in range(0, len(first)):
                    worksheet.write(0, m, first[m])
                    # 写入数据
                    worksheet.write(num, m, itemlist[count])
                    count = count + 1
            except Exception as err:
                print(err)
                break
            num = num + 1
        print("写入完毕...")

    # 这里和上面是写重复了的
    # 不管了，反正都可以用就行
    elif type == 4 or type == 5:
        # 制定查询规则
        rank = "brand"
        if type == 4:
            type = "up"
        else:
            type = "hot"
        # 写入excel的行数
        num = 1
        for i in range(0, page, 20):
            url = "https://top.taobao.com/index.php?topId=TR_FS&leafId=50012010&rank=" + rank + "&type=" + str(
                type) + "&s=" + str(i)
            # 解析网页
            contents = geturl(url)
            # print(contents.decode("unicode_escape", "ignore"))
            # 筛选出字典
            dict = {}
            dict = re.compile(r'("list":\[)(.+?})(\])').search(
                contents.decode("unicode_escape", "ignore").replace("\\", "")).group(2)
            # print(dict)
            # 将字符串转化为字典
            product = eval(dict)
            print("正在抓取...")

            for item in product:
                # 取字典里面的值呀
                # 排名
                itemlist.append(item['col1']['text'])
                # 关键词
                itemlist.append(item['col2']['text'])
                # 详情页
                itemlist.append("https:" + item['col2']['url'])
                # 关注指数
                itemlist.append(item['col4']['num'])
                # 升降位次
                itemlist.append(item['col5']['text'])
                # 判断升降，1是升
                itemlist.append(item['col5']['upOrDown'])
                # 升降幅度
                itemlist.append(item['col6']['text'])
            # 每一次下载都暂停1-3秒
            loadtime = random.randint(1, 3)
            print("暂停" + str(loadtime) + "秒")
            time.sleep(loadtime)
        print("正在写入...")

        # 一个队itemlist的计数器
        count = 0
        # 记录第一行
        first = ['排名', '关键词', '详情页','关注指数', '升降位次', '0为降1为升2为空', '升降幅度']
        while 1:
            try:
                # 写入第一行
                for m in range(0, len(first)):
                    worksheet.write(0, m, first[m])
                    # 写入数据
                    worksheet.write(num, m, itemlist[count])
                    count = count + 1
            except Exception as err:
                print(err)
                break
            num = num + 1
        print("写入完毕...")

    workbook.close()
    b = time.clock()
    print('运行时间：' + timetochina(b - a))
