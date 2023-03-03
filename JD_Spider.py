from User_Agents import user_agents
import proxy_operation
import asyncio
import aiohttp
import random
import re
import json
import openpyxl

headers = {
    ":authority": "club.jd.com",
    ":scheme": "https",
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Connection": "keep-alive",
    "Cookie":"shshshfp=39d05cc62e4144f4819fbce8291319de; shshshfpa=03a5fd05-bfb6-1c58-2aaf-d6d2191f0d52-1677826698; shshshfpx=03a5fd05-bfb6-1c58-2aaf-d6d2191f0d52-1677826698; jsavif=1; __jda=122270672.1677826698982337596164.1677826699.1677826699.1677826699.1; __jdc=122270672; __jdv=122270672|direct|-|none|-|1677826698982; shshshfpb=vT-SC67pm5Wtth_7LCj9o7g; areaId=5; __jdu=1677826698982337596164; ipLoc-djd=5-274-49708-49891; jwotest_product=99; token=1f96c62f1b4d1de9f62d5e71ddd3793a,2,932125; __tk=rVJfRtXutiTZIDtwK0RfjcX6uchYuvuEK1uWKvMXBaRLuvJwrVJ0RiXztcTZAbRW4dJfjvMd,2,932125; shshshsID=68457572bfc068e606bc137993b4737f_2_1677826733984; __jdb=122270672.2.1677826698982337596164|1.1677826699; 3AB9D23F7A4B3C9B=WIRQKZXZIYUH6KKKXEVXQ3XWN44WJTW5C27VCDAAUYTTQDXTU6GJIGJJMDEQW3JCALDRT3QLUJWK5LRCJBLJVIOADA; JSESSIONID=FCC1409446241543917E064DD6624508.s1",
    "Host": "club.jd.com",
    "Referer": "https://item.jd.com/",
    "sec-ch-ua": '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109',
    'sec-ch-ua-mobile': '?0',
    'Sec-Fetch-Dest': 'script',
    "sec-ch-ua-platform": '"Windows"',
    'Sec-Fetch-Mode': 'no-cors',
    'Sec-Fetch-Site': 'same-site',
    "User-Agent": random.choice(user_agents)
}

# 代理池
proxies = []


async def get_params(ID, start, end, score):
    page_list = []
    for i in range(start, end):
        params = {
            'callback': 'fetchJSON_comment98',  # 固定值
            'productId': ID,  # 产品 ID
            'score': score,  # 0-3 分别对应全部评价,差评，中评，好评等
            'sortType': '5',  # 排序方式 5:默认排序 6:时间排序
            'page': i,  # 当前的评论页 从0开始
            'pageSize': '10',  # 每页评论的数量
            'isShadowSku': '0',
            'rid': '0',
            'fold': '1'
        }
        page_list.append(params)
    return page_list


async def Replenish():
    plist = proxy_operation.get_proxy()
    for proxy in plist:
        proxies.append(proxy)


async def judge_proxy_num():
    num = len(proxies)
    if num <= 5:
        print("代理数量小于5 开始补充")
        await Replenish()


async def fetch(session, params, queue):
    async with semaphore:
        # base_url = "https://club.jd.com/comment/productPageComments.action?" 没有勾选只看当前商品评价 其他参数一样
        base_url = "https://club.jd.com/comment/skuProductPageComments.action?"
        proxy = random.choice(proxies)
        print(f"当前proxy: {proxy},发起请求的page: {params['page']}")
        try:
            async with session.get(base_url, headers=headers, params=params, proxy=proxy) as resp:
                if resp.status == 200:
                    html = await resp.text()
                    print(f"当前proxy: {proxy},已返回数据的page: {params['page']}")
                    if html:
                        print(f"第{params['page']}页的html进入queue")
                        await queue.put(html)
                    else:
                        print(f'没有返回内容 删除{proxy} 重新选择代理')
                        proxies.remove(proxy)
                        await judge_proxy_num()
                        await fetch(session, params, queue)
                else:
                    print(f'代理池中代理无效 删除{proxy} 重新选择代理')
                    proxies.remove(proxy)
                    await judge_proxy_num()
                    await fetch(session, params, queue)
        except:
            print(f"第{params['page']}页502了!!! 更换代理 重新发起请求啦 ")
            proxies.remove(proxy)
            await judge_proxy_num()
            await fetch(session, params, queue)


async def producer(queue, page_list):
    async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(ssl=False)) as session:
        tasks = [asyncio.create_task(fetch(session, params, queue)) for params in page_list]
        await asyncio.wait(tasks)


async def create_sheet():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.cell(row=1, column=1).value = 'Amount'
    sheet.cell(row=1, column=2).value = 'UserName'
    sheet.cell(row=1, column=3).value = 'ID'
    sheet.cell(row=1, column=4).value = 'ReferenceName'
    sheet.cell(row=1, column=5).value = 'Score'
    sheet.cell(row=1, column=6).value = 'Comment'
    workbook.save('comments.xlsx')


async def consumer(queue):
    wb = openpyxl.load_workbook('comments.xlsx')
    i = 1
    sheet = wb.active
    while True:
        html = await queue.get()
        if html is None:
            break
        else:
            html_json = re.search('(?<=fetchJSON_comment98\().*(?=\);)', html).group(0)
            comments_data = json.loads(html_json)["comments"]
            for comments in comments_data:
                sheet.cell(row=i + 1, column=1).value = i
                sheet.cell(row=i + 1, column=2).value = comments['nickname']
                sheet.cell(row=i + 1, column=3).value = comments['id']
                sheet.cell(row=i + 1, column=4).value = comments['referenceName']
                sheet.cell(row=i + 1, column=5).value = comments['score']
                sheet.cell(row=i + 1, column=6).value = comments['content']
                i += 1
    wb.save('comments.xlsx')

# 定义主函数
async def main():
    ID = 848820  # 商品ID
    start = 0  # 起始页
    end = 100  # 尾页
    score = 1  # 评价等级 321好中坏
    page_list = await get_params(ID, start, end, score)  # 获取params, 每个是一页
    queue = asyncio.Queue()  # 创建队列
    await Replenish()  # 获取代理列表
    # await create_sheet()  # 创建excel表格
    print(f"params以及代理ip获取完毕 进入生产者消费者模式！\n代理列表（{len(proxies)}）个", proxies)
    consumer_task = asyncio.create_task(consumer(queue))
    await producer(queue, page_list)
    await queue.put(None)  # 最后留None让Consumer停止
    await consumer_task
    print("生产者消费者都已完成！")


if __name__ == '__main__':
    semaphore = asyncio.Semaphore(10)
    asyncio.run(main())
    # 注
    # 1.这么多几乎并发的要求从一个Cooike发出可能会被反爬即便使用了代理 当频繁更新代理仍然返回空 说明cookie黑了
    # 2. 代理白名单设置 3.excel创建只用创建一次 4.多次换参数执行 i,ID,start,end,score
    # score=0 num = [0,99] 整100页, 10条/页
    # 以下均为500ml*15瓶
    # 青柑普洱茶  ID=100019151854 score=1 -> max(end) = 8   score=2 -> max(end) = 8 score=3 -> max(end) = 100
    # 茉莉花茶   ID=848824 score=1 -> max(end) = 21   score=2 -> max(end) = 23 score=3 -> max(end) = 99
    # 乌龙茶     ID=848811 score=1 -> max(end) = 9   score=2 -> max(end) = 6 score=3 -> max(end) = 100
    # 绿茶      ID=848815 score=1 -> max(end) = 9   score=2 -> max(end) = 17 score=3 -> max(end) = 100
    # 红茶      ID=848820 score=1 -> max(end) = 6   score=2 -> max(end) = 8 score=3 -> max(end) = 99
    # 中差评 106页 大约1000条
    # 好评 498页 共4980条
