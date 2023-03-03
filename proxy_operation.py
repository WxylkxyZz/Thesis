import requests
from User_Agents import user_agents
import random


def generate_proxy_api_link(num):
    start_url = "http://v2.api.juliangip.com/dynamic/getips"
    params = {
        "trade_no": 1648382495550539,  # 业务编号
        "num": num,  # 数量
        "pt": 1,  # 代理类型
        "result_type": "json",  # 返回类型
        "split": 1,  # 结果分隔符
        "city_name": 1,  # 地区名称
        "ip_remain": 1,  # 剩余可用时长(s)
        "filter": 1,  # IP去重
    }
    new_params = {key: params[key] for key in sorted(params.keys())}
    new_params["key"] = "67f84c93ef7642aa8c59b28065fb9846"
    res = requests.get("http://httpbin.org/get", new_params)
    original = res.url.split("?")[1]
    sign = md5(original.encode())
    link = start_url + "?" + original + "&" + f"sign={sign}"
    return link


def md5(string):
    import hashlib
    m = hashlib.md5()
    # update(arg)传入arg对象来更新hash的对象。必须注意的是，该方法只接受byte类型，否则会报错。这就是要在参数前添加b来转换类型的原因。
    m.update(string)
    return m.hexdigest()


def get_proxy():
    num = 3
    link = generate_proxy_api_link(num)
    headers = {'User_Agents': random.choice(user_agents)}
    res = requests.get(url=link, headers=headers).json()
    proxy_list = res["data"]["proxy_list"]
    proxys = []
    for proxy in proxy_list:
        proxys.append("http://" + proxy.split(",", 1)[0])
    return proxys


if __name__ == "__main__":
    print("hello world!")
