from __future__ import annotations

import random
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

random.seed(20260227)

out_file = Path('/Users/yuanyi.li/Documents/Scrub/excel-processor-6/demo/scrub-demo-orders-v2.xlsx')

headers = [
    '门店', '渠道', '下单日期', '订单号', '客户姓名', '手机号', '省市', '商品名称', '规格', '数量', '单价', '订单金额',
    '支付状态', '发货状态', '物流公司', '运单号', '售后状态', '备注', '内部标签', '录入人', '更新时间'
]

stores = ['上海静安店', '杭州西湖店', '北京朝阳店', '深圳南山店', '成都高新店']
channels = ['天猫旗舰店', '京东自营', '抖音小店', '小程序商城', '企业团购']
regions = ['上海-浦东', '浙江-杭州', '江苏-苏州', '广东-深圳', '北京-海淀', '四川-成都']
products = [
    ('燕麦拿铁速溶粉', '30条/盒', 79),
    ('冷萃咖啡液', '18包/箱', 129),
    ('坚果能量棒', '24支/盒', 99),
    ('椰乳蛋白饮', '12瓶/箱', 118),
    ('冻干草莓脆', '20袋/箱', 86),
    ('低糖黑巧麦片', '500g/袋', 59),
]
logis = ['顺丰', '中通', '圆通', '京东物流', '韵达']
clerks = ['王敏', '李晨', '赵倩', '陈涛', '孙悦']
customers = ['张伟', '王芳', '李娜', '刘洋', '陈杰', '赵琳', '黄涛', '周静', '吴昊', '杨雪']

base = datetime(2026, 2, 1, 9, 0, 0)
rows = []

for i in range(1, 131):
    store = random.choice(stores)
    channel = random.choice(channels)
    dt = base + timedelta(hours=random.randint(0, 24 * 22), minutes=random.randint(0, 59))
    order_no = f'ORD202602{200000 + i}'
    customer = random.choice(customers)
    phone = f"1{random.randint(3000000000, 9999999999)}"
    region = random.choice(regions)

    pname, spec, price = random.choice(products)
    qty = random.choice([1, 1, 2, 2, 3, 4])
    amount = qty * price

    pay = random.choice(['已支付', '已支付', '已支付', '待支付'])
    ship = random.choice(['已发货', '已发货', '待发货'])
    company = random.choice(logis)
    waybill = f'WB{random.randint(10000000, 99999999)}' if ship == '已发货' else ''
    after_sale = random.choice(['无', '无', '无', '退款处理中'])

    note = random.choice([
        '首次下单',
        '老客复购',
        '请尽快发货',
        '地址需电话确认',
        '赠送试用装',
        '客户反馈包装破损',
    ])
    tag = random.choice(['正常', '活动单', '高优先级', '内部'])

    if i % 16 == 0:
        note = note.replace('发货', '髮货')
    if i % 21 == 0:
        pname = pname + ' '
    if i % 19 == 0:
        note = '测试单-流程回归验证'
        tag = '测试'
    if i % 23 == 0:
        note = '退款申请-客户重复下单'
        after_sale = '退款处理中'

    updated = (dt + timedelta(hours=random.randint(1, 30))).strftime('%Y-%m-%d %H:%M:%S')

    rows.append([
        store, channel, dt.strftime('%Y-%m-%d'), order_no, customer, phone, region,
        pname, spec, qty, price, amount, pay, ship, company, waybill, after_sale,
        note, tag, random.choice(clerks), updated
    ])

for idx in [10, 27, 49, 72, 95, 114]:
    src = rows[idx][:]
    dup = src[:]
    dup[0] = random.choice(stores)
    dup[1] = random.choice(channels)
    dup[2] = (datetime.strptime(src[2], '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')
    rows.insert(idx + 1, dup)

wb = Workbook()
ws = wb.active
ws.title = '订单数据'
ws.append(headers)
for r in rows:
    ws.append(r)

wb.save(out_file)
print(out_file)
print(f'rows={len(rows)}, cols={len(headers)}')
