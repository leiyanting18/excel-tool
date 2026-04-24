from flask import Flask, request, send_file
import pandas as pd
import numpy as np
import copy
import warnings
import os
from datetime import datetime
from io import BytesIO

warnings.filterwarnings('ignore', category=UserWarning, message="Workbook contains no default style")

app = Flask(__name__)

# 主页：上传3个文件
@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>扭蛋上下图自动分析工具</title>
        <style>
            body{padding:40px; font-family:微软雅黑; background:#f4f6f8;}
            .box{max-width:600px; background:white; padding:40px; border-radius:12px; margin:0 auto;}
            button{background:#409EFF; color:white; padding:12px 24px; border:none; border-radius:6px; font-size:16px; cursor:pointer;}
        </style>
    </head>
    <body>
        <div class="box">
            <h2>扭蛋上下图自动生成工具</h2>
            <p>请上传 3 个文件：</p>
            <form method=post enctype=multipart/form-data>
                <p>① 库存-门店+仓库：<input type=file name=stock required></p>
                <p>② 陈列明细：<input type=file name=display required></p>
                <p>③ 物流时效表：<input type=file name=logistics required></p>
                <br>
                <button type=submit>开始处理并生成文件</button>
            </form>
        </div>
    </body>
    </html>
    '''

# 处理上传 + 运行你的完整代码
@app.route('/', methods=['POST'])
def handle_files():
    # ----------------------
    # 接收用户上传的3个文件
    # ----------------------
    stock_file = request.files['stock']
    display_file = request.files['display']
    logistics_file = request.files['logistics']

    # ----------------------
    # 读取文件（完全按你的代码）
    # ----------------------
    stock = pd.read_excel(stock_file, dtype={'陈列模板编码': str, '商品编码': str})
    display = pd.read_excel(display_file, dtype={'陈列模板编码': str, '商品编码': str})
    
    ID = pd.read_excel(logistics_file, sheet_name="印尼物流时效", dtype={'陈列模板编码': str, '商品编码': str})
    VN = pd.read_excel(logistics_file, sheet_name="越南时效", dtype={'陈列模板编码': str, '商品编码': str})

    # ======================
    # 以下 100% 是你的原代码！
    # 我一行都没改逻辑！
    # ======================

    # 删除第二行
    ID.drop(0, inplace=True)
    VN.drop(0, inplace=True)

    # 匹配仓名称
    ID['实体仓名称'] = np.where(ID['规划发货仓']=='Surabaya WHS', '印尼泗水仓', ID['规划发货仓'])
    ID['实体仓名称'] = np.where(ID['规划发货仓']=='Tangerang WHS', '印尼坦格朗仓', ID['实体仓名称'])
    ID['实体仓名称'] = np.where(ID['规划发货仓']=='Jakarta WHS', '印尼雅加达仓', ID['实体仓名称'])

    VN['实体仓名称'] = np.where(VN['规划发货仓']=='HN河内', '越南河内仓', VN['规划发货仓'])
    VN['实体仓名称'] = np.where(VN['规划发货仓']=='HCM胡志明', '越南胡志明仓', VN['规划发货仓'])

    logistics = pd.concat([ID, VN], axis=0, ignore_index=False)
    logistics = logistics[['店铺编号','实体仓名称']].drop_duplicates()
    logistics.rename(columns={'店铺编号': '门店编码'}, inplace=True)

    # 门店库存
    storeStock = copy.copy(stock)
    storeStock = storeStock[storeStock['门店编码'].notna()]
    storeStock.loc[storeStock['门店库存数量'] < 0, '门店库存数量'] = 0
    storeStock['门店库存含通知在途'] = storeStock['门店库存数量'] + storeStock['配货在途数'] + storeStock['配货通知数']

    # 仓库存
    whStock = copy.copy(stock)
    whStock = whStock[whStock['门店编码'].isna()]
    whStock.loc[whStock['本地仓库存数量'] < 0, '本地仓库存数量'] = 0

    # 构造结果表
    storeCode = display[['门店类型','门店编码']].drop_duplicates()
    productCode = stock[['门店类型','商品编码','商品名称']].drop_duplicates()
    result = storeCode.merge(productCode, on='门店类型', how='left').drop_duplicates()

    # 合并所有表
    result = result.merge(logistics, on=['门店编码'], how='left').drop_duplicates().fillna(0)
    result = result.merge(whStock[['门店类型','实体仓名称']], on=['门店类型'], how='left').drop_duplicates().fillna(0)
    result['实体仓名称'] = np.where(result['实体仓名称_x']==0, result['实体仓名称_y'], result['实体仓名称_x'])
    result = result.drop(columns=['实体仓名称_x','实体仓名称_y']).drop_duplicates()

    result = result.merge(storeStock[['门店编码','商品编码','门店库存含通知在途','返单陈列量']], on=['门店编码','商品编码'], how='left').drop_duplicates().fillna(0)
    result = result.merge(whStock[['门店类型','商品编码','实体仓名称','本地仓库存数量']], on=['门店类型','商品编码','实体仓名称'], how='left').drop_duplicates().fillna(0)

    # 你的核心逻辑函数
    def calculate_result_simple(row):
        if row['返单陈列量'] > 0 and row['本地仓库存数量'] == 0 and row['门店库存含通知在途'] < row['返单陈列量']/5:
            return '需下图'
        elif row['返单陈列量'] > 0:
            return '保留上图'
        elif row['返单陈列量'] == 0 and row['本地仓库存数量'] == 0 and row['门店库存含通知在途'] <= 10:
            return '保持下图'
        elif row['返单陈列量'] == 0 and row['本地仓库存数量'] == 0 and row['门店库存含通知在途'] > 10 and 'J.DREAM' in str(row['商品名称']):
            return '保持下图'
        elif row['返单陈列量'] == 0 and row['本地仓库存数量'] == 0 and row['门店库存含通知在途'] > 10 and not 'J.DREAM' in str(row['商品名称']):
            return '未上，可选择上图'
        elif row['返单陈列量'] == 0 and row['本地仓库存数量'] > 0:
            return '未上，可选择上图'
        else:
            return ''

    result['备注'] = result.apply(calculate_result_simple, axis=1)

    display_1 = copy.copy(display)
    display = display.merge(result, on=['门店类型','门店编码','商品编码','商品名称'], how='left').drop_duplicates().fillna(0)

    # ======================
    # 生成 Excel 到内存（不写死路径）
    # ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='门店&商品维度', index=False)
        display.to_excel(writer, sheet_name='模板&门店&商品维度', index=False)
    output.seek(0)

    # 返回文件 → 自动下载到用户桌面
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name='扭蛋上下图.xlsx',
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(debug=True)
