import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

def generate_sales_data(num_days=100, records_per_day=(3, 8), output_dir="data"):
    """
    生成随机销售数据并保存为Excel文件
    
    参数:
    num_days: 要生成的天数
    records_per_day: 每天生成的记录数范围，元组格式(最小值, 最大值)
    output_dir: 输出目录
    """
    # 设置随机种子确保可重复性
    np.random.seed(42)
    
    # 生成日期序列：从当前日期往前推num_days天
    end_date = datetime(2025, 4, 20)  # 当前日期
    start_date = end_date - timedelta(days=num_days-1)
    
    # 产品相关数据
    products = [
        # 荤菜
        '毛肚', '羊肉', '牛肉', '虾滑', '鱼片', '肥牛', '猪肉片', 
        # 素菜
        '金针菇', '香菇', '娃娃菜', '油麦菜', '土豆片', '藕片',
        # 丸滑类
        '牛肉丸', '虾丸', '鱼丸', '墨鱼丸',
        # 特色菜品
        '毛血旺', '腰花', '黄喉', '鸭肠',
        # 主食
        '粉条', '面条', '年糕', '米饭'
    ]
    
    categories = ['荤菜', '素菜', '丸滑类', '特色菜品', '主食']

    regions = ['华北', '华南', '华东', '华西']

     # 产品类别映射
    product_category_map = {
        '毛肚': '荤菜', '羊肉': '荤菜', '牛肉': '荤菜', '虾滑': '荤菜',
        '鱼片': '荤菜', '肥牛': '荤菜', '猪肉片': '荤菜',
        
        '金针菇': '素菜', '香菇': '素菜', '娃娃菜': '素菜',
        '油麦菜': '素菜', '土豆片': '素菜', '藕片': '素菜',
        
        '牛肉丸': '丸滑类', '虾丸': '丸滑类', '鱼丸': '丸滑类', '墨鱼丸': '丸滑类',
        
        '毛血旺': '特色菜品', '腰花': '特色菜品', '黄喉': '特色菜品', '鸭肠': '特色菜品',
        
        '粉条': '主食', '面条': '主食', '年糕': '主食', '米饭': '主食'
    }

    # 新增城市数据
    cities = {
        '华北': ['北京', '天津', '石家庄', '太原'],
        '华南': ['广州', '深圳', '珠海', '厦门'],
        '华东': ['上海', '南京', '杭州', '苏州'],
        '华西': ['成都', '重庆', '西安', '昆明']
    }
    
    # 新增销售渠道
    channels = ['直营店',"加盟店"]
    
    # 创建空列表存储所有记录
    all_records = []
    
    # 为每一天生成多条记录
    for day in range(num_days):
        current_date = start_date + timedelta(days=day)
        # 随机决定这一天生成多少条记录
        daily_records = np.random.randint(records_per_day[0], records_per_day[1] + 1)
        
        for _ in range(daily_records):
            # 随机选择产品和其他属性
            product = np.random.choice(products)
            category = product_category_map[product]  # 根据产品确定品类
            region = np.random.choice(regions)
            city = np.random.choice(cities[region])
            channel = np.random.choice(channels)
            quantity = np.random.randint(1, 5)
            unit_price = np.random.uniform(15, 88)  # 火锅菜品价格一般在15-88元之间
            unit_price = round(unit_price, 2)
            unit_cost = round(unit_price * np.random.uniform(0.3, 0.6), 2)  # 火锅店的成本比例通常在30-60%
            total_sales = round(quantity * unit_price, 2)
            total_cost = round(quantity * unit_cost, 2)
            profit = round(total_sales - total_cost, 2)
            
            # 添加记录
            all_records.append({
                'date': current_date,
                'product': product,
                'category': category,
                'region': region,
                'city': city,
                'channel': channel,
                'quantity': quantity,
                'unit_price': unit_price,
                'unit_cost': unit_cost,
                'total_sales': total_sales,
                'total_cost': total_cost,
                'profit': profit
            })
    
    # 将所有记录转换为DataFrame
    df = pd.DataFrame(all_records)
    
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 生成输出文件路径
    output_file = os.path.join(output_dir, f'sales_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    
    # 保存数据
    df.to_excel(output_file, index=False)
    
    return df, output_file

def print_data_summary(df):
    """打印数据摘要信息"""
    print("\n=== 数据预览 ===")
    print(df.head())
    
    print("\n=== 基本统计信息 ===")
    print(df.describe())
    
    print("\n=== 每日销售记录数量统计 ===")
    print(df.groupby('date').size().describe())

if __name__ == "__main__":
    # 生成销售数据，设置100天的数据，每天3-8条记录
    df, output_file = generate_sales_data(num_days=100, records_per_day=(10,16))
    
    # 打印数据摘要
    print_data_summary(df)
    
    print(f"\n数据已保存到: {output_file}")