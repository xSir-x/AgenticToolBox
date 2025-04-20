import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

def generate_sales_data(num_days=100, records_per_day=(3, 8), output_dir="data"):
    """
    生成随机销售数据并保存为CSV文件
    
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
    products = ['笔记本电脑', '手机', '平板电脑', '耳机', '智能手表']
    categories = ['电子产品', '配件']
    regions = ['华北', '华南', '华东', '华西']
    
    # 新增城市数据
    cities = {
        '华北': ['北京', '天津', '石家庄', '太原'],
        '华南': ['广州', '深圳', '珠海', '厦门'],
        '华东': ['上海', '南京', '杭州', '苏州'],
        '华西': ['成都', '重庆', '西安', '昆明']
    }
    
    # 新增销售渠道
    channels = ['线下门店', '电商平台', '代理商', '直营网站', '团购渠道']
    
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
            category = np.random.choice(categories)
            region = np.random.choice(regions)
            city = np.random.choice(cities[region])
            channel = np.random.choice(channels)
            quantity = np.random.randint(1, 50)
            unit_price = np.random.uniform(100, 10000)
            unit_price = round(unit_price, 2)
            unit_cost = round(unit_price * np.random.uniform(0.4, 0.8), 2)
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
    output_file = os.path.join(output_dir, f'sales_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
    
    # 保存数据
    df.to_csv(output_file, index=False)
    
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
    df, output_file = generate_sales_data(num_days=100, records_per_day=(3, 8))
    
    # 打印数据摘要
    print_data_summary(df)
    
    print(f"\n数据已保存到: {output_file}")