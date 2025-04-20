import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from typing import Dict, Any
import os
import matplotlib as mpl

def set_chinese_font():
    """设置中文字体"""
    plt.rcParams['font.sans-serif'] = ['SimSun', 'DejaVu Sans', 'Arial Unicode MS']  # 按优先级尝试不同字体
    plt.rcParams['axes.unicode_minus'] = False    # 用来正常显示负号
    return None  # 不返回font对象，直接使用全局设置

# 初始化中文字体
set_chinese_font()

def analyze_sales_data(df: pd.DataFrame, output_dir: str = "reports") -> Dict[str, Any]:
    """
    对销售数据进行多维度分析
    
    参数:
    df: 包含销售数据的DataFrame
    output_dir: 可视化图表的输出目录
    
    返回:
    包含各维度分析结果的字典
    """
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 确保date列是datetime类型
    df['date'] = pd.to_datetime(df['date'])
    df['month'] = df['date'].dt.strftime('%Y-%m')
    
    # 计算关键指标
    df['profit_margin'] = (df['profit'] / df['total_sales'] * 100).round(2)
    df['cost_ratio'] = (df['total_cost'] / df['total_sales'] * 100).round(2)
    
    # 1. 按月度维度分析
    monthly_analysis = df.groupby('month').agg({
        'total_sales': 'sum',
        'total_cost': 'sum',
        'profit': 'sum',
        'profit_margin': 'mean',
        'quantity': 'sum'
    }).round(2)
    
    # 2. 按城市维度分析
    city_analysis = df.groupby('city').agg({
        'total_sales': 'sum',
        'total_cost': 'sum',
        'profit': 'sum',
        'profit_margin': 'mean',
        'quantity': 'sum'
    }).round(2).sort_values('total_sales', ascending=False)
    
    # 3. 按地区维度分析
    region_analysis = df.groupby('region').agg({
        'total_sales': 'sum',
        'total_cost': 'sum',
        'profit': 'sum',
        'profit_margin': 'mean',
        'quantity': 'sum'
    }).round(2).sort_values('total_sales', ascending=False)
    
    # 4. 按渠道维度分析
    channel_analysis = df.groupby('channel').agg({
        'total_sales': 'sum',
        'total_cost': 'sum',
        'profit': 'sum',
        'profit_margin': 'mean',
        'quantity': 'sum'
    }).round(2).sort_values('total_sales', ascending=False)
    
    # 5. 产品分析
    product_analysis = df.groupby('product').agg({
        'total_sales': 'sum',
        'total_cost': 'sum',
        'profit': 'sum',
        'profit_margin': 'mean',
        'quantity': 'sum'
    }).round(2).sort_values('total_sales', ascending=False)
    
    # 创建可视化图表
    create_visualizations(
        monthly_analysis=monthly_analysis,
        region_analysis=region_analysis,
        channel_analysis=channel_analysis,
        city_analysis=city_analysis,
        output_dir=output_dir
    )
    
    # 创建高级可视化图表
    create_advanced_visualizations(df, output_dir)
    
    return {
        'monthly_analysis': monthly_analysis,
        'city_analysis': city_analysis,
        'region_analysis': region_analysis,
        'channel_analysis': channel_analysis,
        'product_analysis': product_analysis
    }

def create_visualizations(
    monthly_analysis: pd.DataFrame,
    region_analysis: pd.DataFrame,
    channel_analysis: pd.DataFrame,
    city_analysis: pd.DataFrame,
    output_dir: str
):
    """创建销售数据可视化图表"""
    # 使用默认样式
    plt.style.use('default')
    
    # 确保每个新图表都使用中文字体
    set_chinese_font()
    
    # 1. 月度销售趋势图
    plt.figure(figsize=(15, 6))
    plt.plot(monthly_analysis.index.astype(str), monthly_analysis['total_sales'], marker='o')
    plt.title('月度销售趋势')
    plt.xlabel('月份')
    plt.ylabel('销售额')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'monthly_sales_trend.png'))
    plt.close()
    
    # 2. 地区销售分布图
    plt.figure(figsize=(10, 6))
    sns.barplot(data=region_analysis.reset_index(), x='region', y='total_sales')
    plt.title('各地区销售额分布')
    plt.xlabel('地区')
    plt.ylabel('销售额')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'region_sales_distribution.png'))
    plt.close()
    
    # 3. 渠道销售占比饼图
    plt.figure(figsize=(10, 8))
    plt.pie(channel_analysis['total_sales'], labels=channel_analysis.index, autopct='%1.1f%%')
    plt.title('各渠道销售额占比')
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'channel_sales_pie.png'))
    plt.close()
    
    # 4. 城市TOP10销售额对比图
    plt.figure(figsize=(12, 6))
    top10_cities = city_analysis.head(10)
    sns.barplot(data=top10_cities.reset_index(), x='city', y='total_sales')
    plt.title('销售额TOP10城市')
    plt.xlabel('城市')
    plt.ylabel('销售额')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'top10_cities_sales.png'))
    plt.close()

def create_advanced_visualizations(df: pd.DataFrame, output_dir: str):
    """创建高级可视化图表"""
    
    # 确保设置中文字体
    set_chinese_font()
    
    # 1. 产品利润矩阵（气泡图）
    plt.figure(figsize=(12, 8))
    product_analysis = df.groupby('product').agg({
        'total_sales': 'sum',
        'profit': 'sum',
        'quantity': 'sum'
    }).reset_index()
    
    # 计算销售额占比和利润率
    total_sales = product_analysis['total_sales'].sum()
    product_analysis['sales_ratio'] = (product_analysis['total_sales'] / total_sales * 100)
    product_analysis['profit_ratio'] = (product_analysis['profit'] / product_analysis['total_sales'] * 100)
    
    plt.scatter(product_analysis['sales_ratio'], 
                product_analysis['profit_ratio'],
                s=product_analysis['quantity']/30,
                alpha=0.6)
    
    # 添加产品标签，优化标签位置
    for i, row in product_analysis.iterrows():
        plt.annotate(row['product'], 
                    (row['sales_ratio'], row['profit_ratio']),
                    xytext=(5, 5), 
                    textcoords='offset points',
                    fontsize=10,
                    bbox=dict(facecolor='white', edgecolor='none', alpha=0.7))
    
    plt.title('产品利润矩阵分析', fontsize=12, pad=15)
    plt.xlabel('销售额占比(%)', fontsize=10)
    plt.ylabel('利润率(%)', fontsize=10)
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'product_profit_matrix.png'), dpi=300, bbox_inches='tight')
    plt.close()
    
    # 2. 区域渠道热力图
    plt.figure(figsize=(12, 8))
    pivot_table = pd.pivot_table(df, 
                                values='profit',
                                index='region',
                                columns='channel',
                                aggfunc='sum')
    
    sns.heatmap(pivot_table, annot=True, fmt='.0f', cmap='YlOrRd', 
                annot_kws={'size': 10})
    plt.title('区域渠道利润热力图', fontsize=12, pad=15)
    plt.xlabel('销售渠道', fontsize=10)
    plt.ylabel('区域', fontsize=10)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'region_channel_heatmap.png'), dpi=300, bbox_inches='tight')
    plt.close()
    
    # 3. 城市销售额Top10（横向柱状图）
    plt.figure(figsize=(12, 6))
    city_sales = df.groupby(['city', 'region']).agg({
        'total_sales': 'sum'
    }).reset_index().sort_values('total_sales', ascending=True)
    
    top10_cities = city_sales.tail(10)
    
    # 使用不同颜色区分不同区域
    colors = plt.cm.Set3(np.linspace(0, 1, len(df['region'].unique())))
    color_dict = dict(zip(df['region'].unique(), colors))
    bar_colors = [color_dict[region] for region in top10_cities['region']]
    
    bars = plt.barh(top10_cities['city'], top10_cities['total_sales'], 
                   color=bar_colors)
    
    # 添加数值标签
    for bar in bars:
        width = bar.get_width()
        plt.text(width, bar.get_y() + bar.get_height()/2, 
                f'¥{width:,.0f}', 
                ha='left', va='center',
                fontsize=9)
    
    plt.title('城市销售额Top10', fontsize=12, pad=15)
    plt.xlabel('销售额（元）', fontsize=10)
    plt.ylabel('城市', fontsize=10)
    
    # 添加图例
    legend_elements = [plt.Rectangle((0,0),1,1, facecolor=color) 
                      for color in color_dict.values()]
    plt.legend(legend_elements, color_dict.keys(), 
              title='区域', loc='lower right')
    
    plt.grid(True, linestyle='--', alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'city_sales_top10_horizontal.png'), 
                dpi=300, bbox_inches='tight')
    plt.close()
    
    # 4. 渠道利润率雷达图
    plt.figure(figsize=(10, 10))
    channel_metrics = df.groupby('channel').agg({
        'profit': 'sum',
        'total_sales': 'sum'
    })
    channel_metrics['profit_ratio'] = channel_metrics['profit'] / channel_metrics['total_sales'] * 100
    
    angles = np.linspace(0, 2*np.pi, len(channel_metrics.index), endpoint=False)
    values = channel_metrics['profit_ratio'].values
    values = np.concatenate((values, [values[0]]))
    angles = np.concatenate((angles, [angles[0]]))
    
    ax = plt.subplot(111, polar=True)
    ax.plot(angles, values, 'o-', linewidth=2)
    ax.fill(angles, values, alpha=0.25)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(channel_metrics.index, fontsize=10)
    
    # 添加网格和标签
    ax.set_ylim(0, max(values) * 1.2)
    plt.title('各渠道利润率分析', fontsize=12, pad=15)
    
    # 添加利润率标签
    for angle, value in zip(angles[:-1], values[:-1]):
        plt.text(angle, value + 2, f'{value:.1f}%', 
                ha='center', va='center')
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'channel_profit_radar.png'), 
                dpi=300, bbox_inches='tight')
    plt.close()
    
    # 5. 时间趋势分析（双轴图）
    plt.figure(figsize=(15, 6))
    daily_metrics = df.groupby('date').agg({
        'total_sales': 'sum',
        'profit': 'sum',
        'total_cost': 'sum'
    }).reset_index()
    daily_metrics['profit_ratio'] = daily_metrics['profit'] / daily_metrics['total_sales'] * 100
    
    fig, ax1 = plt.subplots(figsize=(15, 6))
    ax2 = ax1.twinx()
    
    bars = ax1.bar(daily_metrics['date'], daily_metrics['total_sales'], 
                  alpha=0.3, color='blue', label='销售额')
    line = ax2.plot(daily_metrics['date'], daily_metrics['profit_ratio'], 
                   color='red', label='利润率', linewidth=2)
    
    ax1.set_xlabel('日期', fontsize=10)
    ax1.set_ylabel('销售额（元）', color='blue', fontsize=10)
    ax2.set_ylabel('利润率(%)', color='red', fontsize=10)
    
    # 合并图例
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
    
    plt.title('销售额和利润率趋势', fontsize=12, pad=15)
    plt.grid(True, linestyle='--', alpha=0.3)
    fig.autofmt_xdate()  # 自动调整x轴日期标签
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'sales_profit_trend.png'), 
                dpi=300, bbox_inches='tight')
    plt.close()
    
    # 6. 产品-成本散点分布
    plt.figure(figsize=(12, 8))
    colors = plt.cm.Set3(np.linspace(0, 1, len(df['product'].unique())))
    color_dict = dict(zip(df['product'].unique(), colors))
    
    for product in df['product'].unique():
        product_data = df[df['product'] == product]
        plt.scatter(product_data['total_cost'], product_data['profit'], 
                   alpha=0.6, label=product, 
                   color=color_dict[product])
        
        # 添加产品均值点和标签
        mean_cost = product_data['total_cost'].mean()
        mean_profit = product_data['profit'].mean()
        plt.scatter(mean_cost, mean_profit, 
                   color=color_dict[product], 
                   s=100, marker='*')
        plt.annotate(f'{product}\n平均值', 
                    (mean_cost, mean_profit),
                    xytext=(10, 10),
                    textcoords='offset points',
                    bbox=dict(facecolor='white', 
                             edgecolor='none', 
                             alpha=0.7),
                    fontsize=9)
    
    plt.xlabel('成本（元）', fontsize=10)
    plt.ylabel('利润（元）', fontsize=10)
    plt.title('产品成本-利润分布', fontsize=12, pad=15)
    plt.grid(True, linestyle='--', alpha=0.3)
    
    # 调整图例位置和字体
    legend = plt.legend(title='产品类型', 
                       bbox_to_anchor=(1.05, 1), 
                       loc='upper left')
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'cost_profit_scatter.png'), 
                dpi=300, bbox_inches='tight')
    plt.close()

def print_analysis_results(analysis_results: Dict[str, pd.DataFrame], output_dir: str):
    """打印分析结果"""
    print("\n=== 月度分析 ===")
    print(analysis_results['monthly_analysis'])
    
    print("\n=== 地区分析 ===")
    print(analysis_results['region_analysis'])
    
    print("\n=== 城市分析（TOP 10）===")
    print(analysis_results['city_analysis'].head(10))
    
    print("\n=== 销售渠道分析 ===")
    print(analysis_results['channel_analysis'])
    
    print("\n=== 产品分析 ===")
    print(analysis_results['product_analysis'])
    
    print(f"\n可视化图表已保存到目录 {output_dir}：")
    print("1. monthly_sales_trend.png - 月度销售趋势图")
    print("2. region_sales_distribution.png - 地区销售分布图")
    print("3. channel_sales_pie.png - 渠道销售占比饼图")
    print("4. top10_cities_sales.png - TOP10城市销售额对比图")
    print("5. product_profit_matrix.png - 产品利润矩阵分析图")
    print("6. region_channel_heatmap.png - 区域渠道利润热力图")
    print("7. city_sales_top10_horizontal.png - 城市销售额Top10横向柱状图")
    print("8. channel_profit_radar.png - 渠道利润率雷达图")
    print("9. sales_profit_trend.png - 销售额和利润率趋势图")
    print("10. cost_profit_scatter.png - 成本-利润分布图")

if __name__ == "__main__":
    # 从CSV文件读取销售数据
    # 注意：这里假设数据文件在data目录下，文件名格式为sales_data_*.csv
    data_dir = "data"
    csv_files = [f for f in os.listdir(data_dir) if f.startswith('sales_data_') and f.endswith('.csv')]
    
    if not csv_files:
        print("错误：未找到销售数据文件！")
        exit(1)
    
    # 使用最新的数据文件
    latest_file = sorted(csv_files)[-1]
    df = pd.read_csv(os.path.join(data_dir, latest_file))
    
    # 进行销售分析
    output_dir = "reports"
    analysis_results = analyze_sales_data(df, output_dir)
    print_analysis_results(analysis_results, output_dir)