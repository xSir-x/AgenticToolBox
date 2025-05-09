#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
上传Excel数据到华为云ElasticSearch服务。

此脚本用于将处理好的Excel数据（特别是processed_data.xlsx文件）上传到华为云的ElasticSearch搜索服务中。
脚本会提取Excel中的关键字段（款号、产品名称、品目），并将其作为一个统一的索引存储在ElasticSearch中，
以便后续进行高效的产品搜索和查询。

功能特点:
1. 自动连接到华为云ElasticSearch服务，支持SSL安全连接
2. 自动创建索引及合适的映射，优化中文搜索体验（使用IK分词器）
3. 批量上传数据，提高处理效率
4. 详细的日志记录，便于跟踪和排错
5. 灵活的命令行参数配置

使用方法:
    python upload_to_es.py --host <ES主机地址> --username <用户名> --password <密码> [选项]

参数说明:
    --data, -d       要上传的Excel文件路径 (默认: data/processed_data.xlsx)
    --host           华为云ElasticSearch主机地址 (必填)
    --port           ElasticSearch端口 (默认: 9200)
    --username       ElasticSearch用户名 (必填)
    --password       ElasticSearch密码 (必填)
    --index          索引名称 (默认: products)
    --no-ssl         不使用SSL连接 (默认使用SSL)
    --chunk-size     批量上传的大小 (默认: 1000)

示例:
    python upload_to_es.py --host es-cn-north-4.myhuaweicloud.com --username admin --password pass123 --index jewelry_products

数据映射:
    款号:    作为keyword类型，用于精确匹配和聚合
    产品名称: 使用IK分词器的text类型，支持中文全文搜索
    品目:    使用IK分词器的text类型，支持中文全文搜索
    上传时间: 自动添加的时间戳记录

注意事项:
    - 请确保已安装所需的依赖包：pip install -r requirements.txt
    - 运行前请确认华为云ElasticSearch服务已开通并可正常访问
    - 中文搜索功能依赖于华为云ES服务中已安装的IK分词器
    - 生产环境中应配置适当的证书验证
"""

import os
import sys
import argparse
import pandas as pd
from elasticsearch import Elasticsearch, helpers
from elasticsearch.exceptions import ConnectionError, NotFoundError
import uuid
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("elasticsearch_upload.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("elasticsearch_uploader")

def connect_to_huaweicloud_es(es_host, es_port, es_username, es_password, use_ssl=True):
    """
    连接到华为云ElasticSearch服务
    
    Args:
        es_host (str): ElasticSearch主机地址
        es_port (int): ElasticSearch端口
        es_username (str): 用户名
        es_password (str): 密码
        use_ssl (bool): 是否使用SSL连接
        
    Returns:
        Elasticsearch: ElasticSearch客户端对象
    """
    try:
        # 构建连接URL
        es_url = f"{'https' if use_ssl else 'http'}://{es_host}:{es_port}"
        logger.info(f"正在连接到华为云ElasticSearch服务: {es_url}")
        
        # 创建ES客户端
        es_client = Elasticsearch(
            [es_url],
            http_auth=(es_username, es_password),
            use_ssl=use_ssl,
            verify_certs=False,  # 在生产环境中应设置为True并提供正确的证书
            ssl_show_warn=False
        )
        
        # 验证连接
        if es_client.ping():
            logger.info("成功连接到华为云ElasticSearch服务")
            es_info = es_client.info()
            logger.info(f"ElasticSearch版本: {es_info['version']['number']}")
            return es_client
        else:
            logger.error("无法连接到华为云ElasticSearch服务")
            return None
    except ConnectionError as e:
        logger.error(f"连接到华为云ElasticSearch服务时出错: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"创建ElasticSearch客户端时出错: {str(e)}")
        return None

def create_index_if_not_exists(es_client, index_name, mapping=None):
    """
    如果索引不存在，则创建索引
    
    Args:
        es_client (Elasticsearch): ElasticSearch客户端
        index_name (str): 索引名称
        mapping (dict): 索引映射
        
    Returns:
        bool: 索引是否创建成功
    """
    try:
        if not es_client.indices.exists(index=index_name):
            logger.info(f"创建索引: {index_name}")
            
            # 默认映射，如果没有提供自定义映射
            if mapping is None:
                mapping = {
                    "mappings": {
                        "properties": {
                            "款号": {"type": "keyword"},
                            "产品名称": {
                                "type": "text",
                                "analyzer": "ik_max_word",
                                "search_analyzer": "ik_smart",
                                "fields": {
                                    "keyword": {"type": "keyword", "ignore_above": 256}
                                }
                            },
                            "品目": {
                                "type": "text",
                                "analyzer": "ik_max_word",
                                "search_analyzer": "ik_smart",
                                "fields": {
                                    "keyword": {"type": "keyword", "ignore_above": 256}
                                }
                            },
                            "上传时间": {"type": "date", "format": "yyyy-MM-dd HH:mm:ss||yyyy-MM-dd||epoch_millis"}
                        }
                    },
                    "settings": {
                        "number_of_shards": 3,
                        "number_of_replicas": 1
                    }
                }
                
            es_client.indices.create(index=index_name, body=mapping)
            logger.info(f"索引 {index_name} 创建成功")
            return True
        else:
            logger.info(f"索引 {index_name} 已存在")
            return True
    except Exception as e:
        logger.error(f"创建索引 {index_name} 时出错: {str(e)}")
        return False

def generate_actions(df, index_name):
    """
    生成批量上传操作
    
    Args:
        df (DataFrame): 包含数据的DataFrame
        index_name (str): 索引名称
        
    Returns:
        generator: 批量上传操作
    """
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for _, row in df.iterrows():
        # 创建文档
        doc = {
            "款号": row["款号"],
            "产品名称": row["产品名称"],
            "品目": row["品目"],
            "上传时间": current_time
        }
        
        # 生成文档ID（可以使用款号和产品名称的组合作为唯一标识）
        doc_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, f"{row['款号']}-{row['产品名称']}"))
        
        # 生成操作
        yield {
            "_index": index_name,
            "_id": doc_id,
            "_source": doc
        }

def upload_data_to_es(es_client, data_file, index_name, chunk_size=1000):
    """
    将Excel数据上传到ElasticSearch
    
    Args:
        es_client (Elasticsearch): ElasticSearch客户端
        data_file (str): Excel文件路径
        index_name (str): 索引名称
        chunk_size (int): 批量上传的大小
        
    Returns:
        bool: 上传是否成功
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(data_file):
            logger.error(f"错误: 找不到Excel文件: {data_file}")
            return False
        
        # 读取Excel文件
        logger.info(f"正在读取Excel文件: {data_file}")
        df = pd.read_excel(data_file)
        
        # 显示数据信息
        logger.info(f"Excel数据尺寸: {df.shape[0]}行 x {df.shape[1]}列")
        logger.info(f"Excel列名: {', '.join(df.columns)}")
        
        # 检查必要的列是否存在
        required_columns = ['款号', '产品名称', '品目']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"错误: Excel文件中缺少以下列: {', '.join(missing_columns)}")
            return False
        
        # 创建索引（如果不存在）
        if not create_index_if_not_exists(es_client, index_name):
            return False
        
        # 批量上传数据
        logger.info(f"开始批量上传数据到索引 {index_name}...")
        success, failed = 0, 0
        
        # 使用helpers.bulk进行批量操作
        actions = generate_actions(df, index_name)
        for ok, result in helpers.streaming_bulk(
            es_client,
            actions,
            chunk_size=chunk_size,
            max_retries=3,
            initial_backoff=2,
            max_backoff=60
        ):
            if ok:
                success += 1
            else:
                failed += 1
                logger.warning(f"上传文档失败: {result}")
            
            # 定期报告进度
            if (success + failed) % 1000 == 0:
                logger.info(f"已处理 {success + failed} 条记录 (成功: {success}, 失败: {failed})")
        
        # 刷新索引，使数据立即可见
        es_client.indices.refresh(index=index_name)
        
        logger.info(f"数据上传完成。总记录数: {df.shape[0]}, 成功: {success}, 失败: {failed}")
        
        # 获取索引文档计数
        count = es_client.count(index=index_name)
        logger.info(f"索引 {index_name} 中的文档数量: {count['count']}")
        
        return True
    except Exception as e:
        logger.error(f"上传数据到ElasticSearch时出错: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return False

def main():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 构建数据目录路径
    data_dir = os.path.join(script_dir, "data")
    # 默认Excel文件路径
    default_data_file = os.path.join(data_dir, "processed_data.xlsx")
    
    # 命令行参数解析
    parser = argparse.ArgumentParser(description="将Excel数据上传到华为云ElasticSearch服务")
    parser.add_argument("--data", "-d", help=f"要上传的Excel文件路径 (默认: {default_data_file})",
                        default=default_data_file)
    parser.add_argument("--host", help="华为云ElasticSearch主机地址", required=True)
    parser.add_argument("--port", help="华为云ElasticSearch端口 (默认: 9200)", type=int, default=9200)
    parser.add_argument("--username", help="华为云ElasticSearch用户名", required=True)
    parser.add_argument("--password", help="华为云ElasticSearch密码", required=True)
    parser.add_argument("--index", help="ElasticSearch索引名称 (默认: products)", default="products")
    parser.add_argument("--no-ssl", help="不使用SSL连接", action="store_true")
    parser.add_argument("--chunk-size", help="批量上传的大小 (默认: 1000)", type=int, default=1000)
    
    args = parser.parse_args()
    
    logger.info(f"Python版本: {sys.version}")
    logger.info(f"当前工作目录: {os.getcwd()}")
    
    # 连接到华为云ElasticSearch服务
    es_client = connect_to_huaweicloud_es(
        args.host,
        args.port,
        args.username,
        args.password,
        not args.no_ssl
    )
    
    if es_client:
        # 上传数据
        upload_data_to_es(
            es_client,
            args.data,
            args.index,
            args.chunk_size
        )
    else:
        logger.error("无法连接到华为云ElasticSearch服务，上传操作终止")
        sys.exit(1)

if __name__ == "__main__":
    main()