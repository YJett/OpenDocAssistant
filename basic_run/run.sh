#!/bin/bash

# 这个脚本用于运行PPT代码的主程序
# 使用方法: ./run.sh <模型名称> <模型ID> <运行模式> <数据集>
# 例如: ./run.sh turbo 1 tf short
#
# 参数说明:
# $1: 模型名称 (如 turbo, gpt4 等)
# $2: 模型ID
# $3: 运行模式 (tf 或 sess)
# $4: 数据集名称

Model="$1"
Model_id="$2"
tf_or_sess="$3" 
dataset="$4"

# 检查是否提供了所有必需的参数
if [ -z "$Model" ] || [ -z "$Model_id" ] || [ -z "$tf_or_sess" ] || [ -z "$dataset" ]; then
    echo "错误: 缺少必需的参数"
    echo "用法: ./run.sh <模型名称> <模型ID> <运行模式> <数据集>"
    exit 1
fi

# 运行主程序
python3 ./DOCXAPI/main.py --test --dataset="$dataset" --model="$Model" --"$tf_or_sess" --resume --second --model_id="$Model_id"
