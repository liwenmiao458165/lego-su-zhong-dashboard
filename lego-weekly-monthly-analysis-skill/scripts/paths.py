import os
# 自包含 skill: 项目根默认=本脚本所在目录(scripts/), 可用环境变量 PROJECT_DIR 覆盖
HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.environ.get("PROJECT_DIR", HERE)
OUTPUTS_DIR = os.path.join(PROJECT_DIR, "outputs")
