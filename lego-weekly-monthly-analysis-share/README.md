# 乐高苏中周月维度分析看板 — 分享包

> 本包包含一份 WorkBuddy Skill（`lego-weekly-monthly-analysis`）+ 它运行所需的全部脚本。
> 同事拿到后按下面步骤即可在自己的环境跑起「周维度 + 月维度」交互分析看板。

---

## ⚠️ 本包不含什么（重要）

- ❌ **没有任何销售数据**（订单 CSV / 销售明细 xlsx / outputs 里的 JSON 与 HTML）。
- ❌ **没有任何密钥/凭证**（`.gh_token`、`.cos_secret*`、`.oss_*` 等均未打包）。
- ✅ 只含 Skill 知识文档 + 纯 Python 脚本（404K）。

数据、密钥、看板产物都需要同事在自己的机器上准备。

---

## 一、目录结构

```
lego-weekly-monthly-analysis-share/
├── README.md
├── skill/                                  # ← 放进 ~/.workbuddy/skills/
│   ├── SKILL.md
│   └── references/weekly_monthly_rules.md
└── scripts/                                # ← 放进你自己的项目目录（如 ~/lego-proj/）
    ├── auto_run_weekly.py                  # 入口：一键跑完 12 脚本流水线 + 部署
    ├── gen_weekly_html.py / _p2/_p3/_p4.py # 周 HTML 生成（4 个）
    ├── gen_monthly_html.py / _p2/_p3.py    # 月 HTML 生成（3 个）
    ├── init_weekly_json.py                 # JSON 基础结构
    ├── extract_weekly_guide.py             # 周导购/门店同比
    ├── extract_weekly_enhanced.py          # 周环比/系列/产品/样品/成交率
    ├── extract_monthly.py                  # 月全部数据
    ├── inject_json.py                      # 把 JSON 注入 HTML 的 const DATA
    ├── wos_attr.py                         # 共享模块（WOS 属性），被 3 个 extract 引用
    ├── deploy_cos.py / deploy_oss.py / deploy_via_api.py  # 三种部署方式
```

---

## 二、安装（2 步）

### 1. 装 Skill
把 `skill/` 整个目录复制到 WorkBuddy 的用户级 skill 目录：
```
cp -R skill ~/.workbuddy/skills/lego-weekly-monthly-analysis
```
重启/刷新 WorkBuddy 后，说「更新周月分析」「周月看板」等即可触发。

### 2. 放脚本
把 `scripts/` 整个目录复制到**同一个项目目录**（例如 `~/lego-proj/`）。
所有脚本必须**同级**，因为 `auto_run_weekly.py` 用 `PROJECT_DIR` 拼接子脚本路径调用。

---

## 三、同事必须改的配置（否则跑不起来）

`scripts/auto_run_weekly.py` 头部有几个**硬编码绝对路径**，请改成同事自己机器的值：

| 行 | 变量 | 改什么 |
|---|---|---|
| ~17 | `PROJECT_DIR` | 改成 `scripts/` 所在目录，如 `/Users/同事名/lego-proj` |
| ~20 | `PYTHON_BIN` | 改成同事的 python（建议 WorkBuddy 托管版 venv：`~/.workbuddy/binaries/python/envs/default/bin/python`，或系统 `python3`） |
| ~22-23 | `DOWNLOADS_DIR` / `DESKTOP_DIR` | 销售订单文件所在目录（默认从 Downloads/Desktop 找） |
| ~109 | `NODE_BIN` | 改成同事的 node（用于部署前 JS 语法校验），无 node 可注释掉校验段 |

> 部署脚本（`deploy_via_api.py` 等）的密钥走环境变量或同目录隐藏文件（`.gh_token` 等），**不要写进脚本**。同事需自行配置 GitHub Personal Access Token（见第五节）。

---

## 四、运行前需自备的数据输入

脚本运行时依赖以下外部文件（**本包不提供**）：

1. **销售订单数据**：`销售订单*.csv`（Downloads）或 `销售订单明细查询*.xlsx`（Desktop）。脚本自动 glob 最新文件。
2. **日看板产物 `dashboard_latest.json`**：`init_weekly_json.py` 会读它拿区域数据。需先跑日看板（或提供同结构 JSON）放到 `PROJECT_DIR/outputs/`。
3. **指标同期 Excel**：如 `M月指标同期.xlsx`（含去年同月、日目标/日同期）。放 Desktop 或脚本约定路径。
4. **客流 Excel**（用于成交率模块）：
   - 月客流：`M月客流.xlsx`（全月累计，7店 / 门店名+客流两列）
   - 周客流：`wkNN周客流.xlsx`（当周，文件名带周号，需同步改 `extract_weekly_enhanced.py` 里的 `WK_FLOW_FILE`）

---

## 五、生成 + 部署

```
cd ~/lego-proj/scripts
python3 auto_run_weekly.py
```

脚本会：生成周/月 HTML → 提取 JSON → 注入 → 跑 JS 校验 → 部署。

**部署方式（三选一，auto_run_weekly.py 按顺序尝试）：**
- **GitHub Pages（推荐）**：配置 `GH_TOKEN`（环境变量或 `.gh_token` 文件）、`GH_OWNER`、`REPO=lego-su-zhong-dashboard`，推 `gh-pages` 分支。自定义域名在 `deploy_via_api.py` 里配 `GH_CUSTOM_DOMAIN`。
- **腾讯云 COS** / **阿里云 OSS**：配置对应 `.cos_*` / `.oss_*` 隐藏文件。

---

## 六、跨区域适配提示

本 skill 是**苏中 7 店定制**（扬州万象汇/京华城/江都金鹰、泰州万象城、宿迁宝龙、淮安新亚/万象城）。
若同事是其他区域：
- 门店列表 + 门店名映射（如「扬州万象汇新店→扬州万象汇」「淮安新亚广场→淮安新亚」）散落在各 `extract_*.py`，需按本区域重改。
- 目标/同期逻辑、配色、重点 SKU 等也可能要调整（详见 `skill/references/weekly_monthly_rules.md`）。

---

## 七、详细规则

所有 KPI 口径、JSON 结构、模块结论/机会点生成逻辑、成交率模块、跨月周取数铁律等，**见 `skill/references/weekly_monthly_rules.md`（第十八章）**，修改看板前务必先读。
