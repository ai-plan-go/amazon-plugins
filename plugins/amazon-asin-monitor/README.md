# Amazon ASIN Monitor v1.2.0

可共享的 Codex 插件与技能包，用于 Amazon 美国、英国、德国、澳洲、法国站 ASIN 前台巡检。

## 本版本能力

- 表格按行选择站点，空白默认美国站。
- 新模板取消“父ASIN网址”，保留旧表兼容读取。
- 检测父 ASIN 变化、子体迁移、子体增减和关联集合变化。
- Buy Box 独立检查 Add to Cart、Buy Now 和 buybox 信号。
- 按站点处理域名、语言、币种和邮编。
- 本地化识别划线价、Coupon、Prime 折扣和多买折扣。
- 类目节点只取顶部 breadcrumb；大类/小类排名只取 Best Sellers Rank。
- Excel 嵌入星级占比图与价格截图，并输出逐行自检结果。

## 仓库结构

- `.codex-plugin/plugin.json`: Codex 插件清单
- `skills/amazon-asin-monitor/`: 技能说明、规则、脚本和测试
- `assets/Amazon-ASIN检查基础信息模板-五站点.xlsx`: 五站点输入模板
- `amazon_frontend_check.py`: 便于直接部署的根目录脚本（GitHub 仓库中保留）

## 部署

1. 将 `skills/amazon-asin-monitor/scripts/amazon_frontend_check.py` 复制到 `D:\Codex\amazon_frontend_check.py`。
2. 将输入工作簿放入 `D:\Codex\各项目链接检查\<项目名称>\1_基础信息`。
3. 用隐藏窗口 PowerShell 方式运行脚本。
4. 从 `D:\Codex\last_run_summary.json` 和项目的 `2_输出信息` 读取结果。

完整路径、执行与判断规则见技能目录中的 `references`。

## GitHub 更新

在本仓库目录执行：

```powershell
git push -u origin main
```

如果尚未绑定远程仓库，先创建 GitHub 仓库并执行：

```powershell
git remote add origin https://github.com/<组织或用户名>/<仓库名>.git
git push -u origin main
```

队友可克隆仓库后安装插件，或直接使用发布页中的 ZIP 包。
