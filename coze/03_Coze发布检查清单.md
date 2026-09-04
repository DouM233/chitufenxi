# Coze 发布检查清单

## Bot / Workflow

- 已创建 Bot：赤兔客服聊天分析师
- 默认模型设为 `gpt5.6sol`
- 已开启文件上传
- 已配置工作流入口参数
- 已配置文件读取节点
- 已配置 Excel/CSV/TXT/LOG 解析
- 已配置历史基准读取
- 已配置 P0/质量/安全二次复核
- 已配置 Excel 生成
- 已配置 Markdown 管理结论生成
- 已配置 result_manifest.json
- 已配置最终 JSON 返回

## 知识/规则

- 已上传 `rules/` 下所有规则文件
- 已上传 `coze/prompts/` 下提示模板
- 已配置新品首次分析模板
- 已配置有历史基准分析模板

## 存储

- 已明确《赤兔历史分析结果》的实际存储位置
- 已确定 Coze 文件 ID 与历史目录之间的映射
- 已配置 `analysis_index`
- 已保留原始聊天记录
- 已保留中间过程文件
- 已保留最终报告文件

## 网页

- 已填写 `web/config.js`
- 已确认接入方式：Chat SDK / Workflow API Proxy / Coze 低代码页面
- 文件上传功能可用
- 开始分析按钮可触发工作流
- 最终结果可展示
- Excel 入口可下载

