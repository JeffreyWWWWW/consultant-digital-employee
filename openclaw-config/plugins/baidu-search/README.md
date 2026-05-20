# Baidu Search Plugin

Openclaw 百度搜索补充工具。它不替换 Tavily，只提供 `baidu_search` 工具用于中文政策、资料和案例的补充检索。

## 配置

把百度千帆 AppBuilder API Key 放到本机环境变量：

```powershell
setx BAIDU_APPBUILDER_KEY "bce-v3/..."
```

重启 Openclaw 后生效。

## 使用

优先使用 Tavily。Tavily 无结果、只有转载页、摘要或低质量结果时，再调用：

```json
{
  "query": "site:gov.cn 国发〔2022〕14号",
  "top_k": 5
}
```

返回结果只作为线索。是否可写为 `已核验原文`，仍以打开官网原文并匹配标题、文号、发文单位或发布时间为准。
