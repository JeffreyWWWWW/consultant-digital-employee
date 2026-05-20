# 搜索源接入说明

本文件定义外部资料检索时可用搜索源的接入方式。搜索源用于“找线索”，核验结论仍以 `source-verification.md` 为准。

## 检索执行顺序

按以下顺序执行：

1. `site:` 官网定向查询
2. Tavily
3. 百度搜索
4. 用户提供链接或手工补充链接

`site:` 是搜索语法，不是独立 API 或搜索源；它应写在 `query` 中，实际执行源仍然是 Tavily 或百度。官网或原始来源定向查询先由 Tavily 执行；Tavily 没有找到原文、只有转载页、结果质量低或工具失败时，再由百度执行同一条 `site:` 查询和中文改写查询。

上一层已经找到并核验原文时，可以停止。上一层没有结果、只有转载页、结果质量低或工具失败时，进入下一层。

## 搜索源登记

每个搜索源应按统一字段记录：

```json
{
  "provider": "Tavily/百度/Tavily+百度/用户提供",
  "query": "实际检索词",
  "status": "success/partial/no_results/failed/not_available/not_run",
  "result_count": 0,
  "notes": "失败原因、不可用原因或结果质量说明"
}
```

状态含义：

- `success`：搜索源返回可用结果。
- `partial`：返回结果但只有转载、摘要或线索页。
- `no_results`：搜索源正常返回但没有结果。
- `failed`：搜索源调用失败。
- `not_available`：当前环境没有该搜索源工具、API key 或权限。
- `not_run`：已有更高优先级结果，未调用该搜索源。

## Tavily

用于快速泛搜、发现可能标题、转载页和官网线索。

适用场景：

- 不确定资料全名；
- 需要快速发现多个候选来源；
- 需要找转载线索后再反查原文。

记录要求：

- `search_provider` 写 `Tavily`。
- 使用 `site:` 定向查询时，`query` 保留完整 `site:` 检索词，`search_provider` 仍写 `Tavily`。
- 如果只找到转载页，`source_type` 不得写 `官网原文`。

## 百度搜索

用于中文覆盖补充和原文反查。百度结果噪音较大，只作为 fallback 和中文反查线索。

适用场景：

- Tavily 没找到中文资料或地方公开信息；
- Tavily 只返回英文、摘要或低质量结果；
- 需要用资料标题、文号、项目主体、发布时间做中文反查；
- 需要从新闻转载、公众号标题反查原始来源。

记录要求：

- Openclaw 中通过 `baidu_search` 工具调用百度搜索；API Key 从本地环境变量 `BAIDU_APPBUILDER_KEY` 读取。
- 单独使用百度时，`search_provider` 写 `百度`。
- Tavily 找线索、百度补充反查后共同确认时，`search_provider` 写 `Tavily+百度`。
- 百度补充时应先执行 Tavily 未命中的同一条 `site:` 查询，再执行中文改写查询。
- 商业媒体、公众号、聚合页结果只能标为 `媒体转载` 或 `线索页`。
- 如果当前环境没有百度搜索工具或 key，记录 `not_available`，不要假装已检索。

## 统一输出

无论使用哪个搜索源，进入 `sections.sources` 前必须归一为以下字段：

```json
{
  "name": "来源名称",
  "agency": "发布单位或来源机构",
  "doc_no": "文号（如有）",
  "date": "发布时间（如有）",
  "url": "原文链接或线索链接",
  "source_type": "官网原文/政府公开平台/机构原文/主体公开页面/政府转载/媒体转载/线索页/未确认",
  "verification_status": "已核验原文/待核验原文/待联网核验/未找到原文",
  "used_for": "支撑内容",
  "query": "实际检索词",
  "search_provider": "Tavily/百度/Tavily+百度/用户提供"
}
```

## 工具不可用处理

如果用户要求使用百度，但当前环境没有对应工具、API key 或权限：

- 不报“已使用百度”。
- 在检索记录中写 `not_available`。
- 继续使用当前可用搜索工具和官方域名定向搜索。
- 在附录B中说明相关来源仍需人工联网核验。
