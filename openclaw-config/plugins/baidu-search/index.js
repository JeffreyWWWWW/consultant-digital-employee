const ENDPOINT = "https://qianfan.baidubce.com/v2/ai_search/web_search";

function jsonResult(payload) {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(payload, null, 2),
      },
    ],
    details: payload,
  };
}

function readConfig(api, ctx) {
  return ctx?.getRuntimeConfig?.() ?? ctx?.runtimeConfig ?? ctx?.config ?? api.config ?? {};
}

function readApiKey(api, ctx) {
  const cfg = readConfig(api, ctx);
  const configured = cfg?.baiduSearch?.apiKey;
  if (typeof configured === "string" && configured.trim()) return configured.trim();
  return process.env.BAIDU_APPBUILDER_KEY || "";
}

function normalizeReference(item) {
  return {
    title: item?.title || "",
    url: item?.url || "",
    date: item?.date || "",
    website: item?.website || "",
    type: item?.type || "",
    snippet: item?.snippet || item?.content || "",
  };
}

function createBaiduSearchTool(api, ctx) {
  return {
    name: "baidu_search",
    label: "Baidu Search",
    description:
      "Search Chinese web results using Baidu Qianfan AI Search. Use as a supplement after Tavily returns no result, low quality results, or only转载线索.",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: ["query"],
      properties: {
        query: {
          type: "string",
          description: "Search query string. Supports site: queries and Chinese policy title/doc-no rewrites.",
        },
        top_k: {
          type: "number",
          minimum: 1,
          maximum: 10,
          description: "Maximum web references to return. Default 5.",
        },
      },
    },
    execute: async (_toolCallId, rawParams) => {
      const query = String(rawParams?.query || "").trim();
      if (!query) throw new Error("baidu_search requires query.");

      const apiKey = readApiKey(api, ctx);
      if (!apiKey) {
        throw new Error("BAIDU_APPBUILDER_KEY is not configured.");
      }

      const topK = Math.max(1, Math.min(10, Number(rawParams?.top_k || 5)));
      const cfg = readConfig(api, ctx);
      const baseUrl = cfg?.baiduSearch?.baseUrl || ENDPOINT;
      const payload = {
        messages: [{ role: "user", content: query }],
        search_source: "baidu_search_v2",
        resource_type_filter: [{ type: "web", top_k: topK }],
      };

      const response = await fetch(baseUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json; charset=utf-8",
          "X-Appbuilder-Authorization": `Bearer ${apiKey}`,
        },
        body: JSON.stringify(payload),
      });

      const text = await response.text();
      let data;
      try {
        data = JSON.parse(text);
      } catch {
        data = { raw: text };
      }

      if (!response.ok) {
        throw new Error(`Baidu search failed: ${response.status} ${JSON.stringify(data)}`);
      }

      return jsonResult({
        provider: "百度",
        query,
        request_id: data?.request_id || "",
        result_count: Array.isArray(data?.references) ? data.references.length : 0,
        results: Array.isArray(data?.references) ? data.references.map(normalizeReference) : [],
      });
    },
  };
}

export default {
  id: "baidu-search",
  name: "Baidu Search",
  description: "Baidu Qianfan AI Search tool for Chinese policy source lookup.",
  register(api) {
    api.registerTool((ctx) => createBaiduSearchTool(api, ctx), { name: "baidu_search" });
  },
};
