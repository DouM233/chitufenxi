window.CHITU_CONFIG = {
  appName: "赤兔客服聊天分析工作台",
  coze: {
    // manual: 本地预览，只生成任务参数
    // chat_sdk: 使用 Coze Chat SDK 嵌入机器人
    // workflow_api_proxy: 通过你自己的后端代理调用 Coze Workflow API
    mode: "manual",
    botId: "填入 Coze Bot ID",
    publicKey: "如使用 OAuth PKCE，在这里填入 public key",
    connectorId: "如使用发布渠道，在这里填入 connector id",
    apiBase: "以 Coze 后台实际配置为准",
    workflowProxyUrl: "例如：https://your-domain.com/api/chitu-analyze"
  },
  analysis: {
    model: "gpt-5.6-sol",
    historyFolderName: "赤兔历史分析结果"
  }
};
