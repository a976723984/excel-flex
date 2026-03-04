import { getConfig, setConfig } from "../services/state.js";

export function initConfigPanel(state) {
  const panel = document.getElementById("config-panel");
  if (!panel) return;

  const config = getConfig(state);

  panel.innerHTML = "";

  const title = document.createElement("div");
  title.className = "config-title";
  title.textContent = "大模型配置";

  const form = document.createElement("div");

  const apiUrlGroup = document.createElement("div");
  apiUrlGroup.className = "config-group";
  const apiUrlLabel = document.createElement("label");
  apiUrlLabel.className = "config-label";
  apiUrlLabel.textContent = "API 地址";
  const apiUrlInput = document.createElement("input");
  apiUrlInput.className = "config-input";
  apiUrlInput.placeholder = "例如：https://api.your-llm.com/v1/chat/completions";
  apiUrlInput.value = config.apiUrl || "";
  apiUrlInput.addEventListener("change", () => {
    setConfig(state, { apiUrl: apiUrlInput.value.trim() });
  });
  apiUrlGroup.appendChild(apiUrlLabel);
  apiUrlGroup.appendChild(apiUrlInput);

  const apiKeyGroup = document.createElement("div");
  apiKeyGroup.className = "config-group";
  const apiKeyLabel = document.createElement("label");
  apiKeyLabel.className = "config-label";
  apiKeyLabel.textContent = "API Key / Secret";
  const apiKeyInput = document.createElement("input");
  apiKeyInput.className = "config-input";
  apiKeyInput.type = "password";
  apiKeyInput.placeholder = "只保存在本地浏览器";
  apiKeyInput.value = config.apiKey || "";
  apiKeyInput.addEventListener("change", () => {
    setConfig(state, { apiKey: apiKeyInput.value.trim() });
  });
  apiKeyGroup.appendChild(apiKeyLabel);
  apiKeyGroup.appendChild(apiKeyInput);

  const modelGroup = document.createElement("div");
  modelGroup.className = "config-group";
  const modelLabel = document.createElement("label");
  modelLabel.className = "config-label";
  modelLabel.textContent = "模型名称（可选）";
  const modelInput = document.createElement("input");
  modelInput.className = "config-input";
  modelInput.placeholder = "例如：gpt-4.1, qwen-max 等";
  modelInput.value = config.model || "";
  modelInput.addEventListener("change", () => {
    setConfig(state, { model: modelInput.value.trim() });
  });
  modelGroup.appendChild(modelLabel);
  modelGroup.appendChild(modelInput);

  const help = document.createElement("div");
  help.className = "config-help";
  help.textContent =
    "注意：示例项目不会将配置上传到服务器，仅存储在浏览器 localStorage 中。";

  form.appendChild(apiUrlGroup);
  form.appendChild(apiKeyGroup);
  form.appendChild(modelGroup);
  form.appendChild(help);

  panel.appendChild(title);
  panel.appendChild(form);
}

