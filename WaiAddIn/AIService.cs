using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WaiAddIn
{
    public class AIService
    {
        private string _baseUrl;
        private string _apiKey;
        private string _model;

        public AIService(string apiKey, string baseUrl = "https://api.deepseek.com/chat/completions", string model = "deepseek-v4-flash")
        {
            _apiKey = apiKey;
            _baseUrl = baseUrl;
            _model = model;
        }

        public void UpdateConfig(string apiKey, string baseUrl, string model)
        {
            _apiKey = apiKey;
            _baseUrl = baseUrl;
            _model = model;
        }

        public Task<string> CallAsync(string systemPrompt, IEnumerable<ChatMessageData> history, string userMessage, double temperature = 0.2, int maxTokens = 2048)
        {
            var messages = new List<ChatMessageData>();
            if (history != null)
            {
                messages.AddRange(history);
            }

            if (!string.IsNullOrWhiteSpace(userMessage))
            {
                messages.Add(new ChatMessageData { Role = "user", Content = userMessage });
            }

            return CallInternalAsync(systemPrompt, messages, temperature, maxTokens);
        }

        public async Task<string> CallAsync(string systemPrompt, string userMessage, double temperature = 0.2, int maxTokens = 2048)
        {
            var messages = new List<ChatMessageData>();
            if (!string.IsNullOrWhiteSpace(userMessage))
            {
                messages.Add(new ChatMessageData { Role = "user", Content = userMessage });
            }

            return await CallInternalAsync(systemPrompt, messages, temperature, maxTokens);
        }

        private async Task<string> CallInternalAsync(string systemPrompt, IEnumerable<ChatMessageData> messages, double temperature, int maxTokens)
        {
            if (string.IsNullOrWhiteSpace(_apiKey))
                return "[错误] 未配置 API Key。请在设置中填入。";

            var requestMessages = new List<object>();
            if (!string.IsNullOrWhiteSpace(systemPrompt))
            {
                requestMessages.Add(new { role = "system", content = systemPrompt });
            }

            if (messages != null)
            {
                foreach (var message in messages)
                {
                    if (message == null || string.IsNullOrWhiteSpace(message.Content))
                        continue;

                    requestMessages.Add(new { role = string.IsNullOrWhiteSpace(message.Role) ? "user" : message.Role, content = message.Content });
                }
            }

            var body = new
            {
                model = _model,
                messages = requestMessages,
                temperature = temperature,
                max_tokens = maxTokens
            };

            var json = JsonConvert.SerializeObject(body);

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(_baseUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Headers.Add("Authorization", $"Bearer {_apiKey}");

                var requestBytes = Encoding.UTF8.GetBytes(json);
                using (var stream = await request.GetRequestStreamAsync())
                {
                    await stream.WriteAsync(requestBytes, 0, requestBytes.Length);
                }

                string responseBody;
                using (var response = (HttpWebResponse)await request.GetResponseAsync())
                using (var responseStream = response.GetResponseStream())
                using (var reader = new System.IO.StreamReader(responseStream, Encoding.UTF8))
                {
                    responseBody = await reader.ReadToEndAsync();
                    if ((int)response.StatusCode >= 200 && (int)response.StatusCode < 300)
                    {
                        var obj = JObject.Parse(responseBody);
                        return obj["choices"]?[0]?["message"]?["content"]?.ToString() ?? responseBody;
                    }
                    return $"[错误 {response.StatusCode}] {responseBody}";
                }
            }
            catch (WebException webEx)
            {
                try
                {
                    using (var response = (HttpWebResponse)webEx.Response)
                    using (var responseStream = response?.GetResponseStream())
                    using (var reader = responseStream == null ? null : new System.IO.StreamReader(responseStream, Encoding.UTF8))
                    {
                        var responseBody = reader?.ReadToEnd() ?? webEx.Message;
                        var status = response != null ? ((int)response.StatusCode).ToString() : "0";
                        return $"[错误 {status}] {responseBody}";
                    }
                }
                catch
                {
                    return $"[错误] 调用失败：{webEx.Message}";
                }
            }
            catch (Exception ex)
            {
                return $"[错误] 调用失败：{ex.Message}";
            }
        }
    }
}
