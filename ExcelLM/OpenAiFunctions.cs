using System;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelDna.Integration;

namespace ExcelLM
{
    public static class OpenAiFunctions
    {
        private static readonly HttpClient HttpClient = new HttpClient();

        [ExcelFunction(Name = "OpenAI.PROMPT", Description = "Calls an OpenAI-compatible endpoint and returns the first response message.")]
        public static object Prompt(string promptText, string model, object baseUrl, object apiKey, object maxTokens)
        {
            var normalizedBaseUrl = NormalizeString(baseUrl);
            var normalizedApiKey = NormalizeString(apiKey);
            var normalizedModel = string.IsNullOrWhiteSpace(model) ? "gpt-3.5-turbo" : model.Trim();
            var normalizedPrompt = promptText ?? string.Empty;
            int? maxTokensValue = NormalizeInt(maxTokens);

            if (string.IsNullOrWhiteSpace(normalizedBaseUrl))
                normalizedBaseUrl = "https://api.openai.com/v1/chat/completions";

            if (string.IsNullOrWhiteSpace(normalizedApiKey))
                return "#ERR: API key is required";

            return ExcelAsyncUtil.Run(
                "OpenAI.PROMPT",
                new object[] { normalizedPrompt, normalizedModel, normalizedBaseUrl, normalizedApiKey, maxTokensValue ?? -1 },
                () => ExecutePrompt(normalizedPrompt, normalizedModel, normalizedBaseUrl, normalizedApiKey, maxTokensValue));
        }

        private static string ExecutePrompt(string promptText, string model, string baseUrl, string apiKey, int? maxTokens)
        {
            try
            {
                var requestPayload = new OpenAiChatRequest
                {
                    Model = model,
                    Messages = new[] { new OpenAiMessage { Role = "user", Content = promptText } },
                    MaxTokens = maxTokens
                };

                var options = new JsonSerializerOptions
                {
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                var json = JsonSerializer.Serialize(requestPayload, options);
                using var requestMessage = new HttpRequestMessage(HttpMethod.Post, baseUrl)
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                };

                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                using var response = HttpClient.SendAsync(requestMessage).GetAwaiter().GetResult();
                var responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                if (!response.IsSuccessStatusCode)
                {
                    var errorMessage = TryExtractError(responseBody);
                    return $"#ERR: {response.StatusCode} {errorMessage}";
                }

                return TryExtractContent(responseBody) ?? "#ERR: No content returned";
            }
            catch (Exception ex)
            {
                return "#ERR: " + ex.Message;
            }
        }

        private static string TryExtractContent(string responseBody)
        {
            try
            {
                using var doc = JsonDocument.Parse(responseBody);
                if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                {
                    var first = choices[0];
                    if (first.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
                    {
                        return content.GetString()?.Trim();
                    }

                    if (first.TryGetProperty("text", out var textElement))
                    {
                        return textElement.GetString()?.Trim();
                    }
                }
            }
            catch
            {
                // fall through to error below
            }

            return null;
        }

        private static string TryExtractError(string responseBody)
        {
            try
            {
                using var doc = JsonDocument.Parse(responseBody);
                if (doc.RootElement.TryGetProperty("error", out var error) && error.TryGetProperty("message", out var message))
                {
                    return message.GetString() ?? responseBody;
                }
            }
            catch
            {
                // ignore JSON parse errors
            }

            return responseBody;
        }

        private static string NormalizeString(object value)
        {
            if (value is ExcelReference reference)
            {
                var coerced = XlCall.Excel(XlCall.xlCoerce, reference);
                return NormalizeString(coerced);
            }

            if (value is ExcelMissing || value is ExcelEmpty || value is null)
                return null;
            if (value is string s)
                return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
            if (value is double d)
                return d.ToString(CultureInfo.InvariantCulture);
            return value.ToString();
        }

        private static int? NormalizeInt(object value)
        {
            if (value is ExcelReference reference)
            {
                var coerced = XlCall.Excel(XlCall.xlCoerce, reference);
                return NormalizeInt(coerced);
            }

            if (value is ExcelMissing || value is ExcelEmpty || value is null)
                return null;
            if (value is double d && !double.IsNaN(d))
                return Convert.ToInt32(d);
            if (int.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                return parsed;
            return null;
        }

        private class OpenAiChatRequest
        {
            public string Model { get; set; }
            public OpenAiMessage[] Messages { get; set; }
            public int? MaxTokens { get; set; }
        }

        private class OpenAiMessage
        {
            public string Role { get; set; }
            public string Content { get; set; }
        }
    }
}
