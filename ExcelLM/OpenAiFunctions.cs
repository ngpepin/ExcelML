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
        private const string DefaultEndpoint = "https://api.openai.com/v1/chat/completions";
        private const double DefaultTemperature = 1.0d;
        private const int DefaultMaxTokens = 256;
        private static readonly HttpClient HttpClient = new HttpClient();

        [ExcelFunction(Name = "OpenAI.PROMPT", Description = "Calls an OpenAI-compatible endpoint and returns the first response message.", IsVolatile = true)]
        public static object Prompt(string promptText, string model, object baseUrl, object apiKey, object temperature, object maxTokens)
        {
            var userBaseUrl = NormalizeString(baseUrl);
            var normalizedApiKey = NormalizeString(apiKey);
            var normalizedModel = string.IsNullOrWhiteSpace(model) ? "gpt-3.5-turbo" : model.Trim();
            var normalizedPrompt = promptText ?? string.Empty;
            var temperatureValue = NormalizeDouble(temperature) ?? DefaultTemperature;
            var maxTokensValue = NormalizeInt(maxTokens) ?? DefaultMaxTokens;
            var endpoint = ResolveEndpoint(userBaseUrl);

            if (string.IsNullOrWhiteSpace(normalizedApiKey))
                return "#ERR: API key is required";

            return ExcelAsyncUtil.Run(
                "OpenAI.PROMPT",
                new object[] { normalizedPrompt, normalizedModel, endpoint, normalizedApiKey, temperatureValue, maxTokensValue },
                () => ExecutePrompt(normalizedPrompt, normalizedModel, endpoint, normalizedApiKey, temperatureValue, maxTokensValue));
        }

        private static string ResolveEndpoint(string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(baseUrl))
                return DefaultEndpoint;

            var trimmed = baseUrl.Trim();
            trimmed = trimmed.TrimEnd('/');

            if (trimmed.IndexOf("/chat/completions", StringComparison.OrdinalIgnoreCase) >= 0)
                return trimmed;

            return trimmed + "/chat/completions";
        }

        private static string ExecutePrompt(string promptText, string model, string endpoint, string apiKey, double temperature, int? maxTokens)
        {
            try
            {
                var requestPayload = new OpenAiChatRequest
                {
                    Model = model,
                    Temperature = temperature,
                    Messages = new[] { new OpenAiMessage { Role = "user", Content = promptText } },
                    MaxTokens = maxTokens
                };

                var options = new JsonSerializerOptions
                {
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                var json = JsonSerializer.Serialize(requestPayload, options);
                using (var requestMessage = new HttpRequestMessage(HttpMethod.Post, endpoint)
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                })
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                    using (var response = HttpClient.SendAsync(requestMessage).GetAwaiter().GetResult())
                    {
                        var responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                        if (!response.IsSuccessStatusCode)
                        {
                            var errorMessage = TryExtractError(responseBody);
                            return $"#ERR: {response.StatusCode} {errorMessage}";
                        }

                        return TryExtractContent(responseBody) ?? "#ERR: No content returned";
                    }
                }
            }
            catch (Exception ex)
            {
                return "#ERR: " + ex.Message;
            }
        }

        private static double? NormalizeDouble(object value)
        {
            if (value is ExcelReference reference)
            {
                var coerced = XlCall.Excel(XlCall.xlCoerce, reference);
                return NormalizeDouble(coerced);
            }

            if (value is ExcelMissing || value is ExcelEmpty || value is null)
                return null;
            if (value is double d && !double.IsNaN(d))
                return d;
            if (double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
                return parsed;
            return null;
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

        private static string TryExtractContent(string responseBody)
        {
            try
            {
                using (var doc = JsonDocument.Parse(responseBody))
                {
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
            }
            catch
            {
                // ignore parsing issues and fall through
            }

            return null;
        }

        private static string TryExtractError(string responseBody)
        {
            try
            {
                using (var doc = JsonDocument.Parse(responseBody))
                {
                    if (doc.RootElement.TryGetProperty("error", out var error) && error.TryGetProperty("message", out var message))
                    {
                        return message.GetString() ?? responseBody;
                    }
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

        private class OpenAiChatRequest
        {
            [JsonPropertyName("model")]
            public string Model { get; set; }

            [JsonPropertyName("temperature")]
            public double Temperature { get; set; }

            [JsonPropertyName("messages")]
            public OpenAiMessage[] Messages { get; set; }

            [JsonPropertyName("max_tokens")]
            public int? MaxTokens { get; set; }
        }

        private class OpenAiMessage
        {
            [JsonPropertyName("role")]
            public string Role { get; set; }

            [JsonPropertyName("content")]
            public string Content { get; set; }
        }
    }
}
