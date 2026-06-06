package com.polestar.claude

import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.OkHttpClient
import okhttp3.Request
import okhttp3.RequestBody.Companion.toRequestBody
import org.json.JSONArray
import org.json.JSONObject
import java.util.concurrent.TimeUnit

data class Message(val role: String, val content: String)

sealed class ClaudeResult {
    data class Success(val text: String) : ClaudeResult()
    data class Error(val message: String) : ClaudeResult()
}

class ClaudeApiClient {

    private val http = OkHttpClient.Builder()
        .connectTimeout(15, TimeUnit.SECONDS)
        .readTimeout(30, TimeUnit.SECONDS)
        .writeTimeout(15, TimeUnit.SECONDS)
        .build()

    private val json = "application/json; charset=utf-8".toMediaType()

    suspend fun sendMessage(
        apiKey: String,
        history: List<Message>,
        systemPrompt: String = SYSTEM_PROMPT
    ): ClaudeResult = withContext(Dispatchers.IO) {
        if (apiKey.isBlank()) {
            return@withContext ClaudeResult.Error("No API key set. Open the app and add your Claude API key in Settings.")
        }

        val messages = JSONArray().apply {
            history.forEach { msg ->
                put(JSONObject().apply {
                    put("role", msg.role)
                    put("content", msg.content)
                })
            }
        }

        val body = JSONObject().apply {
            put("model", "claude-opus-4-8")
            put("max_tokens", 1024)
            put("system", systemPrompt)
            put("messages", messages)
        }.toString().toRequestBody(json)

        val request = Request.Builder()
            .url("https://api.anthropic.com/v1/messages")
            .addHeader("x-api-key", apiKey)
            .addHeader("anthropic-version", "2023-06-01")
            .addHeader("content-type", "application/json")
            .post(body)
            .build()

        runCatching {
            http.newCall(request).execute().use { response ->
                val responseBody = response.body?.string() ?: ""
                if (!response.isSuccessful) {
                    val errorMsg = runCatching {
                        JSONObject(responseBody).optJSONObject("error")?.optString("message") ?: responseBody
                    }.getOrDefault(responseBody)
                    ClaudeResult.Error("Claude API error ${response.code}: $errorMsg")
                } else {
                    val text = JSONObject(responseBody)
                        .getJSONArray("content")
                        .getJSONObject(0)
                        .getString("text")
                    ClaudeResult.Success(text)
                }
            }
        }.getOrElse { e ->
            ClaudeResult.Error("Network error: ${e.message}")
        }
    }

    companion object {
        private const val SYSTEM_PROMPT = """You are a voice assistant integrated into a Polestar 2 electric vehicle.
Your responses are spoken aloud while the driver is in the car, so:
- Keep answers concise and clear — aim for 2-3 sentences unless more detail is genuinely needed.
- Never use markdown, bullet points, or formatting — plain spoken text only.
- For navigation or safety questions, prioritise brevity so the driver can act quickly.
- You can discuss any topic: news, music, directions, general knowledge, or just conversation."""
    }
}
