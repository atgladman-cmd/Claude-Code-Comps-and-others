package com.polestar.claude

class ConversationManager {

    private val _history = mutableListOf<Message>()
    val history: List<Message> get() = _history.toList()

    fun addUserMessage(text: String) {
        _history.add(Message("user", text))
        trimIfNeeded()
    }

    fun addAssistantMessage(text: String) {
        _history.add(Message("assistant", text))
    }

    fun clear() = _history.clear()

    val isEmpty get() = _history.isEmpty()

    // Keep the last 10 exchanges to avoid bloating the context
    private fun trimIfNeeded() {
        if (_history.size > 20) {
            _history.subList(0, _history.size - 20).clear()
        }
    }
}
