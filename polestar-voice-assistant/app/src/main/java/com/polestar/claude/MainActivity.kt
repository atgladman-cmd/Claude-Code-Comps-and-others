package com.polestar.claude

import android.content.Intent
import android.os.Bundle
import android.provider.Settings
import android.view.inputmethod.EditorInfo
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import com.polestar.claude.databinding.ActivityMainBinding

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        loadSavedKey()

        binding.saveButton.setOnClickListener { saveKey() }

        binding.apiKeyInput.setOnEditorActionListener { _, actionId, _ ->
            if (actionId == EditorInfo.IME_ACTION_DONE) { saveKey(); true } else false
        }

        binding.launchButton.setOnClickListener {
            startActivity(Intent(this, AssistantActivity::class.java))
        }

        binding.defaultAssistantButton.setOnClickListener {
            // Opens Android's "Default apps" settings so the user can choose this app
            // as their digital assistant — that's what routes the steering wheel button here.
            startActivity(Intent(Settings.ACTION_VOICE_INPUT_SETTINGS))
        }
    }

    private fun loadSavedKey() {
        val prefs = getSharedPreferences(AssistantActivity.PREFS, MODE_PRIVATE)
        val saved = prefs.getString(AssistantActivity.KEY_API_KEY, BuildConfig.CLAUDE_API_KEY) ?: ""
        if (saved.isNotBlank()) {
            // Show masked key
            binding.apiKeyInput.setText(saved)
        }
    }

    private fun saveKey() {
        val key = binding.apiKeyInput.text?.toString()?.trim() ?: ""
        if (key.isBlank()) {
            Toast.makeText(this, "Please enter your Claude API key", Toast.LENGTH_SHORT).show()
            return
        }
        getSharedPreferences(AssistantActivity.PREFS, MODE_PRIVATE)
            .edit()
            .putString(AssistantActivity.KEY_API_KEY, key)
            .apply()
        Toast.makeText(this, "API key saved", Toast.LENGTH_SHORT).show()
        binding.apiKeyInput.clearFocus()
    }
}
