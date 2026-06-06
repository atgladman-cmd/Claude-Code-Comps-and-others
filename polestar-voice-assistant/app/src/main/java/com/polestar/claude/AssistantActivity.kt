package com.polestar.claude

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.speech.RecognitionListener
import android.speech.RecognizerIntent
import android.speech.SpeechRecognizer
import android.speech.tts.TextToSpeech
import android.speech.tts.UtteranceProgressListener
import android.view.View
import android.view.WindowManager
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.lifecycle.lifecycleScope
import com.polestar.claude.databinding.ActivityAssistantBinding
import kotlinx.coroutines.launch
import java.util.Locale

class AssistantActivity : AppCompatActivity(), TextToSpeech.OnInitListener {

    private lateinit var binding: ActivityAssistantBinding
    private lateinit var speechRecognizer: SpeechRecognizer
    private lateinit var tts: TextToSpeech

    private val claudeClient = ClaudeApiClient()
    private val conversation = ConversationManager()

    private var ttsReady = false
    private var currentState = State.IDLE

    private enum class State { IDLE, LISTENING, THINKING, SPEAKING }

    // -------------------------------------------------------------------------
    // Lifecycle
    // -------------------------------------------------------------------------

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityAssistantBinding.inflate(layoutInflater)
        setContentView(binding.root)

        // Keep screen on while the assistant is active
        window.addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON)

        tts = TextToSpeech(this, this)

        binding.micButton.setOnClickListener {
            when (currentState) {
                State.IDLE -> startListening()
                State.LISTENING -> stopListeningManually()
                State.THINKING, State.SPEAKING -> Unit // ignore taps while busy
            }
        }

        binding.newConversationButton.setOnClickListener {
            conversation.clear()
            setIdle("Conversation cleared. Tap the mic to start.")
        }

        binding.closeButton.setOnClickListener { finish() }

        if (ContextCompat.checkSelfPermission(this, Manifest.permission.RECORD_AUDIO)
            != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                this, arrayOf(Manifest.permission.RECORD_AUDIO), RC_AUDIO
            )
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        // Steering wheel button pressed again while already showing
        if (currentState == State.IDLE) startListening()
    }

    override fun onResume() {
        super.onResume()
        // Auto-start listening when launched by the assistant trigger
        val action = intent.action
        if (action == Intent.ACTION_ASSIST ||
            action == "android.intent.action.VOICE_COMMAND"
        ) {
            if (ttsReady) startListening() // else onInit will call it
        }
    }

    override fun onDestroy() {
        super.onDestroy()
        if (::speechRecognizer.isInitialized) speechRecognizer.destroy()
        tts.shutdown()
    }

    override fun onRequestPermissionsResult(
        requestCode: Int, permissions: Array<String>, grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == RC_AUDIO &&
            grantResults.firstOrNull() == PackageManager.PERMISSION_GRANTED
        ) {
            startListening()
        } else if (requestCode == RC_AUDIO) {
            setIdle("Microphone permission is required.")
        }
    }

    // -------------------------------------------------------------------------
    // TTS
    // -------------------------------------------------------------------------

    override fun onInit(status: Int) {
        if (status == TextToSpeech.SUCCESS) {
            tts.language = Locale.getDefault()
            tts.setOnUtteranceProgressListener(object : UtteranceProgressListener() {
                override fun onStart(utteranceId: String) {}
                override fun onDone(utteranceId: String) {
                    runOnUiThread { onSpokenDone() }
                }
                override fun onError(utteranceId: String) {
                    runOnUiThread { onSpokenDone() }
                }
            })
            ttsReady = true
            // If we were waiting for TTS before listening, start now
            val action = intent.action
            if (action == Intent.ACTION_ASSIST || action == "android.intent.action.VOICE_COMMAND") {
                startListening()
            }
        }
    }

    private fun speak(text: String) {
        setState(State.SPEAKING)
        binding.statusText.text = "Claude"
        binding.responseText.text = text
        tts.speak(text, TextToSpeech.QUEUE_FLUSH, null, UTTERANCE_ID)
    }

    private fun onSpokenDone() {
        // After Claude speaks, wait a beat then start listening again (hands-free loop)
        setIdle("Tap to ask another question")
    }

    // -------------------------------------------------------------------------
    // Speech recognition
    // -------------------------------------------------------------------------

    private fun startListening() {
        if (!ttsReady) return
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.RECORD_AUDIO)
            != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                this, arrayOf(Manifest.permission.RECORD_AUDIO), RC_AUDIO
            )
            return
        }

        tts.stop()
        if (::speechRecognizer.isInitialized) speechRecognizer.destroy()
        speechRecognizer = SpeechRecognizer.createSpeechRecognizer(this)
        speechRecognizer.setRecognitionListener(recognitionListener)

        val recognizerIntent = Intent(RecognizerIntent.ACTION_RECOGNIZE_SPEECH).apply {
            putExtra(RecognizerIntent.EXTRA_LANGUAGE_MODEL, RecognizerIntent.LANGUAGE_MODEL_FREE_FORM)
            putExtra(RecognizerIntent.EXTRA_LANGUAGE, Locale.getDefault())
            putExtra(RecognizerIntent.EXTRA_MAX_RESULTS, 1)
            putExtra(RecognizerIntent.EXTRA_PARTIAL_RESULTS, true)
            putExtra(RecognizerIntent.EXTRA_SPEECH_INPUT_COMPLETE_SILENCE_LENGTH_MILLIS, 1500L)
            putExtra(RecognizerIntent.EXTRA_SPEECH_INPUT_POSSIBLY_COMPLETE_SILENCE_LENGTH_MILLIS, 1500L)
        }

        setState(State.LISTENING)
        speechRecognizer.startListening(recognizerIntent)
    }

    private fun stopListeningManually() {
        if (::speechRecognizer.isInitialized) speechRecognizer.stopListening()
    }

    private val recognitionListener = object : RecognitionListener {
        override fun onReadyForSpeech(params: Bundle?) {
            binding.statusText.text = "Listening…"
        }

        override fun onBeginningOfSpeech() {
            binding.statusText.text = "Listening…"
        }

        override fun onRmsChanged(rmsdB: Float) {
            // Animate mic button scale based on volume
            val scale = 1f + (rmsdB.coerceIn(0f, 10f) / 10f) * 0.3f
            binding.micButton.animate().scaleX(scale).scaleY(scale).setDuration(80).start()
        }

        override fun onBufferReceived(buffer: ByteArray?) {}

        override fun onEndOfSpeech() {
            binding.micButton.animate().scaleX(1f).scaleY(1f).setDuration(100).start()
            binding.statusText.text = "Processing…"
        }

        override fun onError(error: Int) {
            binding.micButton.animate().scaleX(1f).scaleY(1f).setDuration(100).start()
            val msg = when (error) {
                SpeechRecognizer.ERROR_AUDIO -> "Audio error. Try again."
                SpeechRecognizer.ERROR_NO_MATCH -> "Didn't catch that. Tap to try again."
                SpeechRecognizer.ERROR_SPEECH_TIMEOUT -> "No speech detected. Tap to try again."
                SpeechRecognizer.ERROR_NETWORK, SpeechRecognizer.ERROR_NETWORK_TIMEOUT ->
                    "Network error. Check your connection."
                else -> "Didn't catch that. Tap to try again."
            }
            setIdle(msg)
        }

        override fun onResults(results: Bundle?) {
            binding.micButton.animate().scaleX(1f).scaleY(1f).setDuration(100).start()
            val matches = results?.getStringArrayList(SpeechRecognizer.RESULTS_RECOGNITION)
            val spoken = matches?.firstOrNull()?.trim()
            if (spoken.isNullOrBlank()) {
                setIdle("Didn't catch that. Tap to try again.")
            } else {
                sendToClaude(spoken)
            }
        }

        override fun onPartialResults(partialResults: Bundle?) {
            val partial = partialResults
                ?.getStringArrayList(SpeechRecognizer.RESULTS_RECOGNITION)
                ?.firstOrNull() ?: return
            binding.responseText.text = partial
        }

        override fun onEvent(eventType: Int, params: Bundle?) {}
    }

    // -------------------------------------------------------------------------
    // Claude API
    // -------------------------------------------------------------------------

    private fun sendToClaude(userText: String) {
        setState(State.THINKING)
        binding.statusText.text = "Thinking…"
        binding.responseText.text = "\"$userText\""

        conversation.addUserMessage(userText)

        val prefs = getSharedPreferences(PREFS, MODE_PRIVATE)
        val apiKey = prefs.getString(KEY_API_KEY, BuildConfig.CLAUDE_API_KEY) ?: ""

        lifecycleScope.launch {
            when (val result = claudeClient.sendMessage(apiKey, conversation.history)) {
                is ClaudeResult.Success -> {
                    conversation.addAssistantMessage(result.text)
                    speak(result.text)
                }
                is ClaudeResult.Error -> {
                    conversation.clear()
                    setIdle(result.message)
                    if (ttsReady) {
                        tts.speak(result.message, TextToSpeech.QUEUE_FLUSH, null, UTTERANCE_ID)
                    }
                }
            }
        }
    }

    // -------------------------------------------------------------------------
    // UI state helpers
    // -------------------------------------------------------------------------

    private fun setState(state: State) {
        currentState = state
        binding.micButton.animate().scaleX(1f).scaleY(1f).setDuration(0).start()
        when (state) {
            State.IDLE -> {
                binding.listeningIndicator.visibility = View.INVISIBLE
                binding.thinkingIndicator.visibility = View.GONE
                binding.micButton.isEnabled = true
                binding.micButton.alpha = 1f
            }
            State.LISTENING -> {
                binding.listeningIndicator.visibility = View.VISIBLE
                binding.thinkingIndicator.visibility = View.GONE
                binding.micButton.isEnabled = true
                binding.micButton.alpha = 1f
            }
            State.THINKING -> {
                binding.listeningIndicator.visibility = View.INVISIBLE
                binding.thinkingIndicator.visibility = View.VISIBLE
                binding.micButton.isEnabled = false
                binding.micButton.alpha = 0.5f
            }
            State.SPEAKING -> {
                binding.listeningIndicator.visibility = View.INVISIBLE
                binding.thinkingIndicator.visibility = View.GONE
                binding.micButton.isEnabled = false
                binding.micButton.alpha = 0.5f
            }
        }
    }

    private fun setIdle(hint: String) {
        setState(State.IDLE)
        binding.statusText.text = hint
    }

    companion object {
        private const val RC_AUDIO = 100
        private const val UTTERANCE_ID = "claude_response"
        const val PREFS = "polestar_claude_prefs"
        const val KEY_API_KEY = "claude_api_key"
    }
}
