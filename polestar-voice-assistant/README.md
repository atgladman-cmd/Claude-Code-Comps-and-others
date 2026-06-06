# Claude for Polestar 2

A native Android voice assistant that lets you talk to Claude AI by pressing the voice button on your Polestar 2's steering wheel.

## How it works

```
Steering wheel voice button
        │
        ▼
Android Auto sends ASSIST / VOICE_COMMAND intent
        │
        ▼
This app launches AssistantActivity (registered as default digital assistant)
        │
        ▼
Android SpeechRecognizer listens → transcribes your speech
        │
        ▼
Anthropic Claude API (claude-opus-4-8) answers
        │
        ▼
Android TTS speaks the response aloud in the car
```

The conversation is multi-turn — Claude remembers context across exchanges within a session. Tap the reset icon to start a new topic.

---

## Requirements

| Requirement | Version |
|---|---|
| Android Studio | Hedgehog (2023.1.1) or newer |
| Android SDK | API 26+ (Android 8.0) |
| Target device | Android phone with Android Auto |
| Car | Polestar 2 (Android Automotive OS with Android Auto) |
| API key | [Anthropic Claude API key](https://console.anthropic.com/) |

---

## Build & install

### 1. Clone and open

```bash
git clone <this-repo>
cd polestar-voice-assistant
```

Open in Android Studio → *File → Open* → select the `polestar-voice-assistant` folder.

### 2. Supply your API key

**Option A — gradle.properties (local, never committed):**

Add to `gradle.properties`:
```
CLAUDE_API_KEY=sk-ant-your-key-here
```

**Option B — enter it at runtime in the app Settings screen** (recommended for sharing builds).

### 3. Build and install

```bash
./gradlew installDebug
# or use Android Studio's Run button
```

---

## First-time setup on the phone

1. **Open the app** and enter your Anthropic API key on the Settings screen. Tap *Save key*.

2. **Set this app as your default digital assistant:**
   - Tap *"Set as default digital assistant"* in the app — it opens Android Settings directly.
   - Navigate to: *Settings → Apps → Default apps → Digital assistant app*
   - Select **Claude for Polestar**.

3. **Connect to your Polestar 2** via Android Auto (USB or wireless).

4. **Press the steering wheel voice button** — the app opens and immediately starts listening.

5. **Speak your question.** Claude answers aloud. Press the button again for a follow-up.

---

## Steering wheel button — technical detail

The Polestar 2's steering wheel voice button sends `KEYCODE_VOICE_ASSIST` to Android Auto, which forwards an `android.intent.action.ASSIST` intent to whichever app is set as the device's default digital assistant.

`AssistantActivity` declares intent filters for:
- `android.intent.action.ASSIST` — long-press Home / steering wheel
- `android.intent.action.VOICE_COMMAND` — Android Auto voice button

When either intent arrives, the activity launches, acquires the mic, and immediately begins speech recognition — no tapping required.

---

## Project structure

```
app/src/main/
├── AndroidManifest.xml             — permissions + assistant intent filters
└── java/com/polestar/claude/
    ├── MainActivity.kt             — settings screen (API key, setup steps)
    ├── AssistantActivity.kt        — voice UI + state machine
    ├── ClaudeApiClient.kt          — Anthropic API calls via OkHttp
    └── ConversationManager.kt      — multi-turn conversation history
```

---

## Customising Claude's behaviour

Edit the `SYSTEM_PROMPT` constant in `ClaudeApiClient.kt`:

```kotlin
private const val SYSTEM_PROMPT = """You are a voice assistant in a Polestar 2...
```

The default prompt instructs Claude to keep responses brief and avoid markdown so they read naturally when spoken aloud.

---

## Privacy

- Audio is processed on-device by Android's SpeechRecognizer (Google Speech Services by default).
- Transcribed text is sent to Anthropic's API over HTTPS.
- No data is stored or logged by this app beyond your session conversation history (in memory only).
- Your API key is stored in Android's SharedPreferences (encrypted on devices running Android 6+).

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Steering wheel button still opens Google Assistant | Re-check *Settings → Apps → Default apps → Digital assistant* |
| "No API key set" spoken on first use | Open the app on the phone, enter and save your key |
| Speech recognition doesn't start | Grant microphone permission when prompted |
| App not visible in default assistant list | Make sure the APK is installed (not just sideloaded via adb without install) |
| Android Auto doesn't show the app | Check that your phone's Android Auto version is 9.0+ |
