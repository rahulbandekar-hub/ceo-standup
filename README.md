# Prime Harvest Digest

A weekly, fully automated "leadership meeting minutes" generator for PRIMEHARVEST — a fictitious CPG company I invented as a strategy-practice exercise. I play CEO; the script researches real CPG/retail industry news via Claude's web search tool, role-plays a full leadership team's take on it, and writes the result straight to a formatted document every Sunday morning. It's a way to keep commercial-strategy thinking sharp against a fresh competitive narrative every week, without spending my own research time on it.

**Architecture:** one scheduled Claude API call with server-side web search enabled, prompted to generate nine functional leadership perspectives (marketing, supply chain, finance, etc.) plus a CEO synthesis, written directly to a `.docx` on the Desktop. No database, no server, no cloud dependency — a single Python script triggered weekly by a macOS LaunchAgent.

## The interesting problem: cutting cost 10x without cutting quality

The original version made ~20 uncapped web searches and generated long output on every run — $1–3/week, fine in isolation but not worth that for a personal practice tool. The rewrite:

- Switched to Claude's **dynamic-filtering search tool**, which screens search results server-side before they consume context, instead of dumping full pages into the prompt and filtering after.
- Capped web search to 8 uses per run instead of 20.
- Tightened the prompt: 1–2 bullets per persona instead of 2–3, dropped a "what's changed since last week" section (and the previous-digest lookup that fed it — one less thing that could break).
- Lowered `max_tokens` from 16,000 to 6,000 to match the now-shorter target output.

Net result: ~$0.10–0.15/run, down from $1–3, with no real loss in the quality of the exercise.

**A smaller lesson from the same project:** the very first scheduled run silently failed because a Mac sleep cycle dropped the connection mid-request, and the script had no retry logic. Any unattended job making a network call longer than a few seconds needs to assume the machine might sleep mid-request — fixed by wrapping the scheduled job in `caffeinate -i`, which blocks idle sleep for the duration of the call.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure your API key

```bash
cp .env.example .env
```

Edit `.env` with your [Anthropic API key](https://console.anthropic.com/).

### 3. Personalize the prompt

Open `generate.py` and edit `PROMPT_TEMPLATE` — swap in your own background, the fictitious (or real) company you want to frame the exercise around, and the personas/sections you want covered.

### 4. Run it

```bash
python3 generate.py
```

The digest is written to `~/Desktop/Prime Harvest/` as a `.docx`.

### 5. (Optional) Schedule it weekly

```bash
cp com.primeharvestdigest.plist.example ~/Library/LaunchAgents/com.primeharvestdigest.plist
```

Edit the plist to replace `YOUR_USERNAME` with your Mac username, then load it:

```bash
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.primeharvestdigest.plist
```

Runs every Sunday at 9:00 AM local time, wrapped in `caffeinate -i` so it survives idle sleep mid-request.

## Cost

~$0.10–0.15 per run with the current prompt and search cap — well under $1/month at a weekly cadence.

## License

MIT
