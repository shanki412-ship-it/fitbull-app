# FitBull – app code (update channel)

*Strong habits. Smarter training.*

This repository holds only the **app code** for FitBull, a personal coach that runs in your own Google account
(Google Sheets + Apps Script): no personal settings, no data, no keys.

FitBull apps check `version.json` here once a day. When there is a newer version, the owner gets an email and an
**Update now** button in their app; one tap installs `Code.gs` and `Log.html` into their own Apps Script project.
Their Config (goals, plan), Sheet, logs and Gemini key are never touched.

| File | What it is |
|---|---|
| `Code.gs` | the engine |
| `Log.html` | the phone app |
| `appsscript.json` | project settings, including the two permissions the Update button needs |
| `version.json` | the latest version and what's new |
