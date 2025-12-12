# 🎯 Word Autofiller Pro

Automated word completion tool for games like Roblox's "Last Letter".

## 🚀 Features

- ✨ Automatic word completion
- 🔍 OCR screen scanning (Admin mode)
- 🎨 Modern dark mode UI
- 🔊 Text-to-speech
- ⚙️ Customizable settings

## 📦 Automated Builds

Every push to `main` automatically:
1. ✅ Builds Windows EXE
2. 📤 Uploads to GitHub Artifacts
3. 🎯 Sends to Discord via webhook

## 🔧 Setup

### GitHub Secrets

Add this secret to your repository:
- `DISCORD_WEBHOOK`: Your Discord webhook URL

**Steps:**
1. Go to your repository → Settings → Secrets and variables → Actions
2. Click "New repository secret"
3. Name: `DISCORD_WEBHOOK`
4. Value: Your webhook URL (e.g., `https://discord.com/api/webhooks/...`)

### Manual Build (Local)
```bash
python build_config.py
```

## 📥 Download

Latest build: [GitHub Actions Artifacts](../../actions)

Or wait for Discord notification! 🎯

---

Made with ♥️ for "Last Letter"
