# 📊 Mobile Financial System Analyzer — Bangladesh
**বিকাশ · নগদ · রকেট** স্টেটমেন্ট বিশ্লেষণ সফটওয়্যার

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

---

## 🚀 Streamlit Cloud এ Deploy করুন (সহজ)

### ধাপ ১ — GitHub এ আপলোড
1. GitHub.com এ নতুন repository তৈরি করুন (যেমন: `mfs-analyzer`)
2. এই ৩টি ফাইল upload করুন:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`

### ধাপ ২ — Streamlit Cloud Deploy
1. **share.streamlit.io** এ যান
2. GitHub account দিয়ে login করুন
3. **"New app"** ক্লিক করুন
4. Repository: `আপনার-username/mfs-analyzer` নির্বাচন করুন
5. Main file: `app.py`
6. **"Deploy!"** ক্লিক করুন

✅ কয়েক মিনিটে আপনার সাইট live হবে: `https://your-app.streamlit.app`

---

## 💻 লোকালে চালান

```bash
# Install
pip install -r requirements.txt

# Run
streamlit run app.py
```

তারপর: **http://localhost:8501**

---

## ✨ ফিচার
| ফিচার | বিবরণ |
|---|---|
| 🔍 Auto-Detection | বিকাশ / নগদ / রকেট স্বয়ংক্রিয়ভাবে সনাক্ত |
| 📊 Charts | মাসিক, দৈনিক, ঘণ্টাভিত্তিক, Pie chart |
| 💸 Size Analysis | ছোট (২-৮হা) / মধ্যম (৮-২০হা) / উচ্চ (২০হা+) |
| 👥 Top Accounts | শীর্ষ ১০ কাউন্টারপার্ট |
| ⚠️ Alert | রাত্রিকালীন ও সন্দেহজনক কার্যক্রম |
| 📄 HTML Report | প্রিন্টযোগ্য সম্পূর্ণ HTML রিপোর্ট |
| 📝 Word Report | .docx ফরম্যাটে রিপোর্ট |
| 📊 CSV Export | ফিল্টার করা ডেটা CSV তে |

## 📁 File Structure
```
mfs-analyzer/
├── app.py                    ← মূল Streamlit অ্যাপ
├── requirements.txt          ← Python packages
├── README.md
└── .streamlit/
    └── config.toml           ← Theme & settings
```

## 🗂️ Supported Formats
- **নগদ Standard** — Agent/Individual Statement (XLSX)
- **নগদ Business** — Business Statement (XLSX)
- **বিকাশ Business Agent** — Statement (XLSX)
- **রকেট** — DBBL Mobile Banking (XLSX) *(coming soon)*

---
*Developed for Bangladesh MFS Analysis*
