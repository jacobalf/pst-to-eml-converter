# PST to EML Converter

A **free, open-source PST to EML converter for Windows**, built with .NET and Microsoft Outlook.  
Convert Outlook PST files into standard EML files while preserving folder structure — **no paid licenses, no trials, no limitations**.

👉 **Download the installer:**  
https://github.com/kostigas/PstToEmlConverter/releases/

---

## ✨ Features

- ✅ Convert **PST → EML**
- 📂 Supports **single PST file** or **folder with multiple PSTs**
- 🗂️ Preserves Outlook folder structure
- ⚡ Fast and reliable (uses Outlook’s native engine)
- 🧾 Detailed conversion logs
- 🆓 **100% free & open source**
- 🪟 Windows desktop application (WPF)
- 🔒 Runs completely **offline**

---

## ⚠️ Important Requirement

> **Microsoft Outlook must be installed on the same machine**

This tool uses Outlook’s official COM interface to read PST files.  
This guarantees **maximum compatibility** with all PST versions, but Outlook is required.

---

## 📥 Download & Install

1. Go to the **Releases page**  
   👉 https://github.com/kostigas/PstToEmlConverter/releases/
2. Download the latest `.msi` installer
3. Run the installer
4. Launch **PST → EML Converter** from the Start Menu

No additional setup required.

---

## 🛠️ How to Use

1. Choose **Source**
   - Single PST file **or**
   - Folder containing multiple PST files
2. Choose **Destination folder**
3. Select options:
   - Preserve folder structure
   - Skip existing files
4. Click **Start**
5. Wait for conversion to complete  
   *(If it looks stuck — don’t panic, Outlook can be slow on large PSTs)*

Converted emails will be saved as `.eml` files.

---

## 📁 What Gets Converted?

| Item Type | Result |
|---------|-------|
| Emails | ✅ Converted to `.eml` |
| Attachments | ✅ Included in `.eml` |
| Contacts | ⏭️ Skipped |
| Calendar items | ⏭️ Skipped |
| Tasks / Notes | ⏭️ Skipped |

This is intentional — **EML is an email format**.

---

## 🧾 Logging

- Conversion progress is shown inside the app
- A detailed debug log is written to the output directory
- Errors are logged but **do not stop the entire conversion**

---

## 💡 Why This Exists

Most PST converters are:
- Extremely expensive
- Locked behind trials
- Overkill for one-time migrations

This project exists to provide a **simple, transparent, and free alternative** for everyone.

---

## 📜 License

This project is licensed under the **MIT License**.  
You are free to use, modify, and distribute it.

See [`LICENSE`](LICENSE) for details.

---

## 🧑‍💻 Contributing

Contributions are welcome!

- Bug reports
- Feature requests
- Code improvements
- Documentation fixes

Just open an issue or pull request.

---

## 🔍 Keywords (SEO)

PST to EML converter, Outlook PST to EML, free PST converter, open source PST to EML, Outlook PST export, email migration tool, PST email extractor, Windows PST converter

---

## 🚀 Project Status

- ✔️ Actively working
- ✔️ Stable for daily use
- ✔️ Open to improvements

---

**⭐ If you find this useful, please star the repository — it helps others find it!**
