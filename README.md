# PST → EML Converter

> **Fork notice:** This project is a **fork** of the original [kostigas/PstToEmlConverter](https://github.com/kostigas/PstToEmlConverter). It retains the same core functionality while adding a key fix — **sent and received timestamps (date + time) are now preserved in the EML output**. See [What changed in this fork](#-what-changed-in-this-fork) below.

A **free, open-source Windows utility** that migrates Outlook PST archives into standard `.eml` files while mirroring your original folder layout. No paywalls, no trial timers, no feature restrictions.

👉 **Grab the installer from the Releases page (this fork):**  
https://github.com/jacobalf/pst-to-eml-converter/releases/

---

## ✨ Highlights

- ✅ Converts **PST → EML**
- 🗂️ Handles a **single PST file** or a **whole folder** of PSTs
- 📁 Reproduces the Outlook folder hierarchy
- ⚡ Uses Outlook's native engine for broad compatibility & speed
- 🧾 Produces a detailed conversion log
- 🆓 **Completely free and open source**
- 🪟 Windows desktop app built with WPF
- 🔒 Runs fully **offline** — no internet or cloud involved

---

## 📌 A Note Before You Begin

> **Microsoft Outlook must be installed on the same machine.**

The converter talks to PST files through Outlook's official COM interface. This guarantees rock-solid compatibility with virtually every PST version — but it also means Outlook needs to be present locally.

---

## 📦 Installation

1. Head over to the **Releases page**  
   👉 https://github.com/jacobalf/pst-to-eml-converter/releases/
2. Download the latest package (e.g. `PstToEmlConverter-1.0.1-Setup.exe`)
3. Run the installer
4. Launch **PST → EML Converter** from the Start Menu

That's it — there's nothing else to configure.

---

## 🧭 Getting Started

1. Pick your **Source**
   - A single PST file, **or**
   - A folder that contains one or more PST files
2. Choose the **Destination** folder for the results
3. Toggle any options you'd like:
   - Preserve folder structure
   - Skip files that already exist
4. Hit **Start**
5. Let it run _(large PSTs can look frozen — Outlook just needs a moment, be patient)_

When it finishes, your emails will be sitting in the destination as `.eml` files.

---

## 📄 What Ends Up in the Output?

| Item type        | Outcome                   |
| ---------------- | ------------------------- |
| Email messages   | ✅ Saved as `.eml`        |
| Attachments      | ✅ Embedded in the `.eml` |
| Contacts         | ⏭️ Skipped                |
| Calendar entries | ⏭️ Skipped                |
| Tasks and notes  | ⏭️ Skipped                |

This behavior is deliberate — EML is, after all, an email format.

---

## 🧾 Logging & Errors

- On-screen progress keeps you informed while converting
- A debug log is dropped into the output directory for deeper inspection
- Individual errors are recorded but **never halt the whole batch**

---

## 🔁 What Changed in This Fork

Relative to the original repository, this fork introduces the following:

| Area                              | Change                                                                                    |
| --------------------------------- | ----------------------------------------------------------------------------------------- |
| **Date sent preservation 🕓**     | The `Date:` header now reflects the real **sent** timestamp via Outlook MAPI properties.  |
| **Date received preservation 📥** | A `Received:` header is now written with the original **received** date & time.           |
| **Reliability**                   | A new `TryGetPropertyDateTime` helper pulls timestamps from the most reliable source first.|
| **Version**                       | Bumped to **v1.0.1**; installer compiled to `PstToEmlConverter-1.0.1-Setup.exe`.          |

> **Upstream tracking:** The upstream source targets the `.net10.0-windows` framework; this fork currently builds against `.net8.0-windows` for toolchain stability. Full details are logged in [`CHANGELOG.md`](CHANGELOG.md).

You'll find the complete, line-by-line record of every modification in the [`CHANGELOG.md`](CHANGELOG.md) file at the repository root.

---

## 🤔 Why Another PST Converter?

Most PST tools on the market are:

- Painfully expensive
- Crippled by trial modes
- Far more than you need for a one-off move

This project's whole reason for being is a **simple, transparent, and free** option that anyone can use and trust.

---

## 📜 License

Released under the **MIT License** — use it, modify it, distribute it freely.  
See [`LICENSE`](LICENSE) for the full terms.

---

## 🧑💻 Contributing

Open to contributions of all kinds:

- Bug reports
- Feature suggestions
- Code improvements
- Documentation polish

Just open an issue or send in a pull request — on this fork's repository at https://github.com/jacobalf/pst-to-eml-converter.

---

## 🔍 Keywords (SEO)

PST to EML converter, Outlook PST to EML, free PST converter, open source PST to EML, Outlook PST export, email migration tool, PST email extractor, Windows PST converter, Outlook timestamp preservation, PST date received fix

---

## 🚀 Project Status

- ✔️ Actively maintained
- ✔️ Stable for day-to-day use
- ✔️ Open to enhancements

---

**⭐ Found it handy? Consider starring the repository — it helps others discover it!**
