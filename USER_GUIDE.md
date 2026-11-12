# Chitran — User Guide & Command Cheat Sheet

**Chitran** is your local-first, ₹0-cost desktop companion and accountability system built on Windows 11.

You have **two ways** to control the exact same brain:
1. **Ditto the Cyber-Slime (Mouse Hub — Bottom Taskbar):** Click (left-click or right-click) **Ditto** inside your taskbar to open your quick-action menu without typing any commands.
2. **Flow Launcher (`Alt + Space` — Keyboard Fast-Path):** When you're already typing code or notes, press `Alt + Space` and type `cn <command>` without touching your mouse.

---

## 1. Understanding Ditto (Your Taskbar Cyber-Slime)

Ditto is both your **quick launcher** and your **live status & accountability indicator**:

| Ditto's Look | State | What It Means |
| :--- | :--- | :--- |
| **Ice-Cyan glossy slime**, hopping & blinking `^ ^` | `FOCUSED` | All clear. Monitoring active, no overdue deadlines. |
| **Warm Amber-Cream + Cyber-Goggles**, sparks | `WORKING` | Running a background job for you (e.g., Colab execution). |
| **Coral-Pink slime**, wide crimson eyes + **`!`** badge | `ALERT` | A goal deadline is overdue or distraction threshold hit. |
| **Mist-Grey melted puddle**, calm shut eyes | `PAUSED` | Monitoring and reminders are paused. |
| **Sky-Aqua + droplet / Mint-Teal + envelope** | `HYDRATION` / `MAIL` | Water break reminder or email briefing ready. |

### Ditto Controls (On the Taskbar Ditto)
* **Left-Click (single tap) or Right-Click:** Opens the **Ditto Cyber-HUD Menu** (Goals, Ask Mentor, Quick Note, Auto-Fill Form, Mute Voice, User Guide, Minimize to Tray, Exit).
* **Drag (click + hold + move):** Move Ditto anywhere along your taskbar.
* **Double-Click:** Quickly toggle Pause / Resume monitoring.
* **Minimize to Tray (`^`):** Hides the taskbar slime to save space while keeping Ditto alive inside Windows' System Tray arrow (`^` near your clock).
  * **To bring Ditto back anytime:** Either click the **`^`** arrow in your taskbar → right-click Ditto → **Show Taskbar Ditto**, OR press `Alt + Space` → type **`cn ditto`** → press `Enter`.
* **Exit Ditto Completely:** Shuts down both the taskbar Ditto and the tray icon cleanly. (Launching any `cn` command later will bring Ditto back.)

---

## 2. Commands Cheat Sheet (`Alt + Space` → `cn ...`)

### Implemented & Ready Today
| Command | What It Does |
| :--- | :--- |
| **`cn goal <task> @ <HH:MM>`** | Logs a daily goal + deadline in SQLite (accepts `21:30` or `9:30 pm`). |
| **`cn goal <task>`** *(no time)* | Opens a quick **Deadline Required** popup with 1-click presets (`+1 hr`, `+2 hrs`, `21:00`, etc.). |
| **`cn goal`** *(empty)* | Opens the **Daily Goal Planner HUD** to view today's goals, mark items done, or add targets. |
| **`cn done <id or keyword>`** | Marks a goal completed (e.g. `cn done 2` or `cn done DBMS`). |
| **`cn note <text>`** | Appends a timestamped bullet to `C:\Users\Rohit Patil\Daily_Notes\YYYY-MM-DD.md`. |
| **`cn ask <question>`** | Queries your AI mentor (Groq / Gemini free tier) and opens the floating dark HUD. |
| **`cn fill [role]`** | Auto-fills active Chrome form fields from `config/profile_vault.json` (never clicks Submit). |
| **`cn colab [url]`** | Runs active Google Colab cells in the background with Win+L keepalive lock. |
| **`cn say <text>`** / **`cn mute`** | Speaks via headphones (`edge-tts`) or toggles voice mute on/off. |
| **`cn pause`** | Toggles monitoring pause/resume (switches Ditto between `focused` and `paused`). |
| **`cn ditto`** | Launches or restores **Ditto** inside your Windows taskbar. |
| **`cn help`** / **`cn guide`** | Opens this User Guide inside the floating dark HUD. |
| **`cn kill`** | Emergency brake: cancels active background tasks and releases the keepalive lock. |
| **`cn mail`** | Reads today's emails for `rohit.newchallenger@gmail.com` via background Chrome CDP, speaks summary & opens HUD (auto-runs at 10:00 AM & 10:30 PM active). |
| **`cn research <topic>`** | Searches technical sources in parallel, saves a structured Markdown Context Pack to `~/Daily_Notes/context_packs/`, and opens a 1-click **Copy Prompt for Claude/ChatGPT** HUD. |
| **`cn organize`** | Scans `Downloads` and opens a **Dark Cyber-HUD confirmation preview** to sort matching files into `Documents` subfolders (`Certificates`, `Resumes`, `Data Science`, `Power BI Desktop`, `Tableau - Exclr`, `sql -exclr`, `Excel-exclr`, `Data-Analyst_assignment_Excelr`) with **zero deletion**. |
| **`cn undo`** | Safely moves every file from the last `cn organize` batch right back to `Downloads` using `data/undo_log.json`. |
| **`cn doc` / `cn excel` / `cn ppt`** | Two-stage Office generator: researches topic -> opens **Outline Approval HUD** (live slide count, **Revise with AI**, optional hackathon/college `.pptx` template branding, Mermaid flowcharts, & web image search) -> builds `.docx`/`.xlsx`/`.pptx` into `~/Documents/Chitran_Outputs/`. |
| **`cn stats`** | Opens your responsive dark-mode **Analytics Dashboard** (`http://127.0.0.1:7777/dashboard`) with Day/Week/Month focus charts, completed goals, app breakdown, hydration streaks & ₹0 API quota gauges. |
| *Background Monitor* | Runs automatically every 15s (100% local): logs focus vs. distraction, nudges on non-study YouTube before deadlines, and fires 45-min water / 60-min eye breaks. |

### Coming in Upcoming Steps (Roadmap)
* **PPT Design Skill Polish:** Dedicated slide-design & layout proportions skill to bring generated `.pptx` decks closer to Claude/Gemini presentation quality.

---

## 3. Built-In Guardrails
* **₹0 Cost Guarantee:** Quota tracking warns at 80% and hard-stops at 95% of daily free-tier limits.
* **Zero Silent Edits / Submissions:** `cn fill` never clicks Submit; you always review first.
* **Logs for Troubleshooting:** Background daemon and Ditto logs are saved locally in `C:\Users\Rohit Patil\sidekick-os\data\daemon.log` and `data\mascot.log`.
