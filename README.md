# Jarvis-AI
Jarvis-An AI assistant for your computer.
<br>
Author - Arsalan Khan
<br>
First install Ollama into ur PC/Laptop 
<br>
use local/offline version of ollama . not neccesary to sign in
<br>
Download version 3.12 of python
<br>
Use Visual Studio Code 
<br>
Install the following librarires into your Python/Terminal/VS code terminal
<br>
## Libraries Used

### Third-Party Libraries

CustomTkinter<br>
feedparser<br>
MSS<br>
NumPy<br>
Ollama Python<br>
psutil<br>
PyAutoGUI<br>
PyAudio<br>
pyperclip<br>
pytesseract<br>
Requests<br>
screen-brightness-control<br>
SpeechRecognition<br>
Pillow (PIL)<br>
pywin32<br>
pystray<br>
edge-tts<br>

### Python Standard Libraries

ctypes<br>
concurrent.futures<br>
datetime<br>
difflib<br>
io<br>
json<br>
os<br>
re<br>
sqlite3<br>
subprocess<br>
sys<br>
threading<br>
time<br>
webbrowser<br>
pathlib<br>
urllib.parse<br>
tkinter<br>
wave<br>

### External Requirements

Ollama<br>
Tesseract OCR<br>
Windows Speech API (SAPI)<br>
Qwen3 8B model (`qwen3:8b`)<br>
Qwen2.5-VL 7B vision model (`qwen2.5vl:7b`)<br>

# JARVIS — AI-Powered Windows Desktop Assistant

JARVIS is a bilingual, screen-aware Windows 11 desktop assistant that can understand natural voice or typed commands, answer questions using a local AI model, and safely automate supported computer tasks.<br>
The AI runs locally through Ollama, keeping conversations private and removing API costs.<br>

<br>

## 🎙️ Voice and Language Support

- Continuously listens for voice commands while voice control is enabled.<br>
- Supports English and Hindi voice commands.<br>
- Automatically detects whether the user is speaking English or Hindi.<br>
- Understands many naturally worded and Hinglish-style requests through local AI intent recognition.<br>
- Supports wake phrases such as “Hey Jarvis”, “Jarvis”, and “हे जार्विस”.<br>
- Remains awake after the first wake command, so “Hey Jarvis” is not required before every task.<br>
- Can be returned to standby mode when wake-word-only listening is preferred.<br>
- Voice control can be turned ON or OFF from a command, the GUI, or the system-tray menu.<br>
- Speaks responses in English using Windows SAPI.<br>
- Speaks Hindi responses using a Hindi neural voice.<br>
- Also accepts typed commands through the desktop interface.<br>
- Protects the listening loop from individual recognition and Unicode errors.<br>

<br>

## 🧠 Local AI and Natural-Language Understanding

- Uses the local Ollama `qwen3:8b` model for conversations, intent recognition, and task planning.<br>
- Answers general questions without requiring a paid cloud AI API.<br>
- Replies briefly in the same language used by the user.<br>
- Converts different natural phrases with the same meaning into supported actions.<br>
- Creates safe multi-step desktop plans containing up to 14 steps.<br>
- Can re-plan the remaining steps when a screen element is not ready or a step fails.<br>
- Uses only an approved list of desktop actions and never executes AI-generated shell commands.<br>
- Uses the active window and visible screen text to understand words such as “this”, “current”, “here”, and “that”.<br>

<br>

## 🚀 Application Control

- Opens Calculator, Notepad, Paint, Task Manager, File Explorer, Command Prompt, PowerShell, and Camera.<br>
- Finds and opens other applications listed in the Windows Start Menu.<br>
- Uses fuzzy name matching, so the spoken app name does not always need to be exact.<br>
- Recognizes common aliases for Chrome, Edge, Firefox, Brave, Opera, VS Code, Spotify, VLC, WhatsApp, Discord, Teams, Zoom, Word, Excel, PowerPoint, and more.<br>
- Switches to an application that is already running instead of unnecessarily opening another copy.<br>
- Finds and restores minimized applications.<br>
- Matches applications using their window title, process name, aliases, and keywords.<br>

<br>

## 🪟 Window and Desktop Management

- Switches to a named background window and brings it to the front.<br>
- Tracks the active window and the last window selected with the mouse.<br>
- Understands contextual commands such as “close this window”.<br>
- Minimizes the current window or a named window.<br>
- Maximizes the current window or a named window.<br>
- Restores the current window or a named window.<br>
- Preserves a window’s previous maximized state when bringing it back from minimized mode.<br>
- Normally closes a selected app, file, folder, or window while preserving unsaved-work prompts.<br>
- Shows the desktop by minimizing all windows without changing their sizes.<br>
- Restores all windows after showing the desktop.<br>

<br>

## 🌐 Browser and Tab Control

- Works with Chrome, Microsoft Edge, Firefox, Brave, and Opera.<br>
- Opens a new browser tab without resizing the browser window.<br>
- Opens a new browser window.<br>
- Opens a website or search result in a new tab.<br>
- Searches Google in a new tab using a natural command.<br>
- Opens Google, YouTube, ChatGPT, Gmail, Google Drive, GitHub, Wikipedia, Reddit, LinkedIn, Instagram, Facebook, and WhatsApp Web.<br>
- Understands spoken web addresses containing words such as “dot” and “slash”.<br>
- Searches YouTube for a requested song, video, or topic.<br>
- Switches to the next or previous tab.<br>
- Searches through open tabs and switches to a tab using part of its title.<br>
- Closes the current browser tab.<br>
- Searches through open tabs and closes a specifically named tab.<br>
- Reopens the most recently closed tab.<br>
- Opens browser history and browser downloads.<br>
- Opens the bookmark dialog for the current page.<br>
- Zooms in, zooms out, and resets browser zoom.<br>
- Goes backward or forward between pages.<br>
- Refreshes the current page and opens Find on Page.<br>
- Toggles browser full-screen mode.<br>

<br>

## 🖱️ Screen-Aware Computer Control

- Identifies the window the user is currently working on.<br>
- Types dictated English or Unicode/Hindi text into the active application.<br>
- Presses individual keyboard keys such as Enter, Tab, Escape, arrows, Delete, function keys, and more.<br>
- Performs approved keyboard shortcuts and key combinations.<br>
- Copies, cuts, pastes, selects all, saves, undoes, and redoes actions.<br>
- Opens Task View and Windows clipboard history.<br>
- Scrolls up or down by a requested amount.<br>
- Scrolls left or right.<br>
- Scrolls to the top or bottom of a page.<br>
- Moves the mouse to a requested screen element without clicking it.<br>
- Clicks, double-clicks, or right-clicks visible text controls.<br>
- Drags one visible screen element onto another.<br>
- Waits for requested text to appear before continuing a multi-step task.<br>
- Supports screens that span multiple monitors.<br>

<br>

## 👁️ OCR and Local Vision

- Reads visible text from the complete desktop using Tesseract OCR.<br>
- Copies all readable screen text to the clipboard.<br>
- Opens an image selected by the user, extracts its text, and copies it to the clipboard.<br>
- Copies text currently selected in another application.<br>
- Copies a requested text range from one visible word or phrase to another.<br>
- Supports English OCR and Hindi OCR when the Hindi Tesseract language data is installed.<br>
- Uses OCR coordinates to locate and click labels on the screen.<br>
- Can use a local multimodal Ollama model to find icons, images, thumbnails, and controls that contain no readable text.<br>
- Falls back from OCR to local vision when a text-based target cannot be found confidently.<br>
- Vision-based clicking can be turned ON or OFF while Jarvis is running.<br>

<br>

## 🔊 Media, Volume, and Display Controls

- Plays or pauses the currently active media session.<br>
- Moves to the next or previous media track.<br>
- Increases or decreases Windows volume.<br>
- Mutes or unmutes Windows audio.<br>
- Reads the current display brightness.<br>
- Increases or decreases brightness by a default or requested amount.<br>
- Sets display brightness to a specific percentage.<br>

<br>

## 📁 File and Folder Management

- Opens Desktop, Downloads, Documents, Pictures, Music, and Videos folders when available.<br>
- Builds a reusable index of folders inside the user profile.<br>
- Finds and opens folders using approximate names instead of requiring an exact match.<br>
- Searches both normal Windows user folders and OneDrive folders.<br>
- Finds files by partial name inside Desktop, Documents, Downloads, and Pictures.<br>
- Displays matching file results inside the Jarvis interface.<br>
- Opens the best matching file when requested.<br>
- Opens Jarvis’s saved notes and screenshot folders.<br>

<br>

## 📝 Notes, Reminders, and Memory

- Saves timestamped notes to a local text file.<br>
- Opens the saved notes file on command.<br>
- Creates spoken reminders using seconds, minutes, or hours.<br>
- Stores recent conversations in a local SQLite database.<br>
- Permanently remembers information when the user says “remember”.<br>
- Recalls and displays saved memories when requested.<br>
- Adds stored memories as context when answering future local-AI questions.<br>

<br>

## 💻 System Information and Utilities

- Displays and speaks CPU usage.<br>
- Displays and speaks RAM usage.<br>
- Reports disk usage.<br>
- Reports battery percentage and charger connection status.<br>
- Shows live CPU, memory, and battery information in the GUI.<br>
- Reports the current time and date.<br>
- Retrieves current weather, temperature, feels-like temperature, and humidity.<br>
- Retrieves and displays current news headlines.<br>
- Captures timestamped screenshots and saves them in a dedicated folder.<br>

<br>

## 🔐 Windows Power Controls and Safety

- Locks the computer.<br>
- Puts the computer into Sleep mode.<br>
- Hibernates the computer when Windows hibernation is enabled.<br>
- Schedules shutdown with a 15-second delay.<br>
- Schedules restart with a 15-second delay.<br>
- Cancels a scheduled Windows shutdown or restart.<br>
- Requests confirmation before shutdown, restart, sleep, hibernate, and other sensitive actions.<br>
- Rejects negative confirmations even when the response contains another confirmation word.<br>
- Requests confirmation before risky desktop interactions such as deletion, submission, payment, uploading, or sending.<br>
- Treats OCR text as untrusted screen content rather than executable instructions.<br>

<br>

## 🖥️ Desktop Interface and Background Operation

- Includes a futuristic CustomTkinter desktop dashboard.<br>
- Displays commands, answers, actions, plans, OCR results, files, folders, memories, news, and errors in an activity console.<br>
- Provides quick buttons for voice control, standby mode, system status, OCR, screenshots, weather, news, and Windows startup.<br>
- Includes a text box for manually entering commands.<br>
- Runs long operations in worker threads so the interface remains responsive.<br>
- Continues running from the Windows system tray when the main window is closed.<br>
- Provides tray controls to show Jarvis, turn voice ON or OFF, change wake mode, or completely quit.<br>
- Can automatically start silently when the user signs in to Windows.<br>
- Runs independently after VS Code or the launching terminal is closed.<br>
- Prevents multiple Jarvis instances from competing for the microphone.<br>
- Can automatically switch to another installed Python runtime when VS Code selects one without the required packages.<br>
- Safely displays English and Hindi diagnostics in VS Code, PowerShell, and background mode.<br>

<br>

## 🔒 Privacy and Cost

- Uses a locally installed Ollama model for AI conversations and desktop planning.<br>
- Does not require paid OpenAI API credits.<br>
- Stores notes, screenshots, chat history, and memories locally on the computer.<br>
- Uses whitelisted automation actions instead of running arbitrary commands created by the AI.<br>
- Basic successful actions can be performed silently without unnecessary voice confirmations.<br>

<br>

## 💬 Example Commands

`Hey Jarvis`<br>
`Open Notepad and type Hello World`<br>
`Switch to Chrome`<br>
`Close this window`<br>
`Open my Python course folder`<br>
`Open a new tab and search for Python tutorials`<br>
`Switch to the YouTube tab`<br>
`Close this tab`<br>
`Increase the volume`<br>
`Set brightness to 60 percent`<br>
`Copy the selected text`<br>
`Copy text from this word to that word`<br>
`Read the text from this image`<br>
`Click the Settings button`<br>
`Show the desktop`<br>
`Remember that my project name is JARVIS`<br>
`What do you remember?`<br>
`Take a screenshot`<br>
`Tell me the system status`<br>
`नोटपैड खोलो`<br>
`ब्राइटनेस कम करो`<br>
`यूट्यूब टैब बंद करो`<br>

<br>

## ⚠️ Requirements and Limitations

- Designed specifically for Windows 11.<br>
- Ollama and the configured local language model must be installed and running for AI answers and natural-language planning.<br>
- Vision-based icon control requires the configured local multimodal Ollama model.<br>
- Tesseract OCR must be installed for screen and image text recognition.<br>
- Hindi OCR requires the Hindi Tesseract language file.<br>
- Speech recognition, Hindi neural speech, weather, and news features may require an internet connection.<br>
- A reminder works only while Jarvis remains running.<br>
- Screen automation depends on what is currently visible and may require confirmation for sensitive actions.<br>
