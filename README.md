Offline Speech-to-Text (Vosk-Based)
Lightweight, offline, and uncensored speech-to-text tool using Vosk models. Designed for simplicity and flexibility.

Features
Vosk-based speech-to-text (completely offline)
Includes two mini-models (English & Russian)
Drag & drop model support – download any Vosk model and drop it into the /models folder (recognized automatically, no restart needed)
Three text delivery modes:
Interface Mode – Displays transcribed text in the main window
Cursor Mode – Sends transcribed text to where the cursor is active
Window Mode – Sends text directly to a selected application window
Keyword activation (English only for now) – Allows triggering actions based on detected words
Live Mode (Experimental) – Continuous transcription with a slight delay (needs further tuning)
Installation
Install Python 3.8 (included in the package)
Run install.bat – this will create the environment inside the same working folder
A shortcut icon will be placed on your desktop – run it to start the program
Adding New Models
Download additional models from the Vosk website
Extract the model folder (avoid double-folder structure)
Place it inside the /models folder in the application directory
The new model will be recognized automatically, no restart required
Uninstalling
Since the entire program is self-contained, uninstallation is not really necessary.
If needed, running uninstall.bat will remove the environment, but you can just delete the folder manually if you prefer.
Notes
Keyword activation works well, but still basic (currently just pressing "Enter" when detected)
Live Mode is experimental – works but requires fine-tuning for better accuracy
No cloud processing – everything runs offline for privacy and full control
