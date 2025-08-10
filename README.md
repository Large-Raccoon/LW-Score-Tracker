<a href="https://buymeacoffee.com/alargeraccoon" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="41" width="174"></a>

If you find this tool helpful, please consider making a donation by buying me a coffee!

# Last War Score Tracker
Gets scoreboard data from screenshots and writes to text (csv) for use with Excel or Google Sheets. Supports VS (Alliance Duel), TD (Alliance Tech Donations), 
and KS (Kill Score). Screenshots may be taken automatically, manually (guided), or imported.

## Prerequisites
1. Microsoft Windows
    - Development and testing was all performed on Windows 11. Windows 10 should work as well.
    - Minimum Version: Windows 10

2. Powershell
    - The bulk of this tool uses Powershell where version 5.1 is installed by default on Windws 10 and 11. However, Powershell 7 provides a superior experience and can be acquired by installing via winget (see below for command)
    - Source: Install from terminal (open command prompt or PowerShell from the Start Menu)
    - Minimum Version: 5.1 (default in Windows 10 and 11)
    - Recommended Version: 7.5.2 (or greater)
    - Install Command:
      ```
      winget install --id Microsoft.PowerShell --source winget
      ```

3. Python
    - A small portion of this tool requires Python in order to leverage OpenCV (cv2). This is crucial for getting coordinates of image matches.
    - Source: https://www.python.org/downloads/
    - Minimum Version: 3.6.0

4. OpenCV (cv2)
    - Image match recognition tool that will be used by Python for getting coordinates of image matches.
    - Source: Install from terminal (open command prompt or PowerShell from the Start Menu)
    - Minimum Version: 4.12.0.88
    - Install Command:
      ```
      pip install opencv-contrib-python
      ```

5. ImageMagick
    - A Powershell/Windows friendly image processing tool. It is required for the preparation of images which will be read by the OCR.
    - Source: https://imagemagick.org/script/download.php
    - Minimum Version: 7.1.1-47 Q16-HDRI (64-bit)

6. Android Debug Bridge (adb)
    - Copy your ADB files into C:\ADB, C:\Program Files\ADB, or add to PATH if saving elsewhere
    - Source: https://developer.android.com/tools/releases/platform-tools
    - Minimum Version: 36.0.0

7. Tesseract OCR
    - Reads text from images and converts to actual text format for processing. Required for reading player names and scores from screenshots.
    - Source: https://github.com/UB-Mannheim/tesseract/wiki
    - Minimum Version: 5.0.4.2

8. BlueStacks or physical Android device with USB debugging enabled
    - This is only required for automatic collection of screenshots. If you will be supplying screenshots taken manually, you may skip this.
    - Ideal BlueStacks display settings: Portrait orientation at 1080x1920 resolution with 240 DPI pixel density
    - Enabling USB debugging in BlueStacks: https://support.bluestacks.com/hc/en-us/articles/23925869130381-How-to-enable-Android-Debug-Bridge-on-BlueStacks-5
    - Enabling USB debugging on physical Android device: https://developer.android.com/studio/debug/dev-options
    - Take note of the IP and port that will be used.
        - Example: 127.0.0.1:5555

## Initial Setup
Some work will be needed to get everything setup initially. After finishing setup, running the score tracker is quick and easy.

1. Begin by downloading this repository:
    - "<> Code" > "Download ZIP"
    - Extract the downloaded Zip file anywhere convenient to remember.

2. Download, install, and configure any of the missing prerequisites listed in above "Prerequisites" section.

3. Launch PowerShell (can be found in the Start Menu) then navigate to the path you extracted this repository into.
    - Example:
    ```
    cd "C:\Last War\LW-Score-Tracker-main"
    ```

4. Set your PowerShell execution policy to "Bypass" so it's easier to run scripts.
    - Command:
    ```
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass
    ```

5. Populate your roster.csv with player names.
    - Navigate to the config folder and open roster.csv with something like Excel or with Google Sheets
    - In the first column under "Player", add all the player names from your alliance.
    - TIP: Copying and pasting from the game's PC client or BlueStacks makes this easier.
    - Save your updated roster.csv (make sure to keep it in the same config folder)

6. config.json
    - From the same config folder in step 5, you will find "config.json". Open this with notepad or any plaintext editor of your choice.
        - Set your score requirements:
          ```
          "Alliance": {                    
            "RosterName": "roster.csv", <-- The name of your roster file in config
            "VS_DailyMin": "7200000",   <-- Daily VS score requirements. If no daily VS requirements, set to "0".
            "VS_WeeklyMin": "0",        <-- Weekly VS score requirements. Leave at "0" if alliance has daily VS requirements.
            "TD_WeeklyMin": "20000"     <-- Weekly alliance tech donation requirements.
          },
          ```

      - Specify your languages:
        ```
        "ImageProcessing": {
            "Languages": "eng+kor"  <-- Default languages are English and Korean
          },
        ```
          The OCR will read your player names faster and more accurately if you limit the amount of languages it looks for. Alternatively, you may need to add additional languages. Some language options are as follows:
           * ara = Arabic
           * chi_sim = Chinese (simplified)
           * eng = English
           * jpn = Japanese
           * kor = Korean
           * rus = Russian\Slavic
           * tha = Thai
          - You can download more from: https://github.com/tesseract-ocr/tessdata_best. Be sure to add any traineddata you download to your tessdata folder.

      - If using auto mode: Configure your ADB settings:
        ```
        "ADB": {
            "Enabled": 1,                    <-- If using PC client to manually capture screenshots, set this to 0.
            "Emulator": "Bluestacks",        <-- Leave at default for BlueStacks. Set to "Emulator": "" if connecting to phone/tablet.
            "EmulatorExe": "HD-Player.exe",  <-- Executable to launch emulator in auto mode.
            "Device": "127.0.0.1:5555",      <-- IP and port for ADB to connect to. Should work with default Bluestacks settings.
            "Package": "com.fun.lastwar.gp", <-- Android package name to lauunch.
            "VsSwipeDistance": "457",        <-- Define how far to scroll VS daily scoreboard in auto mode
            "VsWklySwipeDistance": "443",    <-- Define how far to scroll VS weekly scoreboard in auto mode
            "TdSwipeDistance": "450"         <-- Define how far to scroll tech donation and kill score scoreboards in auto mode
          },
        ```

        - If using manual mode with the PC client:
          ```
          "PC": {
            "Enabled": 0,               <-- If using PC client to manually capture screenshots, set this to 1.
            "ProcessName": "LastWar",   <-- Process name of game as Windows sees it.
            "LaunchExe": "Launch.exe",  <-- Name of executable to launch game.
            "WindowWidth": "720",       <-- Enforced width of the PC client game window.
            "WindowHeight": "1280"      <-- Enforced height of the PC client game window.
          },
          ```
    
        - If you want the report sent to Google Sheets, configure this section.
          ```
          "GoogleSheets": {
            "Enabled": 0,         <-- If exporting reports to Google Sheets, set this to 1.
            "Account": "",        <-- Service account email needed ONLY if using a .p12 certificate. .p12 required for PowerShell 5.1 users.
            "CertFile": "",       <-- Add the filename of your .json or .p12 certificate here
            "SpreadsheetID": "",  <-- Add your spreadsheet ID here. ID is between https://docs.google.com/spreadsheets/d/ and the /edit?gid= part of your URL.
            "SheetName": "",      <-- Add your sheet name here
            "VS_Day1Cell": "A2",  <-- Set where VS Day 1 report should start. Always 4 columns wide.
            "VS_Day2Cell": "E2",  <-- Set where VS Day 2 report should start. Always 4 columns wide.
            "VS_Day3Cell": "I2",  <-- Set where VS Day 3 report should start. Always 4 columns wide.
            "VS_Day4Cell": "M2",  <-- Set where VS Day 4 report should start. Always 4 columns wide.
            "VS_Day5Cell": "Q2",  <-- Set where VS Day 5 report should start. Always 4 columns wide.
            "VS_Day6Cell": "U2",  <-- Set where VS Day 6 report should start. Always 4 columns wide.
            "VS_WeeklyCell": "",  <-- Set where VS weekly report should start. Always 4 columns wide.
            "TD_Cell": "Y2",      <-- Set where Tech Donation report should start. Always 4 columns wide.
            "KS_Cell": "AC2"      <-- Set where Kill Score report should start. Always 4 columns wide.
          },
          ```

7. Note how many ranks with players can fit in your scoreboard UNOBSTRUCTED at any given time. This will be your PPS used in the next step.
    - For BlueStacks at the recommended 1080x1920 resolution, that should be 5 players

8. If you're NOT using the automatic screenshot feature, add your screenshots of a scoreboard one at a time to the import folder. Skip if automatically capturing screenshots.
    - In the import folder, there are 3 folders you can add your screenshots to:
        - KS: For kill score ranking screenshots
        - TD: For weekly alliance tech donation ranking screenshots.
        - VS: For VS ranking screenshots.
    - Note that the screenshots must be added in the correct order one at a time. Screenshot of ranks 1-5 would be added first, ranks 6-10 would be added second, etc.

8. Run the Score Tracker script
- In PowerShell, make sure you're still in the correct folder where you extracted this repository to.
    - Example:
    ```
    cd "C:\Last War\LW-Score-Tracker-main"
    ```
- Now run the script with ONE of these example commands. Remember to update your -PPS number if different:
    - Get VS Day 1 scores with 5 players per screenshot:
    ```
    .\LW-ScoreTracker.ps1 -Type VS -Day 1 -PPS 5 -Mode Auto
    ```
    - Get VS weekly score with 5 players per screenshot:
    ```
    .\LW-ScoreTracker.ps1 -Type VS -PPS 5 -Mode Auto
    ```
    - Get TD (Tech Donation) score with 5 players per screenshot:
    ```
    .\LW-ScoreTracker.ps1 -Type TD -PPS 5 -Mode Auto
    ```
    - Get TD (Tech Donation) scores with screenshots you will manually take through the script.
    ```
    .\LW-ScoreTracker.ps1 -Type TD -PPS 5 -Mode Manual
    ```
    - Get TD (Tech Donation) scores with imported screenshots.
    ```
    .\LW-ScoreTracker.ps1 -Type TD -PPS 5 -Mode Import
    ```
  
9. Confirm that player names are being reported correctly.
    - Open the "Logs" folder from wherever the Score Tracker is installed. Then open up "NameCorrection.log".
    - For ⚠️ entries, these were guesses made by the script. If it did not guess correctly, add the invalid name as an alias to your roster.csv.
    - For ❌ entries, the player was completely unrecognized. Make sure the player is in your roster.csv. Add the invalid name as an alias to your roster.csv if necessary.

## Usage
- Parameters
    - -Type
        - Determines which scoreboard will be captured.
        - -Type VS = Alliance Duel
        - -Type TD = Tech Donations
        - -Type KS = Kill Score
     
    - -Day
        - Specifies which days of VS scores are to be captured.
        - Required when processing VS scores with daily score requirements.
        - Not required for weekly score requirements.
     
    - -PPS
        - Specifies how many players are clearly depicted (with rank number fully intact) in each scoreboard screenshot.
        - Required when automatically collecting screenshots.
        - Will attempt to automatically set when `-Mode Import` is used.
        - For best results, always specify your PPS.
     
    - -Mode
        - Determines how screenshots will be captured.
        - -Type Auto   = Automatically capture screenshots
        - -Type Manual = Manually capture screenshots through a guided process
        - -Type Import = Import screenshots you've already captured (will attempt to use existing screenshots if no images added to import folder)

- Examples
    - Get VS scores from days 1-4 with automatic screenshot capture:
    ```
    .\LW-ScoreTracker.ps1 -Type VS -Day 1,2,3,4 -PPS 5 -Mode Auto
    ```
    
    - Get VS scores from days 1-4 and take screenshots manually:
    ```
    .\LW-ScoreTracker.ps1 -Type VS -Day 1,2,3,4 -PPS 5 -Mode Manual
    ```
   
    - Get VS scores from days 1-4 with imported screenshots:
    ```
    .\LW-ScoreTracker.ps1 -Type VS -Day 1,2,3,4 -PPS 5 -Mode Manual
    ```
    
    - Get VS scores when your alliance has weekly score requirements with automatic screenshot capture:
    ```
    .\LW-ScoreTracker.ps1 -Type VS -PPS 5 -Mode Manual
    ```
    
    - Get weekly tech donation scores with automatic screenshot capture:
    ```
    .\LW-ScoreTracker.ps1 -Type TD -PPS 5 -Mode Auto
    ```
    
    - Get kill scores with automatic screenshot capture:
    ```
    .\LW-ScoreTracker.ps1 -Type KS -PPS 5 -Mode Auto
    ```
    
## Troubleshooting
- Scoreboard is scrolling too far or not far enough
    - Open up your config.json and modify the following settings as needed:
      ```
      "VsSwipeDistance": "457",      <-- VS daily scrolling. Lower to scroll further down. Raise to scroll further up.
      "VsWklySwipeDistance": "443",  <-- VS weekly scrolling. Lower to scroll further down. Raise to scroll further up.
      "TdSwipeDistance": "450"       <-- TD and KS scrolling. Lower to scroll further down. Raise to scroll further up.
      ```
