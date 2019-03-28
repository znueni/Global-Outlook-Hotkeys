# Global Outlook Hotkeys

Its an application, that, when running (in background) provides hotkeys for important Microsoft Outlook features.
These hotkeys work with the Windows-Key on your keyboard and work globally across all applications in Windows.

For example, just press <kbd>Windows</kbd>+<kbd>NUM 0</kbd> in whatever application you are right now, to start a new mail in Outlook.

The keys only work, if Outlook is running and if the Global Outlook Hotkeys application is running.
So make sure you autostart the application when starting Windows!

# Hotkeys
All hotkeys are combinations of keys together with your <kbd>Windows</kbd> key. It is usually on your keyboard left of the <kbd>SPACE</kbd> bar and has the Microsoft Windows logo on it.

* Switch to the main Outlook window:
  * <kbd>Windows</kbd>+<kbd>F12</kbd>
* Switch to the Outlook calendar window (if there is a separate calendar window open):
   * <kbd>Windows</kbd>+<kbd>F11</kbd>
* Compose a new Mail:
   * <kbd>Windows</kbd>+<kbd>NUMPAD 0</kbd>
   * <kbd>Ctrl</kbd>+<kbd>Windows</kbd>+<kbd>SPACE</kbd>   
   * <kbd>Ctrl</kbd>+<kbd>Windows</kbd>+<kbd>F12</kbd>
* Compose a new Appointment/Meeting:
   * <kbd>Windows</kbd>+<kbd>NUMPAD 1</kbd>
   * <kbd>Shift</kbd>+<kbd>Windows</kbd>+<kbd>SPACE</kbd>   
   * <kbd>Shift</kbd>+<kbd>Windows</kbd>+<kbd>F12</kbd>
* Close the Global Outlook Hotkeys application:
  * <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Alt</kbd>+<kbd>Windows</kbd>+<kbd>F12</kbd>  

# Download and Run
You can just download the files in the BIN folder and run GlobalOutlookHotKeys.exe

If you start it, nothing happens. The App has no visible window. It just stays in memory to provide and handle the hotkeys.
So make sure you autostart the application when starting Windows to have the hotkeys working for you all the time.

[Download](bin/)

# Dependencies
The application obviously requires Microsoft Outlook. As it does not use API but rather command line arguments of Outlook.exe, it is not tight to any version and should work with all Outlook versions.
You can pretty easily recode it to support any other mail application.
All code is VB.NET, core version does not matter.
The hotkeys are hardcoded in the code. Please see the comments in the [Globalhotkeys.vb](Globalhotkeys.vb)
