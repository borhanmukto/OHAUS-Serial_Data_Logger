⚖️ Ohaus Scale Serial Data Logger
=================================

This application logs data from a serial device (specifically configured for Ohaus scales) and saves it to an Excel file in real-time. It features a live data dashboard, batch buffering for high-speed data, and automatic sheet switching in Excel to handle large datasets.

📋 Prerequisites
----------------

Before running this application, ensure you have the following installed on your PC:

1.  **Python (version 3.8 or higher)**
    
    *   Download from [python.org](https://www.python.org/downloads/).
        
    *   _Note: During installation, it is highly recommended to check the box "Add Python to PATH"._
        
2.  **A Serial Device**
    
    *   Connect your scale/device via USB or RS232-to-USB adapter.
        

⚙️ Installation
---------------

1.  **Create the Project Folder**
    
    *   Create a new folder on your computer (e.g., SerialLogger).
        
    *   Save the Python script provided in this package inside that folder as **app.py**.
        
2.  Bashpip install streamlit pyserial pandas openpyxl⚠️ If the command above fails (Command not found)If you see an error saying 'pip' is not recognized, it means Python was not added to your system PATH. Use this command instead:Bashpy -m pip install streamlit pyserial pandas openpyxl_(Note: We use pyserial, not serial. If you mistakenly installed serial, uninstall it first.)_
    
    *   Open your Command Prompt (Windows) or Terminal (Mac/Linux).
        
    *   Navigate to your folder (e.g., cd Desktop/SerialLogger).
        
    *   Run the following command:
        

🚀 How to Run
-------------

1.  Open your Command Prompt/Terminal.
    
2.  Navigate to the folder where you saved app.py.
    
3.  Bashstreamlit run app.py⚠️ If the command above failsIf your computer does not recognize streamlit, use the Python launcher directly:Bashpy -m streamlit run app.py
    
4.  A new tab should automatically open in your default web browser displaying the application.
    

📖 User Guide
-------------

### 1\. Configuration (Sidebar)

*   **Refresh Ports:** Click this if you plugged in your device _after_ starting the app.
    
*   **Serial Port:** Select the COM port your device is connected to (e.g., COM3, /dev/ttyUSB0).
    
*   **Baud Rate:** Match this to your device's settings (Default is 9600).
    
*   **Output Filename:** Name of the Excel file to be generated.
    
*   **Batch Size:** Determines how many rows are held in memory before writing to Excel. Higher numbers (e.g., 2000) offer better performance for fast data streams.
    

### 2\. Operation

1.  Click **"🚀 Start Logging"**.
    
2.  The status will change to "Active" and the app will attempt to send a CP (Continuous Print) command to the scale.
    
3.  Watch the **Live Data Preview** to ensure data is coming in correctly.
    
4.  Click **"🛑 Stop Logging"** to safely finish the session. The app will ensure any remaining data in the buffer is written to the file before closing.
    

⚠️ Troubleshooting
------------------

**"Permission Denied" Error:**

*   This usually happens if you have the Desorption\_Data.xlsx file **open in Excel** while the app is trying to write to it. Close the Excel file and try again.
    

**"No COM ports detected":**

*   Ensure your USB cable is plugged in securely.
    
*   You may need to install drivers for your specific RS232-to-USB adapter (common brands are FTDI or Prolific).
    

**Data is lagging:**

*   Increase the **Batch Size** in the sidebar settings. Writing to Excel is slow; doing it less frequently improves performance.
    

### Developed with Python & Streamlit
