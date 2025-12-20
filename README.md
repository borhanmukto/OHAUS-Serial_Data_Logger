⚖️ Ohaus Scale Serial Data Logger
=================================

This application logs data from a serial device (specifically configured for Ohaus scales) and saves it to an Excel file in real-time. It is designed for reliability and ease of use, featuring:

*   **📊 Live Data Dashboard:** View incoming data streams instantly.
    
*   **⚡ Batch Buffering:** Optimized for high-speed data to prevent lag.
    
*   **📑 Auto-Sheet Switching:** Automatically manages large datasets within Excel.
    

📋 Prerequisites
----------------

Before running this application, ensure your setup meets the following requirements:

### 1\. Software

*   **⚠️ Important:** During installation, ensure you check the box labeled **"Add Python to PATH"**.
    

### 2\. Hardware

*   **Serial Device:** Your Ohaus scale.
    
*   **Connection:** A USB cable or an RS232-to-USB adapter.
    

⚙️ Installation
---------------

Follow these steps to set up the environment:

### 1\. Create the Project Folder

1.  Create a new folder on your computer (e.g., SerialLogger).
    
2.  Save the provided Python script inside this folder and name it **app.py**.
    

### 2\. Install Required Libraries

Open your Command Prompt (Windows) or Terminal (Mac/Linux), navigate to your folder, and run the following command:

Bash

`   pip install streamlit pyserial pandas openpyxl   `

> **❗ Critical Note:** We use pyserial, **not** serial. If you have previously installed the package named serial, you must uninstall it first using pip uninstall serial to avoid conflicts.

🚀 How to Run
-------------

1.  Open your Command Prompt or Terminal.
    
2.  Navigate to the folder containing app.py.
    
3.  Bashstreamlit run app.py
    
4.  A new tab will automatically open in your default web browser displaying the dashboard.
    

📖 User Guide
-------------

### 1\. Configuration (Sidebar)

Adjust these settings before starting your session:

*   **🔄 Refresh Ports:** Click this if you plugged in your device _after_ the app was already running.
    
*   **🔌 Serial Port:** Select the specific COM port your device is connected to (e.g., COM3, /dev/ttyUSB0).
    
*   **📡 Baud Rate:** Match this to your device's internal settings (Default is 9600).
    
*   **📁 Output Filename:** Enter the desired name for the generated Excel file.
    
*   **📦 Batch Size:** Controls how many rows are held in memory before saving to the file.
    
    *   _Tip:_ Higher numbers (e.g., 2000) improve performance for fast data streams.
        

### 2\. Operation

1.  Click **🚀 Start Logging**.
    
2.  The status indicator will change to "Active," and the app will send a CP (Continuous Print) command to the scale.
    
3.  Monitor the **Live Data Preview** to verify data integrity.
    
4.  Click **🛑 Stop Logging** to end the session.
    
    *   _Note:_ The app will automatically flush any remaining buffered data to the Excel file before closing.
        

⚠️ Troubleshooting
------------------

**IssueSolution"Permission Denied" Error**This occurs if the Excel file is currently **open**. Close the file and try again.**"No COM ports detected"**

1\. Ensure the USB cable is secure.

2\. Install drivers for your RS232 adapter (common brands: FTDI, Prolific).

**Data is lagging**Increase the **Batch Size** in the sidebar. Writing to Excel is resource-intensive; doing it less frequently improves speed.

_Developed with Python & Streamlit_
