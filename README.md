\#⚖️ High-Speed Serial Data Logger



!\[Python](https://img.shields.io/badge/Python-3.8%2B-blue)

!\[Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B)



A robust, high-performance Streamlit application designed to log data from serial devices (like Ohaus scales, Arduino, etc.) directly to Excel. Optimized for \*\*high-frequency data streams (up to 500Hz)\*\*, it features real-time visualization, automatic port detection, and intelligent memory management.



\## 🚀 Key Features



\* \*\*⚡ High-Speed Performance:\*\* Optimized buffering system capable of handling \*\*500+ readings per second\*\* without UI lag or freezing.

\* \*\*📊 Real-Time Dashboard:\*\* Live data preview and session statistics updated efficiently via throttled UI refresh.

\* \*\*🔌 Smart Connectivity:\*\* Automatically detects and lists available COM ports.

\* \*\*💾 Intelligent Excel Logging:\*\*

&nbsp;   \* \*\*Batch Writing:\*\* Writes data in chunks to prevent disk I/O bottlenecks.

&nbsp;   \* \*\*Auto-Pagination:\*\* Automatically creates new Excel sheets when the row limit (default 800k) is reached.

&nbsp;   \* \*\*Resume Capability:\*\* Appends to existing files safely without overwriting data.

\* \*\*🎨 User-Friendly Interface:\*\* Clean, centered, and responsive design powered by Streamlit.



\## 🛠️ Installation



\*\*1. Clone the repository\*\*

```bash

git clone \[https://github.com/yourusername/serial-data-logger.git](https://github.com/yourusername/serial-data-logger.git)

cd serial-data-logger



\*\*2. Create a Virtual Environment (Recommended)\*\*



\# Windows

python -m venv venv

venv\\Scripts\\activate



\# Mac/Linux

python3 -m venv venv

source venv/bin/activate


\*\*3. Install Dependencies\*\*


pip install -r requirements.txt



\## 🖥️ Usage



1\. Connect your device (e.g., Scale, Arduino) via USB/Serial to your computer.



2\. Run the application:



streamlit run app.py



3\. Configure \& Start:



Select your Serial Port from the sidebar dropdown.



Set the Baud Rate (default 9600).



Click 🚀 Start Logging.

