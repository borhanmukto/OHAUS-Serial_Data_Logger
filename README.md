⚖️ High-Speed Serial Data LoggerA robust, high-performance Streamlit application designed to log data from serial devices (like Ohaus scales) directly to Excel. Optimized for high-frequency data streams (up to 500Hz), it features real-time visualization, automatic port detection, and intelligent memory management.🚀 Key Features⚡ High-Speed Performance: Optimized buffering system capable of handling 500+ readings per second without UI lag or freezing.📊 Real-Time Dashboard: Live data preview and session statistics updated efficiently (throttled UI refresh rate).🔌 Smart Connectivity: automatically detects available COM ports.💾 Intelligent Excel Logging:Writes data in batches to prevent disk I/O bottlenecks.Automatically creates new sheets when the row limit (default 800k) is reached.Resume capability: Appends to existing files without overwriting data.🎨 User-Friendly Interface: Clean, centered, and responsive design powered by Streamlit.🛠️ InstallationClone the repositorygit clone [https://github.com/yourusername/serial-data-logger.git](https://github.com/yourusername/serial-data-logger.git)
cd serial-data-logger
Create a Virtual Environment (Recommended)# Windows
python -m venv venv
venv\Scripts\activate

# Mac/Linux
python3 -m venv venv
source venv/bin/activate
Install Dependenciespip install -r requirements.txt
🖥️ UsageConnect your device (e.g., Scale, Arduino) via USB/Serial to your computer.Run the application:streamlit run app.py
Configure & Start:Select your Serial Port from the sidebar dropdown.Set the Baud Rate (default 9600).Click 🚀 Start Logging.⚙️ ConfigurationSettingDefaultDescriptionBaud Rate9600Communication speed. Must match your device settings.Batch Size2000Number of rows to buffer in memory before writing to disk. Higher values = better performance for fast data.Max Rows800,000Maximum rows per Excel sheet before creating a new sheet.⚠️ Important Notes on High-Speed DataThrottling: The UI updates every 0.5 seconds regardless of data speed to keep the computer responsive. All data is captured in the background.Batching: Do not lower the Batch Size below 500 if you are logging high-speed data (500Hz), or the disk writing process may slow down the logging loop.🤝 ContactBorhan Uddin RabbaniBuilt with ❤️ using Streamlit
