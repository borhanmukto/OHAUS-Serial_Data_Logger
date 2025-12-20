import streamlit as st
import serial
import serial.tools.list_ports
import pandas as pd
import time
import os
import re
from datetime import datetime

# --- Page Config ---
st.set_page_config(
    page_title="Serial Data Logger",
    page_icon="⚖️",
    layout="wide"
)

# --- Custom CSS for Centering ---
st.markdown("""
    <style>
    /* 1. Center all standard headings */
    h1, h2, h3 {
        text-align: center !important;
    }

    /* 2. Center text inside all alert boxes (Info, Warning, Success, Error) */
    .stAlert > div {
        justify-content: center;
        text-align: center;
    }

    /* 3. Custom Card Style for Metrics (Replaces st.metric) */
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 10px;
        padding: 20px 10px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .metric-label {
        font-size: 16px;
        color: #6c757d;
        margin-bottom: 5px;
        font-weight: 500;
    }
    .metric-value {
        font-size: 28px;
        font-weight: bold;
        color: #212529;
        margin: 0;
    }
    .metric-sub {
        font-size: 14px;
        color: #495057;
        margin-top: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Session State Initialization ---
if 'logging_active' not in st.session_state:
    st.session_state['logging_active'] = False
if 'data_buffer' not in st.session_state:
    st.session_state['data_buffer'] = []
if 'total_rows' not in st.session_state:
    st.session_state['total_rows'] = 0

# --- Helper Functions ---
def get_available_ports():
    """Scans and returns a list of available COM ports."""
    ports = serial.tools.list_ports.comports()
    # Simply return the device name (e.g., 'COM3' or '/dev/ttyUSB0')
    return [port.device for port in ports]

def make_card_html(label, value, sub_text=""):
    """Generates HTML for a centered metric card."""
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub_text}</div>
    </div>
    """

def save_to_excel(filename, data, max_rows):
    """Handles appending data to Excel, including sheet management."""
    if not data:
        return
    
    df_batch = pd.DataFrame(data)
    
    # Determine Write Mode and Sheet
    sheet_index = 1
    start_row = 0
    write_header = True
    mode = 'w'
    if_sheet_exists = None

    if os.path.exists(filename):
        mode = 'a'
        if_sheet_exists = 'overlay'
        try:
            with pd.ExcelFile(filename) as xls:
                last_sheet_name = xls.sheet_names[-1]
                # Extract index from "SheetX"
                match = re.search(r'Sheet(\d+)', last_sheet_name)
                if match:
                    sheet_index = int(match.group(1))
                
                # Read last sheet to check rows
                df_last = pd.read_excel(xls, sheet_name=last_sheet_name)
                current_rows = len(df_last)
                
                if current_rows >= max_rows:
                    sheet_index += 1
                    start_row = 0
                    write_header = True
                else:
                    start_row = current_rows + 1 # +1 for header
                    write_header = False
        except Exception as e:
            st.error(f"Error reading existing Excel: {e}")
            return

    sheet_name = f"Sheet{sheet_index}"

    try:
        with pd.ExcelWriter(filename, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
            df_batch.to_excel(
                writer, 
                sheet_name=sheet_name, 
                index=False, 
                header=write_header, 
                startrow=start_row
            )
        return True
    except Exception as e:
        st.error(f"Failed to save to Excel: {e}")
        return False

# --- Sidebar Configuration ---
st.sidebar.header("⚙️ Configuration")

if st.sidebar.button("Refresh Ports"):
    st.rerun()

# Simplified Port Selection Logic
available_ports = get_available_ports()
serial_port = None

if available_ports:
    # Just show the list of ports found (e.g. ['COM3', 'COM4'])
    serial_port = st.sidebar.selectbox("Serial Port", available_ports, index=0)
else:
    # If no ports found, allow manual entry
    st.sidebar.warning("No COM ports detected automatically.")
    serial_port = st.sidebar.text_input("Manually Enter Port (e.g., COM5)", "COM5")

baud_rate = st.sidebar.number_input("Baud Rate", value=9600, step=100)
output_filename = st.sidebar.text_input("Output Filename", "Desorption_Data.xlsx")

st.sidebar.subheader("Advanced Settings")
batch_size = st.sidebar.number_input("Batch Size (Rows to buffer)", value=50, min_value=1)
max_rows_per_sheet = st.sidebar.number_input("Max Rows Per Sheet", value=800000)

# --- Main UI ---
st.title("⚖️ Ohaus Scale Data Logger")

# Updated: Normal text (div) instead of heading (h5)
st.markdown("""
    <div style='text-align: center; margin-bottom: 20px;'>
        Developed by Borhan Uddin Rabbani || <a href='https://www.linkedin.com/in/borhan-uddin-rabbani/' target='_blank'>Connect on LinkedIn</a>
    </div>
    """, unsafe_allow_html=True)

st.write("") 

# Centered Buttons using Columns
_, btn_col1, btn_col2, _ = st.columns([3, 2, 2, 3])
with btn_col1:
    start_btn = st.button("🚀 Start Logging", type="primary", use_container_width=True)
with btn_col2:
    stop_btn = st.button("🛑 Stop Logging", type="secondary", use_container_width=True)

st.write("")

# --- Custom Metrics Area (Replaced standard metrics with HTML Cards) ---
metrics_container = st.container()
with metrics_container:
    m_col1, m_col2, m_col3 = st.columns(3)
    status_placeholder = m_col1.empty()
    count_placeholder = m_col2.empty()
    last_val_placeholder = m_col3.empty()

st.subheader("Live Data Preview (Last 10)")
table_placeholder = st.empty()

# --- Logic Flow ---

if stop_btn:
    st.session_state['logging_active'] = False
    st.warning("Stop request sent. The loop will exit shortly.")

if start_btn:
    st.session_state['logging_active'] = True
    st.session_state['data_buffer'] = []
    st.session_state['total_rows'] = 0

if st.session_state['logging_active']:
    
    # Initial status display
    status_placeholder.markdown(make_card_html("Status", "Connecting...", "Initializing"), unsafe_allow_html=True)
    count_placeholder.markdown(make_card_html("Rows Logged", "0", "Waiting for data"), unsafe_allow_html=True)
    last_val_placeholder.markdown(make_card_html("Last Reading", "-", "--"), unsafe_allow_html=True)
    
    try:
        # 1. Establish Connection
        ser = serial.Serial(
            port=serial_port,
            baudrate=baud_rate,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout=1
        )
        
        # 2. Send Command
        time.sleep(1) 
        command = "CP\r\n"
        ser.write(command.encode('utf-8'))
        
        status_placeholder.markdown(make_card_html("Status", "Active", "Connected"), unsafe_allow_html=True)
        st.toast(f"Connected to {serial_port}")

        # 3. Main Loop
        buffer = []
        rows_saved = 0
        
        while True:
            try:
                if ser.in_waiting > 0:
                    raw_response = ser.readline().decode('utf-8', errors='ignore')
                    response = re.sub(r'[\x00-\x1F\x7F]', '', raw_response).strip()
                    
                    if response:
                        timestamp = datetime.now()
                        entry = {'Timestamp': timestamp, 'Response': response}
                        buffer.append(entry)
                        
                        # -- UI Updates (Frequent) --
                        current_total = rows_saved + len(buffer)
                        
                        # Update Custom HTML Cards
                        last_val_placeholder.markdown(make_card_html("Last Reading", response, "Raw Value"), unsafe_allow_html=True)
                        count_placeholder.markdown(make_card_html("Rows Logged", str(current_total), "Session Total"), unsafe_allow_html=True)
                        
                        # Update table every 5 readings
                        if len(buffer) % 5 == 0:
                            preview_df = pd.DataFrame(buffer[-10:])
                            preview_df['Timestamp'] = preview_df['Timestamp'].dt.strftime('%H:%M:%S')
                            table_placeholder.dataframe(preview_df, use_container_width=True)

                        # -- Batch Saving --
                        if len(buffer) >= batch_size:
                            status_placeholder.markdown(make_card_html("Status", "Saving...", "Writing to Disk"), unsafe_allow_html=True)
                            success = save_to_excel(output_filename, buffer, max_rows_per_sheet)
                            if success:
                                rows_saved += len(buffer)
                                buffer = [] 
                                status_placeholder.markdown(make_card_html("Status", "Active", "Collecting Data"), unsafe_allow_html=True)
                            else:
                                st.error("Failed to save data. Stopping.")
                                break
                
                else:
                    time.sleep(0.1)

            except serial.SerialException:
                st.error("Device disconnected abruptly.")
                break
                
    except Exception as e:
        st.error(f"Connection Error: {e}")
        
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            
        if 'buffer' in locals() and buffer:
            st.info(f"Saving {len(buffer)} remaining rows...")
            save_to_excel(output_filename, buffer, max_rows_per_sheet)
            
        status_placeholder.markdown(make_card_html("Status", "Stopped", "Connection Closed"), unsafe_allow_html=True)
        st.session_state['logging_active'] = False
        st.success("Logging session finished.")

else:
    # Idle State Display
    status_placeholder.markdown(make_card_html("Status", "Idle", "Ready to Start"), unsafe_allow_html=True)
    count_placeholder.markdown(make_card_html("Rows Logged", "0", "-"), unsafe_allow_html=True)
    last_val_placeholder.markdown(make_card_html("Last Reading", "-", "-"), unsafe_allow_html=True)
    
    st.info("Configure settings in the sidebar and press Start.")