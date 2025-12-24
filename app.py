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

    /* 2. Center text inside all alert boxes */
    .stAlert > div {
        justify-content: center;
        text-align: center;
    }

    /* 3. Custom Card Style for Metrics */
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

# We track row count to display progress
if 'row_count' not in st.session_state:
    st.session_state['row_count'] = 0

# --- Helper Functions ---
def get_available_ports():
    """Scans and returns a list of available COM ports."""
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]

def make_card_html(label, value, sub_text=""):
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub_text}</div>
    </div>
    """

def save_to_csv_append(filename, data):
    """
    Appends data to CSV. 
    Writes header only if file does not exist.
    """
    if not data:
        return False
    
    try:
        df_batch = pd.DataFrame(data)
        
        # Check if file exists to determine if we need a header
        file_exists = os.path.isfile(filename)
        
        # Append mode 'a', no index, header only if new file
        df_batch.to_csv(filename, mode='a', index=False, header=not file_exists)
        
        return True
    except Exception as e:
        st.error(f"Failed to save to CSV: {e}")
        return False

# --- Sidebar Configuration ---
st.sidebar.header("⚙️ Configuration")

if st.sidebar.button("Refresh Ports"):
    st.rerun()

available_ports = get_available_ports()
if available_ports:
    serial_port = st.sidebar.selectbox("Serial Port", available_ports, index=0)
else:
    st.sidebar.warning("No COM ports detected.")
    serial_port = st.sidebar.text_input("Manually Enter Port", "COM5")

baud_rate = st.sidebar.number_input("Baud Rate", value=9600, step=100)

# UPDATED: Default to CSV
output_filename = st.sidebar.text_input("Output Filename", "Desorption_Data.csv")

st.sidebar.subheader("Performance Settings")
batch_size = st.sidebar.number_input("Batch Size (Rows to buffer)", value=500, min_value=100, step=100, help="Number of readings to collect before writing to CSV.")

# --- Main UI ---
st.title("⚖️ Ohaus Scale Data Logger")

st.markdown("""
    <div style='text-align: center; margin-bottom: 20px;'>
        Developed by Borhan Uddin Rabbani || <a href='https://www.linkedin.com/in/borhan-uddin-rabbani/' target='_blank'>Connect on LinkedIn</a>
    </div>
    """, unsafe_allow_html=True)

st.write("") 

_, btn_col1, btn_col2, _ = st.columns([3, 2, 2, 3])
with btn_col1:
    start_btn = st.button("🚀 Start Logging", type="primary", use_container_width=True)
with btn_col2:
    stop_btn = st.button("🛑 Stop Logging", type="secondary", use_container_width=True)

st.write("")

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
    st.session_state['row_count'] = 0
    
    # UPDATED: Check existing CSV file to resume count
    if os.path.exists(output_filename):
        try:
            # Fast line count without loading the whole file into memory
            with open(output_filename, 'rb') as f:
                row_count = sum(1 for _ in f)
            # Subtract 1 for header if count > 0
            st.session_state['row_count'] = max(0, row_count - 1) if row_count > 0 else 0
        except Exception:
            pass # Start fresh if error reading

if st.session_state['logging_active']:
    
    status_placeholder.markdown(make_card_html("Status", "Connecting...", "Initializing"), unsafe_allow_html=True)
    
    try:
        ser = serial.Serial(
            port=serial_port,
            baudrate=baud_rate,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout=0.1 # Short timeout for non-blocking feel
        )
        
        # Send Command
        time.sleep(1) 
        ser.write("CP\r\n".encode('utf-8'))
        
        status_placeholder.markdown(make_card_html("Status", "Active", "Connected"), unsafe_allow_html=True)
        st.toast(f"Connected to {serial_port}")

        buffer = []
        last_ui_update_time = time.time()
        
        while True:
            if not st.session_state['logging_active']:
                break

            try:
                # 1. READ ALL AVAILABLE DATA
                if ser.in_waiting > 0:
                    raw_lines = ser.readlines() 
                    
                    for raw_line in raw_lines:
                        try:
                            decoded_line = raw_line.decode('utf-8', errors='ignore')
                            clean_line = re.sub(r'[\x00-\x1F\x7F]', '', decoded_line).strip()
                            
                            if clean_line:
                                buffer.append({
                                    'Timestamp': datetime.now(), 
                                    'Response': clean_line
                                })
                        except:
                            continue

                    # 2. THROTTLE UI UPDATES
                    current_time = time.time()
                    if current_time - last_ui_update_time > 0.5:
                        
                        current_total = st.session_state['row_count'] + len(buffer)
                        
                        last_reading = buffer[-1]['Response'] if buffer else "-"
                        
                        # Update Cards
                        last_val_placeholder.markdown(make_card_html("Last Reading", last_reading, "Raw Value"), unsafe_allow_html=True)
                        count_placeholder.markdown(make_card_html("Rows Logged", str(current_total), "Session Total"), unsafe_allow_html=True)
                        
                        # Update Table
                        if buffer:
                            preview_df = pd.DataFrame(buffer[-10:])
                            preview_df['Timestamp'] = preview_df['Timestamp'].dt.strftime('%H:%M:%S.%f').str[:-3]
                            table_placeholder.dataframe(preview_df, use_container_width=True)
                        
                        last_ui_update_time = current_time

                    # 3. SAVE BATCH (CSV Logic)
                    if len(buffer) >= batch_size:
                        status_placeholder.markdown(make_card_html("Status", "Saving...", "Writing to CSV"), unsafe_allow_html=True)
                        
                        success = save_to_csv_append(output_filename, buffer)
                        
                        if success:
                            st.session_state['row_count'] += len(buffer)
                            buffer = [] # Clear buffer
                            status_placeholder.markdown(make_card_html("Status", "Active", "Collecting Data"), unsafe_allow_html=True)
                        else:
                            st.error("Failed to save data. Stopping to prevent data loss.")
                            st.session_state['logging_active'] = False
                            break
                else:
                    time.sleep(0.01)

            except serial.SerialException:
                st.error("Device disconnected abruptly.")
                st.session_state['logging_active'] = False
                break
                
    except Exception as e:
        st.error(f"Connection Error: {e}")
        st.session_state['logging_active'] = False
        
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            
        # Save remaining data
        if 'buffer' in locals() and buffer:
            st.info(f"Saving {len(buffer)} remaining rows...")
            success = save_to_csv_append(output_filename, buffer)
            if success:
                st.session_state['row_count'] += len(buffer)
            
        status_placeholder.markdown(make_card_html("Status", "Stopped", "Connection Closed"), unsafe_allow_html=True)

else:
    status_placeholder.markdown(make_card_html("Status", "Idle", "Ready to Start"), unsafe_allow_html=True)
    count_placeholder.markdown(make_card_html("Rows Logged", str(st.session_state['row_count']), "Session Total"), unsafe_allow_html=True)
    last_val_placeholder.markdown(make_card_html("Last Reading", "-", "-"), unsafe_allow_html=True)
    
    st.info("Configure settings in the sidebar and press Start.")
