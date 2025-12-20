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
    h1, h2, h3 { text-align: center !important; }
    .stAlert > div { justify-content: center; text-align: center; }
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 10px;
        padding: 20px 10px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .metric-label { font-size: 16px; color: #6c757d; margin-bottom: 5px; font-weight: 500; }
    .metric-value { font-size: 28px; font-weight: bold; color: #212529; margin: 0; }
    .metric-sub { font-size: 14px; color: #495057; margin-top: 5px; }
    </style>
    """, unsafe_allow_html=True)

# --- Session State Initialization ---
if 'logging_active' not in st.session_state:
    st.session_state['logging_active'] = False

# --- Helper Functions ---
def get_available_ports():
    """Scans and returns a list of available COM ports."""
    return serial.tools.list_ports.comports()

def make_card_html(label, value, sub_text=""):
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub_text}</div>
    </div>
    """

# --- Sidebar Configuration ---
st.sidebar.header("⚙️ Configuration")

if st.sidebar.button("Refresh Ports"):
    st.rerun()

available_ports = get_available_ports()
force_manual = st.sidebar.checkbox("My port is not listed / Enter manually")

if available_ports and not force_manual:
    default_index = 0
    # Heuristic: Prefer ports with "USB" in description
    for i, port in enumerate(available_ports):
        if "USB" in port.description:
            default_index = i
            
    selected_port_obj = st.sidebar.selectbox(
        "Serial Port", 
        available_ports, 
        index=default_index,
        format_func=lambda p: f"{p.device} ({p.description})"
    )
    serial_port = selected_port_obj.device
else:
    if not available_ports:
        st.sidebar.warning("No COM ports detected automatically.")
    
    raw_port = st.sidebar.text_input("Manually Enter Port (e.g., COM5)", "COM5")
    serial_port = raw_port.strip().upper()

baud_rate = st.sidebar.number_input("Baud Rate", value=9600, step=100)
output_filename = st.sidebar.text_input("Output Filename", "Desorption trial 1.xlsx")

st.sidebar.subheader("Performance Settings")
batch_size = st.sidebar.number_input("Batch Size", value=500, min_value=10, step=10)
max_rows_per_sheet = st.sidebar.number_input("Max Rows Per Sheet", value=800000)

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

if st.session_state['logging_active']:
    
    status_placeholder.markdown(make_card_html("Status", "Connecting...", "Initializing"), unsafe_allow_html=True)
    
    ser = None
    try:
        # 1. Establish Connection (Exact logic from snippet)
        ser = serial.Serial(
            port=serial_port,
            baudrate=baud_rate,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout=1  # Reverted to 1s as per snippet
        )
        
        status_placeholder.markdown(make_card_html("Status", "Active", f"Connected to {serial_port}"), unsafe_allow_html=True)
        st.toast(f"✅ Successfully connected to port {ser.portstr}")

        # 2. Send Command (Exact logic from snippet)
        if ser and ser.is_open:
            try:
                command = "CP\r\n"
                ser.write(command.encode('utf-8'))
                time.sleep(0.5) # Wait for device
            except Exception as e:
                st.error(f"❌ Error sending command: {e}")

        # 3. Initialization for File Handling
        data_batch = []
        sheet_index = 1
        row_count = 0
        
        # Check existing file to resume logging (Logic from snippet)
        if os.path.exists(output_filename):
            try:
                with pd.ExcelFile(output_filename) as xls:
                    last_sheet_name = xls.sheet_names[-1]
                    # Extract number from "Sheet1", "Sheet2" etc
                    match = re.search(r'Sheet(\d+)', last_sheet_name)
                    if match:
                        sheet_index = int(match.group(1))
                    
                    df_last = pd.read_excel(xls, sheet_name=last_sheet_name)
                    row_count = len(df_last)
                    
                    if row_count >= max_rows_per_sheet:
                        sheet_index += 1
                        row_count = 0
                st.toast(f"Resuming log on Sheet{sheet_index} at row {row_count + 1}")
            except Exception as e:
                st.warning(f"Could not inspect existing file. Starting fresh. Error: {e}")

        # 4. Main Loop
        last_ui_update = time.time()
        
        while st.session_state['logging_active']:
            try:
                # Read line (blocking up to timeout=1s)
                raw_response = ser.readline().decode('utf-8', errors='ignore')
                
                # Regex cleaning (Exact logic from snippet)
                response = re.sub(r'[\x00-\x1F\x7F]', '', raw_response).strip()

                if response:
                    timestamp = datetime.now()
                    data_batch.append({'Timestamp': timestamp, 'Response': response})
                    
                    # Update UI (Throttled to prevent lag, similar to 'status update interval' but faster for UI)
                    if time.time() - last_ui_update > 0.5:
                        current_total_estimate = row_count + len(data_batch)
                        last_val_placeholder.markdown(make_card_html("Last Reading", response, "Raw Value"), unsafe_allow_html=True)
                        count_placeholder.markdown(make_card_html("Rows Logged", f"~{current_total_estimate}", f"Sheet{sheet_index}"), unsafe_allow_html=True)
                        
                        # Update Table
                        preview_df = pd.DataFrame(data_batch[-10:])
                        if not preview_df.empty:
                            preview_df['Timestamp'] = preview_df['Timestamp'].dt.strftime('%H:%M:%S')
                            table_placeholder.dataframe(preview_df, use_container_width=True)
                        
                        last_ui_update = time.time()

                # Batch Saving (Exact logic from snippet)
                if len(data_batch) >= batch_size:
                    status_placeholder.markdown(make_card_html("Status", "Saving...", "Writing to Disk"), unsafe_allow_html=True)
                    
                    current_sheet_name = f"Sheet{sheet_index}"
                    df_batch = pd.DataFrame(data_batch)
                    
                    write_mode = 'a' if os.path.exists(output_filename) else 'w'
                    if_sheet_exists = 'overlay' if write_mode == 'a' else None
                    write_header = not os.path.exists(output_filename) or row_count == 0

                    with pd.ExcelWriter(output_filename, engine='openpyxl', mode=write_mode, if_sheet_exists=if_sheet_exists) as writer:
                        df_batch.to_excel(
                            writer, 
                            sheet_name=current_sheet_name, 
                            index=False, 
                            header=write_header, 
                            startrow=row_count if write_header else row_count + 1
                        )
                    
                    row_count += len(data_batch)
                    data_batch.clear()
                    
                    status_placeholder.markdown(make_card_html("Status", "Active", "Collecting Data"), unsafe_allow_html=True)

                    # Check for Sheet Overflow
                    if row_count >= max_rows_per_sheet:
                        st.toast(f"Sheet{sheet_index} is full. Switching to Sheet{sheet_index + 1}.")
                        sheet_index += 1
                        row_count = 0

            except serial.SerialException as e:
                st.error(f"Serial connection error: {e}")
                break
                
    except Exception as e:
        st.error(f"Connection Error: {e}")
        
    finally:
        # Cleanup block (Saving remaining data)
        if 'data_batch' in locals() and data_batch:
            st.info(f"Saving {len(data_batch)} remaining records...")
            current_sheet_name = f"Sheet{sheet_index}"
            df_batch = pd.DataFrame(data_batch)
            write_mode = 'a' if os.path.exists(output_filename) else 'w'
            if_sheet_exists = 'overlay' if write_mode == 'a' else None
            write_header = not os.path.exists(output_filename) or row_count == 0
            
            try:
                with pd.ExcelWriter(output_filename, engine='openpyxl', mode=write_mode, if_sheet_exists=if_sheet_exists) as writer:
                    df_batch.to_excel(
                        writer, 
                        sheet_name=current_sheet_name, 
                        index=False, 
                        header=write_header, 
                        startrow=row_count if write_header else row_count + 1
                    )
            except Exception as e:
                st.error(f"Error saving final batch: {e}")
        
        if ser and ser.is_open:
            ser.close()
            
        status_placeholder.markdown(make_card_html("Status", "Stopped", "Connection Closed"), unsafe_allow_html=True)
        st.session_state['logging_active'] = False
        st.success("Logging session finished.")

else:
    status_placeholder.markdown(make_card_html("Status", "Idle", "Ready to Start"), unsafe_allow_html=True)
    count_placeholder.markdown(make_card_html("Rows Logged", "0", "-"), unsafe_allow_html=True)
    last_val_placeholder.markdown(make_card_html("Last Reading", "-", "-"), unsafe_allow_html=True)
    
    st.info("Configure settings in the sidebar and press Start.")
