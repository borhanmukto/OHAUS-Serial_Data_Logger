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
if 'sheet_index' not in st.session_state:
    st.session_state['sheet_index'] = 1
if 'row_count' not in st.session_state:
    st.session_state['row_count'] = 0

# --- Helper Functions ---
def get_available_ports():
    """Scans and returns a list of available COM ports."""
    # Return the full port objects, not just names, for better detection info
    return serial.tools.list_ports.comports()

def make_card_html(label, value, sub_text=""):
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub_text}</div>
    </div>
    """

def save_to_excel_optimized(filename, data, max_rows, current_sheet_idx, current_row_cnt):
    """Optimized Saver: Uses passed-in row counts instead of reading the file."""
    if not data:
        return False, current_sheet_idx, current_row_cnt
    
    df_batch = pd.DataFrame(data)
    batch_len = len(df_batch)
    
    if current_row_cnt + batch_len > max_rows:
        current_sheet_idx += 1
        current_row_cnt = 0
        write_header = True
        start_row = 0
    else:
        write_header = (current_row_cnt == 0)
        start_row = current_row_cnt + (1 if not write_header else 0)

    sheet_name = f"Sheet{current_sheet_idx}"
    mode = 'a' if os.path.exists(filename) else 'w'
    if_sheet_exists = 'overlay' if mode == 'a' else None

    try:
        with pd.ExcelWriter(filename, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
            df_batch.to_excel(
                writer, 
                sheet_name=sheet_name, 
                index=False, 
                header=write_header, 
                startrow=start_row
            )
        return True, current_sheet_idx, current_row_cnt + batch_len
    except Exception as e:
        st.error(f"Failed to save to Excel: {e}")
        return False, current_sheet_idx, current_row_cnt

# --- Sidebar Configuration ---
st.sidebar.header("⚙️ Configuration")

if st.sidebar.button("Refresh Ports"):
    st.rerun()

available_ports = get_available_ports()
force_manual = st.sidebar.checkbox("My port is not listed / Enter manually")

if available_ports and not force_manual:
    # Auto-detect logic: Try to select the most likely candidate (USB Serial)
    default_index = 0
    # Heuristic: Prefer ports with "USB" in description, or the last one in the list (usually newest)
    for i, port in enumerate(available_ports):
        if "USB" in port.description:
            default_index = i
    
    # If no specific "USB" port found, but list exists, default to the last one (often the external device)
    if default_index == 0 and len(available_ports) > 1 and "USB" not in available_ports[0].description:
        default_index = len(available_ports) - 1

    selected_port_obj = st.sidebar.selectbox(
        "Serial Port", 
        available_ports, 
        index=default_index,
        format_func=lambda p: f"{p.device} ({p.description})"
    )
    serial_port = selected_port_obj.device
else:
    # Manual Entry Mode
    if not available_ports:
        st.sidebar.warning("No COM ports detected automatically.")
    
    raw_port = st.sidebar.text_input("Manually Enter Port (e.g., COM4)", "COM4")
    # CRITICAL FIX: Strip spaces and ensure uppercase for Windows consistency
    serial_port = raw_port.strip().upper() 

baud_rate = st.sidebar.number_input("Baud Rate", value=9600, step=100)
output_filename = st.sidebar.text_input("Output Filename", "Desorption_Data.xlsx")

st.sidebar.subheader("Performance Settings")
batch_size = st.sidebar.number_input("Batch Size (Rows to buffer)", value=2000, min_value=100, step=100)
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
    st.session_state['sheet_index'] = 1
    st.session_state['row_count'] = 0
    
    if os.path.exists(output_filename):
        try:
            with pd.ExcelFile(output_filename) as xls:
                last_sheet = xls.sheet_names[-1]
                match = re.search(r'Sheet(\d+)', last_sheet)
                if match:
                    st.session_state['sheet_index'] = int(match.group(1))
                df_last = pd.read_excel(xls, sheet_name=last_sheet)
                st.session_state['row_count'] = len(df_last)
        except:
            pass 

if st.session_state['logging_active']:
    
    status_placeholder.markdown(make_card_html("Status", "Connecting...", "Initializing"), unsafe_allow_html=True)
    
    try:
        # Try to open the port
        ser = serial.Serial(
            port=serial_port,
            baudrate=baud_rate,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout=0.1
        )
        
        # Send Command
        time.sleep(1) 
        ser.write("CP\r\n".encode('utf-8'))
        
        status_placeholder.markdown(make_card_html("Status", "Active", "Connected"), unsafe_allow_html=True)
        st.toast(f"Connected to {serial_port}")

        buffer = []
        last_ui_update_time = time.time()
        total_session_rows = 0
        
        while True:
            try:
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

                    current_time = time.time()
                    if current_time - last_ui_update_time > 0.5:
                        current_total = total_session_rows + len(buffer)
                        last_reading = buffer[-1]['Response'] if buffer else "-"
                        
                        last_val_placeholder.markdown(make_card_html("Last Reading", last_reading, "Raw Value"), unsafe_allow_html=True)
                        count_placeholder.markdown(make_card_html("Rows Logged", str(current_total), "Session Total"), unsafe_allow_html=True)
                        
                        if buffer:
                            preview_df = pd.DataFrame(buffer[-10:])
                            preview_df['Timestamp'] = preview_df['Timestamp'].dt.strftime('%H:%M:%S.%f').str[:-3]
                            table_placeholder.dataframe(preview_df, use_container_width=True)
                        
                        last_ui_update_time = current_time

                    if len(buffer) >= batch_size:
                        status_placeholder.markdown(make_card_html("Status", "Saving...", "Writing to Disk"), unsafe_allow_html=True)
                        
                        success, new_idx, new_cnt = save_to_excel_optimized(
                            output_filename, 
                            buffer, 
                            max_rows_per_sheet,
                            st.session_state['sheet_index'],
                            st.session_state['row_count']
                        )
                        
                        if success:
                            total_session_rows += len(buffer)
                            st.session_state['sheet_index'] = new_idx
                            st.session_state['row_count'] = new_cnt
                            buffer = [] 
                            status_placeholder.markdown(make_card_html("Status", "Active", "Collecting Data"), unsafe_allow_html=True)
                        else:
                            st.error("Failed to save data. Stopping to prevent data loss.")
                            break
                else:
                    time.sleep(0.01)

            except serial.SerialException as e:
                st.error(f"Device disconnected: {e}")
                break
                
    except serial.SerialException as e:
        # Specific error handling for port issues
        if "Access is denied" in str(e):
            st.error(f"🚫 Access Denied to {serial_port}. It might be open in another program (like Jupyter Notebook). Please close other apps and try again.")
        elif "FileNotFound" in str(e) or "No such file" in str(e):
             st.error(f"❌ Could not find port '{serial_port}'. Please check the name and ensure the device is plugged in.")
        else:
            st.error(f"Connection Error: {e}")
        
    except Exception as e:
        st.error(f"Unexpected Error: {e}")
        
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            
        if 'buffer' in locals() and buffer:
            st.info(f"Saving {len(buffer)} remaining rows...")
            save_to_excel_optimized(
                output_filename, buffer, max_rows_per_sheet,
                st.session_state['sheet_index'], st.session_state['row_count']
            )
            
        status_placeholder.markdown(make_card_html("Status", "Stopped", "Connection Closed"), unsafe_allow_html=True)
        st.session_state['logging_active'] = False

else:
    status_placeholder.markdown(make_card_html("Status", "Idle", "Ready to Start"), unsafe_allow_html=True)
    count_placeholder.markdown(make_card_html("Rows Logged", "0", "-"), unsafe_allow_html=True)
    last_val_placeholder.markdown(make_card_html("Last Reading", "-", "-"), unsafe_allow_html=True)
    
    st.info("Configure settings in the sidebar and press Start.")
