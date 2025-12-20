import streamlit as st
import serial
import pandas as pd
from datetime import datetime
import time
import os
import re

# --- Page Configuration ---
st.set_page_config(
    page_title="Serial Data Logger",
    page_icon="⚖️",
    layout="wide"
)

# --- Title and Description ---
st.title("⚖️ Serial Data Logger")
st.markdown("""
This application logs data from a serial device (e.g., Ohaus scale) to an Excel file.
**Note:** This app must be run locally to access your computer's COM ports.
""")

# --- Session State Initialization ---
if 'ser' not in st.session_state:
    st.session_state.ser = None

# --- Sidebar Configuration ---
st.sidebar.header("🔌 Connection Settings")

# Default values from the original script
default_port = 'COM5'
default_baud = 9600
default_filename = "Desorption_trial_1.xlsx"

serial_port = st.sidebar.text_input("Serial Port (e.g., COM5 or /dev/ttyUSB0)", value=default_port)
baud_rate = st.sidebar.number_input("Baud Rate", value=default_baud, step=100)

st.sidebar.header("📁 File Settings")
output_filename = st.sidebar.text_input("Output Filename (.xlsx)", value=default_filename)
batch_size = st.sidebar.number_input("Write Batch Size (Rows)", value=500, help="Number of readings to collect before saving to Excel.")
max_rows_per_sheet = st.sidebar.number_input("Max Rows per Sheet", value=800000)

st.sidebar.header("⏱️ Update Settings")
ui_update_rate = st.sidebar.slider("Chart Update Rate (readings)", 1, 100, 10, help="Update the chart every N readings to save performance.")

# --- Helper Functions ---
def clean_response(raw_response):
    """Cleans the string of illegal control characters and whitespace."""
    try:
        decoded = raw_response.decode('utf-8', errors='ignore')
        # Remove control characters
        cleaned = re.sub(r'[\x00-\x1F\x7F]', '', decoded).strip()
        return cleaned
    except Exception:
        return None

def get_existing_file_info(filename, max_rows):
    """Checks existing Excel file to determine start sheet and row."""
    sheet_index = 1
    row_count = 0
    
    if os.path.exists(filename):
        try:
            # We only read the sheet names and the last sheet to save time
            xls = pd.ExcelFile(filename)
            last_sheet_name = xls.sheet_names[-1]
            
            # Extract number from "Sheet1", "Sheet2", etc.
            if "Sheet" in last_sheet_name:
                try:
                    sheet_index = int(last_sheet_name.replace("Sheet", ""))
                except ValueError:
                    sheet_index = 1
            
            # Read just the last sheet to get row count
            df_last = pd.read_excel(xls, sheet_name=last_sheet_name)
            row_count = len(df_last)
            
            if row_count >= max_rows:
                sheet_index += 1
                row_count = 0
            
            st.success(f"📂 Found existing file. Resuming on **Sheet{sheet_index}**, Row **{row_count}**.")
            return sheet_index, row_count
        except Exception as e:
            st.warning(f"⚠️ Could not inspect existing file. Starting fresh. Error: {e}")
            return 1, 0
    else:
        return 1, 0

# --- Main Logic ---

# Create two columns for the workflow controls
col1, col2 = st.columns(2)

# === STEP 1: CONNECTION ===
with col1:
    st.subheader("Step 1: Connection")
    
    # Check if we are already connected
    if st.session_state.ser is None:
        if st.button("🔌 Connect to Port", key="btn_connect"):
            try:
                # Open Serial Connection
                ser = serial.Serial(
                    port=serial_port,
                    baudrate=baud_rate,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    bytesize=serial.EIGHTBITS,
                    timeout=1
                )
                
                # Send Start Command (CP) immediately upon connection
                command = "CP\r\n"
                ser.write(command.encode('utf-8'))
                time.sleep(0.5)
                
                # Store in session state
                st.session_state.ser = ser
                st.success(f"✅ Successfully connected to {serial_port}")
                st.rerun() # Rerun to update the UI state
                
            except serial.SerialException as e:
                st.error(f"❌ Connection Error: {e}")
    else:
        # If connected, show details and disconnect button
        st.success(f"✅ Connected to **{st.session_state.ser.port}**")
        if st.button("❌ Disconnect", key="btn_disconnect"):
            if st.session_state.ser.is_open:
                st.session_state.ser.close()
            st.session_state.ser = None
            st.rerun()

# === STEP 2: LOGGING ===
start_logging = False
with col2:
    st.subheader("Step 2: Logging")
    if st.session_state.ser is not None and st.session_state.ser.is_open:
        start_logging = st.button("🚀 Start Logging", type="primary", key="btn_start")
    else:
        st.info("Please complete Step 1 to enable logging.")

# === LOGGING LOOP ===
if start_logging and st.session_state.ser is not None:
    ser = st.session_state.ser # Retrieve connection
    
    # 1. Initialization
    data_batch = []
    
    # UI Placeholders for real-time updates
    st.divider()
    status_container = st.container()
    metrics_col1, metrics_col2, metrics_col3 = st.columns(3)
    chart_placeholder = st.empty()
    log_placeholder = st.empty()
    
    with status_container:
        st.info(f"Logging active on {ser.port}. Click 'Stop' in the top-right toolbar to end session.")

    # Determine start point
    sheet_index, total_rows_written = get_existing_file_info(output_filename, max_rows_per_sheet)
    
    # Session state for chart data
    chart_data = []

    try:
        # 3. Main Loop
        current_sheet_rows = total_rows_written
        readings_count = 0
        
        while True:
            # Check if port is still open (safety check)
            if not ser.is_open:
                raise Exception("Serial port unexpectedly closed.")

            # Read line
            if ser.in_waiting > 0:
                raw_response = ser.readline()
                response = clean_response(raw_response)

                if response:
                    timestamp = datetime.now()
                    data_batch.append({'Timestamp': timestamp, 'Response': response})
                    
                    # Update metrics logic
                    readings_count += 1
                    current_sheet_rows += 1
                    
                    # Try to parse response as float for charting
                    try:
                        # Extract first number found in string for the chart
                        val_match = re.search(r"[-+]?\d*\.\d+|\d+", response)
                        if val_match:
                            val = float(val_match.group())
                            chart_data.append(val)
                            if len(chart_data) > 100: # Keep chart buffer small
                                chart_data.pop(0)
                    except:
                        pass

                    # -- UI UPDATES --
                    # Only update UI every N readings to prevent lag
                    if readings_count % ui_update_rate == 0:
                        with metrics_col1:
                            st.metric("Current Sheet", f"Sheet{sheet_index}")
                        with metrics_col2:
                            st.metric("Rows in Sheet", f"{current_sheet_rows:,}")
                        with metrics_col3:
                            st.metric("Latest Reading", response)
                        
                        if chart_data:
                            chart_placeholder.line_chart(chart_data)

                    # -- BATCH SAVE LOGIC --
                    if len(data_batch) >= batch_size:
                        df_batch = pd.DataFrame(data_batch)
                        
                        current_sheet_name = f"Sheet{sheet_index}"
                        write_mode = 'a' if os.path.exists(output_filename) else 'w'
                        if_sheet_exists = 'overlay' if write_mode == 'a' else None
                        
                        # Check if file exists to determine header writing
                        file_exists = os.path.exists(output_filename)
                        write_header = (not file_exists) or (current_sheet_rows <= len(data_batch))

                        with pd.ExcelWriter(output_filename, engine='openpyxl', mode=write_mode, if_sheet_exists=if_sheet_exists) as writer:
                            # Calculate start row logic
                            start_row = current_sheet_rows - len(data_batch) 
                            
                            # Reverting to exact logic from original script for safety
                            start_row_arg = start_row if write_header else start_row + 1
                            if not file_exists:
                                start_row_arg = 0 # Fresh file

                            df_batch.to_excel(
                                writer, 
                                sheet_name=current_sheet_name, 
                                index=False, 
                                header=write_header, 
                                startrow=start_row_arg
                            )
                        
                        log_placeholder.text(f"✅ Saved {len(data_batch)} rows to {current_sheet_name} at {datetime.now().strftime('%H:%M:%S')}")
                        data_batch.clear()

                        # Check for sheet overflow
                        if current_sheet_rows >= max_rows_per_sheet:
                            sheet_index += 1
                            current_sheet_rows = 0
                            st.warning(f"Sheet full. Switching to Sheet{sheet_index}")
            
            # Small sleep to prevent maxing out CPU if no data
            time.sleep(0.01)

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
    
    finally:
        # Graceful Exit Logic
        # If the user stops the script, we close the port and reset session state
        # so they can connect again cleanly next time.
        if st.session_state.ser and st.session_state.ser.is_open:
            st.session_state.ser.close()
            st.session_state.ser = None
            st.warning("🔌 Serial port closed. Connect again to restart.")
        
        # Save remaining data
        if 'data_batch' in locals() and data_batch:
            st.info(f"Saving {len(data_batch)} remaining records...")
            df_batch = pd.DataFrame(data_batch)
            current_sheet_name = f"Sheet{sheet_index}"
            write_mode = 'a' if os.path.exists(output_filename) else 'w'
            if_sheet_exists = 'overlay' if write_mode == 'a' else None
            # Simplistic save for remainder
            with pd.ExcelWriter(output_filename, engine='openpyxl', mode=write_mode, if_sheet_exists=if_sheet_exists) as writer:
                df_batch.to_excel(writer, sheet_name=current_sheet_name, index=False, header=False, startrow=current_sheet_rows - len(data_batch) + 1)
        
        st.success("👋 Program finished.")
