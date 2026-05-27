import streamlit as st
import pandas as pd
import time
import os
import io
from datetime import datetime

EXCEL_FILE = "inventory_temper.xlsx"

# ฟังก์ชันเซฟไฟล์แบบกันแอปพัง (ดักคนเปิดไฟล์ค้าง)
def save_data(df_items, df_logs):
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_items.to_excel(writer, sheet_name="items", index=False)
            df_logs.to_excel(writer, sheet_name="transaction_logs", index=False)
        return True
    except PermissionError:
        st.error("❌ บันทึกไม่ได้! เนื่องจากมีคนกำลังเปิดไฟล์ 'inventory_temper.xlsx' ค้างไว้ กรุณาปิดไฟล์ Excel ก่อนทำรายการ")
        return False
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        return False

def init_database():
    if not os.path.exists(EXCEL_FILE):
        initial_stock = [
            {"barcode_id": "1001", "item_name": "ถุงมือทนความร้อนสูง (ไลน์ Temper)", "quantity": 30, "min_stock": 10},
            {"barcode_id": "1002", "item_name": "ผ้าชามัวร์เช็ดกระจก", "quantity": 50, "min_stock": 20},
            {"barcode_id": "1003", "item_name": "สเปรย์ทำความสะอาดแม่พิมพ์", "quantity": 15, "min_stock": 5},
            {"barcode_id": "1004", "item_name": "คัตเตอร์ตัดฟิล์ม/ขูดกระจก", "quantity": 25, "min_stock": 10},
            {"barcode_id": "1005", "item_name": "แว่นตานิรภัยสำหรับหน้าเตา", "quantity": 20, "min_stock": 5}
        ]
        df_items = pd.DataFrame(initial_stock)
        df_logs = pd.DataFrame(columns=["timestamp", "action_type", "user_name", "item_name", "qty"])
        save_data(df_items, df_logs)
    else:
        try:
            df_items = pd.read_excel(EXCEL_FILE, sheet_name="items", dtype={"barcode_id": str})
            if 'min_stock' not in df_items.columns:
                df_items['min_stock'] = 5 
                df_logs = pd.read_excel(EXCEL_FILE, sheet_name="transaction_logs")
                save_data(df_items, df_logs)
        except Exception:
            pass # ข้ามไปถ้าไฟล์ถูกเปิดอยู่ตอนรันครั้งแรก

def get_all_items():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="items", dtype={"barcode_id": str})
        df.columns = ['โค้ด', 'ชื่ออุปกรณ์', 'จำนวน', 'ขั้นต่ำ']
        df['สถานะ'] = df.apply(lambda x: '🔴 ต่ำกว่าเกณฑ์' if x['จำนวน'] <= x['ขั้นต่ำ'] else '🟢 ปกติ', axis=1)
        return df
    except:
        return pd.DataFrame(columns=['โค้ด', 'ชื่ออุปกรณ์', 'จำนวน', 'ขั้นต่ำ', 'สถานะ'])

def get_transaction_logs():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="transaction_logs", dtype={"qty": int})
        if df.empty:
            return pd.DataFrame(columns=['วัน-เวลา', 'ประเภท', 'ผู้ทำรายการ', 'ชื่ออุปกรณ์', 'จำนวน'])
        mapping = {'WITHDRAW': '🔴 เบิก', 'RESTOCK': '🟢 เติม', 'EDIT_MASTER': '⚙️ แก้ไข Master'}
        df['action_type'] = df['action_type'].map(mapping).fillna('อื่นๆ')
        df.columns = ['วัน-เวลา', 'ประเภท', 'ผู้ทำรายการ', 'ชื่ออุปกรณ์', 'จำนวน']
        return df.iloc[::-1]
    except:
        return pd.DataFrame(columns=['วัน-เวลา', 'ประเภท', 'ผู้ทำรายการ', 'ชื่ออุปกรณ์', 'จำนวน'])

def highlight_low_stock(row):
    if row['จำนวน'] <= row['ขั้นต่ำ']:
        return ['background-color: #ffe6e6; color: #cc0000'] * len(row) 
    return [''] * len(row)

# CSS สำหรับฟอนต์ Barcode
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Libre+Barcode+39&display=swap');
.barcode-text {
    font-family: 'Libre Barcode 39', cursive;
    font-size: 48px;
    color: black;
    padding: 10px;
    background-color: white;
    border-radius: 5px;
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="ระบบคลังอุปกรณ์แผนก Temper", page_icon="🔥", layout="centered")
init_database()

st.title("🔥 TEMPER INVENTORY SYSTEM")
st.markdown("**ระบบคลังอุปกรณ์ไลน์ Temper (ใช้งานง่าย | แจ้งเตือน | กันไฟล์ล็อก)**")
st.divider()

tab1, tab2, tab3 = st.tabs(["👋 เบิกอุปกรณ์", "🔐 ผู้ดูแลระบบ (Admin)", "🖨️ สร้างบาร์โค้ด"])

# ==========================================
# TAB 1: สำหรับพนักงานเบิกของ
# ==========================================
with tab1:
    df_stock = get_all_items()
    
    if not df_stock.empty:
        low_stock_df = df_stock[df_stock['จำนวน'] <= df_stock['ขั้นต่ำ']]
        if not low_stock_df.empty:
            low_stock_top10 = low_stock_df.sort_values(by='จำนวน', ascending=False).head(10)
            st.error("⚠️ **แจ้งเตือนสต๊อกต่ำกว่าเกณฑ์**")
            st.dataframe(low_stock_top10.style.apply(highlight_low_stock, axis=1), use_container_width=True, hide_index=True)
            st.divider()

        with st.expander("📦 ดูตารางคลังอุปกรณ์ทั้งหมด (คลิกเพื่อเปิด/ปิด)", expanded=False):
            st.dataframe(df_stock.style.apply(highlight_low_stock, axis=1), use_container_width=True, hide_index=True)
    
    st.subheader("📝 ยิงบาร์โค้ดบันทึกการเบิกอุปกรณ์")
    user_name = st.text_input("1. ชื่อพนักงานผู้เบิก :", key="w_user")
    barcode_input = st.text_input("2. คลิกช่องนี้แล้วยิง บาร์โค้ด :", key="w_barcode")
    request_qty = st.number_input("3. จำนวนที่เบิก (ชิ้น) :", min_value=1, step=1, value=None, key="w_qty")

    if st.button("💾 ยืนยันการเบิกสินค้า", type="primary", use_container_width=True):
        if not user_name or not barcode_input or request_qty is None:
            st.warning("⚠️ กรุณากรอกข้อมูลให้ครบถ้วนก่อนยืนยัน")
        else:
            try:
                df_items = pd.read_excel(EXCEL_FILE, sheet_name="items", dtype={"barcode_id": str})
                target_barcode = barcode_input.strip()
                row_match = df_items[df_items['barcode_id'] == target_barcode]
                
                if row_match.empty:
                    st.error(f"❌ ไม่พบรหัสโค้ด '{target_barcode}' นี้ในระบบ!")
                else:
                    idx = row_match.index[0]
                    item_name = df_items.loc[idx, 'item_name']
                    current_qty = int(df_items.loc[idx, 'quantity'])
                    min_stock = int(df_items.loc[idx, 'min_stock'])
                    
                    if request_qty > current_qty:
                        st.error(f"❌ สต๊อกไม่พอ! มีอยู่ {current_qty} ชิ้น")
                    else:
                        new_qty = current_qty - request_qty
                        df_items.loc[idx, 'quantity'] = new_qty
                        
                        df_logs = pd.read_excel(EXCEL_FILE, sheet_name="transaction_logs")
                        new_log = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "action_type": "WITHDRAW", "user_name": user_name, "item_name": item_name, "qty": request_qty}
                        df_logs = pd.concat([df_logs, pd.DataFrame([new_log])], ignore_index=True)
                        
                        if save_data(df_items, df_logs):
                            st.success(f"✅ เบิกสำเร็จ! {item_name} ออกไป {request_qty} ชิ้น (เหลือ {new_qty} ชิ้น)")
                            if new_qty <= min_stock:
                                st.warning(f"⚠️ **เตือน:** '{item_name}' เหลือเพียง {new_qty} ชิ้น โปรดแจ้งแอดมินให้เติมสต๊อก!")
                                time.sleep(3)
                            else:
                                time.sleep(1)
                            st.rerun()
            except PermissionError:
                st.error("❌ บันทึกไม่ได้! มีคนกำลังเปิดไฟล์ Excel ค้างไว้ กรุณาปิดไฟล์ก่อนทำรายการ")

# ==========================================
# TAB 2: สำหรับ ADMIN
# ==========================================
with tab2:
    st.subheader("🔐 พื้นที่ควบคุมแอดมิน")
    ADMIN_PASSWORD = "1234" 
    
    password_input = st.text_input("กรุณากรอกรหัสผ่านแอดมิน :", type="password", key="admin_pwd")
    login_button = st.button("🔓 เข้าสู่ระบบแอดมิน", use_container_width=True)
    
    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False
        
    if login_button:
        if password_input == ADMIN_PASSWORD:
            st.session_state.admin_logged_in = True
            st.success("🔓 เข้าสู่ระบบแอดมินสำเร็จ")
            time.sleep(0.5)
            st.rerun()
        else:
            st.error("❌ รหัสผ่านไม่ถูกต้อง")
            st.session_state.admin_logged_in = False

    if st.session_state.admin_logged_in:
        st.divider()
        admin_menu = st.radio("เลือกการทำงาน :", ["🟢 เติมสต๊อก", "✨ จัดการตาราง Master", "📜 ดูประวัติ"], horizontal=True)
        st.divider()
        
        if admin_menu == "🟢 เติมสต๊อก":
            admin_name = st.text_input("1. ชื่อแอดมินผู้เติมสต๊อก :", key="a_user")
            barcode_admin = st.text_input("2. ยิง บาร์โค้ด :", key="a_barcode")
            add_qty = st.number_input("3. จำนวนที่นำมาเติม (ชิ้น) :", min_value=1, step=1, value=None, key="a_qty")

            if st.button("🚀 อัปเดตสต๊อก", use_container_width=True, type="primary"):
                if not admin_name or not barcode_admin or add_qty is None:
                    st.warning("⚠️ กรุณากรอกข้อมูลให้ครบถ้วน")
                else:
                    try:
                        df_items = pd.read_excel(EXCEL_FILE, sheet_name="items", dtype={"barcode_id": str})
                        target_barcode = barcode_admin.strip()
                        row_match = df_items[df_items['barcode_id'] == target_barcode]
                        
                        if row_match.empty:
                            st.error("❌ ไม่พบรหัสโค้ดนี้ใน Master!")
                        else:
                            idx = row_match.index[0]
                            item_name = df_items.loc[idx, 'item_name']
                            df_items.loc[idx, 'quantity'] = int(df_items.loc[idx, 'quantity']) + add_qty
                            
                            df_logs = pd.read_excel(EXCEL_FILE, sheet_name="transaction_logs")
                            new_log = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "action_type": "RESTOCK", "user_name": admin_name, "item_name": item_name, "qty": add_qty}
                            df_logs = pd.concat([df_logs, pd.DataFrame([new_log])], ignore_index=True)
                            
                            if save_data(df_items, df_logs):
                                st.success(f"🟢 เติมสต๊อก '{item_name}' เพิ่มอีก {add_qty} ชิ้น")
                                time.sleep(1)
                                st.rerun()
                    except PermissionError:
                        st.error("❌ บันทึกไม่ได้! มีคนกำลังเปิดไฟล์ Excel ค้างไว้ กรุณาปิดไฟล์ก่อน")

        elif admin_menu == "✨ จัดการตาราง Master":
            admin_name = st.text_input("ชื่อแอดมินผู้ทำการบันทึก Master :", key="m_user_edit")
            df_current = get_all_items().drop(columns=['สถานะ'])
            edited_df = st.data_editor(df_current, num_rows="dynamic", use_container_width=True, key="master_editor")
            
            if st.button("💾 บันทึกตารางลงไฟล์ Excel (Save)", type="primary", use_container_width=True):
                if not admin_name:
                    st.warning("⚠️ กรุณากรอกชื่อแอดมินก่อนกดบันทึก")
                else:
                    has_error = False
                    for idx, row in edited_df.iterrows():
                        if pd.isna(row['โค้ด']) or str(row['โค้ด']).strip() == "" or pd.isna(row['ชื่ออุปกรณ์']) or str(row['ชื่ออุปกรณ์']).strip() == "":
                            has_error = True
                            
                    if has_error:
                        st.error("❌ ข้อมูลไม่สมบูรณ์! รหัสโค้ดและชื่ออุปกรณ์ห้ามปล่อยว่าง")
                    else:
                        edited_df_save = edited_df.copy()
                        edited_df_save.columns = ['barcode_id', 'item_name', 'quantity', 'min_stock']
                        edited_df_save['barcode_id'] = edited_df_save['barcode_id'].astype(str).str.strip()
                        edited_df_save['item_name'] = edited_df_save['item_name'].astype(str).str.strip()
                        
                        try:
                            df_logs = pd.read_excel(EXCEL_FILE, sheet_name="transaction_logs")
                            new_log = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "action_type": "EDIT_MASTER", "user_name": admin_name, "item_name": "อัปเดตตาราง Master & ขั้นต่ำ", "qty": 0}
                            df_logs = pd.concat([df_logs, pd.DataFrame([new_log])], ignore_index=True)
                            
                            if save_data(edited_df_save, df_logs):
                                st.success("✅ อัปเดตข้อมูลและตั้งค่าขั้นต่ำลงไฟล์ Excel เรียบร้อยแล้ว!")
                                time.sleep(1)
                                st.rerun()
                        except PermissionError:
                            st.error("❌ บันทึกไม่ได้! มีคนกำลังเปิดไฟล์ Excel ค้างไว้")

        elif admin_menu == "📜 ดูประวัติ":
            df_logs = get_transaction_logs()
            if not df_logs.empty:
                st.dataframe(df_logs, use_container_width=True, hide_index=True)
            else:
                st.caption("ยังไม่มีประวัติการทำรายการ")
                
        if st.button("🔒 ออกจากระบบ Admin", type="secondary"):
            st.session_state.admin_logged_in = False
            st.rerun()

# ==========================================
# TAB 3: สร้างและดาวน์โหลด Excel Barcode
# ==========================================
with tab3:
    st.subheader("🖨️ สร้างบาร์โค้ดสำหรับนำไปปรินต์")
    df_bc = get_all_items()
    if not df_bc.empty:
        df_bc = df_bc[['โค้ด', 'ชื่ออุปกรณ์']]
        df_bc['Barcode_Print'] = "*" + df_bc['โค้ด'].astype(str) + "*"
        
        st.caption("👀 ตัวอย่างหน้าตาบาร์โค้ด (ดึงจากรหัส Master ปัจจุบัน)")
        for i, row in df_bc.iterrows():
            col1, col2 = st.columns([1, 2])
            with col1:
                st.markdown(f"**{row['ชื่ออุปกรณ์']}**<br>รหัส: {row['โค้ด']}", unsafe_allow_html=True)
            with col2:
                st.markdown(f"<span class='barcode-text'>*{row['โค้ด']}*</span>", unsafe_allow_html=True)
            st.divider()
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_bc.to_excel(writer, index=False, sheet_name='Barcodes')
        processed_data = output.getvalue()
        
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel บาร์โค้ด (Click here)", data=processed_data, file_name="temper_barcodes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)