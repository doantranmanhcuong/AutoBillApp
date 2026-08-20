import streamlit as st

def display_zalo_message(company_name, total_amount):
    """Vẽ khối giao diện chứa tin nhắn Zalo mẫu để người dùng Copy"""
    if company_name and total_amount > 0:
        # Format số tiền dùng dấu chấm cho đẹp (VD: 4.143.960)
        amount_str = f"{total_amount:,.0f}".replace(",", ".")
        
        msg = f"Gửi sếp hồ sơ thanh toán cho {company_name}.\nTổng tiền: {amount_str} đ.\nSếp xem file đính kèm nhé!"
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.info("💡 **Mẫu tin nhắn Zalo (Bấm biểu tượng ở góc phải hộp thoại để Copy)**")
        st.code(msg, language="text")