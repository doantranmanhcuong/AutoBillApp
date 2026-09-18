import streamlit as st

def apply_office_theme():
    """Áp dụng CSS giao diện văn phòng: đơn giản, trong sáng, thanh lịch, dễ đọc"""
    st.markdown("""
    <style>
        /* Tối ưu không gian và phông chữ */
        .block-container {
            padding-top: 1.8rem;
            padding-bottom: 2rem;
            max-width: 96%;
        }
        
        /* Tiêu đề thanh lịch */
        h1, h2, h3 {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            color: #1e293b;
            font-weight: 600;
        }
        
        /* Thẻ tóm tắt tài chính chuẩn kế toán */
        .metric-card {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 14px 18px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        }
        .metric-label {
            font-size: 0.85rem;
            color: #64748b;
            font-weight: 500;
            margin-bottom: 4px;
        }
        .metric-value {
            font-size: 1.35rem;
            font-weight: 700;
            color: #0f172a;
        }
        .metric-highlight {
            color: #1d4ed8;
        }
        
        /* Chỉnh nút bấm trang nhã */
        .stButton button[kind="primary"] {
            background-color: #2563eb;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            padding: 0.5rem 1rem;
        }
        .stButton button[kind="primary"]:hover {
            background-color: #1d4ed8;
            color: #ffffff;
        }
        
        /* Expander thanh lịch */
        .streamlit-expanderHeader {
            font-size: 0.95rem;
            font-weight: 600;
            color: #334155;
            background-color: #f8fafc;
            border-radius: 6px;
        }
    </style>
    """, unsafe_allow_html=True)

def display_financial_summary(tong_tien_hang, tien_thue, tong_thanh_toan, chu_so_tien, is_vat_included, thue_suat=None, tax_breakdown_str=None):
    """Hiển thị bảng tóm tắt số liệu thanh toán theo chuẩn kế toán (hỗ trợ cả đơn thuế và đa thuế 5%, 8%, 10%)"""
    if tax_breakdown_str:
        vat_status_badge = "(Đã gồm VAT)" if is_vat_included else "(Chưa gồm VAT)"
        tax_label = "Tổng tiền thuế GTGT:"
        sub_tax_info = f'<div style="font-size: 0.78rem; color: #64748b; margin-top: 5px; line-height: 1.3;">Chi tiết: {tax_breakdown_str}</div>'
    else:
        vat_status_badge = f"(Đã gồm VAT {thue_suat}%)" if is_vat_included else f"(Chưa gồm VAT, thuế {thue_suat}%)"
        tax_label = f"Tiền thuế GTGT ({thue_suat}%):" if thue_suat is not None else "Tiền thuế GTGT:"
        sub_tax_info = ""
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Cộng tiền hàng (chưa thuế):</div>
            <div class="metric-value">{tong_tien_hang:,.0f} đ</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">{tax_label}</div>
            <div class="metric-value">{tien_thue:,.0f} đ</div>
            {sub_tax_info}
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card" style="border-left: 4px solid #2563eb;">
            <div class="metric-label">Tổng cộng thanh toán {vat_status_badge}:</div>
            <div class="metric-value metric-highlight">{tong_thanh_toan:,.0f} đ</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown(f"""
    <div style="margin-top: 10px; padding: 8px 12px; background-color: #f1f5f9; border-radius: 6px; font-size: 0.9rem; color: #334155;">
        <b>Bằng chữ:</b> <i>{chu_so_tien}</i>
    </div>
    """, unsafe_allow_html=True)

def display_zalo_message(company_name, total_amount):
    """Vẽ khối giao diện chứa tin nhắn Zalo mẫu để người dùng sao chép nhanh"""
    if company_name and total_amount > 0:
        amount_str = f"{total_amount:,.0f}".replace(",", ".")
        msg = f"Kính gửi Sếp hồ sơ thanh toán cho {company_name}.\nTổng số tiền: {amount_str} đ.\nEm đã lập đầy đủ biểu mẫu đính kèm, kính chuyển Sếp xem duyệt ạ!"
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.caption("📋 **Mẫu tin nhắn gửi lãnh đạo / đối tác qua Zalo (Bấm nút sao chép ở góc phải):**")
        st.code(msg, language="text")