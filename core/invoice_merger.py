import re

class InvoiceMerger:
    @staticmethod
    def _normalize_str(s: str) -> str:
        if not s:
            return ""
        return re.sub(r'[\s\.\,\-_]+', '', str(s)).lower().strip()

    @classmethod
    def merge_invoices(cls, invoice_results: list[dict]) -> dict:
        """
        Hợp nhất danh sách kết quả bóc tách từ nhiều hóa đơn của cùng một Nhà Cung Cấp.
        :param invoice_results: Danh sách các dict dạng:
               [{"filename": "hd1.pdf", "data": {...}}, {"filename": "hd2.pdf", "data": {...}}]
        :return: Dict dữ liệu hợp nhất theo đúng schema chuẩn của AIExtractor
        """
        if not invoice_results:
            return {}

        valid_invoices = [inv for inv in invoice_results if inv.get("data") and "error" not in inv.get("data", {})]
        failed_invoices = [inv for inv in invoice_results if not inv.get("data") or "error" in inv.get("data", {})]

        if not valid_invoices:
            # Nếu tất cả đều lỗi, trả về lỗi của hóa đơn đầu tiên
            first_err = invoice_results[0].get("data", {}).get("error", "Không có dữ liệu hóa đơn hợp lệ để gộp.")
            return {"error": first_err}

        # 1. Thu thập và kiểm tra tính đồng nhất của Nhà Cung Cấp
        canh_bao_tong = []

        # Cảnh báo nếu có hóa đơn bị lỗi trong quá trình bóc tách
        if failed_invoices:
            failed_details = [f"'{inv.get('filename')}'" for inv in failed_invoices]
            canh_bao_tong.append(
                f"🚨 CẢNH BÁO QUAN TRỌNG: Bạn đã tải lên {len(invoice_results)} hóa đơn nhưng chỉ bóc tách thành công {len(valid_invoices)} hóa đơn. "
                f"Có {len(failed_invoices)} hóa đơn chưa được gộp ({', '.join(failed_details)}). Vui lòng kiểm tra lại các file này!"
            )
        ncc_names = []
        ncc_msts = []
        base_ncc = {
            "ten_cong_ty": "",
            "dia_chi": "",
            "dien_thoai": "",
            "ma_so_thue": "",
            "email": "",
            "so_tai_khoan": "",
            "ten_ngan_hang": "",
            "nguoi_dai_dien": "",
            "chuc_vu": ""
        }
        base_khach_hang = {
            "ten_khach_hang": "",
            "dia_chi": "",
            "ma_so_thue": "",
            "so_tai_khoan": "",
            "ten_ngan_hang": ""
        }

        for idx, item in enumerate(valid_invoices):
            inv_data = item.get("data") or {}
            fname = item.get("filename", f"Hóa đơn {idx + 1}")
            
            # Kiểm tra NCC
            ncc_info = inv_data.get("thong_tin_nha_cung_cap") or {}
            ten = str(ncc_info.get("ten_cong_ty") or "").strip()
            mst = str(ncc_info.get("ma_so_thue") or "").strip()

            if ten:
                ncc_names.append((fname, ten))
            if mst:
                ncc_msts.append((fname, mst))

            # Bổ sung các trường còn khuyết cho NCC từ các hóa đơn khác nhau
            for k in base_ncc:
                val = str(ncc_info.get(k) or "").strip()
                if val and not base_ncc[k]:
                    base_ncc[k] = val

            # Bổ sung thông tin khách hàng
            kh_info = inv_data.get("thong_tin_khach_hang") or {}
            for k in base_khach_hang:
                val = str(kh_info.get(k) or "").strip()
                if val and not base_khach_hang[k]:
                    base_khach_hang[k] = val

        # BIỆN PHÁP CHỐNG RÒ RỈ THÔNG TIN BÊN MUA (GRAPHENE) SANG BÊN BÁN (NCC):
        kh_mst = str(base_khach_hang.get("ma_so_thue") or "").strip()
        kh_stk = str(base_khach_hang.get("so_tai_khoan") or "").strip()
        if "graphene" in str(base_ncc.get("ten_cong_ty") or "").lower():
            base_ncc["ten_cong_ty"] = ""
        if str(base_ncc.get("ma_so_thue") or "").strip() in ("4101649609", kh_mst):
            base_ncc["ma_so_thue"] = ""
        if str(base_ncc.get("so_tai_khoan") or "").strip() in ("884249867", kh_stk):
            base_ncc["so_tai_khoan"] = ""

        for idx, item in enumerate(valid_invoices):
            inv_data = item.get("data") or {}
            fname = item.get("filename", f"Hóa đơn {idx + 1}")
            # Gom cảnh báo từng hóa đơn
            for cb in inv_data.get("danh_sach_canh_bao", []):
                if cb and str(cb).strip():
                    canh_bao_tong.append(f"[{fname}] {cb}")

        # Kiểm tra cảnh báo nếu có dấu hiệu khác Nhà cung cấp
        if len(ncc_msts) > 1:
            first_mst_norm = cls._normalize_str(ncc_msts[0][1])
            mismatches = [f"{fname} (MST: {mst})" for fname, mst in ncc_msts if cls._normalize_str(mst) != first_mst_norm]
            if mismatches:
                canh_bao_tong.insert(0, f"⚠️ Lưu ý: Phát hiện các hóa đơn có Mã số thuế khác nhau: {', '.join(mismatches)}. Vui lòng kiểm tra lại xem có đúng cùng một Nhà cung cấp không.")
        elif len(ncc_names) > 1:
            first_name_norm = cls._normalize_str(ncc_names[0][1])
            mismatches = [f"{fname} ({ten})" for fname, ten in ncc_names if cls._normalize_str(ten) != first_name_norm]
            if mismatches:
                canh_bao_tong.insert(0, f"⚠️ Lưu ý: Tên nhà cung cấp giữa các hóa đơn có sự khác biệt: {', '.join(mismatches)}.")

        # 2. Hợp nhất số chứng từ và ngày tháng
        so_chung_tu_list = []
        ngay_thang_list = []

        for idx, item in enumerate(valid_invoices):
            inv_data = item.get("data") or {}
            tt_chung = inv_data.get("thong_tin_chung") or {}
            sct = str(tt_chung.get("so_chung_tu") or "").strip()
            ntn = str(tt_chung.get("ngay_thang_nam") or "").strip()
            
            if sct and sct not in so_chung_tu_list:
                so_chung_tu_list.append(sct)
            if ntn and ntn not in ngay_thang_list:
                ngay_thang_list.append(ntn)

        merged_so_chung_tu = ", ".join(so_chung_tu_list) if so_chung_tu_list else ""
        merged_ngay_thang = ", ".join(ngay_thang_list) if ngay_thang_list else ""

        # 3. Hợp nhất danh sách hàng hóa
        merged_hang_hoa = []
        current_stt = 1

        from core.utils.helpers import safe_float

        for idx, item in enumerate(valid_invoices):
            inv_data = item.get("data") or {}
            fname = item.get("filename", f"HĐ {idx + 1}")
            sct = str((inv_data.get("thong_tin_chung") or {}).get("so_chung_tu") or "").strip()
            inv_ref = f"HĐ {sct}" if sct else fname

            ds_hh = inv_data.get("danh_sach_hang_hoa") or []
            if not isinstance(ds_hh, list):
                continue

            for hh in ds_hh:
                if not isinstance(hh, dict):
                    continue

                ten_hh = str(hh.get("ten_hang_hoa") or "").strip()
                if not ten_hh:
                    continue

                sl = safe_float(hh.get("so_luong"), 0.0)
                don_gia = safe_float(hh.get("don_gia"), 0.0)
                thanh_tien = safe_float(hh.get("thanh_tien"), sl * don_gia)
                if thanh_tien == 0 and sl > 0 and don_gia > 0:
                    thanh_tien = sl * don_gia

                ghi_chu_goc = str(hh.get("ghi_chu") or "").strip()
                
                # Ghi chú nguồn gốc từ hóa đơn nào nếu có nhiều hóa đơn gộp
                if len(valid_invoices) > 1:
                    note_tag = f"[{inv_ref}]"
                    ghi_chu_final = f"{note_tag} {ghi_chu_goc}".strip() if ghi_chu_goc else note_tag
                else:
                    ghi_chu_final = ghi_chu_goc

                ten_ncc_inv = str((inv_data.get("thong_tin_nha_cung_cap") or {}).get("ten_cong_ty") or base_ncc.get("ten_cong_ty") or "").strip()
                muc_dich_inv = str(hh.get("muc_dich") or hh.get("muc_dich_su_dung") or "").strip()
                ton_kho_inv = str(hh.get("ton_kho") or hh.get("ton_kh") or "").strip()

                # Thuế suất từng mặt hàng
                item_thue = hh.get("thue_suat")
                if item_thue is None or item_thue == "":
                    item_thue = (inv_data.get("thong_tin_vat") or {}).get("thue_suat", 8.0)
                item_thue_float = safe_float(item_thue, 8.0)

                merged_hang_hoa.append({
                    "stt": current_stt,
                    "ten_hang_hoa": ten_hh,
                    "don_vi_tinh": str(hh.get("don_vi_tinh") or ""),
                    "so_luong": sl,
                    "don_gia": don_gia,
                    "thanh_tien": thanh_tien,
                    "thue_suat": item_thue_float,
                    "muc_dich": muc_dich_inv,
                    "muc_dich_su_dung": muc_dich_inv,
                    "ton_kho": ton_kho_inv,
                    "ton_kh": ton_kho_inv,
                    "nha_cung_cap": ten_ncc_inv,
                    "ghi_chu": ghi_chu_final
                })
                current_stt += 1

        # 4. Hợp nhất thông tin VAT
        vat_modes = [bool(((inv.get("data") or {}).get("thong_tin_vat") or {}).get("da_bao_gom_vat", False)) for inv in valid_invoices]
        vat_rates = [safe_float(((inv.get("data") or {}).get("thong_tin_vat") or {}).get("thue_suat", 8.0), 8.0) for inv in valid_invoices]
        
        # Chọn chế độ VAT phổ biến nhất
        da_bao_gom_vat_final = max(set(vat_modes), key=vat_modes.count) if vat_modes else False
        thue_suat_final = max(set(vat_rates), key=vat_rates.count) if vat_rates else 8.0

        if len(set(vat_rates)) > 1:
            canh_bao_tong.append(f"⚠️ Lưu ý: Các hóa đơn có mức thuế suất GTGT khác nhau ({list(set(vat_rates))}%). Mức thuế mặc định được chọn là {thue_suat_final}%.")

        # 5. Hợp nhất các thẻ thông tin động (thong_tin_dong)
        merged_thong_tin_dong = {}
        for item in valid_invoices:
            tt_dong = (item.get("data") or {}).get("thong_tin_dong") or {}
            if isinstance(tt_dong, dict):
                for k, v in tt_dong.items():
                    if v and not merged_thong_tin_dong.get(k):
                        merged_thong_tin_dong[k] = v

        # Điền các trường tự động gợi ý cho đề nghị thanh toán / hợp đồng
        if merged_so_chung_tu:
            merged_thong_tin_dong.setdefault("ly_do_thanh_toan", f"Thanh toán theo các hóa đơn số: {merged_so_chung_tu}")
            merged_thong_tin_dong.setdefault("ly_do_de_nghi", f"Thanh toán theo các hóa đơn số: {merged_so_chung_tu}")

        first_chung = (valid_invoices[0].get("data") or {}).get("thong_tin_chung") or {}
        return {
            "thong_tin_nha_cung_cap": base_ncc,
            "thong_tin_chung": {
                "loai_chung_tu": "Bảng kê hóa đơn tổng hợp" if len(valid_invoices) > 1 else first_chung.get("loai_chung_tu", "Hóa đơn"),
                "so_chung_tu": merged_so_chung_tu,
                "ngay_thang_nam": merged_ngay_thang
            },
            "thong_tin_khach_hang": base_khach_hang,
            "danh_sach_hang_hoa": merged_hang_hoa,
            "thong_tin_vat": {
                "da_bao_gom_vat": da_bao_gom_vat_final,
                "thue_suat": thue_suat_final
            },
            "thong_tin_dong": merged_thong_tin_dong,
            "danh_sach_canh_bao": canh_bao_tong,
            "_multi_invoice_meta": {
                "total_uploaded": len(invoice_results),
                "total_invoices": len(valid_invoices),
                "failed_count": len(failed_invoices),
                "failed_files": [inv.get("filename") for inv in failed_invoices],
                "invoice_files": [item.get("filename") for item in valid_invoices]
            }
        }
