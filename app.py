import streamlit as st
import pandas as pd

# 1. Cấu hình giao diện trang Web
st.set_page_config(page_title="Trang Web Báo Cáo Pivot 3 Sheet & Drill-down", layout="wide")
st.title("📊 Trang tổng hợp dữ liệu KPI")

# Hàm tự động xử lý trùng tên cột nếu file Excel có 2 cột giống hệt tên nhau
def deduplicate_columns(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

# 2. Thanh bên trái: Tải file dữ liệu
uploaded_file = st.sidebar.file_uploader("Tải lên file dữ liệu Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        # Kiểm tra xem file có đủ 3 sheet cố định không
        required_sheets = ["Baocao", "BVDR B1", "KHCN"]
        found_sheets = [s for s in required_sheets if s in sheet_names]

        if "Baocao" not in sheet_names:
            st.error("⚠️ File Excel thiếu trang (sheet) 'Baocao'. Vui lòng kiểm tra lại file!")
        else:
            # 3. ĐỌC DỮ LIỆU NGUYÊN BẢN TỪ 3 SHEET CỐ ĐỊNH
            df_baocao = pd.read_excel(uploaded_file, sheet_name="Baocao")
            df_baocao.columns = [str(col).strip() for col in df_baocao.columns]
            df_baocao = deduplicate_columns(df_baocao)

            df_bvdr = pd.read_excel(uploaded_file, sheet_name="BVDR B1") if "BVDR B1" in sheet_names else pd.DataFrame()
            if not df_bvdr.empty:
                df_bvdr.columns = [str(col).strip() for col in df_bvdr.columns]
                df_bvdr = deduplicate_columns(df_bvdr)

            df_khcn = pd.read_excel(uploaded_file, sheet_name="KHCN") if "KHCN" in sheet_names else pd.DataFrame()
            if not df_khcn.empty:
                df_khcn.columns = [str(col).strip() for col in df_khcn.columns]
                df_khcn = deduplicate_columns(df_khcn)

            st.sidebar.success(f"Đã tải thành công 3 Sheet: {', '.join(found_sheets)}")

            # --- NHẬN DIỆN CỘT CÁN BỘ TRÊN SHEET BAOCAO (CỘT DÒNG DUY NHẤT) ---
            row_col = "Cán bộ giải quyết bồi thường"
            for c in df_baocao.columns:
                if "cán bộ giải quyết bồi thường" in c.lower() or "can bo giai quyet boi thuong" in c.lower():
                    row_col = c
                    break

            # Tên cột cán bộ ở 2 Sheet phụ (Dùng để dò tìm, KHÔNG ĐỔI TÊN HAY THÊM VÀO PIVOT)
            bvdr_cb_col = next((c for c in df_bvdr.columns if "cb đầu mối" in c.lower() or "cb dau moi" in c.lower()), None) if not df_bvdr.empty else None
            khcn_cb_col = next((c for c in df_khcn.columns if "cán bộ bt" in c.lower() or "can bo bt" in c.lower()), None) if not df_khcn.empty else None

            # --- NHẬN DIỆN VÀ ĐỔI TÊN CỘT 'Mã HS' THÀNH 'C37/CI' ---
            cols_baocao = list(df_baocao.columns)
            loai_hoso_col = None
            totrinh_col = None
            ngoaigiao_col = None
            ma_hs_col = None
            dung_han_col = None  # Cột phân loại Đúng/Trễ hạn

            for c in cols_baocao:
                c_lower = c.lower()
                if ("loại hồ sơ" in c_lower or "loai ho so" in c_lower) and not loai_hoso_col:
                    loai_hoso_col = c
                elif ("tờ trình bttđ" in c_lower or "to trinh bttd" in c_lower or "bttđ" in c_lower) and not totrinh_col:
                    totrinh_col = c
                elif ("ngoại giao" in c_lower or "ngoai giao" in c_lower) and not ngoaigiao_col:
                    ngoaigiao_col = c
                elif ("mã hs" in c_lower or "ma hs" in c_lower or "mahs" in c_lower) and not ma_hs_col:
                    ma_hs_col = c
                elif any(kw in c_lower for kw in ["đúng/trễ", "dung/tre", "đúng hạn", "dung han", "tiến độ", "tien do", "trễ hạn"]) and not dung_han_col:
                    dung_han_col = c

            # Đổi tên cột 'Mã HS' sang 'C37/CI' trên DataFrame df_baocao
            c37_col_name = "C37/CI"
            if ma_hs_col:
                df_baocao.rename(columns={ma_hs_col: c37_col_name}, inplace=True)
                # Cập nhật lại danh sách cột
                cols_baocao = list(df_baocao.columns)
                if ma_hs_col == loai_hoso_col: loai_hoso_col = c37_col_name
                ma_hs_col = c37_col_name

            # Quy đổi mã 1, 2, 3 của Loại hồ sơ
            mapping_dict = {
                1: "Ngoại trú", "1": "Ngoại trú", "1.0": "Ngoại trú",
                2: "Nội trú", "2": "Nội trú", "2.0": "Nội trú",
                3: "Tử vong", "3": "Tử vong", "3.0": "Tử vong"
            }

            if loai_hoso_col:
                df_baocao[loai_hoso_col + "_mapped"] = df_baocao[loai_hoso_col].map(mapping_dict).fillna(df_baocao[loai_hoso_col].astype(str))

            # --- CẤU HÌNH PIVOT TABLE ---
            st.sidebar.header("⚙️ Cấu hình Pivot Table")
            
            # Danh sách các cột giá trị có thể chọn ở Sheet Baocao (Loại trừ cột Dòng row_col)
            available_val_cols = [c for c in cols_baocao if c != row_col and not c.endswith("_mapped")]
            default_vals = available_val_cols[:2] if len(available_val_cols) >= 2 else available_val_cols

            val_cols = st.sidebar.multiselect(
                "Chọn các cột Giá trị cần đếm (từ Sheet Baocao):", 
                available_val_cols, 
                default=default_vals
            )

            # --- DÒ TÌM & TÍNH TOÁN DỮ LIỆU TỪ 2 SHEET KHCN VÀ BVDR B1 ---
            
            # 1. Logic Sheet KHCN (Dò theo cột 'Cán bộ BT')
            khcn_metrics = {}
            if not df_khcn.empty and khcn_cb_col:
                nghiepvu_col = next((c for c in df_khcn.columns if "nghiệp vụ" in c.lower() or "nghiep vu" in c.lower()), None)

                if nghiepvu_col:
                    list_atsk = ["KHN", "ATS"]
                    list_pawci = ["CPA", "WCI", "TNCN.HSP", "TNCN.GVP", "PAI"]
                    list_dulich = ["DQT", "DLVN", "FLE", "YDL", "DTN", "NND"]

                    for cb, group in df_khcn.groupby(khcn_cb_col):
                        cb_str = str(cb).strip()
                        nv_series = group[nghiepvu_col].astype(str).str.strip().str.upper()
                        
                        khcn_metrics[cb_str] = {
                            "KHCN/ATSK": int(nv_series.isin(list_atsk).sum()),
                            "PA/WCI": int(nv_series.isin(list_pawci).sum()),
                            "Du lịch": int(nv_series.isin(list_dulich).sum())
                        }

            # 2. Logic Sheet BVDR B1 (Dò theo cột 'CB đầu mối')
            bvdr_metrics = {}
            bvdr_col_name = "D99 cứng/bổ sung mềm B1 (nội+ngoại)"
            if not df_bvdr.empty and bvdr_cb_col:
                hsmem_col = next((c for c in df_bvdr.columns if "hsmem" in c.lower()), None)
                hsbs_col = next((c for c in df_bvdr.columns if "hsbs" in c.lower()), None)
                tcbt_col = next((c for c in df_bvdr.columns if "tcbt" in c.lower()), None)

                if hsmem_col and hsbs_col and tcbt_col:
                    for cb, group in df_bvdr.groupby(bvdr_cb_col):
                        cb_str = str(cb).strip()
                        
                        cond1 = (group[hsmem_col].astype(str).str.strip().str.upper() == "HSMEM") & \
                                (group[hsbs_col].astype(str).str.strip().str.upper() == "BS")
                        
                        cond2 = (group[hsmem_col].astype(str).str.strip().str.upper() == "HSCUNG") & \
                                (group[tcbt_col].astype(str).str.strip() == "Khac TCBT D99")
                        
                        bvdr_metrics[cb_str] = {
                            bvdr_col_name: int((cond1 | cond2).sum())
                        }

            # --- TẠO BẢNG PIVOT TỔNG HỢP ---
            def is_not_blank(series):
                return series.notna() & (series.astype(str).str.strip() != "") & (series.astype(str).str.lower() != "nan")

            has_totrinh = is_not_blank(df_baocao[totrinh_col]) if totrinh_col else pd.Series(False, index=df_baocao.index)
            has_ngoaigiao = is_not_blank(df_baocao[ngoaigiao_col]) if ngoaigiao_col else pd.Series(False, index=df_baocao.index)
            valid_for_loai_hoso = ~(has_totrinh | has_ngoaigiao)

            pivot_data = []
            grouped = df_baocao.groupby(row_col, dropna=False)
            final_display_columns = []

            # Tên cột tính Tỷ lệ % Đúng hạn
            pct_dunghan_col = "% Thời gian GQ đúng hạn"

            # Biến lưu tổng đúng hạn và tổng số hồ sơ toàn bộ để tính dòng TỔNG CỘNG
            total_dung_han_count = 0
            total_all_hoso_count = 0

            for name, group in grouped:
                display_name = str(name).strip() if pd.notna(name) and str(name).strip() != "" else "(Blank)"
                row_dict = {row_col: display_name}

                # 1. Các cột chọn từ Sheet Baocao
                for c in val_cols:
                    if loai_hoso_col and c == loai_hoso_col:
                        group_indices = group.index
                        valid_group_mask = valid_for_loai_hoso.loc[group_indices]
                        mapped_series = group.loc[valid_group_mask, c + "_mapped"]

                        c_ngoai, c_noi, c_tuvong = f"{c} [Ngoại trú]", f"{c} [Nội trú]", f"{c} [Tử vong]"
                        row_dict[c_ngoai] = int((mapped_series == "Ngoại trú").sum())
                        row_dict[c_noi] = int((mapped_series == "Nội trú").sum())
                        row_dict[c_tuvong] = int((mapped_series == "Tử vong").sum())

                        for sub_c in [c_ngoai, c_noi, c_tuvong]:
                            if sub_c not in final_display_columns:
                                final_display_columns.append(sub_c)

                    elif c == c37_col_name:
                        # Đếm số lượng giá trị C37 trong cột C37/CI
                        is_c37 = group[c].astype(str).str.strip().str.upper() == "C37"
                        row_dict[c] = int(is_c37.sum())
                        if c not in final_display_columns:
                            final_display_columns.append(c)
                    else:
                        row_dict[c] = int(is_not_blank(group[c]).sum())
                        if c not in final_display_columns:
                            final_display_columns.append(c)

                # TÍNH % THỜI GIAN GQ ĐÚNG HẠN DÀNH CHO CÁN BỘ
                if dung_han_col:
                    dung_han_series = group[dung_han_col].astype(str).str.strip().str.lower()
                    is_dung_han = dung_han_series.str.contains("đúng", na=False) | dung_han_series.str.contains("dung", na=False)
                    valid_hoso = is_not_blank(group[dung_han_col])

                    count_dung_han = is_dung_han.sum()
                    count_total = valid_hoso.sum() if valid_hoso.sum() > 0 else len(group)

                    total_dung_han_count += count_dung_han
                    total_all_hoso_count += count_total

                    rate = (count_dung_han / count_total * 100) if count_total > 0 else 0.0
                    row_dict[pct_dunghan_col] = f"{rate:.1f}%"
                else:
                    row_dict[pct_dunghan_col] = "N/A"

                if pct_dunghan_col not in final_display_columns:
                    final_display_columns.append(pct_dunghan_col)

                # 2. Ghép chỉ tiêu đọc từ Sheet KHCN (đối chiếu qua tên Cán bộ)
                kh_data = khcn_metrics.get(display_name, {"KHCN/ATSK": 0, "PA/WCI": 0, "Du lịch": 0})
                for kh_col in ["KHCN/ATSK", "PA/WCI", "Du lịch"]:
                    row_dict[kh_col] = kh_data.get(kh_col, 0)
                    if kh_col not in final_display_columns:
                        final_display_columns.append(kh_col)

                # 3. Ghép chỉ tiêu đọc từ Sheet BVDR B1 (đối chiếu qua tên Cán bộ)
                bv_data = bvdr_metrics.get(display_name, {bvdr_col_name: 0})
                row_dict[bvdr_col_name] = bv_data.get(bvdr_col_name, 0)
                if bvdr_col_name not in final_display_columns:
                    final_display_columns.append(bvdr_col_name)

                pivot_data.append(row_dict)

            pivot_df = pd.DataFrame(pivot_data)
            
            # Đảm bảo chỉ có DUY NHẤT 1 cột row_col ở đầu bảng
            ordered_cols = [row_col] + [c for c in final_display_columns if c != row_col]
            pivot_df = pivot_df[ordered_cols]

            # Dòng TỔNG CỘNG ở cuối bảng
            total_row = {row_col: "--- TỔNG CỘNG ---"}
            for c in final_display_columns:
                if c == pct_dunghan_col:
                    if total_all_hoso_count > 0:
                        overall_rate = (total_dung_han_count / total_all_hoso_count * 100)
                        total_row[c] = f"{overall_rate:.1f}%"
                    else:
                        total_row[c] = "N/A"
                else:
                    total_row[c] = pivot_df[c].sum()

            display_df = pd.concat([pivot_df, pd.DataFrame([total_row])], ignore_index=True)

            # --- HIỂN THỊ BẢNG PIVOT ---
            st.subheader("1. Bảng tổng hợp")
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            st.write("---")

            # --- PHẦN DRILL-DOWN CHI TIẾT ---
            st.subheader("2. Xem chi tiết hồ sơ")
            st.caption("💡 *Chọn Cán bộ và Cột chỉ tiêu để xem danh sách hồ sơ chi tiết tương ứng từ File nguồn.*")

            col_select1, col_select2 = st.columns(2)

            with col_select1:
                unique_persons = pivot_df[row_col].unique().tolist()
                selected_person = st.selectbox(f"🎯 Chọn `{row_col}`:", unique_persons)

            with col_select2:
                selected_target_col = st.selectbox("📌 Chọn Cột chỉ tiêu muốn xem chi tiết:", final_display_columns)

            # Lọc dữ liệu chi tiết khi người dùng chọn
            if selected_person and selected_target_col:
                detail_df = pd.DataFrame()
                source_sheet_name = ""

                # CASE 1: CÁC CỘT THUỘC SHEET KHCN
                if selected_target_col in ["KHCN/ATSK", "PA/WCI", "Du lịch"]:
                    source_sheet_name = "KHCN"
                    if not df_khcn.empty and khcn_cb_col:
                        person_mask = df_khcn[khcn_cb_col].astype(str).str.strip() == str(selected_person).strip()
                        nghiepvu_col = next((c for c in df_khcn.columns if "nghiệp vụ" in c.lower() or "nghiep vu" in c.lower()), None)

                        if nghiepvu_col:
                            nv_series = df_khcn[nghiepvu_col].astype(str).str.strip().str.upper()
                            if selected_target_col == "KHCN/ATSK":
                                target_mask = nv_series.isin(["KHN", "ATS"])
                            elif selected_target_col == "PA/WCI":
                                target_mask = nv_series.isin(["CPA", "WCI", "TNCN.HSP", "TNCN.GVP", "PAI"])
                            else:
                                target_mask = nv_series.isin(["DQT", "DLVN", "FLE", "YDL", "DTN", "NND"])
                            
                            detail_df = df_khcn[person_mask & target_mask]

                # CASE 2: CỘT THUỘC SHEET BVDR B1
                elif selected_target_col == bvdr_col_name:
                    source_sheet_name = "BVDR B1"
                    if not df_bvdr.empty and bvdr_cb_col:
                        person_mask = df_bvdr[bvdr_cb_col].astype(str).str.strip() == str(selected_person).strip()
                        hsmem_col = next((c for c in df_bvdr.columns if "hsmem" in c.lower()), None)
                        hsbs_col = next((c for c in df_bvdr.columns if "hsbs" in c.lower()), None)
                        tcbt_col = next((c for c in df_bvdr.columns if "tcbt" in c.lower()), None)

                        if hsmem_col and hsbs_col and tcbt_col:
                            cond1 = (df_bvdr[hsmem_col].astype(str).str.strip().str.upper() == "HSMEM") & \
                                    (df_bvdr[hsbs_col].astype(str).str.strip().str.upper() == "BS")
                            
                            cond2 = (df_bvdr[hsmem_col].astype(str).str.strip().str.upper() == "HSCUNG") & \
                                    (df_bvdr[tcbt_col].astype(str).str.strip() == "Khac TCBT D99")
                            
                            detail_df = df_bvdr[person_mask & (cond1 | cond2)]

                # CASE 3: CÁC CỘT THUỘC SHEET BAOCAO
                else:
                    source_sheet_name = "Baocao"
                    if selected_person == "(Blank)":
                        person_mask = df_baocao[row_col].isna() | (df_baocao[row_col].astype(str).str.strip() == "")
                    else:
                        person_mask = df_baocao[row_col].astype(str).str.strip() == str(selected_person).strip()

                    if "[Ngoại trú]" in selected_target_col:
                        type_mask = (df_baocao[loai_hoso_col + "_mapped"] == "Ngoại trú") & valid_for_loai_hoso
                        detail_df = df_baocao[person_mask & type_mask]
                    elif "[Nội trú]" in selected_target_col:
                        type_mask = (df_baocao[loai_hoso_col + "_mapped"] == "Nội trú") & valid_for_loai_hoso
                        detail_df = df_baocao[person_mask & type_mask]
                    elif "[Tử vong]" in selected_target_col:
                        type_mask = (df_baocao[loai_hoso_col + "_mapped"] == "Tử vong") & valid_for_loai_hoso
                        detail_df = df_baocao[person_mask & type_mask]
                    elif selected_target_col == c37_col_name:
                        c37_mask = df_baocao[c37_col_name].astype(str).str.strip().str.upper() == "C37"
                        detail_df = df_baocao[person_mask & c37_mask]
                    elif selected_target_col == pct_dunghan_col:
                        if dung_han_col:
                            dung_han_series = df_baocao[dung_han_col].astype(str).str.strip().str.lower()
                            dung_mask = dung_han_series.str.contains("đúng", na=False) | dung_han_series.str.contains("dung", na=False)
                            detail_df = df_baocao[person_mask & dung_mask]
                        else:
                            detail_df = df_baocao[person_mask]
                    else:
                        col_mask = is_not_blank(df_baocao[selected_target_col])
                        detail_df = df_baocao[person_mask & col_mask]

                st.success(
                    f"📋 Kết quả (Trích xuất từ Sheet **{source_sheet_name}**): Tìm thấy **{len(detail_df):,}** hồ sơ cho **{row_col}** = `{selected_person}` tại chỉ tiêu **{selected_target_col}**:"
                )

                st.dataframe(detail_df, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {e}")

else:
    st.info("👋 Chào bạn! Vui lòng tải file Excel (.xlsx) ở thanh bên trái để bắt đầu.")