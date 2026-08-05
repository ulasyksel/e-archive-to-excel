import streamlit as st

from excel_exporter import build_excel
from processor import count_pdfs, process_files


st.set_page_config(
    page_title="E-Arşiv Fatura Dökümü",
    layout="wide",
)

st.title("E-Arşiv Fatura Dökümü")
st.write(
    "PDF veya ZIP dosyalarını yükleyin. "
    "Uygulama fatura bilgilerini okuyup Excel dökümü oluşturur."
)

uploaded_files = st.file_uploader(
    "Dosyaları seçin",
    type=["pdf", "zip"],
    accept_multiple_files=True,
)

if uploaded_files:
    try:
        pdf_count = count_pdfs(uploaded_files)
        st.success(f"{pdf_count} PDF yüklendi.")
    except Exception as error:
        pdf_count = 0
        st.error(f"Dosyalar okunamadı: {error}")

    if pdf_count > 0 and st.button("İşlemi Başlat", type="primary"):
        progress = st.progress(0, text="Faturalar hazırlanıyor...")

        def update_progress(current, total):
            ratio = current / total if total else 1
            progress.progress(
                ratio,
                text=f"İşlenen PDF: {current} / {total}",
            )

        try:
            main_df, control_df, error_df = process_files(
                uploaded_files,
                progress_callback=update_progress,
            )

            excel_bytes = build_excel(
                main_df=main_df,
                control_df=control_df,
                error_df=error_df,
                total_pdf=pdf_count,
            )

            st.session_state["main_df"] = main_df
            st.session_state["control_df"] = control_df
            st.session_state["error_df"] = error_df
            st.session_state["excel_bytes"] = excel_bytes
            st.session_state["processed_pdf_count"] = pdf_count

            st.success("Fatura işleme ve Excel oluşturma tamamlandı.")

        except Exception as error:
            st.error(f"İşlem tamamlanamadı: {error}")

        finally:
            progress.empty()

else:
    st.info("İşleme başlamak için PDF veya ZIP yükleyin.")


if "main_df" in st.session_state:
    main_df = st.session_state["main_df"]
    control_df = st.session_state["control_df"]
    error_df = st.session_state["error_df"]
    excel_bytes = st.session_state["excel_bytes"]

    control_count = (
        int((control_df["DURUM"] == "KONTROL").sum())
        if not control_df.empty
        else 0
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("Ödeme satırı", len(main_df))
    col2.metric("Fatura", len(control_df))
    col3.metric("Kontrol gereken", control_count)

    st.download_button(
        label="Excel'i İndir",
        data=excel_bytes,
        file_name="E_ARSIV_MAGAZA_FATURA_DOKUM.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )

    summary_tab, control_tab, error_tab = st.tabs(
        ["E-Fatura Özet", "Kontrol", "Okuma Hataları"]
    )

    with summary_tab:
        st.dataframe(
            main_df,
            width="stretch",
            hide_index=True,
        )

    with control_tab:
        st.dataframe(
            control_df,
            width="stretch",
            hide_index=True,
        )

    with error_tab:
        if error_df.empty:
            st.success("Okuma hatası bulunmadı.")
        else:
            st.dataframe(
                error_df,
                width="stretch",
                hide_index=True,
            )
