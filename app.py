import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Container Tracking",
    page_icon="📦",
    layout="wide"
)

# ---------------------------
# TITLE
# ---------------------------

st.title("📦 Container Tracking System")

st.write(
    "Upload the container Excel file and search by container number."
)

st.divider()

# ---------------------------
# FILE UPLOAD
# ---------------------------

uploaded_file = st.file_uploader(
    "Upload Container Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:

        df = pd.read_excel(
            uploaded_file,
            dtype=str
        )

        # Clean column names
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.upper()
        )

        # Clean data
        for column in df.columns:
            df[column] = (
                df[column]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        st.success(
            f"Excel loaded successfully - {len(df)} containers"
        )

        # ---------------------------
        # CHECK CONTAINER COLUMN
        # ---------------------------

        if "CONTAINER" not in df.columns:

            st.error(
                "CONTAINER column was not found in the Excel file."
            )

            st.write("Excel columns found:")

            st.write(list(df.columns))

        else:

            # Normalize container numbers
            df["CONTAINER"] = (
                df["CONTAINER"]
                .str.upper()
                .str.replace(" ", "", regex=False)
            )

            st.divider()

            # ---------------------------
            # SEARCH
            # ---------------------------

            container_number = st.text_input(
                "🔎 Container Number",
                placeholder="Example: MRKU1234567"
            )

            container_number = (
                container_number
                .upper()
                .replace(" ", "")
                .strip()
            )

            if container_number:

                result = df[
                    df["CONTAINER"] == container_number
                ]

                if result.empty:

                    st.error(
                        f"❌ Container {container_number} not found!"
                    )

                else:

                    container = result.iloc[0]

                    # ---------------------------
                    # FOUND
                    # ---------------------------

                    st.success(
                        "✅ CONTAINER FOUND"
                    )

                    st.header(container_number)

                    # ---------------------------
                    # LINE
                    # ---------------------------

                    if "LINE" in df.columns:

                        line = container["LINE"]

                        st.subheader("Shipping Line")

                        st.markdown(
                            f"# 🚢 {line}"
                        )

                    st.divider()

                    # ---------------------------
                    # MAIN INFO
                    # ---------------------------

                    col1, col2, col3 = st.columns(3)

                    if "SIZE_TYPE" in df.columns:

                        col1.metric(
                            "Size / Type",
                            container["SIZE_TYPE"]
                        )

                    if "STATUS" in df.columns:

                        col2.metric(
                            "Status",
                            container["STATUS"]
                        )

                    if "LOCATION" in df.columns:

                        col3.metric(
                            "Location",
                            container["LOCATION"]
                        )

                    st.divider()

                    # ---------------------------
                    # ALL INFORMATION
                    # ---------------------------

                    st.subheader(
                        "Container Information"
                    )

                    information = []

                    for column in df.columns:

                        information.append(
                            {
                                "FIELD": column,
                                "VALUE": container[column]
                            }
                        )

                    info_df = pd.DataFrame(information)

                    st.dataframe(
                        info_df,
                        hide_index=True,
                        use_container_width=True
                    )

    except Exception as error:

        st.error(
            "The Excel file could not be read."
        )

        st.exception(error)

else:

    st.info(
        "Please upload the container Excel file."
    )
