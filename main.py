import streamlit as st
import xarray as xr
import pandas as pd
import zipfile

from base64 import b64encode
from io import BytesIO

def convert(files, drop_nan, progress_bar):
    outputs = []
    progress_count = 0

    for file in files:
        try:
            file_bytes = BytesIO(file.getvalue())

            with xr.open_dataset(file_bytes) as ds:
                df = ds.to_dataframe()
                buffer = BytesIO()
                buffer.seek(0)

                if drop_nan:
                    df = df.dropna(how='all')

                st.write(df.head(10))

                status_text = st.empty()
                total_rows = len(df)

                with pd.ExcelWriter(buffer, engine='openpyxl') as writer: # type: ignore
                    chunk_size = max(1, total_rows // 100)

                    for i in range(0, total_rows, chunk_size):
                        chunk = df.iloc[i:min(i + chunk_size, total_rows)]

                        if i == 0:
                            chunk.to_excel(writer, index=True)
                        else:
                            chunk.to_excel(writer, index=True, startrow=i+1, header=False)

                        progress_percentage = min(100, int((i + chunk_size) / total_rows * 100))
                        status_text.text(f"Converting {file.name} to Excel: {progress_percentage}%")

                buffer.seek(0)
                outputs.append((file.name.replace('.nc', '.xlsx'), buffer))

                progress_count += 1
                progress_bar.progress(progress_count / len(files), f"Processing {file.name}")

        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
            return None

    if outputs:
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for filename, excel_buffer in outputs:
                zip_file.writestr(filename, excel_buffer.getvalue())

        return zip_buffer
    else:
        st.error("No data could be converted to Excel files")
        return None

def download(buffer):
    buffer.seek(0)
    b64 = b64encode(buffer.read()).decode()
    return f'<a href="data:application/zip;base64,{b64}" download="converted_netcdf.zip">Download ZIP file</a>'

def main():
    st.title("NetCDF to Excel Converter")

    files = st.file_uploader(
        "Upload NetCDF files",
        type='.nc',
        accept_multiple_files=True
    )

    drop_nan = st.checkbox("Drop rows with all NaN values", value=False)

    if files:
        st.write(f"Number of files uploaded: {len(files)}")
        if st.button("Convert to Excel"):
            progress_bar = st.progress(0, "Starting conversion...")
            buffer = convert(files, drop_nan, progress_bar)

            if buffer:
                st.markdown(download(buffer), unsafe_allow_html=True)
                st.success("Conversion completed! Click the link above to download.")

            progress_bar.progress(1.0, "Conversion complete!")

if __name__ == "__main__":
    main()
