import streamlit as st
#import pythoncom
import docx
import time
from docx2pdf import convert
import os

# Initialize COM
#pythoncom.CoInitialize()
st.set_page_config(page_title='Docx to PDF', page_icon='📄')
st.title('Multiple Word to PDF Converter 📄')

# File upload section for multiple DOCX files
docx_files = st.file_uploader('Word file upload kara kaam chor 🤣👌', type=['docx'], accept_multiple_files=True)

# Output folder for PDF files
output_folder = "output_pdf"

if not os.path.exists(output_folder):
    os.makedirs(output_folder)
st.snow()
# Convert and display result for each uploaded DOCX file
if docx_files:
    for docx_file in docx_files:
        try:
            # Save the uploaded Word file using a unique filename
            uploaded_file_path = os.path.join(output_folder, f"uploaded_file_{int(time.time())}.docx")
            with open(uploaded_file_path, "wb") as f:
                f.write(docx_file.read())

            # Convert Word to PDF
            convert(uploaded_file_path)

            # Display success message and offer PDF download
            pdf_file = os.path.join(output_folder, f"output_{os.path.splitext(os.path.basename(uploaded_file_path))[0]}.pdf")
            st.success(f"Conversion successful for {os.path.basename(uploaded_file_path)}. Download your PDF:")
            st.balloons()
            st.download_button(label=f"Download PDF ({os.path.basename(uploaded_file_path)})", data=pdf_file, file_name=os.path.basename(pdf_file))

        except Exception as e:
            st.error(f"An error occurred during conversion for {os.path.basename(uploaded_file_path)}: {e}")

        finally:
            # Remove the uploaded Word file after conversion
            os.remove(uploaded_file_path)

else:
    # Display a warning if no file is uploaded
    st.warning("Word file upload nhi kiya ?")

audiofile = open('audio.mp3','rb')
st.audio(audiofile.read())
videourl = open('Video.mp4','rb')
st.video(videourl)


