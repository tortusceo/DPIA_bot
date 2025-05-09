import streamlit as st
import os
import tempfile
from dpia_bot import (
    load_google_api_key,
    extract_text,
    convert_file_to_markdown,
    construct_prompt_for_gemini, # New prompt construction function
    generate_dpia_from_prompt,
    convert_markdown_to_original_format,
    get_file_extension,
    REFERENCE_DPIA_FILENAME
)

st.set_page_config(page_title="DPIA Bot", layout="wide")

st.title("DPIA Bot 🤖")
st.markdown("Upload a customer's DPIA template (e.g., .docx or .xlsx), and the bot will attempt to fill it out using a reference DPIA.")

# --- File Paths --- 
# Ensure REFERENCE_DPIA_FILENAME is accessible. 
# For Streamlit Cloud, this file would need to be part of your deployed app repository.
if not os.path.exists(REFERENCE_DPIA_FILENAME):
    st.error(f"Critical Error: Reference DPIA file '{REFERENCE_DPIA_FILENAME}' not found. The application cannot proceed.")
    st.stop() # Stop the app if the reference DPIA is missing

# --- File Upload --- 
uploaded_customer_file = st.file_uploader(
    "Upload Customer DPIA Template (DOCX or XLSX)", 
    type=["docx", "xlsx"],
    help="Select the customer's DPIA template file you want to process."
)

if uploaded_customer_file is not None:
    st.info(f"File uploaded: {uploaded_customer_file.name} (Size: {uploaded_customer_file.size} bytes)")

    if st.button("🚀 Process DPIA", help="Click to start processing the uploaded DPIA template."):
        # Create a temporary directory to store uploaded and generated files
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_customer_filepath = os.path.join(temp_dir, uploaded_customer_file.name)

            # Save uploaded file to the temporary path
            with open(temp_customer_filepath, "wb") as f:
                f.write(uploaded_customer_file.getbuffer())
            # st.write(f"Uploaded file saved temporarily to: {temp_customer_filepath}") # For debugging

            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                status_text.info("Loading API Key...")
                api_key = load_google_api_key()
                progress_bar.progress(5)

                status_text.info(f"Loading Reference DPIA ('{REFERENCE_DPIA_FILENAME}')...")
                reference_dpia_text = extract_text(REFERENCE_DPIA_FILENAME)
                if not reference_dpia_text:
                    st.error(f"Could not load reference DPIA content from {REFERENCE_DPIA_FILENAME}.")
                    st.stop()
                progress_bar.progress(10)

                customer_original_file_extension = get_file_extension(temp_customer_filepath)
                status_text.info(f"Converting customer template ('{uploaded_customer_file.name}') to Markdown...")
                customer_template_markdown = convert_file_to_markdown(temp_customer_filepath)
                if not customer_template_markdown or "Error during" in customer_template_markdown:
                    st.error(f"Could not convert customer file '{uploaded_customer_file.name}' to Markdown.")
                    st.stop()
                progress_bar.progress(25)

                status_text.info("Constructing prompt for AI model...")
                prompt_text_for_gemini = construct_prompt_for_gemini(reference_dpia_text, customer_template_markdown)
                progress_bar.progress(30)

                status_text.info("Calling AI model (Gemini)... This may take a few minutes.")
                new_dpia_markdown = generate_dpia_from_prompt(prompt_text_for_gemini, api_key)
                if "Error during Gemini API call" in new_dpia_markdown or "Error: Gemini API returned no text" in new_dpia_markdown:
                    st.error(f"Failed to generate DPIA from AI model: {new_dpia_markdown}")
                    st.stop()
                progress_bar.progress(75)

                status_text.info("Converting AI-generated Markdown back to original format...")
                output_filename_base = os.path.splitext(uploaded_customer_file.name)[0]
                prepared_doc_filepath_temp = os.path.join(temp_dir, f"{output_filename_base}_completed{customer_original_file_extension}")

                final_output_path_temp = convert_markdown_to_original_format(
                    new_dpia_markdown,
                    customer_original_file_extension,
                    prepared_doc_filepath_temp,
                    style_reference_filepath=temp_customer_filepath
                )
                progress_bar.progress(90)

                if "_conversion_error.txt" in os.path.basename(final_output_path_temp).lower(): # Check basename for safety
                    st.error(f"Error converting Markdown to the original format.")
                    with open(final_output_path_temp, "rb") as fp:
                        st.download_button(
                            label="Download Conversion Error Log (TXT)",
                            data=fp,
                            file_name=os.path.basename(final_output_path_temp),
                            mime="text/plain"
                        )
                    st.stop()
                else:
                    status_text.success("DPIA processed successfully!")
                    st.balloons()
                    with open(final_output_path_temp, "rb") as fp:
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if customer_original_file_extension == ".docx" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        if customer_original_file_extension not in [".docx", ".xlsx"]:
                             mime_type = "application/octet-stream" # Fallback for other types like .md, .txt

                        st.download_button(
                            label=f"Download Completed DPIA ({os.path.basename(final_output_path_temp)})",
                            data=fp,
                            file_name=os.path.basename(final_output_path_temp),
                            mime=mime_type
                        )
                    progress_bar.progress(100)

            except Exception as e:
                st.error(f"An unexpected error occurred during processing.")
                st.exception(e) # Displays the full traceback in the Streamlit app for debugging
                if progress_bar is not None: progress_bar.empty()
                if status_text is not None: status_text.empty()
            finally:
                # The problematic lines that caused the TypeError are removed.
                # Cleanup for unexpected errors is handled in the except block.
                # On success, we want the success messages and progress bar to persist.
                pass # finally block is now empty or can be removed if not needed for other resources

# Add some instructions or information at the bottom
st.markdown("---")
st.markdown("**Instructions:**")
st.markdown("1. Ensure your `reference_dpia.txt` file is in the same directory as this application.")
st.markdown("2. Ensure your `.env` file with the `GOOGLE_API_KEY` is present for local execution.")
st.markdown("3. Upload the customer's DPIA template file.")
st.markdown("4. Click 'Process DPIA'.")
st.markdown("5. Download the completed document.") 