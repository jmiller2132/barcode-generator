import streamlit as st
import pandas as pd
import io
import zipfile
from PIL import Image
import barcode
from barcode.writer import ImageWriter
import segno
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
import tempfile
import os

# Page configuration
st.set_page_config(
    page_title="Barcode Generator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Barcode Generator")
st.markdown("Upload an Excel or CSV file and convert a column into barcodes")

# Initialize session state
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None
if 'df' not in st.session_state:
    st.session_state.df = None

# Sidebar configuration
with st.sidebar:
    st.header("File Settings")
    has_header = st.checkbox(
        "Excel file has a header row",
        value=True,
        help="If unchecked, columns will be named 'Column 1', 'Column 2', etc."
    )
    
    st.header("Barcode Settings")
    enable_checksum = st.checkbox(
        "Enable Checksum (Check Digit)",
        value=False,
        help="A checksum adds a mathematical digit to the end of your barcode to prevent mis-scans. Turn this OFF if you want the scanned result to match your Excel data exactly. Turn this ON for high-security or shipping labels."
    )
    show_text = st.checkbox(
        "Show Human-Readable Text",
        value=True,
        help="Turn this off to generate barcodes without the numbers/text underneath."
    )
    
    # User Guide at the bottom
    with st.expander("📖 User Guide & Help"):
        st.markdown("""
#### 1. Prepare Your Excel/CSV
* **Text Formatting:** Format your barcode column as **'Text'** in Excel before saving. If it's a 'Number', long IDs may turn into scientific notation (like 1.0E+10), which breaks the barcode.
* **Headers:** Uncheck 'File has header row' if your data starts on Line 1.

#### 2. The Checksum (Check Digit)
* **OFF (Recommended):** Use this so the scanned result matches your Excel data exactly.
* **ON:** Use for high-security or shipping labels where a mathematical 'check digit' is required.
* *Note: EAN and UPC codes always require a checksum.*

#### 3. Export Options
* **ZIP:** Best for importing into labeling software (Zebra, Bartender).
* **Excel:** Best for scannable inventory lists.
* **PDF:** Best for immediate printing on Avery 5160 sheets.
        """)


def clean_barcode_value(value):
    """Clean barcode value by stripping '.0' from the end if it's a string.
    
    This prevents Excel's 'Number' format from turning '101' into '101.0'.
    """
    value_str = str(value).strip()
    if value_str.endswith('.0'):
        value_str = value_str[:-2]
    return value_str


def validate_numeric_data(data, symbology):
    """Validate that data is numeric for EAN-13 and UPC-A symbologies."""
    if symbology in ['EAN-13', 'UPC-A']:
        invalid_rows = []
        for idx, value in enumerate(data):
            try:
                # Check if value is numeric (after cleaning)
                cleaned_val = clean_barcode_value(value)
                if not cleaned_val.isdigit():
                    invalid_rows.append((idx + 1, value))
            except:
                invalid_rows.append((idx + 1, value))
        return invalid_rows
    return []


def generate_barcode_image(value, symbology, add_checksum=False, show_text=True):
    """Generate a barcode image in memory.
    
    Args:
        value: The value to encode in the barcode
        symbology: The barcode type
        add_checksum: Whether to add a checksum (for Code 128 and Code 39 only)
        show_text: Whether to show human-readable text below the barcode
    """
    try:
        # Clean the value (strip '.0' if present)
        value_str = clean_barcode_value(value)
        
        if symbology == 'QR Code':
            qr = segno.make(value_str, error='M')
            img_buffer = io.BytesIO()
            qr.save(img_buffer, kind='png', scale=5, border=2)
            img_buffer.seek(0)
            return Image.open(img_buffer)
        
        elif symbology == 'Data Matrix':
            dm = segno.make(value_str, error='M')
            img_buffer = io.BytesIO()
            dm.save(img_buffer, kind='png', scale=5, border=2)
            img_buffer.seek(0)
            return Image.open(img_buffer)
        
        else:
            # 1D barcodes using python-barcode
            barcode_class_map = {
                'Code 128': barcode.get_barcode_class('code128'),
                'Code 39': barcode.get_barcode_class('code39'),
                'EAN-13': barcode.get_barcode_class('ean13'),
                'UPC-A': barcode.get_barcode_class('upca'),
            }
            
            barcode_class = barcode_class_map.get(symbology)
            if not barcode_class:
                return None
            
            # EAN-13 and UPC-A always include checksum (mandatory in standard)
            # Code 128 and Code 39 can have checksum toggled
            if symbology in ['Code 128', 'Code 39']:
                code = barcode_class(value_str, writer=ImageWriter(), add_checksum=add_checksum)
            else:
                # EAN-13 and UPC-A always have checksum
                code = barcode_class(value_str, writer=ImageWriter())
            
            img_buffer = io.BytesIO()
            # Control text visibility via font_size option (0 = hide, >0 = show)
            options = {'font_size': 0 if not show_text else 10}
            code.write(img_buffer, options=options)
            img_buffer.seek(0)
            return Image.open(img_buffer)
    
    except Exception as e:
        st.error(f"Error generating barcode for value '{value}': {str(e)}")
        return None


def create_excel_with_barcodes(df, selected_column, symbology, add_checksum=False, show_text=True, barcode_column='Barcode'):
    """Create an Excel file with embedded barcode images."""
    output = io.BytesIO()
    temp_files = []  # Collect temp file paths to clean up after workbook closes
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Create a copy of the dataframe and add barcode column
            df_export = df.copy()
            df_export[barcode_column] = ''  # Add empty barcode column
            df_export.to_excel(writer, sheet_name='Sheet1', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Get barcode column index
            barcode_col_idx = len(df.columns)  # Index of the barcode column (0-indexed)
            
            # Set column width for barcode column (30 units)
            worksheet.set_column(barcode_col_idx, barcode_col_idx, 30)
            
            # Set row height for data rows (75 points)
            for row_num in range(1, len(df_export) + 1):
                worksheet.set_row(row_num, 75)
            
            # Generate barcodes and insert images
            for idx, value in enumerate(df[selected_column]):
                barcode_img = generate_barcode_image(value, symbology, add_checksum, show_text)
                if barcode_img:
                    # Save image to temporary file
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
                        barcode_img.save(tmp_file.name, 'PNG')
                        tmp_path = tmp_file.name
                        temp_files.append(tmp_path)
                        
                        # Calculate position (Excel is 0-indexed for insert_image)
                        row = idx + 1  # +1 for header row
                        col = barcode_col_idx
                        
                        # Insert image with scaling to fit cell
                        worksheet.insert_image(row, col, tmp_path, {
                            'x_scale': 0.6,
                            'y_scale': 0.6,
                            'x_offset': 5,
                            'y_offset': 5
                        })
    finally:
        # Clean up temp files after workbook is closed (or on error)
        for tmp_path in temp_files:
            try:
                os.unlink(tmp_path)
            except:
                pass
    
    output.seek(0)
    return output


def create_pdf_label_sheet(df, selected_column, symbology, add_checksum=False, show_text=True):
    """Create a PDF with barcodes in Avery 5160 style (3 columns, 10 rows).
    
    Note: Uses original values (not checksummed) for text display.
    """
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=letter)
    
    page_width, page_height = letter
    margin_left = 0.2 * inch
    margin_top = 0.5 * inch
    label_width = 2.625 * inch
    label_height = 1 * inch
    cols = 3
    rows = 10
    
    total_labels = 0
    max_labels_per_page = cols * rows
    
    for idx, value in enumerate(df[selected_column]):
        if total_labels % max_labels_per_page == 0 and total_labels > 0:
            c.showPage()  # New page
        
        # Calculate position
        page_pos = total_labels % max_labels_per_page
        col_pos = page_pos % cols
        row_pos = page_pos // cols
        
        x = margin_left + col_pos * label_width
        y = page_height - margin_top - (row_pos + 1) * label_height
        
        # Store original value for text display (not checksummed)
        original_value = value
        
        # Generate barcode image (may include checksum)
        barcode_img = generate_barcode_image(value, symbology, add_checksum, show_text)
        if barcode_img:
            # Calculate target size in points (for ReportLab)
            max_width_pts = label_width - 0.2 * inch
            max_height_pts = label_height * 0.7
            
            # Get image dimensions in pixels
            img_width_px, img_height_px = barcode_img.size
            aspect_ratio = img_width_px / img_height_px
            
            # Calculate target dimensions in points, maintaining aspect ratio
            if img_width_px / img_height_px > max_width_pts / max_height_pts:
                target_width_pts = max_width_pts
                target_height_pts = max_width_pts / aspect_ratio
            else:
                target_height_pts = max_height_pts
                target_width_pts = max_height_pts * aspect_ratio
            
            # Save to temp file for ReportLab
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
                barcode_img.save(tmp_file.name, 'PNG')
                
                # Draw barcode (ReportLab will scale it)
                img_reader = ImageReader(tmp_file.name)
                barcode_x = x + (label_width - target_width_pts) / 2
                barcode_y = y + label_height * 0.25
                c.drawImage(img_reader, barcode_x, barcode_y, width=target_width_pts, height=target_height_pts)
                
                # Draw text below barcode (use original value, not checksummed) if enabled
                # Maintain spacing even when text is hidden to keep label layout consistent
                if show_text:
                    text_y = y + 0.05 * inch
                    c.setFont("Helvetica", 8)
                    text_width = c.stringWidth(str(original_value), "Helvetica", 8)
                    text_x = x + (label_width - text_width) / 2
                    c.drawString(text_x, text_y, str(original_value))
                
                os.unlink(tmp_file.name)
        
        total_labels += 1
    
    c.save()
    output.seek(0)
    return output


def create_zip_of_pngs(df, selected_column, symbology, add_checksum=False, show_text=True):
    """Create a ZIP file containing PNG barcode images.
    
    Note: Uses original values (not checksummed) for filenames.
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for value in df[selected_column]:
            # Store original value for filename (not checksummed)
            original_value = value
            
            # Generate barcode image (may include checksum)
            barcode_img = generate_barcode_image(value, symbology, add_checksum, show_text)
            if barcode_img:
                img_buffer = io.BytesIO()
                barcode_img.save(img_buffer, format='PNG')
                img_buffer.seek(0)
                
                # Clean filename using original value (not checksummed)
                clean_value = str(original_value).strip().replace('/', '_').replace('\\', '_')
                filename = f"{clean_value}.png"
                zip_file.writestr(filename, img_buffer.read())
    
    zip_buffer.seek(0)
    return zip_buffer


# File upload
uploaded_file = st.file_uploader(
    "Upload Excel or CSV file",
    type=['xlsx', 'csv'],
    help="Supported formats: .xlsx, .csv"
)

if uploaded_file is not None:
    try:
        # Read file based on header setting
        if uploaded_file.name.endswith('.xlsx'):
            if has_header:
                df = pd.read_excel(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, header=None)
                # Rename columns to Column 1, Column 2, etc.
                df.columns = [f'Column {i+1}' for i in range(len(df.columns))]
        else:
            if has_header:
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_csv(uploaded_file, header=None)
                # Rename columns to Column 1, Column 2, etc.
                df.columns = [f'Column {i+1}' for i in range(len(df.columns))]
        
        st.session_state.df = df
        st.session_state.uploaded_data = uploaded_file
        
        # Display preview
        st.subheader("Data Preview")
        st.dataframe(df.head(10), use_container_width=True)
        
        # Column selection
        st.subheader("Configuration")
        col1, col2 = st.columns(2)
        
        with col1:
            selected_column = st.selectbox(
                "Select column for barcode generation",
                options=df.columns.tolist(),
                help="Choose the column that contains the data to convert to barcodes"
            )
        
        with col2:
            symbology_options = ['Code 128', 'Code 39', 'EAN-13', 'UPC-A', 'QR Code', 'Data Matrix']
            symbology = st.radio(
                "Select barcode symbology",
                options=symbology_options,
                help="Choose the type of barcode to generate"
            )
            
            # Show symbology description
            symbology_descriptions = {
                'Code 128': 'Code 128: High-density, compact barcode for packaging and shipping. Supports all ASCII.',
                'Code 39': 'Code 39: Alphanumeric code used in healthcare/electronics.',
                'EAN-13': 'EAN-13: 13-digit retail barcode for worldwide products (Numeric only).',
                'UPC-A': 'UPC-A: 12-digit standard barcode for US/Canada retail (Numeric only).',
                'QR Code': 'QR Code: 2D matrix for URLs or large data.',
                'Data Matrix': 'Data Matrix: Square 2D code for small spaces (electronics/healthcare).'
            }
            st.caption(f"📋 **About {symbology}:** {symbology_descriptions.get(symbology, '')}")
        
        # Export format selection
        export_format = st.radio(
            "Export format",
            options=['ZIP of PNGs', 'Excel (Embedded Images)', 'PDF (Label Sheet)'],
            horizontal=True
        )
        
        # Handle checksum logic for EAN-13 and UPC-A
        # These symbologies always require checksum (mandatory in standard)
        actual_checksum_setting = enable_checksum
        if symbology in ['EAN-13', 'UPC-A']:
            actual_checksum_setting = True  # Force checksum for these types
            if not enable_checksum:
                st.info(f"ℹ️ **Note:** {symbology} barcodes always include a checksum digit as part of the global standard. The barcode will be generated with a checksum even if the setting is disabled.")
        
        # Validation for numeric barcodes
        if symbology in ['EAN-13', 'UPC-A']:
            invalid_rows = validate_numeric_data(df[selected_column], symbology)
            if invalid_rows:
                error_msg = f"Invalid data found for {symbology} (must be numeric):\n"
                for row_num, value in invalid_rows[:10]:  # Show first 10 errors
                    error_msg += f"  - Row {row_num}: '{value}'\n"
                if len(invalid_rows) > 10:
                    error_msg += f"  ... and {len(invalid_rows) - 10} more rows"
                st.error(error_msg)
                st.stop()
        
        # Generate button
        if st.button("Generate Barcodes", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                status_text.text("Generating barcodes...")
                progress_bar.progress(20)
                
                if export_format == 'ZIP of PNGs':
                    status_text.text("Creating ZIP file...")
                    progress_bar.progress(60)
                    zip_buffer = create_zip_of_pngs(df, selected_column, symbology, actual_checksum_setting, show_text)
                    progress_bar.progress(100)
                    status_text.text("Complete!")
                    
                    st.success("Barcode ZIP file generated successfully!")
                    st.download_button(
                        label="Download ZIP file",
                        data=zip_buffer,
                        file_name=f"barcodes_{symbology.replace(' ', '_')}.zip",
                        mime="application/zip"
                    )
                
                elif export_format == 'Excel (Embedded Images)':
                    status_text.text("Creating Excel file with embedded images...")
                    progress_bar.progress(60)
                    excel_buffer = create_excel_with_barcodes(df, selected_column, symbology, actual_checksum_setting, show_text)
                    progress_bar.progress(100)
                    status_text.text("Complete!")
                    
                    st.success("Excel file with barcodes generated successfully!")
                    st.download_button(
                        label="Download Excel file",
                        data=excel_buffer,
                        file_name=f"barcodes_{symbology.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                elif export_format == 'PDF (Label Sheet)':
                    status_text.text("Creating PDF label sheet...")
                    progress_bar.progress(60)
                    pdf_buffer = create_pdf_label_sheet(df, selected_column, symbology, actual_checksum_setting, show_text)
                    progress_bar.progress(100)
                    status_text.text("Complete!")
                    
                    st.success("PDF label sheet generated successfully!")
                    st.download_button(
                        label="Download PDF file",
                        data=pdf_buffer,
                        file_name=f"barcodes_{symbology.replace(' ', '_')}.pdf",
                        mime="application/pdf"
                    )
            
            except Exception as e:
                st.error(f"Error generating barcodes: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
            finally:
                progress_bar.empty()
                status_text.empty()
    
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
