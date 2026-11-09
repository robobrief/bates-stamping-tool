import os
import argparse
import subprocess
import shutil
import tempfile
import uuid
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, DictionaryObject, ArrayObject, FloatObject, DecodedStreamObject, NumberObject, TextStringObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from io import BytesIO

STAMP_FONT = "Helvetica-Bold"
STAMP_ANNOTATION_AUTHOR = "BatesStamper"
STAMP_ANNOTATION_SUBJECT = "Bates Stamp"
STAMP_ANNOTATION_NAME_PREFIX = "BatesStamper-"
COLOR_PRESETS = {
    "black": {
        "label": "Black, no background",
        "text_color": "#000000",
        "background_box": False,
        "box_color": None,
        "use_case": "Most common"
    },
    "white": {
        "label": "White, no background",
        "text_color": "#FFFFFF",
        "background_box": False,
        "box_color": None,
        "use_case": "Dark PDFs"
    },
    "red": {
        "label": "Red, no background",
        "text_color": "#FF0000",
        "background_box": False,
        "box_color": None,
        "use_case": "Privileged / review"
    },
    "black-box": {
        "label": "Black on white",
        "text_color": "#000000",
        "background_box": True,
        "box_color": "#FFFFFF",
        "use_case": "Maximum readability"
    },
    "white-box": {
        "label": "White on black",
        "text_color": "#FFFFFF",
        "background_box": True,
        "box_color": "#000000",
        "use_case": "Dark drawings"
    },
    "red-box": {
        "label": "Red on white",
        "text_color": "#FF0000",
        "background_box": True,
        "box_color": "#FFFFFF",
        "use_case": "Review copies"
    },
    "reverse-red": {
        "label": "White on red",
        "text_color": "#FFFFFF",
        "background_box": True,
        "box_color": "#FF0000",
        "use_case": "Very visible"
    },
    "blue-box": {
        "label": "Blue on white",
        "text_color": "#0000FF",
        "background_box": True,
        "box_color": "#FFFFFF",
        "use_case": "Production sets"
    },
    "reverse-blue": {
        "label": "White on blue",
        "text_color": "#FFFFFF",
        "background_box": True,
        "box_color": "#0000FF",
        "use_case": "Internal numbering"
    },
    "highlight": {
        "label": "Black on yellow",
        "text_color": "#000000",
        "background_box": True,
        "box_color": "#FFFF00",
        "use_case": "High visibility"
    }
}

def get_text_metrics(text, font_name=STAMP_FONT, size=12):
    """Get accurate text width and height metrics for reportlab drawing."""
    text_width = pdfmetrics.stringWidth(text, font_name, size)
    ascent = pdfmetrics.getAscent(font_name, size)
    descent = pdfmetrics.getDescent(font_name, size)
    text_height = ascent - descent
    return text_width, text_height, ascent, descent

def parse_hex_color(color, fallback_rgb, warning_label):
    """Parse hex color to normalized RGB tuple (0..1), with fallback on invalid values."""
    try:
        parsed = HexColor(color)
        return (parsed.red, parsed.green, parsed.blue)
    except:
        print(f"Warning: Invalid {warning_label} '{color}', using fallback value instead")
        return fallback_rgb

def resolve_pdf_object(obj):
    """Resolve a PyPDF2 indirect object if needed."""
    if hasattr(obj, "get_object"):
        return obj.get_object()
    return obj

def to_roman(number):
    """Convert an integer to Roman numerals."""
    values = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
    ]
    result = []
    remaining = number
    for value, numeral in values:
        while remaining >= value:
            result.append(numeral)
            remaining -= value
    return "".join(result)

def to_alpha(number):
    """Convert a 1-based integer to alphabetical page notation (A, B... AA...)."""
    if number <= 0:
        return str(number)
    result = []
    current = number
    while current > 0:
        current -= 1
        result.append(chr(ord("A") + (current % 26)))
        current //= 26
    return "".join(reversed(result))

def format_page_label_number(number, style):
    """Format a page label number according to PDF page label style."""
    if style == "D":
        return str(number)
    if style == "R":
        return to_roman(number)
    if style == "r":
        return to_roman(number).lower()
    if style == "A":
        return to_alpha(number)
    if style == "a":
        return to_alpha(number).lower()
    return str(number)

def get_page_labels(reader):
    """Get page labels for each page from PDF /PageLabels, falling back to 1..N."""
    page_count = len(reader.pages)
    fallback = [str(index + 1) for index in range(page_count)]

    try:
        root = resolve_pdf_object(reader.trailer["/Root"])
        page_labels_tree = resolve_pdf_object(root.get("/PageLabels")) if root else None
        if not page_labels_tree:
            return fallback

        nums = resolve_pdf_object(page_labels_tree.get("/Nums"))
        if not nums:
            return fallback

        label_ranges = []
        for i in range(0, len(nums), 2):
            start_index = int(resolve_pdf_object(nums[i]))
            label_dict = resolve_pdf_object(nums[i + 1])
            label_ranges.append((start_index, label_dict))
        label_ranges.sort(key=lambda item: item[0])

        labels = []
        active_start = 0
        active_dict = {}
        range_cursor = 0

        for page_index in range(page_count):
            while range_cursor < len(label_ranges) and label_ranges[range_cursor][0] <= page_index:
                active_start, active_dict = label_ranges[range_cursor]
                range_cursor += 1

            prefix = str(active_dict.get("/P", ""))
            style_raw = active_dict.get("/S")
            style = str(style_raw).lstrip("/") if style_raw else None
            start_value = int(active_dict.get("/St", 1))
            sequence_value = start_value + (page_index - active_start)
            number_part = format_page_label_number(sequence_value, style) if style else ""
            label_text = f"{prefix}{number_part}"
            labels.append(label_text if label_text else str(page_index + 1))

        return labels
    except (KeyError, TypeError, ValueError, IndexError) as e:
        print(f"  -> Warning: Could not read PDF page labels ({e}); using sequential page numbers")
        return fallback

def copy_outlines(reader, writer):
    """Copy PDF outlines/bookmarks from reader to writer."""
    try:
        source_outline = reader.outline
    except Exception as e:
        print(f"  -> Warning: Could not read source bookmarks ({e})")
        return

    if not source_outline:
        return

    page_number_by_ref = {}
    for index, page in enumerate(reader.pages):
        reference = getattr(page, "indirect_reference", None)
        if reference and hasattr(reference, "idnum") and hasattr(reference, "generation"):
            page_number_by_ref[(reference.idnum, reference.generation)] = index

    def walk(items, parent=None):
        previous_bookmark = None
        for item in items:
            if isinstance(item, list):
                walk(item, previous_bookmark or parent)
                continue

            title = getattr(item, "title", None)
            if not title and hasattr(item, "get"):
                title = item.get("/Title")
            if not title:
                previous_bookmark = None
                continue

            page_number = None
            page_ref = getattr(item, "page", None)
            if page_ref and hasattr(page_ref, "idnum") and hasattr(page_ref, "generation"):
                page_number = page_number_by_ref.get((page_ref.idnum, page_ref.generation))
            try:
                if page_number is None:
                    page_number = reader.get_destination_page_number(item)
            except Exception:
                previous_bookmark = None
                continue

            try:
                previous_bookmark = writer.add_outline_item(str(title), int(page_number), parent=parent)
            except Exception as e:
                print(f"  -> Warning: Could not copy bookmark '{title}' ({e})")
                previous_bookmark = None

    walk(source_outline)

def copy_page_labels(reader, writer):
    """Copy PDF /PageLabels metadata from reader to writer."""
    try:
        source_root = resolve_pdf_object(reader.trailer.get("/Root"))
        if not source_root:
            return
        page_labels = source_root.get("/PageLabels")
        if not page_labels:
            return
        page_labels_dict = resolve_pdf_object(page_labels)
        writer._root_object[NameObject("/PageLabels")] = page_labels_dict.clone(writer)
    except Exception as e:
        print(f"  -> Warning: Could not copy page labels ({e})")

def set_page_labels_from_values(writer, labels):
    """Set output PDF page labels explicitly from a list of label strings."""
    try:
        nums = ArrayObject()
        for index, label in enumerate(labels):
            label_dict = DictionaryObject()
            label_dict[NameObject("/P")] = TextStringObject(str(label))
            nums.append(NumberObject(index))
            nums.append(label_dict)
        writer._root_object[NameObject("/PageLabels")] = DictionaryObject({
            NameObject("/Nums"): nums
        })
    except Exception as e:
        print(f"  -> Warning: Could not set page labels from stamp values ({e})")

def add_page_label_bookmarks(writer, labels, parent_title="NAVIGATE BY PAGE LABEL"):
    """Add one bookmark per page under a parent outline item."""
    if not labels:
        return
    try:
        parent_bookmark = writer.add_outline_item(parent_title, 0)
        for page_index, label in enumerate(labels):
            writer.add_outline_item(str(label), page_index, parent=parent_bookmark)
    except Exception as e:
        print(f"  -> Warning: Could not add page-label bookmarks ({e})")

def create_stamp(text, x, y, size=12, color="#000000", width=None, height=None,
                background_box=False, box_color="#FFFFFF", box_padding=2):
    """Create a PDF with a text stamp at the specified position"""
    # Use provided dimensions or default to letter size
    pagesize = (width, height) if width and height else letter
    
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=pagesize)
    box_padding = max(0, box_padding)
    text_width, text_height, _, descent = get_text_metrics(text, size=size)
    text_x = x
    text_y = y

    if background_box:
        try:
            c.setFillColor(HexColor(box_color))
        except:
            print(f"Warning: Invalid box color '{box_color}', using white instead")
            c.setFillColorRGB(1, 1, 1)
        box_x = x - box_padding
        box_y = y + descent - box_padding
        box_width = text_width + (box_padding * 2)
        box_height = text_height + (box_padding * 2)
        c.rect(box_x, box_y, box_width, box_height, fill=1, stroke=0)
        text_x = box_x + ((box_width - text_width) / 2)
        text_y = box_y + ((box_height - text_height) / 2) - descent

    try:
        # Convert hex color to RGB
        c.setFillColor(HexColor(color))
    except:
        # Fallback to black if color is invalid
        print(f"Warning: Invalid color '{color}', using black instead")
        c.setFillColorRGB(0, 0, 0)
    c.setFont(STAMP_FONT, size)
    c.drawString(text_x, text_y, text)
    c.save()
    packet.seek(0)
    return packet

def create_stamp_annotation(text, x, y, size=12, color="#000000",
                           background_box=False, box_color="#FFFFFF", box_padding=2,
                           writer=None):
    """Create a removable free-text annotation."""
    box_padding = max(0, box_padding)
    text_width, text_height, _, descent = get_text_metrics(text, size=size)
    rect_width = text_width + (box_padding * 2)
    rect_height = text_height + (box_padding * 2)
    rect = [
        x - box_padding,
        y + descent - box_padding,
        x - box_padding + rect_width,
        y + descent - box_padding + rect_height
    ]
    text_rgb = parse_hex_color(color, (0, 0, 0), "text color")
    annotation = {
        "/Type": "/Annot",
        "/Subtype": "/FreeText",
        "/Rect": rect,
        "/Contents": text,
        "/T": STAMP_ANNOTATION_AUTHOR,
        "/Subj": STAMP_ANNOTATION_SUBJECT,
        "/DA": f"/Helv {size} Tf {text_rgb[0]:.6f} {text_rgb[1]:.6f} {text_rgb[2]:.6f} rg",
        "/Q": 1,
        "/F": 4,
        "/BS": {"/W": 0},
        "/NM": f"{STAMP_ANNOTATION_NAME_PREFIX}{uuid.uuid4().hex}"
    }
    if background_box:
        box_rgb = parse_hex_color(box_color, (1, 1, 1), "box color")
        annotation["/C"] = [box_rgb[0], box_rgb[1], box_rgb[2]]

    if writer is not None:
        appearance_packet = BytesIO()
        c = canvas.Canvas(appearance_packet, pagesize=(rect_width, rect_height))
        if background_box:
            c.setFillColorRGB(*annotation["/C"])
            c.rect(0, 0, rect_width, rect_height, fill=1, stroke=0)
        c.setFillColorRGB(*text_rgb)
        c.setFont(STAMP_FONT, size)
        text_x = (rect_width - text_width) / 2
        text_y = ((rect_height - text_height) / 2) - descent
        c.drawString(text_x, text_y, text)
        c.save()
        appearance_packet.seek(0)

        appearance_pdf = PdfReader(appearance_packet)
        appearance_page = appearance_pdf.pages[0]
        contents = resolve_pdf_object(appearance_page.get("/Contents"))
        if isinstance(contents, list):
            stream_bytes = b"".join(resolve_pdf_object(item).get_data() for item in contents)
        else:
            stream_bytes = contents.get_data()
        resources = resolve_pdf_object(appearance_page.get("/Resources"))
        appearance_stream = DecodedStreamObject()
        appearance_stream.set_data(stream_bytes)
        appearance_stream[NameObject("/Type")] = NameObject("/XObject")
        appearance_stream[NameObject("/Subtype")] = NameObject("/Form")
        appearance_stream[NameObject("/BBox")] = ArrayObject([
            FloatObject(0), FloatObject(0),
            FloatObject(rect_width), FloatObject(rect_height)
        ])
        if resources:
            appearance_stream[NameObject("/Resources")] = resources.clone(writer)
        appearance_ref = writer._add_object(appearance_stream)
        annotation["/AP"] = DictionaryObject({
            NameObject("/N"): appearance_ref
        })

    return annotation

def should_remove_annotation(annotation_obj, annotation_removal):
    """Return True when an annotation should be removed by removal mode."""
    subtype = annotation_obj.get("/Subtype")
    if subtype not in ("/FreeText", "/Stamp"):
        return False
    if annotation_removal == "all-freetext" and subtype == "/FreeText":
        return True
    author = str(annotation_obj.get("/T", ""))
    subject = str(annotation_obj.get("/Subj", ""))
    name = str(annotation_obj.get("/NM", ""))
    if author == STAMP_ANNOTATION_AUTHOR or subject == STAMP_ANNOTATION_SUBJECT:
        return True
    if name.startswith(STAMP_ANNOTATION_NAME_PREFIX):
        return True

    # Backward compatibility for files stamped before /NM markers were added.
    da = str(annotation_obj.get("/DA", ""))
    border = resolve_pdf_object(annotation_obj.get("/BS")) or {}
    border_width = border.get("/W")
    legacy_freetext_match = (
        annotation_obj.get("/Q") == 1 and
        annotation_obj.get("/F") == 4 and
        border_width == 0 and
        "/Helv" in da and
        "/AP" in annotation_obj
    )
    if legacy_freetext_match:
        return True

    return (
        subtype == "/Stamp" and
        annotation_obj.get("/F") == 4 and
        annotation_obj.get("/Name") == "/BatesStamp" and
        "/AP" in annotation_obj
    )

def remove_annotations_from_pdf(input_path, output_path, annotation_removal='bates'):
    """Remove selected annotations from a PDF while preserving core structure."""
    reader = PdfReader(input_path)
    writer = PdfWriter()

    metadata = reader.metadata
    if metadata:
        safe_metadata = {key: str(value) for key, value in metadata.items() if value is not None}
        if safe_metadata:
            writer.add_metadata(safe_metadata)

    removed_count = 0

    for page in reader.pages:
        writer.add_page(page)
        output_page = writer.pages[-1]
        annots = resolve_pdf_object(output_page.get("/Annots"))
        if not annots:
            continue

        kept = []
        for annot_ref in annots:
            annot_obj = resolve_pdf_object(annot_ref)
            if should_remove_annotation(annot_obj, annotation_removal):
                removed_count += 1
            else:
                kept.append(annot_ref)

        if kept:
            output_page[NameObject("/Annots")] = ArrayObject(kept)
        elif "/Annots" in output_page:
            del output_page[NameObject("/Annots")]

    copy_outlines(reader, writer)
    copy_page_labels(reader, writer)

    with open(output_path, 'wb') as output_file:
        writer.write(output_file)

    return removed_count

def flatten_pdf(input_path, output_path, pdf_settings="/prepress", allow_fallback=True):
    """Flatten/rewrite a PDF using Ghostscript, with optional PyPDF2 fallback."""
    try:
        gs_executable = shutil.which("gs") or shutil.which("gswin64c") or shutil.which("gswin32c")
        if not gs_executable:
            gs_root = r"C:\Program Files\gs"
            if os.path.isdir(gs_root):
                gs_versions = sorted(
                    [path for path in os.listdir(gs_root) if os.path.isdir(os.path.join(gs_root, path))],
                    reverse=True
                )
                for version_dir in gs_versions:
                    for binary_name in ("gswin64c.exe", "gswin32c.exe"):
                        candidate = os.path.join(gs_root, version_dir, "bin", binary_name)
                        if os.path.exists(candidate):
                            gs_executable = candidate
                            break
                    if gs_executable:
                        break
        if not gs_executable:
            raise FileNotFoundError("Ghostscript executable not found")

        # Try using Ghostscript for best results
        gs_command = [gs_executable, "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER",
                     "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                     f"-dPDFSETTINGS={pdf_settings}", f"-sOutputFile={output_path}",
                     input_path]

        subprocess.run(gs_command, check=True)
        if pdf_settings != "/prepress":
            print(f"  -> Rewrote PDF using Ghostscript with {pdf_settings} compression profile")
        else:
            print(f"  -> Flattened PDF using Ghostscript")
        return True
    except (subprocess.SubprocessError, FileNotFoundError):
        if not allow_fallback:
            print(f"  -> Ghostscript not available; requested operation requires Ghostscript")
            return False
        print(f"  -> Ghostscript not available, using PyPDF2 fallback for flattening")
        
        # PyPDF2 fallback method
        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()
            
            for page in reader.pages:
                # Force content stream compression to help with flattening
                writer.add_page(page)
                
            with open(output_path, 'wb') as f:
                writer.write(f)
            
            return True
        except Exception as e:
            print(f"  -> Error in PyPDF2 fallback flattening: {e}")
            # If flattening fails, copy the original file to output
            with open(input_path, 'rb') as infile, open(output_path, 'wb') as outfile:
                outfile.write(infile.read())
            return False

def apply_bates_stamp(input_path, output_path, bates_prefix, start_number, position='bottom-right',
                     color="#000000", margin=10, size=12, flatten_output=False,
                     number_source='sequential', page_labels=None, compress_output='none',
                     background_box=False, box_color="#FFFFFF", box_padding=2,
                     stamp_mode='merge', set_page_labels_from_stamp=False,
                     bookmark_page_labels=False,
                     page_label_bookmark_parent="NAVIGATE BY PAGE LABEL"):
    """Apply Bates stamp to a PDF file"""
    reader = PdfReader(input_path)
    writer = PdfWriter()
    metadata = reader.metadata
    if metadata:
        safe_metadata = {key: str(value) for key, value in metadata.items() if value is not None}
        if safe_metadata:
            writer.add_metadata(safe_metadata)
    stamp_labels = []
    
    # Process each page
    for i, page in enumerate(reader.pages):
        # Get page dimensions
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        crop_left = float(page.cropbox.left)
        crop_bottom = float(page.cropbox.bottom)
        crop_right = float(page.cropbox.right)
        crop_top = float(page.cropbox.top)
        visible_width = crop_right - crop_left
        visible_height = crop_top - crop_bottom
        
        # Calculate text width (approximate)
        if number_source == 'page-label' and page_labels:
            label_value = page_labels[i] if i < len(page_labels) else str(i + 1)
            bates_number = f"{bates_prefix}{label_value}"
        else:
            bates_number = f"{bates_prefix}{start_number + i:06d}"
        stamp_labels.append(bates_number)
        text_width, _, _, _ = get_text_metrics(bates_number, size=size)
        
        # Define positions based on visible crop area so stamps remain visible on cropped pages.
        positions = {
            'bottom-right': (crop_left + visible_width - margin - text_width, crop_bottom + margin),
            'bottom-left': (crop_left + margin, crop_bottom + margin),
            'bottom-middle': (crop_left + visible_width / 2 - text_width / 2, crop_bottom + margin),
            'top-right': (crop_left + visible_width - margin - text_width, crop_bottom + visible_height - margin - size),
            'top-left': (crop_left + margin, crop_bottom + visible_height - margin - size),
            'top-middle': (crop_left + visible_width / 2 - text_width / 2, crop_bottom + visible_height - margin - size),
            'center': (crop_left + visible_width / 2 - text_width / 2, crop_bottom + visible_height / 2)
        }
        x, y = positions.get(position, positions['bottom-right'])
        
        if stamp_mode == 'annotation':
            writer.add_page(page)
            writer.add_annotation(
                i,
                create_stamp_annotation(
                    bates_number,
                    x,
                    y,
                    size=size,
                    color=color,
                    background_box=background_box,
                    box_color=box_color,
                    box_padding=box_padding,
                    writer=writer
                )
            )
        else:
            # Create stamp for this page with correct page size
            stamp_packet = create_stamp(
                bates_number,
                x,
                y,
                size=size,
                color=color,
                width=page_width,
                height=page_height,
                background_box=background_box,
                box_color=box_color,
                box_padding=box_padding
            )

            # Create a new reader for the stamp
            stamp_pdf = PdfReader(stamp_packet)
            stamp_page = stamp_pdf.pages[0]

            # Apply stamp as an overlay
            page.merge_page(stamp_page)
            writer.add_page(page)

    copy_outlines(reader, writer)
    if set_page_labels_from_stamp and number_source == 'sequential':
        set_page_labels_from_values(writer, stamp_labels)
        navigation_labels = stamp_labels
    else:
        copy_page_labels(reader, writer)
        navigation_labels = page_labels if page_labels else stamp_labels

    if bookmark_page_labels:
        add_page_label_bookmarks(writer, navigation_labels, page_label_bookmark_parent)
    
    temp_output_path = output_path
    needs_postprocess = flatten_output or compress_output != "none"
    if stamp_mode == 'annotation' and needs_postprocess:
        print("  -> Warning: flatten/compress disabled in annotation mode to keep stamps removable")
        needs_postprocess = False

    # Write the output file (either final or temporary)
    if needs_postprocess:
        temp_output_path = output_path + ".temp.pdf"
    
    with open(temp_output_path, 'wb') as output_file:
        writer.write(output_file)
    
    # If flattening/compression is requested, perform the post-process operation
    if needs_postprocess:
        compression_profiles = {
            "screen": "/screen",
            "ebook": "/ebook",
            "printer": "/printer",
            "prepress": "/prepress",
            "default": "/default"
        }
        gs_setting = compression_profiles.get(compress_output, "/prepress")
        allow_fallback = compress_output == "none"
        postprocess_ok = flatten_pdf(
            temp_output_path,
            output_path,
            pdf_settings=gs_setting,
            allow_fallback=allow_fallback
        )
        if not postprocess_ok:
            print("  -> Warning: Output compression was requested but Ghostscript is unavailable; keeping uncompressed output")
            shutil.copyfile(temp_output_path, output_path)
        try:
            os.remove(temp_output_path)  # Clean up temp file
        except:
            pass  # Ignore errors in cleanup
    
    return start_number + len(reader.pages)

def process_directory(input_dir, output_dir, bates_prefix, start_number, position='bottom-right',
                     color="#000000", margin=10, size=12, flatten_input=False, flatten_output=False,
                     number_source='sequential', compress_output='none',
                     background_box=False, box_color="#FFFFFF", box_padding=2,
                     stamp_mode='merge', set_page_labels_from_stamp=False,
                     bookmark_page_labels=False,
                     page_label_bookmark_parent="NAVIGATE BY PAGE LABEL"):
    """Process all PDF files in a directory and its subdirectories"""
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    current_number = start_number
    total_files_processed = 0
    total_files_failed = 0
    
    # Create temp directory for flattened input files if needed
    temp_dir = None
    if flatten_input:
        temp_dir = tempfile.mkdtemp(prefix="bates_temp_")
    
    # Walk through the directory
    for root, _, files in os.walk(input_dir):
        # Create corresponding output subdirectory
        rel_path = os.path.relpath(root, input_dir)
        current_output_dir = os.path.join(output_dir, rel_path) if rel_path != '.' else output_dir
        
        if not os.path.exists(current_output_dir):
            os.makedirs(current_output_dir)
        
        # Process PDF files
        for file in files:
            if file.lower().endswith('.pdf'):
                input_path = os.path.join(root, file)
                output_path = os.path.join(current_output_dir, file)
                original_input_path = input_path
                
                print(f"Processing: {input_path}")
                page_labels = None
                if number_source == 'page-label':
                    try:
                        labels_reader = PdfReader(original_input_path)
                        page_labels = get_page_labels(labels_reader)
                    except Exception as e:
                        print(f"  -> Warning: Could not load page labels: {e}")
                
                # Flatten input if requested
                if flatten_input:
                    flattened_input_path = os.path.join(temp_dir, file)
                    print(f"  -> Pre-flattening input PDF...")
                    try:
                        flatten_pdf(input_path, flattened_input_path)
                        input_path = flattened_input_path
                    except Exception as e:
                        print(f"  -> Error flattening input: {e}, using original file")
                
                try:
                    # Check if PDF is readable
                    try:
                        test_reader = PdfReader(input_path)
                        num_pages = len(test_reader.pages)
                        print(f"  -> Found {num_pages} pages")
                    except Exception as e:
                        print(f"  -> Warning: Could not read PDF structure: {e}")
                        print(f"  -> Will attempt processing anyway")
                    
                    current_number = apply_bates_stamp(
                        input_path, output_path, bates_prefix, current_number,
                        position=position, color=color, margin=margin, size=size,
                        flatten_output=flatten_output, number_source=number_source,
                        page_labels=page_labels, compress_output=compress_output,
                        background_box=background_box, box_color=box_color, box_padding=box_padding,
                        stamp_mode=stamp_mode,
                        set_page_labels_from_stamp=set_page_labels_from_stamp,
                        bookmark_page_labels=bookmark_page_labels,
                        page_label_bookmark_parent=page_label_bookmark_parent
                    )
                    print(f"  -> Successfully saved to: {output_path}")
                    total_files_processed += 1
                    
                except Exception as e:
                    print(f"  -> Error processing {input_path}: {e}")
                    total_files_failed += 1
    
    # Clean up temp directory if created
    if temp_dir:
        try:
            shutil.rmtree(temp_dir)
        except:
            print(f"  -> Note: Could not remove temporary directory: {temp_dir}")
    
    total_pages = current_number - start_number
    
    # Print summary
    print(f"\nBates stamping summary:")
    print(f"  - Files successfully processed: {total_files_processed}")
    print(f"  - Files failed: {total_files_failed}")
    print(f"  - Total pages stamped: {total_pages}")
    
    return total_pages

def process_directory_remove_annotations(input_dir, output_dir, annotation_removal='bates'):
    """Process all PDFs and remove annotations according to the selected mode."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    total_files_processed = 0
    total_files_failed = 0
    total_annotations_removed = 0

    for root, _, files in os.walk(input_dir):
        rel_path = os.path.relpath(root, input_dir)
        current_output_dir = os.path.join(output_dir, rel_path) if rel_path != '.' else output_dir

        if not os.path.exists(current_output_dir):
            os.makedirs(current_output_dir)

        for file in files:
            if file.lower().endswith('.pdf'):
                input_path = os.path.join(root, file)
                output_path = os.path.join(current_output_dir, file)
                print(f"Processing: {input_path}")
                try:
                    removed = remove_annotations_from_pdf(
                        input_path,
                        output_path,
                        annotation_removal=annotation_removal
                    )
                    total_annotations_removed += removed
                    total_files_processed += 1
                    print(f"  -> Removed {removed} annotations")
                    print(f"  -> Successfully saved to: {output_path}")
                except Exception as e:
                    print(f"  -> Error processing {input_path}: {e}")
                    total_files_failed += 1

    print(f"\nAnnotation removal summary:")
    print(f"  - Files successfully processed: {total_files_processed}")
    print(f"  - Files failed: {total_files_failed}")
    print(f"  - Total annotations removed: {total_annotations_removed}")
    return total_annotations_removed

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Batch Bates stamp PDF files recursively')
    parser.add_argument('input_directory', help='Input directory containing PDF files')
    parser.add_argument('output_directory', help='Output directory for stamped PDF files')
    parser.add_argument('--operation', default='stamp',
                       choices=['stamp', 'remove-annotations'],
                       help='Operation mode: stamp pages or remove annotations')
    parser.add_argument('--annotation-removal', default='bates',
                       choices=['bates', 'all-freetext'],
                       help='When removing annotations: remove only BatesStamper annotations or all FreeText')
    parser.add_argument('--prefix', default='BATES-', help='Prefix for Bates numbers')
    parser.add_argument('--start', type=int, default=1, help='Starting Bates number')
    parser.add_argument('--position', default='bottom-right', 
                        choices=['bottom-right', 'bottom-left', 'bottom-middle', 'top-right', 'top-left', 'top-middle', 'center'],
                        help='Position of the Bates stamp')
    parser.add_argument('--color-preset', default=None,
                       choices=list(COLOR_PRESETS.keys()),
                       help='Preset color combination for text/background box styling')
    parser.add_argument('--color', default='#000000', help='Hex color code for the Bates stamp (e.g., #FF0000 for red)')
    parser.add_argument('--margin', type=int, default=10, help='Margin from edge in points (1/72 inch)')
    parser.add_argument('--size', type=int, default=12, help='Font size for the Bates stamp')
    parser.add_argument('--stamp-mode', default='merge',
                       choices=['merge', 'annotation'],
                       help='Stamp mode: merge (permanent) or annotation (removable)')
    parser.add_argument('--background-box', action='store_true',
                       help='Draw a filled box behind the stamp text to obscure underlying content')
    parser.add_argument('--box-color', default='#FFFFFF',
                       help='Hex color for the background box (e.g., #FFFFFF for white)')
    parser.add_argument('--box-padding', type=int, default=2,
                       help='Padding around stamp text inside the background box (points)')
    parser.add_argument('--flatten-input', action='store_true', 
                       help='Flatten input PDFs before stamping (recommended for complex PDFs with images)')
    parser.add_argument('--flatten-output', action='store_true', 
                       help='Flatten output PDFs after stamping (for better Adobe Reader compatibility)')
    parser.add_argument('--compress-output', default='none',
                       choices=['none', 'screen', 'ebook', 'printer', 'prepress', 'default'],
                       help='Compress output PDF via Ghostscript profile (requires Ghostscript)')
    parser.add_argument('--number-source', default='sequential',
                       choices=['sequential', 'page-label'],
                       help='Source for the stamp number text: sequential counter or PDF page label')
    parser.add_argument('--set-page-labels-from-stamp', action='store_true',
                       help='When number-source is sequential, set PDF page labels to the stamped Bates values')
    parser.add_argument('--bookmark-page-labels', action='store_true',
                       help='Add one bookmark per page label under a parent bookmark folder')
    parser.add_argument('--page-label-bookmark-parent', default='NAVIGATE BY PAGE LABEL',
                       help='Parent bookmark title used by --bookmark-page-labels')
    
    args = parser.parse_args()
    
    if args.operation == 'remove-annotations':
        print("Starting annotation removal")
        print(f"Annotation removal mode: {args.annotation_removal}")
        removed_count = process_directory_remove_annotations(
            args.input_directory,
            args.output_directory,
            args.annotation_removal
        )
        print(f"\nAnnotation removal complete!")
        print(f"Total annotations removed: {removed_count}")
    else:
        effective_color = args.color
        effective_background_box = args.background_box
        effective_box_color = args.box_color
        if args.color_preset:
            preset = COLOR_PRESETS[args.color_preset]
            effective_color = preset["text_color"]
            effective_background_box = preset["background_box"]
            if preset["background_box"] and preset["box_color"]:
                effective_box_color = preset["box_color"]

        print(f"Starting Bates stamping with prefix '{args.prefix}' from number {args.start}")
        print(f"Position: {args.position}, Color: {effective_color}, Font size: {args.size}")
        print(f"Stamp mode: {args.stamp_mode}")
        if args.color_preset:
            print(f"Color preset: {args.color_preset} ({COLOR_PRESETS[args.color_preset]['use_case']})")
        if effective_background_box:
            print(f"Background box: Enabled (color={effective_box_color}, padding={max(0, args.box_padding)})")
        
        if args.flatten_input:
            print(f"Input PDF flattening: Enabled (for better processing of complex PDFs)")
        if args.flatten_output:
            print(f"Output PDF flattening: Enabled (for better Adobe Reader compatibility)")
        if args.compress_output != 'none':
            print(f"Output compression: Enabled ({args.compress_output})")
        if args.number_source == 'page-label':
            print(f"Number source: PDF page labels")
        else:
            print(f"Number source: Sequential ({args.start:06d} and up)")
        if args.set_page_labels_from_stamp and args.number_source == 'sequential':
            print("Set page labels from stamp: Enabled")
        elif args.set_page_labels_from_stamp:
            print("Set page labels from stamp: Ignored (only valid with --number-source sequential)")
        if args.bookmark_page_labels:
            print(f"Bookmark by page label: Enabled (parent='{args.page_label_bookmark_parent}')")
        
        pages_stamped = process_directory(
            args.input_directory, 
            args.output_directory,
            args.prefix,
            args.start,
            args.position,
            effective_color,
            args.margin,
            args.size,
            args.flatten_input,
            args.flatten_output,
            args.number_source,
            args.compress_output,
            effective_background_box,
            effective_box_color,
            args.box_padding,
            args.stamp_mode,
            args.set_page_labels_from_stamp,
            args.bookmark_page_labels,
            args.page_label_bookmark_parent
        )
        
        print(f"\nBates stamping complete!")
        print(f"Total pages stamped: {pages_stamped}")
        print(f"Last Bates number: {args.prefix}{args.start + pages_stamped - 1:06d}")
