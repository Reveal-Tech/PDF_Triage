import os
import glob
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import io
import argparse
import sys
from collections import Counter

def extract_pdf_statistics(pdf_path, output_dir, image_format='tiff'):
    """Extract statistics from a PDF file and split it into individual pages."""
    
    # Open the PDF
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)
    base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # Map format string to file extension
    format_extensions = {
        'tiff': '.tiff',
        'png': '.png',
        'jpeg': '.jpg'
    }
    ext = format_extensions.get(image_format.lower(), '.tiff')
    
    # Prepare statistics data
    stats_data = []
    
    # Process each page
    for page_num, page in enumerate(pdf_document, 1):
        # Create output filename for this page
        output_filename = f"{base_filename}_{page_num}_of_{total_pages}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        
        # Extract the page as a new PDF
        new_pdf = fitz.open()
        new_pdf.insert_pdf(pdf_document, from_page=page_num-1, to_page=page_num-1)
        new_pdf.save(output_path)
        new_pdf.close()
        
        # Get page statistics
        page_stats = get_page_statistics(page, pdf_document, page_num-1)
        page_stats["Original File"] = os.path.basename(pdf_path)
        page_stats["Output File"] = output_filename
        page_stats["Page Number"] = page_num
        page_stats["Total Pages"] = total_pages
        page_stats["File Size (KB)"] = os.path.getsize(output_path) / 1024
        
        # Extract largest image if available
        if page_stats["Raster Count"] > 0:
            img_filename = f"{base_filename}_{page_num}_of_{total_pages}_largest_image{ext}"
            img_path = os.path.join(output_dir, img_filename)
            extract_largest_image(page, img_path, image_format)
            page_stats["Largest Image File"] = img_filename
        else:
            page_stats["Largest Image File"] = "N/A"
        
        stats_data.append(page_stats)
    
    pdf_document.close()
    return stats_data

def get_page_statistics(page, pdf_document, page_index):
    """Get statistics for a specific page including vector objects and colors."""
    stats = {
        "Point Count": 0,
        "Line Count": 0,
        "Polygon Count": 0,
        "Raster Count": 0,
        "Vector Colors": []
    }
    
    # Get page dictionary
    xref = pdf_document.xref_object(page.xref)
    
    # Extract vector graphics (approximate counts)
    paths = page.get_drawings()
    for path in paths:
        if "items" in path:
            for item in path["items"]:
                if item[0] == "l":  # Line
                    stats["Line Count"] += 1
                elif item[0] == "re":  # Rectangle (polygon)
                    stats["Polygon Count"] += 1
                elif item[0] == "c" or item[0] == "v" or item[0] == "y":  # Curves (polygons)
                    stats["Polygon Count"] += 1
                elif item[0] == "m":  # Move (could be point)
                    stats["Point Count"] += 1
        
        # Extract colors used in vectors
        if "color" in path and path["color"]:
            color_str = str(path["color"])
            if color_str not in stats["Vector Colors"]:
                stats["Vector Colors"].append(color_str)
    
    # Count raster images
    images = page.get_images(full=True)
    stats["Raster Count"] = len(images)
    
    # Convert colors list to string
    stats["Vector Colors"] = ", ".join(stats["Vector Colors"]) if stats["Vector Colors"] else "None"
    
    return stats

def extract_largest_image(page, output_path, image_format='tiff'):
    """Extract the largest image from a page and save it in the specified format."""
    largest_img = None
    max_size = 0
    
    for img_index, img in enumerate(page.get_images(full=True)):
        xref = img[0]
        try:
            base_image = page.parent.extract_image(xref)
            image_bytes = base_image["image"]
            
            # Get image size
            img_size = len(image_bytes)
            if img_size > max_size:
                max_size = img_size
                largest_img = image_bytes
        except Exception as e:
            print(f"Warning: Failed to extract image: {e}")
    
    if largest_img:
        try:
            # Save the largest image
            img = Image.open(io.BytesIO(largest_img))
            
            # Convert image mode based on target format
            if image_format.lower() == 'jpeg':
                # JPEG needs RGB mode (no alpha)
                if img.mode in ('RGBA', 'LA'):
                    # Create a white background
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    # Paste the image on the background using alpha as mask
                    if 'A' in img.mode:  # Has alpha channel
                        background.paste(img, mask=img.split()[3])  # 3 is the alpha channel
                    else:
                        background.paste(img)
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                
                img.save(output_path, "JPEG", quality=90)
            
            elif image_format.lower() == 'png':
                # PNG can handle all modes including transparency
                img.save(output_path, "PNG")
            
            else:  # default to TIFF
                # TIFF can handle all modes including transparency
                img.save(output_path, "TIFF", compression="tiff_lzw")
            
            print(f"Saved image to {output_path}")
            
        except Exception as e:
            print(f"Warning: Failed to save image as {image_format}: {e}")
            # Try to save as TIFF as fallback (most versatile format)
            try:
                tiff_path = os.path.splitext(output_path)[0] + '.tiff'
                img = Image.open(io.BytesIO(largest_img))
                img.save(tiff_path, "TIFF", compression="tiff_lzw")
                print(f"Saved image as TIFF instead: {tiff_path}")
            except Exception as e2:
                print(f"Error: Could not save image in any format: {e2}")

def process_all_pdfs(input_dir, output_dir, image_format):
    """Process all PDF files in a directory and create statistics CSV."""
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    all_stats = []
    
    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return
    
    print(f"Using image format: {image_format}")
    
    for pdf_file in pdf_files:
        print(f"Processing {pdf_file}...")
        try:
            stats = extract_pdf_statistics(pdf_file, output_dir, image_format)
            all_stats.extend(stats)
        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")
    
    # Create and save statistics CSV
    if all_stats:
        stats_df = pd.DataFrame(all_stats)
        csv_path = os.path.join(output_dir, "pdf_statistics.csv")
        stats_df.to_csv(csv_path, index=False)
        print(f"Statistics saved to {csv_path}")
        print(f"Processed {len(pdf_files)} PDF files with a total of {len(all_stats)} pages")
    else:
        print("No statistics were collected. Check for errors above.")

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Process PDF files: split pages, extract images, and gather statistics')
    parser.add_argument('-i', '--input', required=True, help='Input directory containing PDF files')
    parser.add_argument('-o', '--output', required=True, help='Output directory for processed files')
    parser.add_argument('-f', '--format', choices=['tiff', 'png', 'jpeg'], default='tiff', 
                        help='Format for extracted images (default: tiff)')
    
    return parser.parse_args()

if __name__ == "__main__":
    args = parse_arguments()
    
    # Validate input directory
    if not os.path.isdir(args.input):
        print(f"Error: Input directory '{args.input}' does not exist or is not a directory")
        sys.exit(1)
    
    print(f"Processing PDFs from: {args.input}")
    print(f"Saving output to: {args.output}")
    
    process_all_pdfs(args.input, args.output, args.format)