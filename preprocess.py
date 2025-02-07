import os
import cv2
import numpy as np
from pathlib import Path
from matplotlib import pyplot as plt

root_dir = os.path.dirname(__file__)
pdf_folder = os.path.join(root_dir, "temp", "ALL2022")
ocr_preprocessed_folder = os.path.join(root_dir, "temp", "OCR_Preprocessed")

os.makedirs(ocr_preprocessed_folder, exist_ok=True)
print(f"📂 OCR Preprocessed Folder: {ocr_preprocessed_folder}")

all_pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith(".pdf")]

def display(image, title="Image"):
    """Display an image using Matplotlib."""
    plt.figure(figsize=(10, 5))
    plt.imshow(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    plt.axis("off")
    plt.title(title)
    plt.show()

def grayscale(image):
    """Convert image to grayscale."""
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

def binarize(image):
    """Apply binary thresholding to enhance text visibility."""
    _, binary_image = cv2.threshold(image, 210, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary_image

def noise_removal(image):
    """Remove small noise from the image using morphological operations."""
    kernel = np.ones((1, 1), np.uint8)
    image = cv2.dilate(image, kernel, iterations=1)
    image = cv2.erode(image, kernel, iterations=1)
    image = cv2.medianBlur(image, 3)
    return image

def deskew(image):
    """Correct the skew of the image using the Hough Transform."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    lines = cv2.HoughLines(edges, 1, np.pi/180, 200)

    if lines is not None:
        angles = []
        for rho, theta in lines[:, 0]:
            angle = (theta * 180 / np.pi) - 90
            angles.append(angle)

        median_angle = np.median(angles)
        (h, w) = image.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, median_angle, 1.0)
        image = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    return image

def remove_borders(image):
    """Remove unnecessary borders from the image."""
    contours, _ = cv2.findContours(image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cntsSorted = sorted(contours, key=lambda x: cv2.contourArea(x))
    
    if cntsSorted:
        x, y, w, h = cv2.boundingRect(cntsSorted[-1])
        return image[y:y+h, x:x+w]
    return image

for pdf_filename in all_pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_filename)
    print(f"\n🔍 Processing: {pdf_path}")

    image_prefix = Path(pdf_filename).stem
    image_output = os.path.join(ocr_preprocessed_folder, f"preprocessed_{image_prefix}")
    os.system(f'pdftoppm -png "{pdf_path}" "{image_output}"')

    generated_images = [f for f in os.listdir(ocr_preprocessed_folder) if f.startswith(f"preprocessed_{image_prefix}")]
    
    if not generated_images:
        print(f"❌ ERROR: No images were generated for {pdf_filename}.")
        continue

    for img_file in generated_images:
        img_path = os.path.join(ocr_preprocessed_folder, img_file)
        print(f"🖼 Processing Image: {img_path}")

        image = cv2.imread(img_path)
        if image is None:
            print(f"❌ ERROR: Cannot read image {img_path}")
            continue

        gray = grayscale(image)
        binary = binarize(gray)
        denoised = noise_removal(binary)
        deskewed = deskew(image)
        final_image = remove_borders(deskewed)

        processed_img_path = os.path.join(ocr_preprocessed_folder, f"final_{img_file}")
        cv2.imwrite(processed_img_path, final_image)
        print(f"✅ Saved: {processed_img_path}")

print("\n🎉 Preprocessing Complete! All images are saved in:", ocr_preprocessed_folder)
