import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import random

# Function to list available fonts and get user selection
def choose_font(directory):
    available_fonts = [f for f in os.listdir(directory) if f.endswith('.ttf')]
    if len(available_fonts) == 0:
        raise ValueError("No font files found in the directory.")
    
    print("Available fonts:")
    for i, font in enumerate(available_fonts):
        print(f"{i + 1}. {font}")
    
    while True:
        try:
            choice = int(input(f"Choose a font (enter the number): "))
            if 1 <= choice <= len(available_fonts):
                return os.path.join(directory, available_fonts[choice - 1])
            else:
                print(f"Invalid choice. Please enter a number between 1 and {len(available_fonts)}.")
        except ValueError:
            print("Invalid input. Please enter a number.")

# Define your paths
EXCEL_FILE = 'motivated seller list.xlsx'  # Path to the uploaded Excel file
SAVE_DIR = 'handwritten_texts'
FONTS_DIR = 'fonts'  # Directory containing font files
BACKGROUND_IMAGE = 'bg.png'  # Path to the uploaded background image

# Create save directory if it doesn't exist
os.makedirs(SAVE_DIR, exist_ok=True)

# Get font path from user
FONT_PATH = choose_font(FONTS_DIR)

# Read Excel file
df = pd.read_excel(EXCEL_FILE)

# Function to draw text on an image
def draw_text(text, file_name, font_path, background_image):
    # Load a font
    font_size = 90
    font = ImageFont.truetype(font_path, font_size)  # Increased font size for better readability
    
    # Load the background image
    img = Image.open(background_image).convert("RGB")  # Convert to RGB mode
    img_width, img_height = img.size
    d = ImageDraw.Draw(img)

    # Split text into lines
    lines = text.split('\n')
    max_line_width = 0
    total_height = 0
    
    # Calculate the dimensions required for each line
    for line in lines:
        text_bbox = d.textbbox((0, 0), line, font=font)
        line_width = text_bbox[2] - text_bbox[0]
        line_height = text_bbox[3] - text_bbox[1]
        max_line_width = max(max_line_width, line_width)
        total_height += line_height + 10  # Adding line spacing

    # Calculate starting x and y positions for centering text
    x_start = (img_width - max_line_width) // 2
    y_start = (img_height - total_height) // 2

    # Initial position for the text
    x, y = x_start, y_start

    # Draw each line of text with slight variations
    text_color = (0, 0, 139)  # Dark blue color for the text (R, G, B values)
    for line in lines:
        # Apply random variations to position
        variation_x = random.randint(-10, 10)
        variation_y = random.randint(-5, 5)
        d.text((x + variation_x, y + variation_y), line, font=font, fill=text_color)
        line_height = d.textbbox((0, 0), line, font=font)[3] - d.textbbox((0, 0), line, font=font)[1]
        y += line_height + 10

    # Save the image
    img.save(os.path.join(SAVE_DIR, f'{file_name}.jpeg'), 'JPEG')

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    addressee = row['Adressee']
    address = row['Address']
    city = row['City']
    state = row['State']
    zip_code = row['Zip']
    text = f"{addressee}\n{address}\n{city}, {state} {zip_code}"
    file_name = f'document_{index + 1}'
    draw_text(text, file_name, FONT_PATH, BACKGROUND_IMAGE)

print('Handwritten images have been saved.')
