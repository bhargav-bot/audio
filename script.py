import os
from pyexcel import get_sheet
from pydub import AudioSegment

# Get current working directory
base_dir = os.getcwd()
excel_file = os.path.join(base_dir, 'Record.xlsx')
print("excel file",excel_file)
french_folder = os.path.join(base_dir, 'french')
english_folder = os.path.join(base_dir, 'english')
output_folder = os.path.join(base_dir, 'merged_audio')
os.makedirs(output_folder, exist_ok=True)

sheet = get_sheet(file_name=excel_file)
rows = sheet.row[1:]  # Skip header row

for index, row in enumerate(rows, start=2):  # start=2 because row[0] is header
    try:
        filename = str(row[0]).strip() 
        if not filename or filename==' ': 
            continue


        french_path = os.path.join(french_folder, filename)
        
        english_path = os.path.join(english_folder, filename)
        print(f"[Row {index}] French path: {french_path}")
        print(f"[Row {index}] English path: {english_path}")

        if not os.path.exists(french_path):
            print(f"[Row {index}] Missing French file: {french_path}")
            
        if not os.path.exists(english_path):
            print(f"[Row {index}] Missing English file: {english_path}")
            continue


        print(f"[Row {index}] Processing French: {french_path}, English: {english_path}")
        french_audio = AudioSegment.from_file(french_path)
        english_audio = AudioSegment.from_file(english_path)

        pause = AudioSegment.silent(duration=500)
        combined = english_audio+pause + french_audio 

        output_path = os.path.join(output_folder, filename)
        combined.export(output_path, format="mp3")
        print(f"[Row {index}] ✅ Merged and saved: {output_path}")
    except Exception as e:
        print(f"[Row {index}] ❌ Error processing '{filename}': {e}")
