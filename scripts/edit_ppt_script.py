import os
import pyautogui
import time
import sys
from google.cloud import texttospeech
from pptx import Presentation




# Load the PowerPoint presentation
presentation = Presentation('C:\\Keysight\\Hackathon\\Demo2.pptx')

print('Starting')

# Specify the slide index and text to be replaced
slide_index = 0  # Change this to the index of the slide containing the text
old_text = sys.argv[1]  # Specify the text you want to replace
new_text = sys.argv[2]  # Specify the new text to replace with

# Replace text on the specified slide
slide = presentation.slides[slide_index]
replace_text_with_animation(slide, old_text, new_text)

# Save the modified presentation
modified_pptx_file = 'modified_presentation.pptx'
presentation.save(modified_pptx_file)

# Export the modified presentation to a video (MP4)
output_video_file = 'output_video.mp4'
pyautogui.hotkey('alt', 'f')
pyautogui.press('v')
pyautogui.press('v')
pyautogui.press('enter')
pyautogui.typewrite(modified_pptx_file, interval=0.1)
pyautogui.press('enter')
pyautogui.press('enter')

# Wait for the video export to complete (you can adjust the sleep duration as needed)
time.sleep(10)

# Check if the video file 'presentation.mp4' was created
if os.path.exists('presentation.mp4'):
    # Rename the exported video to the desired filename
    os.rename('presentation.mp4', output_video_file)
    print(f'Video saved as {output_video_file}')
else:
    print('Error: The video file "presentation.mp4" was not created.')

# Clean up the temporary modified PowerPoint file
#os.remove(modified_pptx_file)

# Optional: You can also delete the temporary folder created by PowerPoint during the export
# Make sure to replace 'C:\\Users\\your_username\\AppData\\Local\\Temp\\ppt_save\\'
# with the actual path to the temporary folder if needed.
temp_folder = 'C:\\Users\\your_username\\AppData\\Local\\Temp\\ppt_save\\'
for filename in os.listdir(temp_folder):
    file_path = os.path.join(temp_folder, filename)
    try:
        if os.path.isfile(file_path):
            os.unlink(file_path)
    except Exception as e:
        print(f"Error deleting {file_path}: {e}")

print(f'Video saved as {output_video_file}')