import os
import base64
from io import BytesIO
import math
import tempfile
import cv2
import numpy as np
import aspose.slides as slides
import aspose.pydrawing as drawing
from PIL import Image
from pydub import AudioSegment
import docx
import PyPDF2
from dotenv import load_dotenv
import openai
from fastapi import UploadFile
from pathlib import Path
from datetime import datetime
import pandas as pd
load_dotenv()

def get_unique_filename(original_filename: str) -> str:
    """Generate a unique filename using timestamp."""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    name, ext = os.path.splitext(original_filename)
    return f"{name}_{timestamp}{ext}"

def process_image(file: UploadFile):
    """Process and convert image to base64."""
    try:
        # Open image using PIL
        image = Image.open(file.file)
        
        # Convert to RGB if necessary
        if image.mode != 'RGB':
            image = image.convert('RGB')
            
        # Resize if image is too large (optional)
        max_size = 1600
        if max(image.size) > max_size:
            ratio = max_size / max(image.size)
            new_size = tuple(int(dim * ratio) for dim in image.size)
            image = image.resize(new_size, Image.Resampling.LANCZOS)
            
        # Convert to base64
        buffer = BytesIO()
        image.save(buffer, format="JPEG", quality=85)
        base64_image = base64.b64encode(buffer.getvalue()).decode()
        
        return [base64_image]  # Return as list for consistency with other functions
    except Exception as e:
        raise ValueError(f"Error processing image: {str(e)}")

def audio_to_text(file: UploadFile):
    """Convert audio to text using OpenAI's Whisper API."""
    openai.api_key = os.getenv("OPENAI_API_KEY")
    audio = AudioSegment.from_file(file.file)
    five_minutes = 5 * 60 * 1000
    truncated_audio = audio[:min(five_minutes, len(audio))]
    
    unique_filename = get_unique_filename(file.filename)
    with tempfile.NamedTemporaryFile(prefix=unique_filename, suffix=Path(file.filename).suffix, delete=False) as temp_file:
        truncated_audio.export(temp_file.name, format=Path(file.filename).suffix[1:])
        with open(temp_file.name, "rb") as audio_file:
            response = openai.audio.transcriptions.create(model="whisper-1", file=audio_file)
    os.unlink(temp_file.name)
    return response.text
import traceback
def video_to_base64_frames_and_audio(file: UploadFile, samples_per_minute=3):
    """Convert video to an array of base64-encoded frames and extract audio."""
    unique_filename = get_unique_filename(file.filename)
    
    # Create a temporary file to store the video
    with tempfile.NamedTemporaryFile(prefix=unique_filename, suffix=Path(file.filename).suffix, delete=False) as temp_video:
        temp_video.write(file.file.read())
        temp_video_path = temp_video.name
        
    # Extract frames
    cap = cv2.VideoCapture(temp_video_path)
    fps = cap.get(cv2.CAP_PROP_FPS)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    total_minutes = math.ceil(total_frames / (fps * 60))
    base64_frames = []
    
    for minute in range(total_minutes):
        start_frame = minute * fps * 60
        end_frame = min((minute + 1) * fps * 60, total_frames)
        if minute == total_minutes - 1:
            frame_indices = np.linspace(start_frame, end_frame - 1, samples_per_minute, dtype=int)
        else:
            frame_indices = np.linspace(start_frame, end_frame - 1, samples_per_minute, dtype=int)
            
        for frame_idx in frame_indices:
            cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
            ret, frame = cap.read()
            if ret:
                _, buffer = cv2.imencode(".jpg", frame)
                base64_frames.append(base64.b64encode(buffer).decode())
    
    cap.release()
    
    # Extract audio
    try:
        # Load video as audio file
        video = AudioSegment.from_file(temp_video_path)

        # Export audio to temporary file
        audio_unique_filename = get_unique_filename("audio.mp3")
        with tempfile.NamedTemporaryFile(prefix=audio_unique_filename, suffix=".mp3", delete=False) as temp_audio:
            video.export(temp_audio.name, format="mp3", parameters=["-q:a", "0"])

            # IMPORTANT: Create new file object AFTER closing the previous one
            temp_audio_path = temp_audio.name  # Store the path

        # Now open it with a new file object
        with open(temp_audio_path, "rb") as audio_file:
            audio_upload = UploadFile(filename=audio_unique_filename, file=audio_file)
            audio_text = audio_to_text(audio_upload)

        # Now it's safe to delete
        os.unlink(temp_audio_path)
    except Exception as e:
        traceback.print_exc()
        audio_text = f"Error extracting audio: {str(e)}"

        # Clean up temporary video file
    os.unlink(temp_video_path)
    return base64_frames, audio_text

def base64_to_text(base64_images):
    """Generate text description from base64-encoded images using OpenAI's GPT."""
    client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    messages = [
        {
            "role": "user",
            "content": [
                "These are frames from a video, a GIF, or slides from a presentation. Generate a complete description.",
                *[{"image": image, "resize": 768} for image in base64_images],
            ],
        }
    ]
    response = client.chat.completions.create(
        model="gpt-4o-mini-2024-07-18",  # Updated to use vision model
        messages=messages,
        max_tokens=200,
    )
    return response.choices[0].message.content

def gif_to_base64_frames(file: UploadFile, max_frames=10):
    """Convert a GIF into an array of base64-encoded frames."""
    gif = Image.open(file.file)
    total_frames = gif.n_frames
    if total_frames <= max_frames:
        frame_indices = range(total_frames)
    else:
        step = (total_frames - 1) / (max_frames - 1)
        frame_indices = [int(i * step) for i in range(max_frames)]
    base64_frames = []
    for idx in frame_indices:
        gif.seek(idx)
        frame = gif.convert("RGB")
        buffer = BytesIO()
        frame.save(buffer, format="JPEG")
        base64_frames.append(base64.b64encode(buffer.getvalue()).decode())
    return base64_frames

def ppt_to_base64(file: UploadFile, max_slides=5):
    """Convert PowerPoint slides to base64-encoded JPEG images."""
    presentation = slides.Presentation(file.file)
    desired_width, desired_height = 1200, 800
    scale_x = desired_width / presentation.slide_size.size.width
    scale_y = desired_height / presentation.slide_size.size.height
    base64_images = []
    for slide in presentation.slides[:max_slides]:
        image_stream = BytesIO()
        slide.get_thumbnail(scale_x, scale_y).save(image_stream, drawing.imaging.ImageFormat.jpeg)
        base64_images.append(base64.b64encode(image_stream.getvalue()).decode())
    return base64_images

def extract_text_from_document(file: UploadFile):
    """Extract text from a Word document or PDF file."""
    file_extension = Path(file.filename).suffix.lower()
    if file_extension in [".doc", ".docx"]:
        doc = docx.Document(file.file)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs])
    elif file_extension == ".pdf":
        pdf_reader = PyPDF2.PdfReader(file.file)
        return "\n".join([page.extract_text() for page in pdf_reader.pages])
    elif file_extension in [".xls", ".xlsx"]:
        # Read all sheets in the Excel file
        df_dict = pd.read_excel(file.file, sheet_name=None)
        
        # Convert each sheet to string and combine them
        text_parts = []
        for sheet_name, df in df_dict.items():
            # Add sheet name as header
            text_parts.append(f"\n=== Sheet: {sheet_name} ===\n")
            
            # Convert DataFrame to string, handling non-string data
            sheet_text = df.to_string(index=False, na_rep='')
            text_parts.append(sheet_text)
            
        return "\n\n".join(text_parts)
        
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")
    
def file_to_text(file: UploadFile):
    """Convert a file (image, GIF, video, PPT, PDF, Word document, or audio) to text."""
    file_extension = Path(file.filename).suffix.lower()
    if file_extension in [".jpg", ".jpeg", ".png", ".bmp", ".webp",".jfif"]:
        base64_images = process_image(file)
        return base64_to_text(base64_images)
    elif file_extension in [".gif"]:
        base64_frames = gif_to_base64_frames(file)
        return base64_to_text(base64_frames)
    elif file_extension in [".mp4", ".avi", ".mov"]:
        base64_frames, audio_text = video_to_base64_frames_and_audio(file)
        visual_description = base64_to_text(base64_frames)
        return f"""Visual Content: {visual_description}
                Audio Content: {audio_text}"""
    elif file_extension in [".ppt", ".pptx"]:
        base64_slides = ppt_to_base64(file)
        return base64_to_text(base64_slides)
    elif file_extension in [".pdf", ".doc", ".docx",".xls", ".xlsx"]:
        return extract_text_from_document(file)
    elif file_extension in [".mp3", ".wav"]:
        return audio_to_text(file)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

def path_to_uploadfile_sync(file_path: str) -> str:
    """Convert a file path to text description."""
    path = Path(file_path)
    with open(file_path, 'rb') as f:
        upload_file = UploadFile(filename=path.name, file=f)
        result = file_to_text(upload_file)
        f.close()
    return result

if __name__ == "__main__":
    file_path = "file_example_MOV_480_700kB.mov"  # Change this to your file path
    result = path_to_uploadfile_sync(file_path)
    print(result)