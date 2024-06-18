#!/usr/bin/env python
# coding: utf-8

import subprocess
import os
import shutil
import pymongo
import pandas as pd
import openpyxl
from frameioclient import FrameioClient
from requests.exceptions import HTTPError
import argparse
from openpyxl.utils.dataframe import dataframe_to_rows

# MongoDB connection information
mongo_client = pymongo.MongoClient("mongodb://localhost:27017/")
db = mongo_client["your_database_name"]

# Initialize the Frame.io client
frameio_client = FrameioClient("Insert token here")

XY_File = None  # Define XY_File at the module level

def process_collections():
    # Query the database for Baselight collections
    baselight_collection = db["baselight"]
    cursor = baselight_collection.find({}, {"_id": 0, "folder": 1, "frames": 1})

    output_ranges = []

    # Process each document in the Baselight collection
    for document in cursor:
        currentFolder = document["folder"]
        parseline = [str(frame) for frame in document["frames"]]  # Convert frames to strings

        parseFolder = currentFolder.split("/")  # Split current folder by "/"
        if len(parseFolder) > 1:
            parseFolder.pop(1)  # Remove the second element if exists
        newFolder = "/".join(parseFolder)  # Reconstruct the folder path

        for techfile in XY_File:
            if newFolder in techfile:
                currentFolder = techfile.strip()

        tempStart = None
        tempLast = None
        for number in parseline:
            if not number.isdigit():
                continue
            number = int(number)
            if tempStart is None:
                tempStart = number
                tempLast = number
            elif number == tempLast + 1:
                tempLast = number
            else:
                if tempStart == tempLast:
                    output_ranges.append((currentFolder, tempStart))
                else:
                    output_ranges.append((currentFolder, f"{tempStart}-{tempLast}"))
                tempStart = number
                tempLast = number
        if tempStart is not None:
            if tempStart == tempLast:
                output_ranges.append((currentFolder, tempStart))
            else:
                output_ranges.append((currentFolder, f"{tempStart}-{tempLast}"))

    return output_ranges


# Function to populate Baselight collection
def populate_baselight_collection(baselight_file):
    baselight_collection = db["baselight"]

    output_data = []

    for line in baselight_file:
        line = line.strip()
        if not line:
            continue  # Skip empty lines
        parts = line.split()
        folder = parts[0]
        frames = [int(frame) for frame in parts[1:] if frame.isdigit()]
        output_data.append({"folder": folder, "frames": frames})

    # Insert data into Baselight collection
    baselight_collection.insert_many(output_data)

def populate_xytech_collection(xytech_file):
    xytech_collection = db["xytech"]
    output_data = []

    workorder = None
    location = []
    notes = []

    for line in xytech_file:
        line = line.strip()
        if not line:
            continue  # Skip empty lines
        if line.startswith("Xytech Workorder"):
            if workorder is not None:  # Save previous data if any
                output_data.append({"workorder": workorder, "location": location, "notes": "\n".join(notes)})
                location = []  # Reset location for next entry
                notes = []  # Reset notes for next entry
            workorder = line.split()[-1]  # Extract workorder number
        elif line.startswith("Location:"):
            continue  # Skip Location: line
        elif line.startswith("Notes:"):
            continue  # Skip Notes: line
        else:
            # Assume it's a location path or a note
            if line.startswith("/"):  # Check if it's a location path
                location.append(line)  # Add to location list
            else:
                notes.append(line)  # Add to notes list

    # Save the last entry after loop ends
    if workorder is not None:
        output_data.append({"workorder": workorder, "location": location, "notes": "\n".join(notes)})

    # Insert data into Xytech collection
    xytech_collection.insert_many(output_data)

# Function to process collections and check frames against video duration
def process_collectionsVideo(video_file):
    # Get video duration
    video_duration = get_video_duration(video_file)

    # Get valid frames from collections
    valid_frames_with_timecodes = []
    valid_trims = []  # Define a list to store valid trimmed videos

    # Get valid frames from collections
    valid_ranges = process_collections()

    # Filter frames based on video duration
    for folder, frames in valid_ranges:
        if isinstance(frames, int):
            if frames < video_duration:
                continue
        else:
            start_frame, end_frame = map(int, frames.split('-'))
            if start_frame < video_duration:
                print_timecode_range(folder, start_frame, end_frame)
                middle_frame = get_middle_frame(start_frame, end_frame)
                thumbnail_file = generate_thumbnail(video_file, middle_frame)
                trimmed_video_file = trim_video(video_file, start_frame, end_frame)
                if trimmed_video_file:
                    # Upload the trimmed video to Frame.io
                    upload_video_to_frameio(trimmed_video_file, frameio_client, "fb5fdf2d-0463-4b56-8085-1c537f25295d")
                    valid_trims.append((folder, frames, trimmed_video_file))
                if thumbnail_file:
                    valid_frames_with_timecodes.append((folder, frames, frame_to_timecode(middle_frame), thumbnail_file))
    return valid_frames_with_timecodes, valid_trims

# Function to get video duration using ffmpeg
def get_video_duration(video_file):
    command = ['ffprobe', '-v', 'error', '-select_streams', 'v:0', '-show_entries', 'stream=duration,r_frame_rate', '-of', 'csv=p=0', video_file]
    try:
        output = subprocess.check_output(command, stderr=subprocess.STDOUT).decode('utf-8').strip()
        parts = output.split(',')
        if len(parts) > 1:
            duration_str = parts[0]
            if '/' in duration_str:
                numerator, denominator = map(float, duration_str.split('/'))
                duration = numerator / denominator
            else:
                duration = float(duration_str)
            frame_rate_str = parts[1]
            if '/' in frame_rate_str:
                numerator, denominator = map(float, frame_rate_str.split('/'))
                frame_rate = numerator / denominator
            else:
                frame_rate = float(frame_rate_str)
        else:
            duration = float(parts[0])
            frame_rate = 24  # Assuming default frame rate as 24 fps
        video_duration = duration * frame_rate
        return video_duration
    except subprocess.CalledProcessError as e:
        print(f"Error: Failed to get video duration for {video_file}.")
        print(f"Command: {' '.join(command)}")
        print(f"Error message: {e.output.decode('utf-8').strip()}")
        return None

def trim_video(video_file, start_frame, end_frame):
    # Calculate start and end timecodes based on frame numbers (assuming 60 fps)
    start_time = start_frame / 60
    end_time = end_frame / 60

    # Define the output trimmed file path
    trimmed_video_file = f"trimmed_{start_frame}_{end_frame}.mp4"
    
    # Use ffmpeg to trim the video
    command = ['ffmpeg', '-y', '-i', video_file, '-ss', str(start_time), '-to', str(end_time), '-vf', 'fps=60', trimmed_video_file]
    try:
        subprocess.run(command, check=True)
        return trimmed_video_file
    except subprocess.CalledProcessError as e:
        print(f"Error: Failed to trim video from frame {start_frame} to {end_frame}.")
        print(f"Command: {' '.join(command)}")
        print(f"Error message: {e}")
        return None

# Function to upload video to Frame.io
def upload_video_to_frameio(video_file, client, parent_folder_id):
    try:
        # Upload the video to Frame.io
        client.assets.upload(parent_folder_id, video_file)
        print(f"Uploaded {os.path.basename(video_file)} successfully to Frame.io.")
    except HTTPError as e:
        print(f"Failed to upload {os.path.basename(video_file)} to Frame.io. Error: {e}")

# Function to print timecode range
def print_timecode_range(folder, start_frame, end_frame):
    start_timecode = frame_to_timecode(start_frame)
    end_timecode = frame_to_timecode(end_frame)
    print(f"Location: {folder}")
    print(f"Range: {start_frame} - {end_frame}")
    print(f"Timecode Range: {start_timecode} - {end_timecode}\n")
    with open('Print_output.txt', 'a') as file:
        file.write(f"Location: {folder}\n")
        file.write(f"Range: {start_frame} - {end_frame}\n")
        file.write(f"Timecode Range: {start_timecode} - {end_timecode}\n\n")

# Function to convert frame number to timecode
def frame_to_timecode(frame):
    fps = 60  # Assuming 24 frames per second
    total_seconds = frame / fps
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    frames = int(frame % fps)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frames:02d}"

# Function to generate thumbnail
def generate_thumbnail(video_file, frame_number):
    thumbnail_file = f"thumbnail_{frame_number}.jpg"
    command = ['ffmpeg', '-y', '-i', video_file, '-vf', f"select=eq(n\\,{frame_number})", '-vframes', '1', thumbnail_file]
    try:
        subprocess.run(command, check=True)
        return thumbnail_file
    except subprocess.CalledProcessError as e:
        print(f"Error: Failed to generate thumbnail for frame {frame_number}.")
        print(f"Command: {' '.join(command)}")
        print(f"Error message: {e}")
        return None

# Function to get middle frame
def get_middle_frame(range_start, range_end):
    # Calculate the middle frame or closest to
    middle_frame = (range_start + range_end) // 2
    return middle_frame

# Function to export data to XLS
def export_to_xls(data, output_file):
    # Create DataFrame
    df = pd.DataFrame(data, columns=["Location", "Frames", "Timecode", "Thumbnail"])

    # Create a directory to store the images if it doesn't exist
    images_dir = "thumbnails"
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)

    # Iterate through each row and import the image to the thumbnails directory
    for index, row in df.iterrows():
        thumbnail_path = row["Thumbnail"]
        filename = os.path.basename(thumbnail_path)
        destination_path = os.path.join(images_dir, filename)
        # Check if the image exists, if not, skip
        if not os.path.exists(thumbnail_path):
            print(f"Error: Image file {thumbnail_path} does not exist.")
            continue
        # Copy the image to the thumbnails directory
        shutil.copy(thumbnail_path, destination_path)
        # Update the thumbnail path in the DataFrame to the new location
        df.at[index, "Thumbnail"] = destination_path

    # Create a new Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write DataFrame to worksheet
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Add images to the worksheet
    for index, row in df.iterrows():
        img = openpyxl.drawing.image.Image(row["Thumbnail"])
        img.width = 96
        img.height = 74
        ws.add_image(img, f"E{index + 2}")

    # Save the workbook
    wb.save(output_file)

    print(f"Data exported to {output_file}")

# Function to save thumbnail paths to collection
def save_thumbnail_paths_to_collection(data):
    thumbnail_collection = db["thumbnails"]
    thumbnail_paths = []

    for item in data:
        thumbnail_paths.append({"path": item[3]})  # Wrap the thumbnail file path in a dictionary

    # Insert thumbnail paths into the collection
    thumbnail_collection.insert_many(thumbnail_paths)

# Function to process command-line arguments
def parse_arguments():
    parser = argparse.ArgumentParser(description="Process Baselight and Xytech files.")
    parser.add_argument("--baselight", help="Path to the Baselight file")
    parser.add_argument("--xytech", help="Path to the Xytech file")
    parser.add_argument("--process", help="Path to the video file to process")
    parser.add_argument("--output", help="Path to the output Excel file")
    return parser.parse_args()

def main():
    global XY_File  # Declare XY_File as a global variable
    # Parse command-line arguments
    args = parse_arguments()
    
    # Paths to Baselight, Xytech, and video files
    baselight_file_path = args.baselight
    xytech_file_path = args.xytech
    video_file_path = args.process
    output_file_path = args.output

 # Open Baselight file if provided
    if baselight_file_path:
        with open(baselight_file_path, "r") as baselight_file:
            populate_baselight_collection(baselight_file)

    # Open Xytech file if provided
    if xytech_file_path:
        with open(xytech_file_path) as f:
            XY_File = f.read().splitlines() 
        with open(xytech_file_path, "r") as xytech_file:
            populate_xytech_collection(xytech_file)

    # Process collections and check frames against video duration if video file provided
    if video_file_path:
        valid_frames, valid_trims = process_collectionsVideo(video_file_path)

        # Save thumbnail paths to collection
        save_thumbnail_paths_to_collection(valid_frames)

        # Export data to XLS if output file provided
        if output_file_path:
            export_to_xls(valid_frames, output_file_path)
    else:
        print("No video file provided. Skipping processing.")

if __name__ == "__main__":
    main()
