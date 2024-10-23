import cv2
import mediapipe as mp
import pyautogui
import time
import os
import math

# Initialize MediaPipe hand detector
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.7)
mp_drawing = mp.solutions.drawing_utils

# Webcam capture
cap = cv2.VideoCapture(0)

# Constants for gesture detection
prev_gesture_time = 0
gesture_cooldown = 2  # Cooldown in seconds before allowing a new gesture
zoom_trigger_distance = 20  # Threshold to detect zoom circle motion
positions = []  # Store positions of index finger for detecting circular motion

# Function to open PowerPoint (optional)
def open_presentation(file_path):
    """
    Open a PowerPoint presentation file in fullscreen mode.
    """
    os.startfile(file_path)  # Opens the presentation in default application (PowerPoint)

def detect_gesture(landmarks):
    """
    Detect hand gestures based on the thumb's upward or downward movement.
    """
    wrist = landmarks[0]
    thumb_tip = landmarks[4]
    index_finger_tip = landmarks[8]

    # Vertical movement of the thumb relative to the wrist
    dy = thumb_tip.y - wrist.y
    
    if dy < -0.1:  # Thumb is above the wrist
        return 'thumb_up'
    elif dy > 0.1:  # Thumb is below the wrist
        return 'thumb_down'
    
    # Zoom detection (circular motion)
    if detect_circle_motion(index_finger_tip):
        return 'zoom'

    return None

def detect_circle_motion(index_finger_tip):
    """
    Detect circular motion using index finger landmarks.
    Track positions of the index finger and check for circular pattern.
    """
    global positions
    positions.append((index_finger_tip.x, index_finger_tip.y))

    # Keep the list short (only track last few positions)
    if len(positions) > 15:
        positions.pop(0)

    if len(positions) < 10:
        return False  # Not enough data points to detect a circle

    # Check if the points form a circular-like motion
    first_point = positions[0]
    last_point = positions[-1]

    # Euclidean distance between the first and last point
    distance = math.sqrt((last_point[0] - first_point[0])**2 + (last_point[1] - first_point[1])**2)

    if distance < zoom_trigger_distance:  # Circular motion detected
        return True
    return False

def perform_gesture_action(gesture):
    """
    Map detected gestures to actions (next/previous slide or zoom).
    """
    if gesture == 'thumb_up':
        pyautogui.press('right')  # Go to next slide
        print("Next slide")
    elif gesture == 'thumb_down':
        pyautogui.press('left')  # Go to previous slide
        print("Previous slide")
    elif gesture == 'zoom':
        pyautogui.scroll(100)  # Zoom in (scroll up)
        print("Zooming in")

# Optional: Open PowerPoint presentation (change file path as needed)
        open_presentation(r"C:\Users\sanja\OneDrive\Desktop\IN65B5114.pptx") #your own ppt 


while cap.isOpened():
    ret, frame = cap.read()
    if not ret:
        break

    # Flip image for correct orientation (mirror effect)
    frame = cv2.flip(frame, 1)
    
    # Convert the frame to RGB as MediaPipe expects RGB input
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb_frame)

    # If hands are detected
    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

            # Detect gestures using thumb's position relative to wrist
            gesture = detect_gesture(hand_landmarks.landmark)

            # Execute gesture actions with cooldown to avoid multiple triggers
            if gesture:
                current_time = time.time()
                if current_time - prev_gesture_time > gesture_cooldown:
                    perform_gesture_action(gesture)
                    prev_gesture_time = current_time

    # Display the video feed with hand tracking
    cv2.imshow('Gesture Control', frame)

    if cv2.waitKey(5) & 0xFF == 27:  # Press 'Esc' to exit
        break

# Release resources
cap.release()
cv2.destroyAllWindows()
