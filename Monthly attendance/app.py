import streamlit as st
import cv2
from tensorflow.keras.models import load_model
import numpy as np

# Load your pre-trained model
model = load_model("C:/Users/VIKRAM RAAJA K/MY WORK/Class/Monthly attendance/emotion_recognition_model.h5")  # Replace with your model path
emotion_labels = ["Angry", "Disgust", "Fear", "Happy", "Sad", "Surprise", "Neutral"]  # Adjust as per your model

# Set up Streamlit app layout
st.title("Facial Emotion Recognition")
st.write("Click 'Start' to open your webcam and detect emotions in real-time.")

# Start/Stop button
start_button = st.button("Start")
stop_button = st.button("Stop")
run_camera = start_button and not stop_button

if run_camera:
    # Open the video capture
    cap = cv2.VideoCapture(0)

    # Display video feed and detect emotions in real-time
    frame_placeholder = st.empty()  # Placeholder for the video frame

    while run_camera:
        ret, frame = cap.read()
        if not ret:
            st.write("Failed to capture video.")
            break

        # Preprocess the frame for the model
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(48, 48))

        for (x, y, w, h) in faces:
            face = gray[y:y+h, x:x+w]
            face = cv2.resize(face, (48, 48)) / 255.0
            face = face.reshape(1, 48, 48, 1)
            prediction = model.predict(face)
            emotion = emotion_labels[np.argmax(prediction)]

            # Display emotion label
            cv2.putText(frame, emotion, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 0, 0), 2)
            cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)

        # Update the frame in Streamlit
        frame_placeholder.image(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))

    cap.release()
else:
    st.write("Camera is stopped. Press 'Start' to begin again.")
