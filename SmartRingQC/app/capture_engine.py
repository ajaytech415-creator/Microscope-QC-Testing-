import cv2
import os
import time
from datetime import datetime

# Adjust import scope for runtime
try:
    from app import db_handler, excel_writer
except ImportError:
    import db_handler
    import excel_writer

class CaptureEngine:
    def __init__(self):
        self.camera_id = 0
        try:
            from app.paths import get_captures_dir
        except ImportError:
            from paths import get_captures_dir
        self.wait_dir = os.path.join(get_captures_dir(), 'WAITING')
        os.makedirs(self.wait_dir, exist_ok=True)
        
    def start(self):
        cap = cv2.VideoCapture(self.camera_id)
        # Try adjusting resolution for better quality
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        
        if not cap.isOpened():
            print("Error: Could not open camera.")
            return

        print("Capture Engine started. Focus on the 'Live Feed' window and press SPACE to capture.")
        
        while True:
            ret, frame = cap.read()
            if not ret:
                time.sleep(0.1)
                continue
                
            cv2.imshow('Live Microscope Feed', frame)
            
            key = cv2.waitKey(1) & 0xFF
            if key == 32: # Spacebar
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                # Use timestamp explicitly for DB vs file name for clarity
                db_timestamp_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                image_name = f"capture_{timestamp}.jpg"
                image_path = os.path.join(self.wait_dir, image_name)
                
                # Save image
                cv2.imwrite(image_path, frame)
                
                # Add to DB and Excel
                db_handler.insert_capture(image_name, image_path)
                excel_writer.add_capture(db_timestamp_str, image_name, image_path)
                
                print(f"Captured: {image_name}")
                
                # Flash effect to indicate capture success
                flash_frame = cv2.bitwise_not(frame)
                cv2.imshow('Live Microscope Feed', flash_frame)
                cv2.waitKey(100)
                
            elif key == 27: # ESC to quit
                break
                
        cap.release()
        cv2.destroyAllWindows()

def run_capture_engine():
    db_handler.init_db()
    engine = CaptureEngine()
    engine.start()

if __name__ == "__main__":
    run_capture_engine()
