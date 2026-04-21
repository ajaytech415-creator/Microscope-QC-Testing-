import time
import numpy as np

class HRESULTException(Exception):
    pass

class MockCameraInfo:
    def __init__(self, id, displayname):
        self.id = id
        self.displayname = displayname

class Toupcam:
    @staticmethod
    def EnumV2():
        return [MockCameraInfo("mock_1", "Optike UHD 4K (Simulated)")]

    @staticmethod
    def Open(cam_id):
        return ToupcamInstance(cam_id)

class ToupcamInstance:
    def __init__(self, cam_id):
        self.id = cam_id
        self.width = 1280
        self.height = 720
        self._running = False
        self._start_time = time.time()
        self._auto_expo = True
        self._expo_time = 10000
        self._gain = 200
        self._hflip = False

    def put_eSize(self, size_index):
        pass

    def get_Size(self):
        return (self.width, self.height)

    def get_ResolutionNumber(self):
        return 3

    def get_Resolution(self, index):
        resolutions = [(1280, 720), (1920, 1080), (3840, 2160)]
        return resolutions[index] if index < len(resolutions) else (1280, 720)

    def StartPullModeWithCallback(self, callback, context):
        self._running = True

    def PullImageV2(self, buf, bits, padding):
        if not self._running:
            raise HRESULTException("Not running")

        t = time.time()

        # Create a BGR frame with a sweeping animation
        frame = np.zeros((self.height, self.width, 3), dtype=np.uint8)
        frame[:, :] = (80, 40, 20)  # Dark background (BGR)

        # Animated gradient bar sweeping left to right
        sweep_x = int(((t * 200) % self.width))
        bar_width = 40
        for i in range(bar_width):
            x = (sweep_x + i) % self.width
            intensity = int(255 * (1 - i / bar_width))
            frame[:, x] = (0, intensity // 2, intensity)  # Orange sweep in BGR

        # Static center crosshair
        cx, cy = self.width // 2, self.height // 2
        frame[cy-2:cy+2, :] = (0, 180, 0)   # horizontal green line
        frame[:, cx-2:cx+2] = (0, 180, 0)   # vertical green line

        # Add "SIMULATED FEED" text area indicator (colored block at top)
        frame[0:30, :] = (0, 80, 160)  # Dark orange bar at top

        # Copy into the provided bytearray (must match buffer size exactly)
        flat_bytes = frame.tobytes()
        length = min(len(buf), len(flat_bytes))
        buf[:length] = flat_bytes[:length]

    def Close(self):
        self._running = False

    # --- Camera Controls (stubs for mock) ---
    def put_AutoExpoEnable(self, enabled: bool):
        self._auto_expo = enabled

    def put_ExpoTime(self, microseconds: int):
        self._expo_time = microseconds

    def get_ExpoAGainRange(self):
        """Returns (min, max, default) gain."""
        return (100, 1600, 200)

    def put_ExpoAGain(self, value: int):
        self._gain = value

    def AwbOnce(self, cb, ctx):
        """One-shot auto white balance — no-op in mock."""
        pass

    def put_HFlip(self, enabled: bool):
        self._hflip = enabled
