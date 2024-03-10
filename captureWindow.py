from pynput import mouse, keyboard
import pyautogui
import threading


class WindowSelector:
    def __init__(self):
        self.start_x = None
        self.start_y = None
        self.end_x = None
        self.end_y = None
        self.selecting = False
        self.listener = None

    def start_selection(self):
        self.listener = mouse.Listener(on_click=self.on_click)
        self.listener.start()
        self.selecting = True

    def stop_selection(self):
        self.selecting = False
        self.listener.stop()

    def on_click(self, x, y, button, pressed):
        if pressed:
            if self.selecting:
                if self.start_x is None and self.start_y is None:
                    self.start_x = x
                    self.start_y = y
                else:
                    self.end_x = x
                    self.end_y = y
                    self.stop_selection()

    def get_selected_coordinates(self):
        if self.start_x is not None and self.start_y is not None and \
                self.end_x is not None and self.end_y is not None:
            return (self.start_x, self.start_y), (self.end_x, self.end_y)
        else:
            return None


class CaptureWindow:
    def __init__(self, coordinates):
        self.coordinates = coordinates
        self.KeyListener = keyboard.Listener(
            on_press=self.on_press, on_release=self.on_release)
        self.mouseListener = mouse.Listener(on_click=self.on_click)

        self.alt_pressed = False
        self.KeyListener.start()
        self.mouseListener.start()

    def on_press(self, key):
        print(key)
        if key == 's':
            print("S Pressed")
        if key == keyboard.Key.shift_l or key == keyboard.Key.shift_r:
            self.alt_pressed = True
        elif self.alt_pressed and key == 's':
            print("Ctrl+S was pressed.")
            self.stop_capture()

    def on_release(self, key):
        if key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r:
            self.alt_pressed = False

    def stop_capture(self):
        self.KeyListener.stop()
        self.mouseListener.stop()


    def on_click(self, x, y, button, pressed):
        if pressed:
            if self.alt_pressed:
                print(
                    f"Save capture state - {self.coordinates} - Ctrl+Clicks pressed")
                # Example usage


def select_window():
    print("Click on top right and bottom left to select the window.")
    selector = WindowSelector()
    selector.start_selection()
    selector.listener.join()
    coordinates = selector.get_selected_coordinates()
    if coordinates:
        return coordinates
    else:
        raise Exception("No coordinates selected.")


def captureWindow(coordinates):
    capture = CaptureWindow(coordinates)
    capture.KeyListener.join()
    capture.mouseListener.join()


if __name__ == "__main__":
    coordinates = select_window()
    print(f"Coordinates: {coordinates}")
    captureWindow(coordinates)
