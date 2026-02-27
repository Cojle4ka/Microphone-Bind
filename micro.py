import tkinter as tk
import keyboard
import ctypes
import winsound
import os
import comtypes
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER

CONFIG_FILE = "mic_settings.txt"

class MicMuteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Micro on/Off")
        self.root.geometry("350x260")
        
        comtypes.CoInitialize()
        self.volume_control = self.get_volume_control()
        
        tk.Label(root, text="Settings", font=("Arial", 12, "bold")).pack(pady=10)
        
        self.entry = tk.Entry(root, justify='center', font=("Arial", 12))
        self.entry.pack(pady=5)
        self.entry.insert(0, self.load_config())
        self.entry.bind("<KeyPress>", self.capture_key)

        tk.Button(root, text="Apply Settings", command=self.apply_hotkey, width=20).pack(pady=5)

        self.status_label = tk.Label(root, text="Status: No Microfone", font=("Arial", 10, "bold"))
        self.status_label.pack(pady=10)

        self.apply_hotkey()
        self.update_loop()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_volume_control(self):
        try:
            devices = AudioUtilities.GetMicrophone()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            return cast(interface, POINTER(IAudioEndpointVolume))
        except:
            return None

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f: return f.read().strip()
            except: pass
        return "f9"

    def capture_key(self, event):
        key = event.keysym.lower()
        mapping = {"control_l": "ctrl", "control_r": "ctrl", "shift_l": "shift", "shift_r": "shift", "caps_lock": "caps lock"}
        key = mapping.get(key, key)
        self.entry.delete(0, tk.END)
        self.entry.insert(0, key)
        return "break"

    def update_loop(self):
        if self.volume_control:
            try:
                is_muted = self.volume_control.GetMute()
                if is_muted:
                    self.status_label.config(text="Micro Off", fg="red")
                else:
                    self.status_label.config(text="Micro on", fg="green")
            except:
                self.volume_control = self.get_volume_control()
        self.root.after(300, self.update_loop)

    def play_system_sound(self, sound_type):
        try:
            if sound_type == "on":
                winsound.Beep(650, )
            else:
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS | winsound.SND_ASYNC)
        except:
            pass

    def toggle_mic(self):
        current_status = self.status_label.cget("text")
        
        try:
            if not self.volume_control:
                self.volume_control = self.get_volume_control()
            
            if current_status == "Micro on":
                self.volume_control.SetMute(True, None)
                self.play_system_sound("off")
            else:
                self.volume_control.SetMute(False, None)
                self.play_system_sound("on")
                
        except Exception:
            self.volume_control = self.get_volume_control()

    def apply_hotkey(self):
        key = self.entry.get().strip()
        keyboard.unhook_all()
        try:
            keyboard.add_hotkey(key, self.toggle_mic, trigger_on_release=True)
            with open(CONFIG_FILE, "w") as f: f.write(key)
        except: pass

    def on_closing(self):
        keyboard.unhook_all()
        comtypes.CoUninitialize()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MicMuteApp(root)
    root.mainloop()
