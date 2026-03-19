# boot.py — must be on CIRCUITPY root for settings.json to be writable
# Disables USB drive when not in BOOTSEL mode so firmware can write to flash
import storage
storage.remount("/", readonly=False)
