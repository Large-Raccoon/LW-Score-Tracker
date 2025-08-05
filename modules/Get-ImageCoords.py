import cv2
import sys
import os
from datetime import datetime

# --- CLI Args Handling ---
if len(sys.argv) < 3:
    print("Usage: python Get-ImageCoords.py <template_path> <image_path> [--debug] [--threshold=0.90]")
    sys.exit(1)

template_path = sys.argv[1]
image_path = sys.argv[2]

debug_mode = False
threshold = 0.8  # Default match threshold

# Parse flags
for arg in sys.argv[3:]:
    if arg == "--debug":
        debug_mode = True
    elif arg.startswith("--threshold="):
        try:
            threshold = float(arg.split("=")[1])
        except ValueError:
            print("Invalid threshold value. Use a number like 0.90.")
            sys.exit(1)

# --- Load images in grayscale ---
template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

if template is None:
    print("Error: Template image not found:", template_path)
    sys.exit(1)
if image is None:
    print("Error: Source image not found:", image_path)
    sys.exit(1)

# --- Template Matching ---
res = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
match_x, match_y = max_loc

# --- Debug Logging ---
if debug_mode:
    try:
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        template_name = os.path.splitext(os.path.basename(template_path))[0]
        image_name = os.path.splitext(os.path.basename(image_path))[0]

        h, w = template.shape[:2]
        debug_img = cv2.cvtColor(image.copy(), cv2.COLOR_GRAY2BGR)  # Convert to color for drawing
        cv2.rectangle(debug_img, (match_x, match_y), (match_x + w, match_y + h), (0, 0, 255), 2)

        debug_filename = f"{image_name}_match_{template_name}_debug.png"
        debug_path = os.path.join(log_dir, debug_filename)
        cv2.imwrite(debug_path, debug_img)

        log_text = (
            f"[{timestamp}]\n"
            f"Template: {template_path}\n"
            f"Image: {image_path}\n"
            f"Confidence: {max_val:.2f}\n"
            f"Match Location: ({match_x}, {match_y})\n"
            f"Debug Image: {debug_path}\n"
            "------------------------\n"
        )

        with open(os.path.join(log_dir, "debug_log.txt"), "a", encoding="utf-8") as log_file:
            log_file.write(log_text)
    except Exception as e:
        print(f"Debug logging failed: {e}", file=sys.stderr)

# --- Threshold filtering ---
if max_val >= threshold:
    print(f"{match_x},{match_y}")
else:
    print("")  # PowerShell will treat this as $null
    sys.exit(0)