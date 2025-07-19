# BHV Hardware Testing Program

A terminal-based testing program for BHV hardware quality control.

## Setup

1. Install Python 3.7 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Install system dependencies for QR code scanning (macOS):
   ```bash
   brew install zbar
   ```
4. Install Homebrew if not already installed (for macOS)
5. Update the following constants in `hardware_test.py` if needed:
   - `CM5_IP_ADDRESS`: Set to your CM5's IP address (default: "distiller@192.168.0.105")
   - `CM5_PASSWORD`: Set to your CM5's password (default: "one")
   - `RGB_LED_TEST_PATH`: Path of the RGB LED test script on the CM5 (default: "/home/distiller/distiller-cm5-sdk/src/distiller_cm5_sdk/hardware/sam/led_interactive_demo.py")
   - `RGB_LED_NUM_ENTERS`: Number of Enter presses needed for the RGB LED program (default: 10)

## Usage

Run the program:
```bash
python hardware_test.py
```

Follow the on-screen instructions carefully. The program will:
1. Start camera preview window (first run: select camera)
2. Show an interactive menu to select which tests to run (use arrow keys, space to toggle, enter to confirm)
3. Ask for device ID - can scan QR code or type manually
4. Upload firmware to the microcontroller (if T01 is selected)
5. Guide through hardware assembly
6. Perform visual checks (T02-T04)
7. Test UI functionality (T05-T08)
8. Run automated SSH tests with up to 5 retry attempts (T09-T14)
9. Save results to Excel with test durations
10. Save test video recording
11. Offer to safely shutdown the CM5 (press Enter to shutdown, 'n' to skip)

### Features

- **QR Code Scanning**: Automatically scan device QR codes for device ID, version, and manufacture ID
- **Camera Preview**: Live camera view with QR detection and test status overlay
- **Video Recording**: Automatic recording of entire test session
- **Test Selection**: Interactive menu to choose which tests to run (default: all tests)
- **Test Timing**: Each test duration is recorded in seconds
- **SSH Retry**: If SSH connection fails, you get 5 attempts to retry
- **Skip Logic**: Tests marked as skipped show "SKIPPED" in Excel
- **Continue Option**: When a test fails, you can choose to continue or restart
- **Safe Shutdown**: After testing, option to safely shutdown the CM5 via SSH

### QR Code Format

The QR code should contain text in this format:
```
PAMIR_AI_BHV_0.2.4_250708_0195
```
Where:
- `0.2.4` = Version number
- `250708` = Manufacture ID
- `0195` = Device ID (last 4 digits)

## Test Results

Results are saved to `hardware_test_results.xlsx` with:
- Device ID, Version, and Manufacture ID (from QR code if scanned)
- Pass/fail status for each test (PASS/FAIL/SKIPPED/NOT RUN)
- Test duration in seconds for each test
- Overall pass/fail status
- Timestamp
- Any notes or errors
- Failed tests summary
- Clickable link to detailed log file (shows filename, links to full path)
- Clickable link to test video recording (shows filename, links to full path)

Log files are saved in the `logs/` folder.
Video recordings are saved in the `recordings/` folder.

### Excel Color Coding
- Green: PASS
- Red: FAIL
- Yellow: NOT RUN
- Gray: SKIPPED

### Test IDs

Each test has a unique ID:
- T01: Firmware Upload
- T02: CM5 LED Visual Check
- T03: RGB LED Visual Check
- T04: E-ink Display Refresh
- T05: UI Appears
- T06: Button Response
- T07: Voice Transcribed
- T08: Voice Correct
- T09: USB MicroPython Detection
- T10: USB Hub Detection
- T11: USB Media Detection
- T12: RGB LED SSH Test
- T13: SD Card Detection
- T14: Camera Detection

Individual log files are created for each device with the format:
`device_<ID>_log_<timestamp>.txt`

## Troubleshooting

- If firmware upload fails, ensure the board is in bootloader mode
- If SSH fails, check network connection and CM5 credentials
- For USB detection issues, ensure all components are properly connected