# Smart Print Manager

A modern, standalone desktop utility for advanced manual duplexing and intelligent PDF print management on Windows.

## Overview

Standard operating system print dialogs often fall short when dealing with home or office printers that lack automatic double-sided (duplex) capabilities. Attempting to manually print odd and even pages for large documents frequently results in misaligned pages, incorrect orientations, and wasted ink.

Smart Print Manager bridges this gap. It provides a dedicated, intelligent engine that handles the complex logic of manual duplexing, page rotation, and selective printing, all wrapped in a clean, modern user interface.

## Core Features

- **Intelligent Manual Duplexing:** Automatically splits print jobs into sequential odd and even passes. It intelligently pauses your print queue, prompts you to flip the physical paper stack, and automatically pads odd-numbered print jobs with a blank page to ensure the final sheet aligns perfectly.
- **Advanced Range & Step Sizing:** Granular control over your print jobs. Input complex page ranges (e.g., `1-15, 18, 20-22`) and apply specific step sizes to print exactly what you need.
- **Auto-Orientation Engine:** Dynamically analyzes the aspect ratio of each individual PDF page against the physical paper feed dimensions of your selected printer. If a mismatch is detected (e.g., a landscape page feeding into a portrait tray), the software rotates the page on the fly to prevent content cutoff.
- **Dynamic Page Stamping:** An optional toggle that digitally draws the native PDF page number onto the bottom right corner of the physical output. Highly useful for large, un-stapled document stacks.
- **Hardware Duplex Passthrough:** Seamlessly hands off jobs to higher-end printers with built-in duplexing while still utilizing the application's advanced range parsing and page stamping features.
- **Modern Interface:** Built with CustomTkinter for a sleek, responsive, and flat design language.

## Prerequisites

- Windows Operating System
- Python 3.8 or higher
- A valid local or network printer installed on the host machine

## Installation for Developers

If you wish to run the tool from the source code or contribute to the repository:

1. Clone this repository to your local machine:

    git clone https://github.com/yourusername/smart-print-manager.git
    cd smart-print-manager

2. Install the required Python dependencies:

    pip install customtkinter PyMuPDF Pillow pywin32

3. Run the application:

    python smart_print.py

## Compiling to an Executable

To build a standalone `.exe` file that can be run on any Windows machine without requiring a Python environment:

1. Ensure `pyinstaller` is installed:

    pip install pyinstaller

2. Run the build command from the root directory:

    pyinstaller --noconsole --onefile smart_print.py

3. The compiled application will be available in the newly created `dist` folder as `smart_print.exe`.

If you would like to add a custom icon:

    pyinstaller --noconsole --onefile --icon=your_icon.ico smart_print.py

## Usage Guide

1. Launch the application and select your target PDF document.
2. Choose your target printer from the dropdown menu.
3. Define your target pages (leave as `all` for the entire document) and adjust the step size if necessary.
4. Select your preferred orientation logic (Auto-Rotate is recommended for most users).
5. Toggle your duplex settings based on your hardware capabilities. If using manual duplex, ensure **Auto-add blank page** is checked to maintain stack integrity.
6. Click **Print Document**. If using manual duplex mode, follow the on-screen prompt to flip your paper stack when the first pass completes.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.