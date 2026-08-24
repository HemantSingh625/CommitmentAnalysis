#!/bin/bash
# Navigate to the folder containing this script
cd "$(dirname "$0")"

echo "=========================================="
echo " Starting Tata Compliance Generator (Mac) "
echo "=========================================="

# Check if python3 is installed
if ! command -v python3 &> /dev/null
then
    echo "ERROR: Python3 is not installed on this Mac."
    echo "Please download and install Python from https://www.python.org/downloads/mac-osx/"
    echo "Press any key to exit..."
    read -n 1
    exit
fi

# Set up a hidden virtual environment so we don't mess with the Mac's system Python
if [ ! -d ".mac_env" ]; then
    echo "First-time setup detected..."
    echo "Installing required dependencies (this might take a minute)..."
    python3 -m venv .mac_env
    source .mac_env/bin/activate
    pip3 install --upgrade pip --quiet
    pip3 install -r requirements.txt --quiet
    echo "Setup complete!"
else
    source .mac_env/bin/activate
fi

# Run the application
echo "Launching Application..."
cd Codes
python3 app.py

echo "Application closed."
