#!/usr/bin/env python3
"""
Setup script to install and verify all dependencies
"""

import subprocess
import sys
import os

print("=" * 60)
print("SETUP: Installing Dependencies")
print("=" * 60)

# Get the requirements file
script_dir = os.path.dirname(os.path.abspath(__file__))
requirements_file = os.path.join(script_dir, "requirements.txt")

if not os.path.exists(requirements_file):
    print(f"\n✗ requirements.txt not found at {requirements_file}")
    sys.exit(1)

print(f"\nInstalling from: {requirements_file}\n")

# Install requirements
try:
    subprocess.check_call([
        sys.executable, 
        "-m", 
        "pip", 
        "install", 
        "-r", 
        requirements_file,
        "--upgrade"
    ])
    print("\n✓ Dependencies installed successfully!")
except subprocess.CalledProcessError as e:
    print(f"\n✗ Failed to install dependencies: {e}")
    sys.exit(1)

# Verify installation
print("\n" + "=" * 60)
print("VERIFY: Checking Installed Packages")
print("=" * 60 + "\n")

packages = [
    'pptx',
    'google.generativeai',
    'dotenv',
    'flask',
    'flask_cors'
]

all_installed = True
for package in packages:
    try:
        __import__(package)
        print(f"✓ {package} is installed")
    except ImportError:
        print(f"✗ {package} is NOT installed")
        all_installed = False

if all_installed:
    print("\n✓ All required packages are installed!")
    print("\n" + "=" * 60)
    print("NEXT STEPS:")
    print("=" * 60)
    print("\n1. Set up your .env file with GEMINI_API_KEY")
    print("2. Run the app: python Frontend/app.py")
    print("3. Test generation with: python test_fix.py")
else:
    print("\n✗ Some packages are missing!")
    sys.exit(1)
