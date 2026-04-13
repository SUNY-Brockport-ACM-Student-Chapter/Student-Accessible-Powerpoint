#!/usr/bin/env python3
"""
Startup script for the RAG-Enhanced PowerPoint Accessibility App
This script helps start both the ChromaDB API and the PowerPoint app.
"""

import subprocess
import time
import sys
import os
import requests
from threading import Thread

def check_chroma_api():
    """Check if ChromaDB API is running"""
    try:
        response = requests.get("http://localhost:8001/health", timeout=5)
        return response.status_code == 200
    except:
        return False

def start_chroma_api():
    """Start the ChromaDB API in a separate thread"""
    print("🚀 Starting ChromaDB API...")
    try:
        os.chdir("app/chroma-api")
        subprocess.run([sys.executable, "app.py"], check=True)
    except KeyboardInterrupt:
        print("\n⏹️ ChromaDB API stopped")
    except Exception as e:
        print(f"❌ Error starting ChromaDB API: {e}")

def start_powerpoint_app():
    """Start the PowerPoint accessibility app"""
    print("🚀 Starting PowerPoint Accessibility App...")
    try:
        os.chdir("../..")  # Go back to root directory
        subprocess.run([sys.executable, "-m", "streamlit", "run", "app/ppt_notes.py"], check=True)
    except KeyboardInterrupt:
        print("\n⏹️ PowerPoint app stopped")
    except Exception as e:
        print(f"❌ Error starting PowerPoint app: {e}")

def main():
    """Main startup function"""
    print("🎯 RAG-Enhanced PowerPoint Accessibility App")
    print("=" * 50)
    
    # Check if ChromaDB API is already running
    if check_chroma_api():
        print("✅ ChromaDB API is already running on localhost:8001")
        print("🚀 Starting PowerPoint app...")
        start_powerpoint_app()
    else:
        print("⚠️ ChromaDB API not detected")
        print("📋 Choose an option:")
        print("1. Start ChromaDB API only (run this first)")
        print("2. Start PowerPoint app only (if API is already running)")
        print("3. Start both (ChromaDB API + PowerPoint app)")
        
        choice = input("\nEnter your choice (1-3): ").strip()
        
        if choice == "1":
            start_chroma_api()
        elif choice == "2":
            if check_chroma_api():
                start_powerpoint_app()
            else:
                print("❌ ChromaDB API is not running. Please start it first.")
                print("💡 Run: cd app/chroma-api && python app.py")
        elif choice == "3":
            print("🚀 Starting both services...")
            # Start ChromaDB API in background
            api_thread = Thread(target=start_chroma_api, daemon=True)
            api_thread.start()
            
            # Wait a bit for API to start
            print("⏳ Waiting for ChromaDB API to start...")
            time.sleep(3)
            
            # Check if API is ready
            max_attempts = 10
            for i in range(max_attempts):
                if check_chroma_api():
                    print("✅ ChromaDB API is ready!")
                    break
                print(f"⏳ Waiting for API... ({i+1}/{max_attempts})")
                time.sleep(2)
            else:
                print("❌ ChromaDB API failed to start")
                return
            
            # Start PowerPoint app
            start_powerpoint_app()
        else:
            print("❌ Invalid choice. Please run the script again.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n👋 Goodbye!")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
