#!/usr/bin/env python3
"""
Demo test script for Email Filter and Dashboard System
This script demonstrates the application's functionality with sample data.
"""

import sys
import os
from datetime import datetime, timedelta

# Add the current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_demo_data():
    """Test the demo email data generation."""
    print("Testing Email Filter and Dashboard System")
    print("=" * 50)
    
    # Import the main application
    try:
        from email_filter_dashboard import EmailFilterApp
        
        # Create a test instance
        app = EmailFilterApp()
        
        # Test demo data generation
        start_date = datetime.now() - timedelta(days=30)
        end_date = datetime.now()
        
        print(f"Testing demo data generation...")
        print(f"Date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        
        demo_emails = app.get_demo_emails(start_date, end_date)
        
        print(f"Generated {len(demo_emails)} demo emails:")
        print("-" * 30)
        
        for i, email in enumerate(demo_emails, 1):
            print(f"{i}. From: {email['name']} ({email['email']})")
            print(f"   Subject: {email['subject']}")
            print(f"   Body: {email['body'][:50]}...")
            print()
        
        print("Demo data generation successful!")
        print("\nTo run the full application, execute:")
        print("python email_filter_dashboard.py")
        
    except ImportError as e:
        print(f"Error importing application: {e}")
        print("Make sure email_filter_dashboard.py is in the same directory.")
    except Exception as e:
        print(f"Error testing application: {e}")

if __name__ == "__main__":
    test_demo_data()
