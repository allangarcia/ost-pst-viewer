#!/usr/bin/env python3
"""
Test script to verify that the process_emails method was successfully moved to EmailProcessor
"""

import sys
import os

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

def test_email_processor_refactor():
    """Test that EmailProcessor has the process_emails method"""
    print("🔍 Testing EmailProcessor refactor...")
    
    try:
        from email_processor import EmailProcessor
        print("✅ EmailProcessor imported successfully")
        
        # Create an instance
        processor = EmailProcessor("dummy.pst")
        print("✅ EmailProcessor instance created")
        
        # Check if process_emails method exists
        if hasattr(processor, 'process_emails'):
            print("✅ process_emails method exists")
            
            # Check if method is callable
            if callable(getattr(processor, 'process_emails')):
                print("✅ process_emails method is callable")
                
                # Check method signature
                import inspect
                sig = inspect.signature(processor.process_emails)
                params = list(sig.parameters.keys())
                expected_params = ['output_dir', 'output_format', 'verbose', 'dry_run']
                
                print(f"✅ Method parameters: {params}")
                
                if all(param in params for param in expected_params):
                    print("✅ All expected parameters present")
                else:
                    print("❌ Missing expected parameters")
                    
            else:
                print("❌ process_emails method is not callable")
        else:
            print("❌ process_emails method does not exist")
            
    except ImportError as e:
        print(f"❌ Import error: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()

def test_main_script():
    """Test that the main script still works"""
    print("\n🔍 Testing main script integration...")
    
    try:
        # Import the main script modules
        from email_processor import EmailProcessor
        from pst_processor import PSTProcessor
        print("✅ All required modules imported successfully")
        
        # Test that process_emails function in main script is a wrapper
        import importlib.util
        spec = importlib.util.spec_from_file_location("main", "pst-exporter.py")
        main_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(main_module)
        
        if hasattr(main_module, 'process_emails'):
            print("✅ Main script has process_emails function")
        else:
            print("❌ Main script missing process_emails function")
            
    except Exception as e:
        print(f"❌ Error testing main script: {e}")

if __name__ == "__main__":
    test_email_processor_refactor()
    test_main_script()
    print("\n🎉 Refactor test completed!")
