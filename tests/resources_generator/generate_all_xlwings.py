#!/usr/bin/env python3
"""
Generate all Excel files using xlwings with Excel calculations.
This script runs all individual xlwings generators to create Excel files
with proper calculated values for integration testing.

Run this script on Windows with Excel installed.
"""

import os
import sys
from pathlib import Path

# Import all xlwings generators
from xlwings_information import create_information_excel_with_xlwings
from xlwings_logical import create_logical_excel_with_xlwings
from xlwings_math import create_math_excel_with_xlwings
from xlwings_text import create_text_excel_with_xlwings


def generate_all_excel_files(output_dir="generated_excel_files"):
    """Generate all Excel files with xlwings calculations."""
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # List of files to generate
    generators = [
        ("INFORMATION.xlsx", create_information_excel_with_xlwings),
        ("logical.xlsx", create_logical_excel_with_xlwings),
        ("MATH.xlsx", create_math_excel_with_xlwings),
        ("TEXT.xlsx", create_text_excel_with_xlwings),
    ]
    
    print("🚀 Starting Excel file generation with xlwings...")
    print("📋 This will create Excel files with calculated formula values")
    print("⚠️  Requires Windows with Microsoft Excel installed")
    print()
    
    created_files = []
    failed_files = []
    
    for filename, generator_func in generators:
        filepath = os.path.join(output_dir, filename)
        try:
            print(f"📝 Generating {filename}...")
            generator_func(filepath)
            created_files.append(filename)
            print(f"✅ Successfully created {filename}")
        except Exception as e:
            print(f"❌ Failed to create {filename}: {e}")
            failed_files.append((filename, str(e)))
        print()
    
    # Summary
    print("=" * 60)
    print("📊 GENERATION SUMMARY")
    print("=" * 60)
    print(f"✅ Successfully created: {len(created_files)} files")
    for filename in created_files:
        print(f"   - {filename}")
    
    if failed_files:
        print(f"\n❌ Failed to create: {len(failed_files)} files")
        for filename, error in failed_files:
            print(f"   - {filename}: {error}")
    
    print(f"\n📁 Output directory: {os.path.abspath(output_dir)}")
    
    if created_files:
        print("\n📋 NEXT STEPS:")
        print("1. Copy the generated Excel files to your xlcalculator project")
        print("2. Place them in: tests/resources/")
        print("3. Run the integration tests to verify Excel compatibility")
        print("\nExample commands:")
        print(f"   copy {output_dir}\\*.xlsx path\\to\\xlcalculator\\tests\\resources\\")
        print("   python -m pytest tests/xlfunctions_vs_excel/ -v")
    
    return len(created_files), len(failed_files)


def check_requirements():
    """Check if xlwings and Excel are available."""
    try:
        import xlwings as xw
        print("✅ xlwings is installed")
    except ImportError:
        print("❌ xlwings is not installed. Install with: pip install xlwings")
        return False
    
    try:
        # Try to start Excel to check if it's available
        app = xw.App(visible=False)
        app.quit()
        print("✅ Microsoft Excel is available")
        return True
    except Exception as e:
        print(f"❌ Microsoft Excel is not available: {e}")
        print("   This script requires Windows with Microsoft Excel installed")
        return False


if __name__ == "__main__":
    print("🔧 Excel File Generator with xlwings")
    print("=" * 60)
    
    # Check requirements
    if not check_requirements():
        print("\n❌ Requirements not met. Exiting.")
        sys.exit(1)
    
    print()
    
    # Generate files
    try:
        created_count, failed_count = generate_all_excel_files()
        
        if failed_count == 0:
            print(f"\n🎉 All {created_count} Excel files generated successfully!")
            sys.exit(0)
        else:
            print(f"\n⚠️  Generated {created_count} files, {failed_count} failed")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n⏹️  Generation cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n💥 Unexpected error: {e}")
        sys.exit(1)