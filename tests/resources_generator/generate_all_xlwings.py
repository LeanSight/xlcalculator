#!/usr/bin/env python3
"""
Generate all Excel files using xlwings with Excel calculations.
This script runs all individual xlwings generators to create Excel files
with proper calculated values for integration testing.

Run this script on Windows with Excel installed.
"""

import os
import sys
import argparse
from pathlib import Path

# Import xlwings generators for dynamic range and xlookup only
from xlwings_xlookup import create_xlookup_excel_with_xlwings
from xlwings_dynamic_range import create_dynamic_range_excel_with_xlwings
from xlwings_dynamic_ranges_comprehensive import create_comprehensive_dynamic_ranges_excel


def generate_all_excel_files(output_dir):
    """Generate all Excel files with xlwings calculations."""
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # List of files to generate - xlookup and comprehensive dynamic ranges
    generators = [
        ("XLOOKUP.xlsx", create_xlookup_excel_with_xlwings),
        ("DYNAMIC_RANGE.xlsx", create_dynamic_range_excel_with_xlwings),
        ("DYNAMIC_RANGES_COMPREHENSIVE.xlsx", create_comprehensive_dynamic_ranges_excel),
    ]
    
    print("🚀 Starting Excel file generation with xlwings...")
    print("📋 This will create 3 Excel files with calculated formula values")
    print("⚠️  Requires Windows with Microsoft Excel installed")
    print("🔧 Using comprehensive DYNAMIC_RANGES generation with faithful Excel behavior")
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
        print(f"   copy {output_dir}\\*.xlsx ..\\resources\\")
        print("   python -m pytest tests/xlfunctions_vs_excel/ -v")
        print("\n📊 Files generated:")
        print("   - XLOOKUP.xlsx: XLOOKUP function with all match modes")
        print("   - DYNAMIC_RANGE.xlsx: INDEX, OFFSET, INDIRECT functions (legacy)")
        print("   - DYNAMIC_RANGES_COMPREHENSIVE.xlsx: Comprehensive dynamic ranges testing")
    
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


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Generate Excel files with xlwings for xlcalculator integration testing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_all_xlwings.py                    # Generate to 'generated_excel_files'
  python generate_all_xlwings.py ../resources       # Generate directly to tests/resources
  python generate_all_xlwings.py C:\\temp\\excel     # Generate to custom Windows path

Requirements:
  - Windows with Microsoft Excel installed
  - xlwings: pip install xlwings
            OR pip install xlcalculator[excel_generation]

Generated files:
  - XLOOKUP.xlsx: XLOOKUP function with all match modes
  - DYNAMIC_RANGE.xlsx: INDEX, OFFSET, INDIRECT functions
        """
    )
    
    parser.add_argument(
        "output_dir",
        nargs="?",
        default="generated_excel_files",
        help="Output directory for generated Excel files (default: generated_excel_files)"
    )
    
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Only check requirements without generating files"
    )
    
    return parser.parse_args()


if __name__ == "__main__":
    print("🔧 Excel File Generator with xlwings")
    print("=" * 60)
    
    # Parse arguments
    args = parse_arguments()
    
    # Check requirements
    if not check_requirements():
        print("\n❌ Requirements not met. Exiting.")
        sys.exit(1)
    
    # If only checking requirements, exit here
    if args.check_only:
        print("\n✅ All requirements met. Ready to generate Excel files.")
        sys.exit(0)
    
    print()
    print(f"📁 Output directory: {os.path.abspath(args.output_dir)}")
    print()
    
    # Generate files
    try:
        created_count, failed_count = generate_all_excel_files(args.output_dir)
        
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