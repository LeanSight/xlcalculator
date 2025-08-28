# xlcalculator - NumPy 2.0 Compatible Fork

![Python 3.13+](https://img.shields.io/badge/Python-3.13+-blue.svg)
![NumPy 2.0+](https://img.shields.io/badge/NumPy-2.0+-green.svg)
![Status Fork](https://img.shields.io/badge/Status-Unofficial%20Fork-orange.svg)

## 🚨 Important Notice

This is an **UNOFFICIAL FORK** of xlcalculator, maintained by LeanSight for NumPy 2.0 and Python 3.13 compatibility.

**Original Repository:** https://github.com/bradbase/xlcalculator  
**Fork Repository:** https://github.com/LeanSight/xlcalculator

## 🎯 Why This Fork?

The original xlcalculator has a dependency restriction `numpy<2`, making it incompatible with:
- NumPy 2.0+ (released June 2024)  
- Python 3.13+ (latest Python release)
- Modern scientific Python stack

This fork resolves these compatibility issues while maintaining 100% API compatibility.

## ✨ Fork Features

- ✅ **NumPy 2.0+ Support** - Compatible with NumPy 2.1.0 through 2.3.2+
- ✅ **Python 3.13 Validated** - Fully tested on Python 3.13.0
- ✅ **Modern Dependencies** - Updated pandas, scipy, openpyxl to latest versions
- ✅ **YEARFRAC Function** - Includes LeanSight yearfrac fork for complete Excel compatibility
- ✅ **All Tests Pass** - Zero regressions from original functionality
- ✅ **Easy Migration** - Drop-in replacement for original xlcalculator

## 📦 Installation

### Basic Installation (Core Functions)
```bash
pip install git+https://github.com/LeanSight/xlcalculator.git
```

### Full Installation (Including YEARFRAC)
```bash
pip install git+https://github.com/LeanSight/xlcalculator.git[excel_functions]
```

### Development Installation
```bash
git clone https://github.com/LeanSight/xlcalculator.git
cd xlcalculator
pip install -e .[test,excel_functions]
```

## 🔧 Requirements

- **Python:** 3.13+ (validated version)
- **NumPy:** 2.1.0+ 
- **pandas:** 2.3.0+
- **scipy:** 1.14.1+
- **openpyxl:** Latest compatible

## 🧪 Validation Status

This fork has been thoroughly tested with:

| Component | Version | Status |
|-----------|---------|--------|
| Python | 3.13.0 | ✅ Validated |
| NumPy | 2.3.2 | ✅ All tests pass |
| pandas | 2.3.2 | ✅ Compatible |
| scipy | 1.14.1 | ✅ Compatible |
| xlcalculator tests | All | ✅ 100% pass rate |
| YEARFRAC tests | All | ✅ 100% pass rate |

## 🔄 Migration from Original

This fork is a **drop-in replacement**. Simply change your installation:

```bash
# Before (original):
pip install xlcalculator

# After (fork):
pip install git+https://github.com/LeanSight/xlcalculator.git
```

**No code changes required** - all APIs remain identical.

## 🚀 Usage

Usage is identical to original xlcalculator:

```python
from xlcalculator import ModelCompiler
from xlcalculator import Model

# Load Excel file
compiler = ModelCompiler()
model = compiler.read_and_parse_archive("example.xlsx")

# Evaluate cells
result = model.evaluate("Sheet1!A1")
```

## 🔗 Dependencies

This fork includes these updated dependencies:

### Core Dependencies
- `numpy>=2.1.0` (was restricted to `<2`)
- `pandas>=2.3.0` (was `>=2.0.0`)  
- `scipy>=1.14.1` (was unspecified)
- `openpyxl` (latest)
- `numpy-financial` (latest)
- `jsonpickle` (latest)

### Excel Functions (Optional)
- `git+https://github.com/LeanSight/yearfrac.git` (NumPy 2.0 compatible fork)

## 🧩 Related Forks

This xlcalculator fork depends on:
- **LeanSight yearfrac fork:** https://github.com/LeanSight/yearfrac
  - Adds NumPy 2.0 compatibility to yearfrac
  - Enables YEARFRAC Excel function support

## 📋 Testing

Run tests to verify your installation:

```bash
# Basic test
python -c "import xlcalculator; print('✅ xlcalculator imported successfully')"

# Full test suite (if installed with [test])
python -m pytest tests/

# Test NumPy compatibility
python -c "import numpy as np; print(f'✅ NumPy {np.__version__}')"
```

## ⚠️ Known Limitations

- **Python Support:** Only Python 3.13+ is validated (may work on 3.9-3.12 but not tested)
- **Platform:** Primarily validated on Windows, should work on Linux/macOS
- **Excel Functions:** Some advanced Excel functions may not be supported (same as original)

## 🆘 Support

### For Fork-Specific Issues:
- **Issues:** https://github.com/LeanSight/xlcalculator/issues
- **Discussions:** Use GitHub Discussions on the fork repo

### For Original Functionality:
- **Documentation:** Refer to original xlcalculator documentation
- **Excel Functions:** Check original function support list

## 🤝 Contributing

Contributions welcome! Please:
1. Fork this repository (not the original)
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

## 📄 License

MIT License - Same as original xlcalculator

## 🙏 Credits

- **Original Author:** Bradley van Ree
- **Original Repository:** https://github.com/bradbase/xlcalculator  
- **Fork Maintainer:** LeanSight
- **yearfrac Original:** https://github.com/kmedian/yearfrac

## 📈 Version Information

**Fork Version:** `0.5.1+numpy2.python313`  
**Based on Original:** `0.5.1.dev0`  
**Last Updated:** 2025-08-28  
**Validation Date:** 2025-08-28