# 🤖 INSTRUCTIONS: Activate ATDD Configuration in ONA Environment

## 🎯 ACTIVATION STEPS

### 1. **Copy Files to Project**
Place these files in the **project root** where ONA will work:

```bash
# Mandatory files in project root:
ONA_ATDD_CONFIG.md              # Absolute rules
ONA_DECISION_PROTOCOL.md        # Decision protocol  
ONA_SYSTEM_PROMPT_ATDD.md       # Agent configuration
ONA_ATDD_EXAMPLES.md            # Practical examples
.ona-config                     # Automatic configuration
activate-ona-atdd.sh            # Activation script
```

### 2. **Run Activation Script**
```bash
./activate-ona-atdd.sh
```

### 3. **Verify Activation**
```bash
ls ONA_* && echo "✅ ATDD configuration ready"
```

## 🚨 INSTRUCTIONS FOR ONA

### **Upon Starting Session:**
1. **READ MANDATORILY**: `ONA_ATDD_CONFIG.md`
2. **ALWAYS CONSULT**: `ONA_DECISION_PROTOCOL.md` before handling failing tests
3. **FOLLOW EXAMPLES**: `ONA_ATDD_EXAMPLES.md` for correct behavior

### **When You Encounter Failing Test:**
```
🚨 MANDATORY PROTOCOL:
1. STOP - Don't create new test
2. READ - ONA_DECISION_PROTOCOL.md  
3. VALIDATE - Test against Excel documentation
4. FIX - Implementation to make test pass
5. COMMIT - Immediate when test passes
```

### **Verification Commands:**
```bash
# Verify configuration
cat ONA_ATDD_CONFIG.md | head -20

# View decision protocol  
cat ONA_DECISION_PROTOCOL.md | grep -A 10 "PROTOCOL"

# Consult examples
cat ONA_ATDD_EXAMPLES.md | grep -A 5 "CORRECT BEHAVIOR"
```

## 📁 FINAL STRUCTURE

```
/workspace/your-project/
├── ONA_ATDD_CONFIG.md          # ← ONA reads automatically
├── ONA_DECISION_PROTOCOL.md    # ← Mandatory protocol
├── ONA_SYSTEM_PROMPT_ATDD.md   # ← Agent configuration  
├── ONA_ATDD_EXAMPLES.md        # ← Reference examples
├── .ona-config                 # ← Automatic configuration
├── activate-ona-atdd.sh        # ← Activation script
├── README.md                   # ← Update with references
└── docs/
    └── _DEV_STANDARDS.md       # ← Already updated
```

## ✅ FUNCTIONALITY VERIFICATION

### **Configuration Test:**
1. Start ONA in the project
2. Ask: "What should I do if I encounter a failing test?"
3. **Expected response**: ONA should mention validate test and fix implementation, NOT create new test

### **Success Indicators:**
- ✅ ONA consults ONA_* files before changes
- ✅ ONA validates tests against Excel documentation
- ✅ ONA fixes implementation instead of creating new tests
- ✅ ONA commits immediately when test passes

### **Problem Indicators:**
- ❌ ONA creates new test when existing test fails
- ❌ ONA modifies test without validating against Excel
- ❌ ONA avoids fixing implementation

## 🎯 EXPECTED RESULT

With this configuration, ONA will strictly follow ATDD principles:
- **Test fails** → **Validate test** → **Fix implementation** → **Test passes** → **Commit**

---

**Status**: READY FOR ACTIVATION ✅  
**Next step**: Copy files and run `./activate-ona-atdd.sh`