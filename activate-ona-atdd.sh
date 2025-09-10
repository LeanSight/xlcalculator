#!/bin/bash
# Script to activate ATDD configuration in ONA environment

echo "🤖 Activating ATDD configuration for ONA agent..."

# Verify configuration files
echo "📋 Verifying configuration files..."
files=("ONA_ATDD_CONFIG.md" "ONA_DECISION_PROTOCOL.md" "ONA_SYSTEM_PROMPT_ATDD.md" "ONA_ATDD_EXAMPLES.md")

for file in "${files[@]}"; do
    if [ -f "$file" ]; then
        echo "✅ $file - FOUND"
    else
        echo "❌ $file - MISSING"
        exit 1
    fi
done

# Create symbolic link in docs directory if it exists
if [ -d "docs" ]; then
    echo "📁 Creating links in docs directory..."
    for file in "${files[@]}"; do
        ln -sf "../$file" "docs/$file" 2>/dev/null
    done
fi

# Show activation message
echo ""
echo "🚨 ATDD CONFIGURATION ACTIVATED"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "CRITICAL RULE: Test fails → Validate → Fix implementation"
echo "PROHIBITED: Create new test to avoid problem"
echo "MANDATORY: Consult ONA_ATDD_CONFIG.md before changes"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "✅ Configuration ready for ONA agent"