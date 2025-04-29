#!/bin/bash

# Excel MCP v0.2 Upgrade Script
# This script applies the v0.2 features to the Excel MCP:
# 1. Formula calculation fix
# 2. Auto-fit column width feature

echo "Excel MCP v0.2 Upgrade Script"
echo "============================="
echo ""

# Backup original files
echo "Creating backups..."
cp excel-mcp.js excel-mcp.js.v0.1.1.bak
cp package.json package.json.v0.1.1.bak
echo "✓ Backups created"
echo ""

# Update version number in package.json
echo "Updating version in package.json..."
sed -i '' 's/"version": "0.1.1"/"version": "0.2.0"/' package.json
echo "✓ Version updated to 0.2.0 in package.json"
echo ""

# Update version number in excel-mcp.js
echo "Updating version in excel-mcp.js..."
sed -i '' 's/version: "0.1.1"/version: "0.2.0"/' excel-mcp.js
echo "✓ Version updated to 0.2.0 in excel-mcp.js"
echo ""

# Create a temp directory for merging
mkdir -p temp_upgrade
touch temp_upgrade/temp_file.js

# Extract the end part of the file (after the tool definitions)
echo "Preparing for code integration..."
tail -n 20 excel-mcp.js > temp_upgrade/end_part.js

# Extract the start part of the file (before the tool definitions)
grep -B 1000 "// Register the read_sheet_names tool" excel-mcp.js > temp_upgrade/start_part.js

# Create the file with all the tool definitions
echo "Integrating tool definitions..."
grep -A 1000 "// Register the read_sheet_names tool" excel-mcp.js | grep -B 1000 "  // Connect the server to the transport" > temp_upgrade/original_tools.js

# Extract read tools from the original file
grep -A 100 "// Register the read_sheet_names tool" excel-mcp.js | grep -A 100 "// Register the read_sheet_data tool" > temp_upgrade/read_tools.js

# Create new merged file
echo "Assembling new file with v0.2 features..."
cat temp_upgrade/start_part.js > excel-mcp-v0.2.js
cat temp_upgrade/read_tools.js >> excel-mcp-v0.2.js
cat fix-formula-issues.js >> excel-mcp-v0.2.js
cat temp_upgrade/end_part.js >> excel-mcp-v0.2.js

# Replace the original file
echo "Applying changes..."
mv excel-mcp-v0.2.js excel-mcp.js
chmod +x excel-mcp.js

# Clean up temp files
echo "Cleaning up..."
rm -rf temp_upgrade

echo ""
echo "✅ Upgrade to v0.2 completed successfully!"
echo ""
echo "Changes applied:"
echo "1. Formula calculation fix integrated"
echo "2. Auto-fit column width feature added"
echo "3. Version updated to 0.2.0"
echo ""
echo "See README-v0.2-updates.md for documentation on the new features."
