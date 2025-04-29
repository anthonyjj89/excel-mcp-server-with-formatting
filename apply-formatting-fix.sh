#!/bin/bash

# Script to apply the formatting fix to excel-mcp.js
# Created by Claude to fix Excel dark mode formatting issues

# Make a backup of the original file
cp excel-mcp.js excel-mcp.js.bak

echo "Created backup of original file as excel-mcp.js.bak"

# Apply the patch
patch -p0 < fix-formatting.patch

if [ $? -eq 0 ]; then
  echo "Successfully applied formatting fix!"
  echo ""
  echo "The fix addresses the following issues:"
  echo "1. Properly handles color formats with # prefix (like #9BC2E6)"
  echo "2. Ensures colors have the required alpha channel (FF) prefix"
  echo "3. Sets both foreground and background colors for better compatibility"
  echo ""
  echo "When using the Excel MCP, prefer light colors like:"
  echo "- Light blue: #BDD7EE"
  echo "- Light green: #C6E0B4"
  echo "- Light yellow: #FFF2CC"
  echo ""
  echo "To test the fix, try formatting cells with these colors."
else
  echo "Error applying patch. Restoring backup..."
  cp excel-mcp.js.bak excel-mcp.js
  echo "Original file restored."
fi
