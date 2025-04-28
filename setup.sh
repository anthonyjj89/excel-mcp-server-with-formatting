#!/bin/bash

# Install dependencies
echo "Installing dependencies..."
npm install

# Build the project
echo "Building project..."
npm run build

# Display success message
echo "Setup complete! Excel MCP with formatting capabilities is ready to use."
echo ""
echo "Add the following to your Claude Desktop config file:"
echo ""
echo "{
  \"mcpServers\": {
    \"excel\": {
      \"command\": \"node\",
      \"args\": [
        \"$(pwd)/dist/index.js\"
      ],
      \"env\": {
        \"EXCEL_FILES_PATH\": \"/path/to/excel/files\"
      }
    }
  }
}"
echo ""
echo "Replace '/path/to/excel/files' with your preferred Excel files location."
echo "Then restart Claude Desktop to activate the MCP."
