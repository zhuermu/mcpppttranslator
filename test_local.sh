#!/bin/bash

# This script tests the PowerPoint Translator MCP service locally

# Get the directory of this script
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Install dependencies
echo "Installing dependencies..."
pip install -r "$SCRIPT_DIR/requirements.txt"

# Create a test PowerPoint file if it doesn't exist
if [ ! -f "$SCRIPT_DIR/test.pptx" ]; then
  echo "Creating test PowerPoint file..."
  python -c "
from pptx import Presentation
from pptx.util import Inches

# Create a presentation
prs = Presentation()

# Add a slide with a title and content
slide_layout = prs.slide_layouts[1]  # Title and content layout
slide = prs.slides.add_slide(slide_layout)

# Set the title
title = slide.shapes.title
title.text = 'Test PowerPoint for Translation'

# Add content
content = slide.placeholders[1]
content.text = 'This is a test PowerPoint file for translation.\n\nIt contains some text that will be translated to another language.\n\nAWS Bedrock is a powerful service for AI/ML tasks.'

# Save the presentation
prs.save('$SCRIPT_DIR/test.pptx')
print('Test PowerPoint file created at $SCRIPT_DIR/test.pptx')
"
fi

# Run the MCP server
echo "Starting PowerPoint Translator MCP server..."
node "$SCRIPT_DIR/index.js"
