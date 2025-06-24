# PowerPoint Translator MCP Service

A Model Context Protocol (MCP) service that provides PowerPoint translation capabilities using AWS Bedrock models.

## Features

- Translate PowerPoint presentations to multiple languages
- Preserve formatting during translation
- Support for AWS Bedrock Nova models
- Intelligent handling of proper nouns, brand names, and special content
- Optimized MCP implementation with fallback support

## Supported Languages

- Simplified Chinese (zh-CN)
- Traditional Chinese (zh-TW)
- English (en)
- Japanese (ja)
- Korean (ko)
- French (fr)
- German (de)
- Spanish (es)

## Prerequisites

- Python 3.8+
- AWS account with Bedrock access
- AWS credentials configured

## Installation

### Using pip (Python)

```bash
# Clone the repository
git clone https://github.com/yourusername/mcpppttranslator.git
cd mcpppttranslator

# Install dependencies
pip install -r requirements.txt

# Or use the built-in helper
python server.py --install-deps
```

### Using uv (Fast Python Package Manager)

```bash
# Clone the repository
git clone https://github.com/yourusername/mcpppttranslator.git
cd mcpppttranslator

# Install dependencies with uv
python server.py --install-deps --use-uv

# Or create a virtual environment with uv
python server.py --install-deps --use-uv --venv

# Specify a custom virtual environment path
python server.py --install-deps --use-uv --venv --venv-path /path/to/custom/venv
```

### Using Amazon Q Configuration

Add the following to your Amazon Q configuration:

```json
"mcpServers": {
  "ppt-translator": {
    "timeout": 60,
    "type": "stdio",
    "command": "python",
    "args": [
      "/path/to/mcpppttranslator/server.py",
      "--mcp"
    ],
    "env": {
      "AWS_ACCESS_KEY_ID": "${AWS_ACCESS_KEY_ID}",
      "AWS_SECRET_ACCESS_KEY": "${AWS_SECRET_ACCESS_KEY}",
      "AWS_REGION": "us-east-1",
      "DEFAULT_TARGET_LANGUAGE": "zh-CN"
    }
  }
}
```

### Direct Python Execution

You can also run the server directly using Python:

```bash
# Run the server in MCP mode
python server.py --mcp

# Or translate a file directly
python server.py --translate --input-file presentation.pptx --target-language ja

# List supported languages
python server.py --list-languages
```

## Usage

Once the MCP server is running, you can use it with any MCP-compatible client like Amazon Q.

### Available Tools

1. `translate_ppt` - Translate a PowerPoint document
   - Parameters:
     - `input_file`: Path to the input PowerPoint file (required)
     - `target_language`: Target language code (default: zh-CN)
     - `output_file`: Path to save the translated file (optional)
     - `translation_method`: Translation method, 'nova' or 'claude' (default: nova)

2. `list_supported_languages` - List all supported target languages

## Example

```
Translate my presentation.pptx to Japanese using the Nova model
```

## Command Line Arguments

The server.py script supports the following command line arguments:

- `--mcp`: Run in MCP mode
- `--translate`: Translate a PowerPoint file
- `--input-file`: Path to the input PowerPoint file
- `--target-language`: Target language code
- `--output-file`: Path to save the translated file
- `--model-id`: Translation model ID (choices: 'amazon.nova-micro-v1:0', 'amazon.nova-lite-v1:0')
- `--list-languages`: List supported languages
- `--install-deps`: Install required dependencies
- `--use-uv`: Use uv package manager instead of pip
- `--venv`: Create and use a virtual environment with uv
- `--venv-path`: Path for the virtual environment (default: ./venv)

## Development

### Local Development

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/mcpppttranslator.git
   cd mcpppttranslator
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the server:
   ```bash
   python server.py --mcp
   ```

## License

MIT
