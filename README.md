# PowerPoint Translator MCP Service

A Model Context Protocol (MCP) service that provides PowerPoint translation capabilities using AWS Bedrock models.

## Features

- Translate PowerPoint presentations to multiple languages
- Preserve formatting during translation
- Support for AWS Bedrock Nova models
- Intelligent handling of proper nouns, brand names, and special content
- Full MCP protocol compliance with proper initialization and tool handling
- Fallback implementation when MCP library is not available

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
- AWS credentials configured (via `aws configure` or environment variables)

## Installation

### Quick Setup

```bash
# Clone the repository
git clone https://github.com/yourusername/mcpppttranslator.git
cd mcpppttranslator

# Install dependencies
python server.py --install-deps

# Or use uv for faster installation
python server.py --install-deps --use-uv
```

### Virtual Environment Setup (Recommended)

```bash
# Create virtual environment with uv
python server.py --install-deps --use-uv --venv

# Or specify custom path
python server.py --install-deps --use-uv --venv --venv-path /path/to/custom/venv
```

## MCP Configuration

### Option 1: AWS CLI Configuration (Recommended)

First, configure your AWS credentials:
```bash
aws configure
```

Then add to your Amazon Q configuration:

```json
{
  "mcpServers": {
    "ppt-translator": {
      "timeout": 300,
      "type": "stdio",
      "command": "python",
      "args": [
        "/path/to/mcpppttranslator/server.py",
        "--mcp"
      ],
      "env": {
        "AWS_REGION": "us-east-1",
        "DEFAULT_TARGET_LANGUAGE": "zh-CN"
      }
    }
  }
}
```

### Option 2: Direct Credential Configuration

```json
{
  "mcpServers": {
    "ppt-translator": {
      "timeout": 300,
      "type": "stdio",
      "command": "python",
      "args": [
        "/path/to/mcpppttranslator/server.py",
        "--mcp"
      ],
      "env": {
        "AWS_ACCESS_KEY_ID": "your-access-key",
        "AWS_SECRET_ACCESS_KEY": "your-secret-key",
        "AWS_REGION": "us-east-1",
        "DEFAULT_TARGET_LANGUAGE": "zh-CN"
      }
    }
  }
}
```

## Usage

### With Amazon Q

Once configured, you can use natural language commands:

```
Translate my presentation.pptx to Japanese using the Nova model
```

### Command Line Usage

```bash
# Run in MCP mode
python server.py --mcp

# Direct translation
python server.py --translate --input-file presentation.pptx --target-language ja

# List supported languages
python server.py --list-languages
```

## Available Tools

1. **translate_ppt** - Translate a PowerPoint document
   - `input_file`: Path to the input PowerPoint file (required)
   - `target_language`: Target language code (default: zh-CN)
   - `output_file`: Path to save the translated file (optional)
   - `translation_method`: Translation method, 'nova' or 'claude' (default: nova)

2. **list_supported_languages** - List all supported target languages

## Command Line Arguments

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

## Testing

Test the MCP server functionality:

```bash
python test_mcp.py
```

## Troubleshooting

### Common Issues

1. **AWS Credentials**: Ensure AWS credentials are properly configured via `aws configure` or environment variables.

2. **Missing Dependencies**: Run `python server.py --install-deps` to install required packages.

3. **MCP Protocol Errors**: The server implements full MCP protocol compliance. If you encounter initialization errors, ensure you're using a compatible MCP client.

4. **File Permissions**: Ensure the server.py file has execute permissions: `chmod +x server.py`

## Development

### Local Development

1. Clone and setup:
   ```bash
   git clone https://github.com/yourusername/mcpppttranslator.git
   cd mcpppttranslator
   python server.py --install-deps --use-uv --venv
   ```

2. Test the server:
   ```bash
   python test_mcp.py
   ```

### MCP Protocol Compliance

This server implements the full MCP protocol specification including:

- Proper `initialize` method handling
- Standard `tools/list` and `tools/call` methods
- Correct JSON-RPC 2.0 response formatting
- Error handling and graceful degradation
- Backward compatibility with legacy MCP methods

## License

MIT
