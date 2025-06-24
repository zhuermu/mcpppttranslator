# PowerPoint Translator MCP Service

A Model Context Protocol (MCP) service that provides PowerPoint translation capabilities using AWS Bedrock models.

## Features

- Translate PowerPoint presentations to multiple languages
- Preserve formatting during translation
- Support for multiple translation engines (Nova Lite and Claude)
- Intelligent handling of proper nouns, brand names, and special content

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

- Node.js 14+ (for npm package installation)
- Python 3.8+
- AWS account with Bedrock access
- AWS credentials configured

## Installation

### Using npm

```bash
npm install -g @modelcontextprotocol/server-ppt-translator
```

### Using Amazon Q Configuration

Add the following to your Amazon Q configuration:

```json
"mcpServers": {
  "ppt-translator": {
    "timeout": 60,
    "type": "stdio",
    "command": "npx",
    "args": [
      "-y",
      "@modelcontextprotocol/server-ppt-translator@latest"
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

## Usage

Once the MCP server is running, you can use it with any MCP-compatible client like Amazon Q or Claude Desktop.

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
Translate my presentation.pptx to Japanese using the Claude model
```

## Development

### Local Development

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/mcpppttranslator.git
   cd mcpppttranslator
   ```

2. Install dependencies:
   ```bash
   npm install
   pip install -r requirements.txt
   ```

3. Run the server:
   ```bash
   node index.js
   ```

### Publishing to npm

1. Update version in package.json
2. Login to npm:
   ```bash
   npm login
   ```
3. Publish:
   ```bash
   npm publish --access public
   ```

## License

MIT
