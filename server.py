#!/usr/bin/env python3
"""
PPT Translator Server
A server that provides PowerPoint translation capabilities using AWS Bedrock models.
Implements the Model Context Protocol (MCP) specification.

For more information about MCP, see: https://modelcontextprotocol.io/
"""

import os
import sys
import logging
import json
import argparse
import traceback
from typing import Dict, Any, Optional, List, Union
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stderr)
    ]
)
logger = logging.getLogger("ppt-translator-mcp")

# Import required libraries with proper error handling
try:
    from pptx import Presentation
except ImportError:
    logger.error("python-pptx not found. Please install it with 'pip install python-pptx'")
    sys.exit(1)

try:
    import boto3
except ImportError:
    logger.error("boto3 not found. Please install it with 'pip install boto3'")
    sys.exit(1)

try:
    from dotenv import load_dotenv
except ImportError:
    logger.error("python-dotenv not found. Please install it with 'pip install python-dotenv'")
    sys.exit(1)

# Try to import fastmcp
try:
    from mcp.server import MCPServer, tool, parameter
    USING_FASTMCP = True
    logger.info("Using mcp.server for MCP implementation")
except ImportError:
    USING_FASTMCP = False
    logger.warning("mcp.server not found. Using fallback MCP implementation.")
    logger.warning("Consider installing fastmcp with 'pip install mcp-server'")

# Load environment variables
load_dotenv()

# Constants
default_target_language = os.getenv('DEFAULT_TARGET_LANGUAGE', 'zh-CN')
AWS_REGION = os.getenv('AWS_REGION', 'us-east-1')
NOVA_MODEL_ID = "amazon.nova-micro-v1:0"

# Language mapping
LANGUAGE_MAP = {
    'zh-CN': '简体中文',
    'zh-TW': '繁体中文',
    'en': '英语',
    'ja': '日语',
    'ko': '韩语',
    'fr': '法语',
    'de': '德语',
    'es': '西班牙语'
}

# Initialize AWS Bedrock client
bedrock_client = None

def init_bedrock_client() -> bool:
    """Initialize the AWS Bedrock client"""
    global bedrock_client
    try:
        bedrock_client = boto3.client(
            'bedrock-runtime',
            aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
            aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'),
            region_name=AWS_REGION
        )
        return True
    except Exception as e:
        logger.error(f"Failed to initialize AWS Bedrock client: {str(e)}")
        return False


def _translate_text(text: str, model_id: str, target_language: str) -> str:
    """Translate using AWS Bedrock Nova Lite"""
    if not bedrock_client:
        if not init_bedrock_client():
            raise Exception("AWS Bedrock client not initialized")
    
    target_lang_name = LANGUAGE_MAP.get(target_language, target_language)
    
    # Check if it's a proper noun or number
    if text.strip().isdigit() or (len(text.strip()) <= 3 and not any(c.isalpha() for c in text)):
        return text
    
    body = json.dumps({
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "text": f"""Translate the following text to {target_lang_name}.
                        
Please follow these rules:
1. Keep all brand names untranslated, such as Amazon, AWS, Bedrock, Nova, etc.
2. Keep all person names untranslated
3. Keep all company names untranslated
4. Keep all product names untranslated
5. Keep all currency amounts untranslated
6. For content that shouldn't be translated, return the original text
7. Only return the translation result, without any explanations or original text

Original text: {text}"""
                    }
                ]
            }
        ],
        "inferenceConfig": {
            "max_new_tokens": 1000,
            "temperature": 0.1
        }
    })
    
    response = bedrock_client.invoke_model(
        body=body,
        modelId=model_id,
        accept="application/json",
        contentType="application/json"
    )
    
    response_body = json.loads(response.get('body').read())
    translated_text = response_body['output']['message']['content'][0]['text'].strip()
    
    # If the translation result contains "important rules" or "don't translate", return the original text
    if "重要规则" in translated_text or "不要翻译" in translated_text:
        return text
        
    return translated_text

def _update_text_frame_with_formatting(text_frame, new_text):
    """Update text frame content while preserving original formatting"""
    try:
        if not text_frame.paragraphs:
            return
        
        # Save the format of the first paragraph
        first_paragraph = text_frame.paragraphs[0]
        if first_paragraph.runs:
            # Save the format of the first run
            first_run = first_paragraph.runs[0]
            font_name = first_run.font.name
            font_size = first_run.font.size
            font_bold = first_run.font.bold
            font_italic = first_run.font.italic
            
            # Default to black
            from pptx.dml.color import RGBColor
            font_color = RGBColor(0, 0, 0)  # Default black
            
            # Safely get the color
            try:
                if hasattr(first_run.font, 'color'):
                    if hasattr(first_run.font.color, 'rgb') and first_run.font.color.rgb:
                        font_color = first_run.font.color.rgb
                    elif hasattr(first_run.font.color, 'type'):
                        # If it's a theme color, we still use default black
                        pass
            except Exception as e:
                logger.error(f"Error getting font color: {str(e)}")
            
            # Clear all paragraphs
            text_frame.clear()
            
            # Add new text and apply original formatting
            paragraph = text_frame.paragraphs[0]
            # The correct method is to call add_run() directly on the paragraph, not on the runs collection
            run = paragraph.add_run()
            run.text = new_text
            
            # Restore formatting
            if font_name:
                run.font.name = font_name
            if font_size:
                run.font.size = font_size
            if font_bold is not None:
                run.font.bold = font_bold
            if font_italic is not None:
                run.font.italic = font_italic
            
            # Set font color
            run.font.color.rgb = font_color
        else:
            # If there are no runs, set the text directly
            text_frame.text = new_text
    except Exception as e:
        logger.error(f"Error updating text formatting: {str(e)}")
        # If formatting fails, set the text directly
        text_frame.text = new_text

def _translate_ppt(input_file: str, output_file: str, target_language: str, model_id: str) -> Dict[str, Any]:
    """Translate a PowerPoint file"""
    try:
        # Load the PPT
        prs = Presentation(input_file)
        translated_count = 0
        
        # Iterate through all slides
        for slide_idx, slide in enumerate(prs.slides):
            logger.info(f"Processing slide {slide_idx + 1}")
            
            # Iterate through all shapes in the slide
            for shape in slide.shapes:
                try:
                    if hasattr(shape, "text") and shape.text.strip():
                        original_text = shape.text.strip()
                        
                        # Translate the text
                        translated_text = _translate_text(original_text, model_id, target_language)
                        
                        if translated_text and translated_text != original_text:
                            # Update the text while preserving formatting
                            if hasattr(shape, 'text_frame') and shape.text_frame:
                                _update_text_frame_with_formatting(shape.text_frame, translated_text)
                            else:
                                shape.text = translated_text
                            
                            translated_count += 1
                            logger.info(f"Translated: '{original_text[:50]}...' -> '{translated_text[:50]}...'")
                except Exception as e:
                    logger.error(f"Error processing shape: {str(e)}")
                    continue
        
        # Save the translated PPT
        prs.save(output_file)
        logger.info(f"Translation completed, saved to: {output_file}")
        
        return {"translated_count": translated_count}
    except Exception as e:
        logger.error(f"Error translating PPT: {str(e)}")
        raise

# Create MCP server
if USING_FASTMCP:
    mcp = MCPServer(
        name="PowerPoint Translator",
        version="1.0.0",
        description="A service that provides PowerPoint translation capabilities using AWS Bedrock models"
    )

    @mcp.tool()
    @parameter("input_file", description="Path to the input PowerPoint file", required=True)
    @parameter("target_language", description="Target language code (e.g., zh-CN, en, ja, ko, fr, de, es)", default=default_target_language)
    @parameter("output_file", description="Path to save the translated PowerPoint file (if not provided, will be auto-generated)")
    @parameter("translation_method", description="Translation method to use: 'nova' (faster) or 'claude' (higher quality)", 
               default="nova", enum=["nova", "claude"])
    def translate_ppt(
        input_file: str,
        target_language: str = default_target_language,
        output_file: str = None,
        translation_method: str = "nova"
    ) -> Dict[str, Any]:
        """Translate a PowerPoint document to the specified language"""
        try:
            if not input_file:
                return {"error": "No input file path provided"}
            
            if not Path(input_file).exists():
                return {"error": f"File does not exist: {input_file}"}
            
            # Generate output filename
            if not output_file:
                input_path = Path(input_file)
                output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{input_path.suffix}")
            
            # Select model based on translation method
            model_id = NOVA_MODEL_ID if translation_method == "nova" else CLAUDE_MODEL_ID
            
            # Execute translation
            result = _translate_ppt(input_file, output_file, target_language, model_id)
            
            return {
                "success": True,
                "input_file": input_file,
                "output_file": output_file,
                "target_language": target_language,
                "translation_method": translation_method,
                "translated_texts_count": result.get('translated_count', 0),
                "message": "PowerPoint translation completed"
            }
            
        except Exception as e:
            logger.error(f"Error processing request: {str(e)}")
            return {"error": f"Processing failed: {str(e)}"}

    @mcp.tool()
    def list_supported_languages() -> Dict[str, str]:
        """List all supported target languages for translation"""
        return {
            "zh-CN": "Simplified Chinese",
            "zh-TW": "Traditional Chinese",
            "en": "English",
            "ja": "Japanese",
            "ko": "Korean",
            "fr": "French",
            "de": "German",
            "es": "Spanish"
        }

else:
    # Fallback implementation for when mcp.server is not available
    class MCPServer:
        """Fallback MCP Server implementation"""
        
        def __init__(self, name, version, description, tools=None):
            """Initialize the MCP server"""
            self.name = name
            self.version = version
            self.description = description
            self.tools = tools or {}
            
        def create_error_response(self, code: int, message: str, req_id: Any = None) -> Dict[str, Any]:
            """Create a JSON-RPC error response"""
            return {
                "jsonrpc": "2.0",
                "error": {
                    "code": code,
                    "message": message
                },
                "id": req_id
            }
        
        def create_success_response(self, result: Any, req_id: Any) -> Dict[str, Any]:
            """Create a JSON-RPC success response"""
            return {
                "jsonrpc": "2.0",
                "result": result,
                "id": req_id
            }
        
        def send_response(self, response: Dict[str, Any]) -> None:
            """Send a response to stdout"""
            sys.stdout.write(json.dumps(response) + "\n")
            sys.stdout.flush()
        
        def handle_describe(self, req_id: Any) -> None:
            """Handle mcp.describe method"""
            tool_schemas = []
            for tool_name, tool_info in self.tools.items():
                tool_schemas.append({
                    "name": tool_name,
                    "description": tool_info.get("description", ""),
                    "parameters": tool_info.get("parameters", {})
                })
                
            response = self.create_success_response({
                "name": self.name,
                "version": self.version,
                "description": self.description,
                "tools": tool_schemas
            }, req_id)
            self.send_response(response)
        
        def handle_invoke(self, params: Dict[str, Any], req_id: Any) -> None:
            """Handle mcp.invoke method"""
            tool_name = params.get("name")
            tool_params = params.get("parameters", {})
            
            if tool_name not in self.tools:
                response = self.create_error_response(
                    -32601,  # Method not found
                    f"Tool not found: {tool_name}",
                    req_id
                )
                self.send_response(response)
                return
            
            try:
                tool_info = self.tools[tool_name]
                result = tool_info["function"](**tool_params)
                response = self.create_success_response(result, req_id)
                self.send_response(response)
            except Exception as e:
                logger.error(f"Error invoking tool {tool_name}: {str(e)}")
                logger.error(traceback.format_exc())
                response = self.create_error_response(
                    -32603,  # Internal error
                    f"Error invoking tool {tool_name}: {str(e)}",
                    req_id
                )
                self.send_response(response)
        
        def handle_request(self, request_line: str) -> bool:
            """Handle a single MCP request
            
            Args:
                request_line: The JSON-RPC request line
                
            Returns:
                bool: True if the request was handled, False if the server should exit
            """
            if not request_line.strip():
                return False  # End of input
            
            try:
                request = json.loads(request_line)
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON: {str(e)}")
                response = self.create_error_response(-32700, "Parse error: Invalid JSON")
                self.send_response(response)
                return True
            
            # Validate JSON-RPC 2.0 request
            if request.get("jsonrpc") != "2.0":
                response = self.create_error_response(
                    -32600,  # Invalid request
                    "Invalid Request: Not a valid JSON-RPC 2.0 request",
                    request.get("id")
                )
                self.send_response(response)
                return True
            
            method = request.get("method")
            params = request.get("params", {})
            req_id = request.get("id")
            
            try:
                if method == "mcp.describe":
                    self.handle_describe(req_id)
                elif method == "mcp.invoke":
                    self.handle_invoke(params, req_id)
                else:
                    response = self.create_error_response(
                        -32601,  # Method not found
                        f"Method not found: {method}",
                        req_id
                    )
                    self.send_response(response)
            except Exception as e:
                logger.error(f"Error handling request: {str(e)}")
                logger.error(traceback.format_exc())
                response = self.create_error_response(
                    -32603,  # Internal error
                    f"Internal error: {str(e)}",
                    req_id
                )
                self.send_response(response)
            
            return True
        
        def run(self) -> None:
            """Run the MCP server"""
            logger.info(f"Starting {self.name} MCP server v{self.version}")
            
            while True:
                try:
                    request_line = sys.stdin.readline()
                    if not self.handle_request(request_line):
                        break
                except KeyboardInterrupt:
                    logger.info("Server stopped by user")
                    break
                except Exception as e:
                    logger.error(f"Unexpected error: {str(e)}")
                    logger.error(traceback.format_exc())
                    break
            
            logger.info("MCP server stopped")

    def tool(name=None, description=None):
        """Decorator for tool functions (fallback implementation)"""
        def decorator(func):
            func._tool_name = name or func.__name__
            func._tool_description = description or func.__doc__
            func._tool_parameters = getattr(func, '_tool_parameters', {})
            return func
        return decorator

    def parameter(name, description=None, required=False, default=None, enum=None):
        """Decorator for tool parameters (fallback implementation)"""
        def decorator(func):
            if not hasattr(func, '_tool_parameters'):
                func._tool_parameters = {}
            
            param_info = {
                "description": description,
                "required": required
            }
            
            if default is not None:
                param_info["default"] = default
            
            if enum is not None:
                param_info["enum"] = enum
                
            func._tool_parameters[name] = param_info
            return func
        return decorator

    # Create fallback MCP server
    mcp = MCPServer(
        name="PowerPoint Translator",
        version="1.0.0",
        description="A service that provides PowerPoint translation capabilities using AWS Bedrock models"
    )

    @tool()
    @parameter("input_file", description="Path to the input PowerPoint file", required=True)
    @parameter("target_language", description="Target language code (e.g., zh-CN, en, ja, ko, fr, de, es)", default=default_target_language)
    @parameter("output_file", description="Path to save the translated PowerPoint file (if not provided, will be auto-generated)")
    @parameter("translation_method", description="Translation method to use: 'nova' (faster) or 'claude' (higher quality)", 
               default="nova", enum=["nova", "claude"])
    def translate_ppt(
        input_file: str,
        target_language: str = default_target_language,
        output_file: str = None,
        translation_method: str = "nova"
    ) -> Dict[str, Any]:
        """Translate a PowerPoint document to the specified language"""
        try:
            if not input_file:
                return {"error": "No input file path provided"}
            
            if not Path(input_file).exists():
                return {"error": f"File does not exist: {input_file}"}
            
            # Generate output filename
            if not output_file:
                input_path = Path(input_file)
                output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{input_path.suffix}")
            
            # Select model based on translation method
            model_id = NOVA_MODEL_ID if translation_method == "nova" else CLAUDE_MODEL_ID
            
            # Execute translation
            result = _translate_ppt(input_file, output_file, target_language, model_id)
            
            return {
                "success": True,
                "input_file": input_file,
                "output_file": output_file,
                "target_language": target_language,
                "translation_method": translation_method,
                "translated_texts_count": result.get('translated_count', 0),
                "message": "PowerPoint translation completed"
            }
            
        except Exception as e:
            logger.error(f"Error processing request: {str(e)}")
            return {"error": f"Processing failed: {str(e)}"}

    @tool()
    def list_supported_languages() -> Dict[str, str]:
        """List all supported target languages for translation"""
        return {
            "zh-CN": "Simplified Chinese",
            "zh-TW": "Traditional Chinese",
            "en": "English",
            "ja": "Japanese",
            "ko": "Korean",
            "fr": "French",
            "de": "German",
            "es": "Spanish"
        }

    # Register tools for fallback implementation
    if not USING_FASTMCP:
        mcp.tools = {
            "translate_ppt": {
                "function": translate_ppt,
                "description": translate_ppt.__doc__,
                "parameters": {
                    "type": "object",
                    "properties": {
                        "input_file": {
                            "type": "string",
                            "description": "Path to the input PowerPoint file"
                        },
                        "target_language": {
                            "type": "string",
                            "description": "Target language code (e.g., zh-CN, en, ja, ko, fr, de, es)",
                            "default": default_target_language
                        },
                        "output_file": {
                            "type": "string",
                            "description": "Path to save the translated PowerPoint file (if not provided, will be auto-generated)"
                        },
                        "translation_method": {
                            "type": "string",
                            "description": "Translation method to use: 'nova' (faster) or 'claude' (higher quality)",
                            "default": "nova",
                            "enum": ["nova", "claude"]
                        }
                    },
                    "required": ["input_file"]
                }
            },
            "list_supported_languages": {
                "function": list_supported_languages,
                "description": list_supported_languages.__doc__,
                "parameters": {
                    "type": "object",
                    "properties": {}
                }
            }
        }

def main():
    """Main function"""
    parser = argparse.ArgumentParser(description='PowerPoint Translator')
    parser.add_argument('--mcp', action='store_true', help='Run in MCP mode')
    parser.add_argument('--translate', action='store_true', help='Translate a PowerPoint file')
    parser.add_argument('--input-file', help='Path to the input PowerPoint file')
    parser.add_argument('--target-language', default=default_target_language, help='Target language code')
    parser.add_argument('--output-file', help='Path to save the translated file')
    parser.add_argument('--model-id', default=NOVA_MODEL_ID, choices=['amazon.nova-micro-v1:0', 'amazon.nova-lite-v1:0'], help='Translation model ID')
    parser.add_argument('--list-languages', action='store_true', help='List supported languages')
    parser.add_argument('--install-deps', action='store_true', help='Install required dependencies')
    parser.add_argument('--use-uv', action='store_true', help='Use uv package manager instead of pip')
    parser.add_argument('--venv', action='store_true', help='Create and use a virtual environment with uv')
    parser.add_argument('--venv-path', help='Path for the virtual environment (default: ./venv)')
    
    args = parser.parse_args()
    
    if args.install_deps:
        try:
            import subprocess
            import shutil
            
            # Check if uv is available
            use_uv = args.use_uv or shutil.which("uv") is not None
            
            logger.info(f"Installing required dependencies using {'uv' if use_uv else 'pip'}...")
            
            if use_uv:
                # Use uv for installation
                cmd = ["uv", "pip", "install", "mcp-server", "python-pptx", "boto3", "python-dotenv"]
                if args.venv:
                    # Create and use a virtual environment
                    venv_path = args.venv_path or os.path.join(os.path.dirname(os.path.abspath(__file__)), "venv")
                    logger.info(f"Creating virtual environment at {venv_path}")
                    subprocess.check_call(["uv", "venv", venv_path])
                    
                    # Determine the Python executable in the virtual environment
                    if os.name == 'nt':  # Windows
                        venv_python = os.path.join(venv_path, "Scripts", "python.exe")
                    else:  # Unix/Linux/Mac
                        venv_python = os.path.join(venv_path, "bin", "python")
                    
                    cmd = [venv_python, "-m", "uv", "pip", "install", "mcp-server", "python-pptx", "boto3", "python-dotenv"]
            else:
                # Use traditional pip
                cmd = [sys.executable, "-m", "pip", "install", "mcp-server", "python-pptx", "boto3", "python-dotenv"]
            
            subprocess.check_call(cmd)
            logger.info("Dependencies installed successfully!")
            return
        except Exception as e:
            logger.error(f"Failed to install dependencies: {str(e)}")
            sys.exit(1)
    
    # Initialize AWS Bedrock client
    init_bedrock_client()
    
    if args.mcp:
        logger.info("Starting PowerPoint Translator in MCP mode...")
        if USING_FASTMCP:
            mcp.run()
        else:
            mcp.run()
    
    elif args.translate:
        if not args.input_file:
            logger.error("Input file is required for translation")
            parser.print_help()
            sys.exit(1)
        
        if not Path(args.input_file).exists():
            logger.error(f"Input file does not exist: {args.input_file}")
            sys.exit(1)
            
        # Generate output filename if not provided
        output_file = args.output_file
        if not output_file:
            input_path = Path(args.input_file)
            output_file = str(input_path.parent / f"{input_path.stem}_translated_{args.target_language}{input_path.suffix}")
            
        try:
            result = _translate_ppt(args.input_file, output_file, args.target_language, args.model_id)
            
            print(f"Translation completed successfully!")
            print(f"Input file: {args.input_file}")
            print(f"Output file: {output_file}")
            print(f"Target language: {args.target_language}")
            print(f"Translation method: {args.method}")
            print(f"Translated {result.get('translated_count', 0)} text elements")
        except Exception as e:
            logger.error(f"Error during translation: {str(e)}")
            sys.exit(1)
    
    elif args.list_languages:
        languages = list_supported_languages()
        print("Supported languages:")
        for code, name in languages.items():
            print(f"  {code}: {name}")
    
    else:
        # Default to MCP mode when no arguments are provided
        logger.info("Starting PowerPoint Translator in MCP mode...")
        mcp.run(transport='stdio')

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        logger.error(traceback.format_exc())
        sys.exit(1)
