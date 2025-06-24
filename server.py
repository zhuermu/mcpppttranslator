#!/usr/bin/env python3
"""
PPT Translator Server
A server that provides PowerPoint translation capabilities using AWS Bedrock models.
"""

import os
import sys
import logging
import json
import argparse
from typing import Dict, Any, Optional, List
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("Error: python-pptx not found. Please install it with 'pip install python-pptx'")
    sys.exit(1)

try:
    import boto3
except ImportError:
    print("Error: boto3 not found. Please install it with 'pip install boto3'")
    sys.exit(1)

try:
    from dotenv import load_dotenv
except ImportError:
    print("Error: python-dotenv not found. Please install it with 'pip install python-dotenv'")
    sys.exit(1)

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize AWS Bedrock client
try:
    bedrock_client = boto3.client(
        'bedrock-runtime',
        aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
        aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'),
        region_name=os.getenv('AWS_REGION', 'us-east-1')
    )
except Exception as e:
    logger.error(f"Failed to initialize AWS Bedrock client: {str(e)}")
    print(f"Error: Failed to initialize AWS Bedrock client. Please check your AWS credentials.")
    sys.exit(1)

default_target_language = os.getenv('DEFAULT_TARGET_LANGUAGE', 'zh-CN')

def _translate_text(text: str, target_language: str, method: str) -> Optional[str]:
    """Translate a single text"""
    try:
        if method == 'claude':
            return _translate_with_claude(text, target_language)
        else:
            return _translate_with_nova(text, target_language)
    except Exception as e:
        logger.error(f"Failed to translate text: {str(e)}")
        return None

def _translate_with_claude(text: str, target_language: str) -> str:
    """Translate using AWS Bedrock Claude 3.5 Sonnet"""
    language_map = {
        'zh-CN': '简体中文',
        'zh-TW': '繁体中文',
        'en': '英语',
        'ja': '日语',
        'ko': '韩语',
        'fr': '法语',
        'de': '德语',
        'es': '西班牙语'
    }
    
    target_lang_name = language_map.get(target_language, target_language)
    
    body = json.dumps({
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": 1000,
        "messages": [
            {
                "role": "user",
                "content": f"Please translate the following text to {target_lang_name}. Keep proper nouns, brand names, person names, place names, company names, and currencies untranslated. Only return the translation result:\n\n{text}"
            }
        ]
    })
    
    response = bedrock_client.invoke_model(
        body=body,
        modelId="anthropic.claude-3-5-sonnet-20241022-v2:0",
        accept="application/json",
        contentType="application/json"
    )
    
    response_body = json.loads(response.get('body').read())
    return response_body['content'][0]['text'].strip()

def _translate_with_nova(text: str, target_language: str) -> str:
    """Translate using AWS Bedrock Nova Lite"""
    language_map = {
        'zh-CN': '简体中文',
        'zh-TW': '繁体中文',
        'en': '英语',
        'ja': '日语',
        'ko': '韩语',
        'fr': '法语',
        'de': '德语',
        'es': '西班牙语'
    }
    
    target_lang_name = language_map.get(target_language, target_language)
    
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
        modelId="amazon.nova-lite-v1:0",
        accept="application/json",
        contentType="application/json"
    )
    
    response_body = json.loads(response.get('body').read())
    translated_text = response_body['output']['message']['content'][0]['text'].strip()
    
    # If the translation result contains "important rules" or "don't translate", return the original text
    if "重要规则" in translated_text or "不要翻译" in translated_text:
        return text
        
    return translated_text

def _translate_ppt(input_file: str, output_file: str, target_language: str, method: str) -> Dict[str, Any]:
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
                        translated_text = _translate_text(original_text, target_language, method)
                        
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

def translate_ppt(
    input_file: str,
    target_language: str = default_target_language,
    output_file: str = None,
    translation_method: str = "nova"
) -> Dict[str, Any]:
    """Translate a PowerPoint document to the specified language
    
    Args:
        input_file: Path to the input PowerPoint file
        target_language: Target language code, default is zh-CN
        output_file: Path to save the translated file, if not provided it will be auto-generated
        translation_method: Translation method, can be 'nova' or 'claude'
        
    Returns:
        Dict[str, Any]: Translation result information
    """
    try:
        if not input_file:
            return {"error": "No input file path provided"}
        
        if not Path(input_file).exists():
            return {"error": f"File does not exist: {input_file}"}
        
        # Generate output filename
        if not output_file:
            input_path = Path(input_file)
            output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{input_path.suffix}")
        
        # Execute translation
        result = _translate_ppt(input_file, output_file, target_language, translation_method)
        
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

def list_supported_languages() -> Dict[str, str]:
    """List all supported target languages for translation
    
    Returns:
        Dict[str, str]: Dictionary of language codes and their names
    """
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

# MCP protocol implementation
def handle_mcp_request():
    """Handle MCP protocol requests"""
    try:
        # Read request from stdin
        request_line = sys.stdin.readline().strip()
        request = json.loads(request_line)
        
        # Process the request
        if request.get("type") == "invoke":
            tool_name = request.get("name")
            params = request.get("params", {})
            
            if tool_name == "translate_ppt":
                result = translate_ppt(
                    input_file=params.get("input_file"),
                    target_language=params.get("target_language", default_target_language),
                    output_file=params.get("output_file"),
                    translation_method=params.get("translation_method", "nova")
                )
                
                # Send response
                response = {
                    "type": "result",
                    "result": result
                }
                sys.stdout.write(json.dumps(response) + "\n")
                sys.stdout.flush()
                
            elif tool_name == "list_supported_languages":
                result = list_supported_languages()
                
                # Send response
                response = {
                    "type": "result",
                    "result": result
                }
                sys.stdout.write(json.dumps(response) + "\n")
                sys.stdout.flush()
                
            else:
                # Unknown tool
                response = {
                    "type": "error",
                    "error": f"Unknown tool: {tool_name}"
                }
                sys.stdout.write(json.dumps(response) + "\n")
                sys.stdout.flush()
                
        elif request.get("type") == "describe":
            # Send server description
            response = {
                "type": "description",
                "name": "PowerPoint Translator",
                "description": "A service that provides PowerPoint translation capabilities using AWS Bedrock models",
                "tools": [
                    {
                        "name": "translate_ppt",
                        "description": "Translate PowerPoint documents to a specified language",
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
                    {
                        "name": "list_supported_languages",
                        "description": "List all supported target languages for translation",
                        "parameters": {
                            "type": "object",
                            "properties": {}
                        }
                    }
                ]
            }
            sys.stdout.write(json.dumps(response) + "\n")
            sys.stdout.flush()
            
        else:
            # Unknown request type
            response = {
                "type": "error",
                "error": f"Unknown request type: {request.get('type')}"
            }
            sys.stdout.write(json.dumps(response) + "\n")
            sys.stdout.flush()
            
    except Exception as e:
        logger.error(f"Error handling MCP request: {str(e)}")
        response = {
            "type": "error",
            "error": f"Error handling request: {str(e)}"
        }
        sys.stdout.write(json.dumps(response) + "\n")
        sys.stdout.flush()

def main():
    """Main function"""
    parser = argparse.ArgumentParser(description='PowerPoint Translator')
    parser.add_argument('--mcp', action='store_true', help='Run in MCP mode')
    parser.add_argument('--translate', action='store_true', help='Translate a PowerPoint file')
    parser.add_argument('--input-file', help='Path to the input PowerPoint file')
    parser.add_argument('--target-language', default=default_target_language, help='Target language code')
    parser.add_argument('--output-file', help='Path to save the translated file')
    parser.add_argument('--method', default='nova', choices=['nova', 'claude'], help='Translation method')
    parser.add_argument('--list-languages', action='store_true', help='List supported languages')
    
    args = parser.parse_args()
    
    if args.mcp:
        print("Starting PowerPoint Translator in MCP mode...")
        while True:
            try:
                handle_mcp_request()
            except KeyboardInterrupt:
                break
            except Exception as e:
                logger.error(f"Error in MCP mode: {str(e)}")
                break
    
    elif args.translate:
        if not args.input_file:
            print("Error: Input file is required for translation")
            parser.print_help()
            sys.exit(1)
            
        result = translate_ppt(
            args.input_file,
            args.target_language,
            args.output_file,
            args.method
        )
        
        if 'error' in result:
            print(f"Error: {result['error']}")
            sys.exit(1)
        else:
            print(f"Translation completed successfully!")
            print(f"Input file: {result['input_file']}")
            print(f"Output file: {result['output_file']}")
            print(f"Target language: {result['target_language']}")
            print(f"Translation method: {result['translation_method']}")
            print(f"Translated {result['translated_texts_count']} text elements")
    
    elif args.list_languages:
        languages = list_supported_languages()
        print("Supported languages:")
        for code, name in languages.items():
            print(f"  {code}: {name}")
    
    else:
        # Default to MCP mode when no arguments are provided
        print("Starting PowerPoint Translator in MCP mode...")
        # Send initial description
        response = {
            "type": "description",
            "name": "PowerPoint Translator",
            "description": "A service that provides PowerPoint translation capabilities using AWS Bedrock models",
            "tools": [
                {
                    "name": "translate_ppt",
                    "description": "Translate PowerPoint documents to a specified language",
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
                {
                    "name": "list_supported_languages",
                    "description": "List all supported target languages for translation",
                    "parameters": {
                        "type": "object",
                        "properties": {}
                    }
                }
            ]
        }
        sys.stdout.write(json.dumps(response) + "\n")
        sys.stdout.flush()
        
        while True:
            try:
                handle_mcp_request()
            except KeyboardInterrupt:
                break
            except Exception as e:
                logger.error(f"Error in MCP mode: {str(e)}")
                break

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        sys.exit(1)
