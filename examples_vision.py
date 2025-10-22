import os
from pathlib import Path
from dotenv import load_dotenv
from test_vision import VisionModelTester


def test_document_analysis():
    """Test document analysis capabilities"""
    load_dotenv()
    tester = VisionModelTester()
    
    # Example: Testing with a document/screenshot
    image_path = "examples/test_1.png"  # Replace with your image
    
    prompts = [
        """Please treat this image as a document and convert it into a pure Markdown format that is human-readable.

**CRITICAL: Output ONLY the Markdown content. Do NOT include any explanations, introductions, or concluding remarks.**

Requirements:
1. **Text Content**: Extract all visible text exactly as it appears, preserving the original wording.

2. **Text Structure**: Use appropriate Markdown formatting to represent the document structure:
   - Use `#`, `##`, `###` for headers based on hierarchy
   - Use bullet points or numbered lists where applicable
   - Use `**bold**` and `*italic*` to preserve text emphasis
   - Use `>` for quotes or callouts

3. **Tables**: Convert all tables into Markdown table format with proper alignment.

4. **Images and Visual Elements**: 
   - Describe images using the format: `![AI Analysis: detailed description of the image, including key visual elements, colors, and purpose]`
   - For charts/diagrams, provide a detailed description of what they represent

5. **Positional Text Information**: When text position is crucial to meaning (e.g., annotations, labels, captions, side notes):
   - Add context tags like `<!-- Position: top-right corner -->` or `<!-- Label for diagram -->`
   - Use indentation or blockquotes to show spatial relationships
   - For multi-column layouts, clearly separate columns

6. **Layout Preservation**: Maintain the visual hierarchy and spatial relationships:
   - Use horizontal rules `---` to separate major sections
   - Use proper spacing and line breaks
   - Preserve left-to-right, top-to-bottom reading order

**Output Format**: Pure Markdown only, ready to be saved as a .md file. No additional commentary.""",
        
        """Convert this image into a Markdown document. Output ONLY Markdown content without any extra text.

Instructions:
- Extract all text verbatim
- Structure using Markdown headers, lists, tables
- For images: Use `![Description: what the image shows and its purpose in the document]`
- For positioned text (labels, annotations): Add HTML comments like `<!-- Note: positioned at bottom-left -->`
- For tables: Use proper Markdown table syntax
- Preserve document hierarchy and layout through Markdown formatting

Output pure Markdown only.""",
        
        """Transform this image into a pure-text Markdown document.

IMPORTANT: Return ONLY Markdown formatted content. No introductions or explanations.

Processing Rules:
1. Text: Copy exactly as shown
2. Structure: Use `#` headers, lists, and formatting
3. Tables: Convert to Markdown tables
4. Images/Graphics: `![AI Note: description including visual details and context]`
5. Position-dependent text: Add context with `<!-- Position: location and purpose -->`
6. Layout: Use spacing, indentation, and horizontal rules to preserve structure

The goal is a standalone Markdown file that fully represents the original document's information.""",
    ]
    
    print("="*60)
    print("Document Analysis Test - Convert to Pure Text (Markdown)")
    print("="*60)
    
    results = tester.compare_models(
        image_path=image_path,
        prompts=prompts,
        save_results=True,
        output_file="document_analysis_results.json",
        reasoning_effort="low"  # Low effort for efficient text extraction
    )
    
    if results:
        for i, result in enumerate(results):
            model_name = result.get('model', 'unknown')
            response = result.get('response', '')
            
            safe_model_name = model_name.replace('.', '_').replace(':', '_')
            output_filename = f"output_{safe_model_name}_prompt{i+1}.md"
            
            with open(output_filename, 'w', encoding='utf-8') as f:
                f.write(response)
            
            print(f"✓ Saved Markdown output to: {output_filename}")
    
    return results


def test_chart_understanding():
    """Test chart and data visualization understanding"""
    load_dotenv()
    tester = VisionModelTester()
    
    # Example: Testing with charts/graphs
    image_path = "examples/sample_chart.jpg"  # Replace with your image
    
    prompts = [
        "What type of chart is this? Describe the data it represents.",
        "Extract the data points from this chart in a structured format.",
        "What insights or trends can you identify from this visualization?",
        "Convert this chart data into a markdown table.",
    ]
    
    print("="*60)
    print("Chart Understanding Test")
    print("="*60)
    
    results = tester.compare_models(
        image_path=image_path,
        prompts=prompts,
        save_results=True,
        output_file="chart_analysis_results.json"
    )
    
    return results


def test_table_extraction():
    """Test table extraction from images"""
    load_dotenv()
    tester = VisionModelTester()
    
    # Example: Testing with tables
    image_path = "examples/sample_table.jpg"  # Replace with your image
    
    prompt = """
    Extract the table from this image and format it as markdown.
    Preserve the structure, headers, and all data accurately.
    If there are merged cells or complex formatting, describe how to represent them.
    """
    
    print("="*60)
    print("Table Extraction Test")
    print("="*60)
    
    for model_name in ["gpt-4.1", "gpt-5"]:
        result = tester.test_vision_model(
            model_name=model_name,
            image_path=image_path,
            prompt=prompt.strip()
        )
        
        # Save individual result
        output_file = f"table_extraction_{model_name.replace('.', '_')}.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"Model: {model_name}\n")
            f.write(f"{'='*60}\n\n")
            f.write(result.get('response', 'Error occurred'))
        
        print(f"Saved to: {output_file}\n")


def test_multi_language_ocr():
    """Test OCR with multiple languages"""
    load_dotenv()
    tester = VisionModelTester()
    
    # Example: Testing with multi-language text
    image_path = "examples/multilingual_text.jpg"  # Replace with your image
    
    prompts = [
        "Extract all text from this image, preserving the original language.",
        "Identify the languages present in this image.",
        "Extract and translate any non-English text to English.",
    ]
    
    print("="*60)
    print("Multi-language OCR Test")
    print("="*60)
    
    results = tester.compare_models(
        image_path=image_path,
        prompts=prompts,
        save_results=True,
        output_file="multilingual_ocr_results.json"
    )
    
    return results


def test_image_url():
    """Test with a publicly accessible image URL"""
    load_dotenv()
    tester = VisionModelTester()
    
    # Example: Testing with URL
    image_url = "https://example.com/image.jpg"  # Replace with actual URL
    
    prompt = "Describe this image in detail, including objects, colors, and composition."
    
    print("="*60)
    print("Image URL Test")
    print("="*60)
    
    results = []
    for model_name in ["gpt-4.1", "gpt-5"]:
        result = tester.test_vision_model(
            model_name=model_name,
            image_path=image_url,
            prompt=prompt,
            use_url=True
        )
        results.append(result)
    
    return results


def test_custom_scenario():
    """
    Customize this function for your specific testing needs
    """
    load_dotenv()
    tester = VisionModelTester()
    
    # Your custom image path
    image_path = "path/to/your/image.jpg"
    
    # Your custom prompts
    custom_prompts = [
        "Your custom prompt here",
        # Add more prompts
    ]
    
    results = tester.compare_models(
        image_path=image_path,
        prompts=custom_prompts,
        save_results=True,
        output_file="custom_test_results.json"
    )
    
    return results


def main():
    """
    Main function - uncomment the test you want to run
    """
    
    print("\n" + "="*60)
    print("Azure OpenAI Vision Model Testing Suite")
    print("="*60 + "\n")
    
    # Uncomment the test you want to run:
    
    test_document_analysis()
    # test_chart_understanding()
    # test_table_extraction()
    # test_multi_language_ocr()
    # test_image_url()
    # test_custom_scenario()
    
    # Or run all tests
    # test_document_analysis()
    # test_chart_understanding()
    # test_table_extraction()
    # test_multi_language_ocr()
    
    print("\nTo run tests:")
    print("1. Ensure .env file is configured with Azure OpenAI credentials")
    print("2. Replace image paths with your actual test images")
    print("3. Uncomment the test function you want to run")
    print("4. Run: uv run python examples_vision.py")


if __name__ == "__main__":
    main()
