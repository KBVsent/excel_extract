import os
import base64
from pathlib import Path
from openai import AzureOpenAI
from typing import Optional
import json


class VisionModelTester:
    """Test Azure OpenAI vision models with image inputs"""
    
    def __init__(self):
        """Initialize Azure OpenAI client"""
        self.client = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_API_KEY"),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview"),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
        )
        
        # Model deployment names
        self.models = {
            "gpt-4.1": os.getenv("AZURE_GPT4_1_DEPLOYMENT", "gpt-4-vision"),
            "gpt-5": os.getenv("AZURE_GPT5_DEPLOYMENT", "gpt-5-vision")
        }
    
    def encode_image(self, image_path: str) -> str:
        """
        Encode image to base64 string
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Base64 encoded image string
        """
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    
    def test_vision_model(
        self,
        model_name: str,
        image_path: str,
        prompt: str = "What do you see in this image? Please describe it in detail.",
        use_url: bool = False,
        temperature: float = 0.7,
        reasoning_effort: str = "medium"  # For GPT-5: low, medium, high
    ) -> dict:
        """
        Test a vision model with an image input
        
        Args:
            model_name: Name of the model (gpt-4.1 or gpt-5)
            image_path: Path to the image file or URL
            prompt: Custom prompt for the model
            use_url: Whether the image_path is a URL
            temperature: Temperature parameter for generation (not supported by GPT-5)
            max_tokens: Maximum tokens to generate
            reasoning_effort: Reasoning effort for GPT-5 (low, medium, high)
            
        Returns:
            Dictionary containing the response and metadata
        """
        if model_name not in self.models:
            raise ValueError(f"Model {model_name} not found. Available models: {list(self.models.keys())}")
        
        deployment_name = self.models[model_name]
        
        # Prepare image content
        if use_url:
            image_content = {
                "type": "image_url",
                "image_url": {
                    "url": image_path
                }
            }
        else:
            # Encode local image to base64
            base64_image = self.encode_image(image_path)
            image_content = {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/jpeg;base64,{base64_image}"
                }
            }
        
        # Create messages
        messages = [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": prompt
                    },
                    image_content
                ]
            }
        ]
        
        print(f"\n{'='*60}")
        print(f"Testing model: {model_name} (deployment: {deployment_name})")
        print(f"Image: {image_path}")
        print(f"Prompt: {prompt}")
        print(f"{'='*60}\n")
        
        try:
            # Prepare API call parameters based on model
            api_params = {
                "model": deployment_name,
                "messages": messages
            }
            
            # Use appropriate token parameter based on model
            if model_name == "gpt-5":
                # GPT-5 specific parameters
                api_params["reasoning_effort"] = reasoning_effort
                # Don't set temperature for GPT-5, only supports default (1)
            else:
                # Other models support temperature and max_tokens
                api_params["temperature"] = temperature
            
            # Call the API
            response = self.client.chat.completions.create(**api_params)
            
            # Extract usage information
            usage_info = {
                "prompt_tokens": response.usage.prompt_tokens,
                "completion_tokens": response.usage.completion_tokens,
                "total_tokens": response.usage.total_tokens
            }
            
            # Add reasoning tokens info for GPT-5
            if hasattr(response.usage, 'completion_tokens_details') and response.usage.completion_tokens_details:
                details = response.usage.completion_tokens_details
                if hasattr(details, 'reasoning_tokens'):
                    usage_info["reasoning_tokens"] = details.reasoning_tokens
                    usage_info["non_reasoning_tokens"] = response.usage.completion_tokens - details.reasoning_tokens
            
            # Extract response content
            result = {
                "model": model_name,
                "deployment": deployment_name,
                "image_path": image_path,
                "prompt": prompt,
                "response": response.choices[0].message.content,
                "usage": usage_info,
                "finish_reason": response.choices[0].finish_reason
            }
            
            print(f"Response:\n{result['response']}\n")
            print(f"Token usage: {result['usage']}\n")
            
            return result
            
        except Exception as e:
            print(f"Error testing {model_name}: {str(e)}\n")
            return {
                "model": model_name,
                "deployment": deployment_name,
                "image_path": image_path,
                "prompt": prompt,
                "error": str(e)
            }
    
    def compare_models(
        self,
        image_path: str,
        prompts: list[str],
        use_url: bool = False,
        save_results: bool = True,
        output_file: str = "vision_test_results.json",
        reasoning_effort: str = "medium"
    ) -> list[dict]:
        """
        Compare multiple models with different prompts
        
        Args:
            image_path: Path to the image file or URL
            prompts: List of prompts to test
            use_url: Whether the image_path is a URL
            save_results: Whether to save results to a file
            output_file: Output file name for results
            max_tokens: Maximum tokens for each response
            reasoning_effort: Reasoning effort for GPT-5 (low, medium, high)
            
        Returns:
            List of result dictionaries
        """
        all_results = []
        
        for prompt in prompts:
            print(f"\n{'#'*60}")
            print(f"Testing with prompt: {prompt}")
            print(f"{'#'*60}")
            
            for model_name in self.models.keys():
                result = self.test_vision_model(
                    model_name=model_name,
                    image_path=image_path,
                    prompt=prompt,
                    use_url=use_url,
                    reasoning_effort=reasoning_effort
                )
                all_results.append(result)
        
        if save_results:
            output_path = Path(output_file)
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(all_results, f, indent=2, ensure_ascii=False)
            print(f"\nResults saved to: {output_path.absolute()}")
        
        return all_results


def main():
    """Main function to run vision model tests"""
    
    # Initialize tester
    tester = VisionModelTester()
    
    # Example 1: Test with a single image and custom prompts
    image_path = "path/to/your/image.jpg"  # Replace with your image path
    
    # Define test prompts
    test_prompts = [
        "What do you see in this image? Please describe it in detail.",
        "Analyze the objects, colors, and composition of this image.",
        "Extract any text visible in this image.",
        "Describe the mood and atmosphere conveyed by this image.",
        "If this is a chart or diagram, explain what it represents."
    ]
    
    # Example 2: Test with a URL
    # image_url = "https://example.com/image.jpg"
    # results = tester.compare_models(
    #     image_path=image_url,
    #     prompts=test_prompts,
    #     use_url=True
    # )
    
    # Example 3: Test single model with single prompt
    # result = tester.test_vision_model(
    #     model_name="gpt-4.1",
    #     image_path=image_path,
    #     prompt="Describe this image in detail."
    # )
    
    # Example 4: Compare all models with multiple prompts
    # results = tester.compare_models(
    #     image_path=image_path,
    #     prompts=test_prompts,
    #     save_results=True,
    #     output_file="vision_comparison_results.json"
    # )
    
    print("\nVision model testing setup complete!")
    print("\nTo use this script:")
    print("1. Set environment variables:")
    print("   - AZURE_OPENAI_API_KEY")
    print("   - AZURE_OPENAI_ENDPOINT")
    print("   - AZURE_OPENAI_API_VERSION (optional)")
    print("   - AZURE_GPT4_1_DEPLOYMENT")
    print("   - AZURE_GPT5_DEPLOYMENT")
    print("2. Uncomment and modify one of the examples above")
    print("3. Run the script")


if __name__ == "__main__":
    main()
