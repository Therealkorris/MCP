"""
Ollama service for MCP-Visio server.
Provides integration with Ollama for AI capabilities.
"""

import os
import json
import logging
import httpx
from typing import Dict, Any, List, Optional
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure logging
logger = logging.getLogger(__name__)

class OllamaService:
    """Service for interacting with Ollama API."""
    
    def __init__(self):
        """Initialize the Ollama service."""
        self.api_url = os.getenv("OLLAMA_API_URL", "http://localhost:11434")
        self.model = os.getenv("OLLAMA_MODEL", "llama3")
        logger.info(f"Initialized Ollama service with model: {self.model}")
    
    async def generate(self, prompt: str, system_prompt: Optional[str] = None) -> Dict[str, Any]:
        """
        Generate text using Ollama.
        
        Args:
            prompt: The prompt to send to Ollama
            system_prompt: Optional system prompt to provide context
        
        Returns:
            Dictionary with the generated text
        """
        try:
            # Prepare the request payload
            payload = {
                "model": self.model,
                "prompt": prompt,
                "stream": False
            }
            
            # Add system prompt if provided
            if system_prompt:
                payload["system"] = system_prompt
            
            # Send the request to Ollama
            async with httpx.AsyncClient(timeout=60.0) as client:
                response = await client.post(
                    f"{self.api_url}/api/generate",
                    json=payload
                )
                
                # Check if request was successful
                if response.status_code == 200:
                    result = response.json()
                    return {
                        "status": "success",
                        "response": result.get("response", ""),
                        "model": self.model
                    }
                else:
                    logger.error(f"Ollama API error: {response.status_code} - {response.text}")
                    return {
                        "status": "error",
                        "message": f"Ollama API error: {response.status_code}",
                        "details": response.text
                    }
        
        except Exception as e:
            logger.error(f"Error generating text with Ollama: {e}")
            return {
                "status": "error",
                "message": f"Error generating text: {str(e)}"
            }
    
    async def analyze_diagram(self, diagram_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Analyze a Visio diagram using Ollama.
        
        Args:
            diagram_data: Dictionary with diagram information
        
        Returns:
            Dictionary with analysis results
        """
        try:
            # Create a prompt for the AI to analyze the diagram
            prompt = f"""
            Please analyze this Visio diagram data and provide insights:
            
            {json.dumps(diagram_data, indent=2)}
            
            Focus on:
            1. The main components and their relationships
            2. The overall structure and organization
            3. Any potential improvements or issues
            
            Provide a clear and concise analysis.
            """
            
            system_prompt = "You are an expert at analyzing and understanding diagrams. Provide a concise, professional analysis."
            
            # Generate the analysis using Ollama
            result = await self.generate(prompt, system_prompt)
            
            return result
        
        except Exception as e:
            logger.error(f"Error analyzing diagram with Ollama: {e}")
            return {
                "status": "error",
                "message": f"Error analyzing diagram: {str(e)}"
            }
    
    async def suggest_improvements(self, diagram_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Suggest improvements for a Visio diagram using Ollama.
        
        Args:
            diagram_data: Dictionary with diagram information
        
        Returns:
            Dictionary with improvement suggestions
        """
        try:
            # Create a prompt for the AI to suggest improvements
            prompt = f"""
            Please suggest improvements for this Visio diagram:
            
            {json.dumps(diagram_data, indent=2)}
            
            Focus on:
            1. Layout and organization
            2. Clarity and readability
            3. Visual design
            4. Information flow
            
            Provide specific, actionable suggestions.
            """
            
            system_prompt = "You are an expert at designing clear and effective diagrams. Provide practical, specific suggestions for improvement."
            
            # Generate the suggestions using Ollama
            result = await self.generate(prompt, system_prompt)
            
            return result
        
        except Exception as e:
            logger.error(f"Error suggesting improvements with Ollama: {e}")
            return {
                "status": "error",
                "message": f"Error suggesting improvements: {str(e)}"
            } 