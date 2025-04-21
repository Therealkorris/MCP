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
            # Create a specialized prompt based on diagram elements
            shapes_count = 0
            connections_count = 0
            pages_count = diagram_data.get("pages_count", 0)
            
            # Extract shape information for smarter analysis
            page_info = []
            shapes_by_type = {}
            
            # Process pages and shapes
            for page in diagram_data.get("pages", []):
                page_name = page.get("name", "Unnamed")
                shapes = page.get("shapes", [])
                connections = page.get("connections", [])
                
                shapes_count += len(shapes)
                connections_count += len(connections)
                
                # Categorize shapes by type
                for shape in shapes:
                    shape_type = shape.get("type", "Unknown")
                    if shape_type not in shapes_by_type:
                        shapes_by_type[shape_type] = 0
                    shapes_by_type[shape_type] += 1
                
                page_info.append({
                    "name": page_name,
                    "shapes_count": len(shapes),
                    "connections_count": len(connections)
                })
            
            # Create a prompt for the AI to analyze the diagram
            prompt = f"""
            Please analyze this Visio diagram data and provide a detailed breakdown:
            
            Diagram Name: {diagram_data.get("name", "Untitled")}
            Pages: {pages_count}
            Total Shapes: {shapes_count}
            Total Connections: {connections_count}
            
            Page Details:
            {json.dumps(page_info, indent=2)}
            
            Shape Types Distribution:
            {json.dumps(shapes_by_type, indent=2)}
            
            Full Diagram Data:
            {json.dumps(diagram_data, indent=2)}
            
            Please provide the following analysis:
            1. The main purpose and type of this diagram (flowchart, organization chart, etc.)
            2. The key components and their relationships
            3. The overall structure and organization
            4. Any patterns or notable features in the diagram
            5. Suggestions for improvement if applicable
            
            Provide a comprehensive analysis with insights about what this diagram represents.
            """
            
            system_prompt = """
            You are an expert at analyzing and understanding Visio diagrams. 
            You can identify diagram types, understand relationships between shapes, 
            and extract meaning from visual representations.
            
            Provide a professional, structured analysis focusing on:
            - The diagram's purpose and intended audience
            - Key components and their relationships
            - Data flow or hierarchical structure
            - Clarity and effectiveness of the visual representation
            
            Base your analysis only on the provided diagram data.
            """
            
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
            # Extract diagram information for targeted suggestions
            shapes_count = 0
            connections_count = 0
            pages_count = diagram_data.get("pages_count", 0)
            
            # Process pages and shapes
            for page in diagram_data.get("pages", []):
                shapes = page.get("shapes", [])
                connections = page.get("connections", [])
                
                shapes_count += len(shapes)
                connections_count += len(connections)
            
            # Create a prompt for the AI to suggest improvements
            prompt = f"""
            Please suggest specific improvements for this Visio diagram:
            
            Diagram Name: {diagram_data.get("name", "Untitled")}
            Pages: {pages_count}
            Total Shapes: {shapes_count}
            Total Connections: {connections_count}
            
            Full Diagram Data:
            {json.dumps(diagram_data, indent=2)}
            
            Focus on the following improvement areas:
            1. Layout and organization (is the diagram well-structured and balanced?)
            2. Clarity and readability (are shapes and connections clear and easy to follow?)
            3. Visual design (colors, sizes, and styles)
            4. Information flow (does the diagram effectively communicate its purpose?)
            5. Completeness (is any important information missing?)
            6. Technical implementation (use of Visio features and best practices)
            
            For each suggestion, please provide:
            - The specific issue identified
            - Why it's a problem for diagram effectiveness
            - A concrete recommendation for improvement
            - How to implement the change in Visio
            
            Provide actionable, practical suggestions tailored to this specific diagram.
            """
            
            system_prompt = """
            You are an expert Visio diagram designer with deep knowledge of visual communication principles.
            Your role is to provide constructive, specific feedback to improve diagrams for clarity,
            effectiveness, and professional appearance.
            
            When suggesting improvements:
            - Be specific and actionable in your recommendations
            - Prioritize changes that will have the biggest impact
            - Consider the diagram's purpose and target audience
            - Reference Visio capabilities and features when relevant
            - Provide a balance of structural, visual, and content-focused improvements
            
            Your goal is to help transform ordinary diagrams into clear, effective visual communication tools.
            """
            
            # Generate the suggestions using Ollama
            result = await self.generate(prompt, system_prompt)
            
            return result
        
        except Exception as e:
            logger.error(f"Error suggesting improvements with Ollama: {e}")
            return {
                "status": "error",
                "message": f"Error suggesting improvements: {str(e)}"
            }
    
    async def generate_diagram_from_description(self, description: str) -> Dict[str, Any]:
        """
        Generate a plan for creating a Visio diagram from a text description.
        
        Args:
            description: Text description of the desired diagram
            
        Returns:
            Dictionary with diagram creation plan
        """
        try:
            # Create a prompt for the AI to generate a diagram plan
            prompt = f"""
            I need to create a Visio diagram based on this description:
            
            {description}
            
            Please provide a detailed plan for creating this diagram in Visio, including:
            
            1. Diagram Type: The most appropriate Visio diagram type to use
            2. Required Stencils: Specific Visio stencils needed
            3. Page Setup: Recommended orientation and size
            4. Shape Inventory: List of all shapes needed with their purpose
            5. Connection Plan: How shapes should be connected
            6. Layout Strategy: Recommended arrangement of elements
            7. Hierarchical Structure: Any parent-child relationships to represent
            8. Color Scheme: Suggested colors for clarity and meaning
            9. Text Elements: Labels and annotations needed
            10. Step-by-Step Creation Process: Ordered steps to create the diagram
            
            Make your plan specific enough that I can follow it to create the exact diagram needed.
            """
            
            system_prompt = """
            You are an expert Visio diagram designer who can translate text descriptions into precise
            diagram creation plans. You have deep knowledge of Visio's capabilities, stencils, and best practices
            for different diagram types.
            
            When creating diagram plans:
            - Select the most appropriate diagram type for the content
            - Recommend specific Visio stencils and shapes
            - Provide clear, step-by-step instructions
            - Consider layout principles for clarity and flow
            - Include both structure and visual design elements
            
            Your goal is to create plans that anyone can follow to create professional, effective diagrams.
            """
            
            # Generate the diagram plan using Ollama
            result = await self.generate(prompt, system_prompt)
            
            return result
        
        except Exception as e:
            logger.error(f"Error generating diagram plan with Ollama: {e}")
            return {
                "status": "error",
                "message": f"Error generating diagram plan: {str(e)}"
            }
    
    async def explain_diagram_element(self, element_data: Dict[str, Any], diagram_context: Dict[str, Any]) -> Dict[str, Any]:
        """
        Explain a specific element in a Visio diagram.
        
        Args:
            element_data: Data about the specific element (shape or connection)
            diagram_context: General information about the diagram for context
            
        Returns:
            Dictionary with explanation of the element
        """
        try:
            # Determine if this is a shape or connection
            element_type = "shape" if "shape_id" in element_data else "connection"
            
            # Create a prompt for the AI to explain the element
            prompt = f"""
            Please explain the purpose and significance of this {element_type} in the Visio diagram:
            
            Element Data:
            {json.dumps(element_data, indent=2)}
            
            Diagram Context:
            {json.dumps(diagram_context, indent=2)}
            
            Please provide:
            1. The function of this {element_type} in the diagram
            2. Its relationship to connected elements
            3. Its importance in the overall diagram
            4. Any technical or domain-specific meaning it might represent
            
            Focus on explaining what this element represents in the context of the diagram's purpose.
            """
            
            system_prompt = """
            You are an expert at interpreting elements within Visio diagrams. You can determine
            the purpose and significance of individual shapes and connections based on their properties
            and context within the larger diagram.
            
            When explaining diagram elements:
            - Consider the element's shape type and properties
            - Look at connections to other elements
            - Analyze text content associated with the element
            - Consider the diagram type and overall purpose
            - Explain technical significance when appropriate
            
            Your goal is to provide clear, insightful explanations of what specific elements represent.
            """
            
            # Generate the explanation using Ollama
            result = await self.generate(prompt, system_prompt)
            
            return result
        
        except Exception as e:
            logger.error(f"Error explaining diagram element with Ollama: {e}")
            return {
                "status": "error",
                "message": f"Error explaining diagram element: {str(e)}"
            } 