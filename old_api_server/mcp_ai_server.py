import uvicorn
from fastapi import FastAPI, HTTPException, status
from pydantic import BaseModel
import re

# Import the tools and the registry from our tools module
from tools import AVAILABLE_TOOLS

# --- FastAPI App Initialization ---
app = FastAPI(
    title="MCP AI Server",
    description="An intelligent server that interprets natural language prompts to execute file operations.",
    version="1.0.0",
)

# --- Pydantic Models for API Data Structure ---
class InvokeRequest(BaseModel):
    """Defines the structure for an incoming prompt."""
    prompt: str

class InvokeResponse(BaseModel):
    """Defines the structure for the response after executing a tool."""
    tool_used: str
    parameters: dict
    result: str | list[str] # Result can be a string or a list of strings (for list_files)

# --- Core MCP Logic ---
def select_tool(prompt: str) -> tuple[str, dict] | None:
    """
    Parses the user prompt to select the appropriate tool and extract its parameters.
    This is a simplified rule-based implementation of what a real LLM would do.
    """
    prompt = prompt.lower().strip()

    # Rule for "list_files"
    if "list files" in prompt or "show all files" in prompt:
        return "list_files", {}

    # Rule for "read_file"
    match = re.search(r"read (?:file )?([\w\.-]+)", prompt)
    if match:
        return "read_file", {"filename": match.group(1)}

    # Rule for "delete_file"
    match = re.search(r"delete (?:file )?([\w\.-]+)", prompt)
    if match:
        return "delete_file", {"filename": match.group(1)}

    # Rule for "create_file" - more complex, needs filename and content
    # Example prompt: "create file my_file.txt with content: some text here"
    match = re.search(r"create (?:file )?([\w\.-]+) with content: (.*)", prompt, re.DOTALL)
    if match:
        return "create_file", {"filename": match.group(1), "content": match.group(2).strip()}
    
    return None # No tool matched

# --- API Endpoint ---
@app.post("/invoke", response_model=InvokeResponse, summary="Invoke a tool with a natural language prompt")
def invoke_tool(request: InvokeRequest):
    """
    Receives a natural language prompt, selects the appropriate tool,
    executes it, and returns the result.
    """
    selection = select_tool(request.prompt)
    if not selection:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Could not determine the appropriate tool for the given prompt.",
        )

    tool_name, params = selection
    
    if tool_name not in AVAILABLE_TOOLS:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Selected tool {tool_name} is not registered.",
        )

    # Execute the function
    tool_info = AVAILABLE_TOOLS[tool_name]
    func = tool_info["function"]
    
    try:
        # Call the function with the extracted parameters
        result = func(**params)
        return InvokeResponse(tool_used=tool_name, parameters=params, result=result)
    except Exception as e:
        # This handles cases where the function call fails for unexpected reasons
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"An error occurred while executing tool {tool_name}: {e}",
        )

# --- Server Startup Logic ---
if __name__ == "__main__":
    print("MCP AI Server starting...")
    print("Access interactive API docs at http://127.0.0.1:8000/docs")
    # Correctly run uvicorn by referencing the module and app instance as a string
    uvicorn.run("mcp_ai_server:app", host="0.0.0.0", port=8001, reload=True)

