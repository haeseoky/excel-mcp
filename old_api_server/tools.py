import os
from pathlib import Path

# --- Configuration ---
STORAGE_DIR = Path("storage")
STORAGE_DIR.mkdir(exist_ok=True)

# --- Security Helper ---
def _get_safe_path(filename: str) -> Path:
    """
    Validates filename and returns a secure path within STORAGE_DIR.
    Prevents path traversal attacks. A leading underscore indicates a private helper.
    """
    if ".." in filename or "/" in filename or "\\" in filename:
        raise ValueError("Invalid characters in filename. Subdirectories are not allowed.")
    
    absolute_storage_path = STORAGE_DIR.resolve()
    target_path = (absolute_storage_path / filename).resolve()

    if absolute_storage_path not in target_path.parents and target_path != absolute_storage_path:
        raise ValueError("Path traversal attempt detected. Invalid path.")
    return target_path

# --- Tool Definitions ---

def create_file(filename: str, content: str) -> str:
    """
    Creates or overwrites a file with the given content in the storage directory.
    This tool is for writing or creating new files.
    """
    try:
        path = _get_safe_path(filename)
        path.write_text(content, encoding="utf-8")
        return f"File {filename} created/updated successfully."
    except (ValueError, IOError) as e:
        return f"Error: {e}"

def read_file(filename: str) -> str:
    """
    Reads and returns the content of a specified file from the storage directory.
    This tool is for retrieving the contents of an existing file.
    """
    try:
        path = _get_safe_path(filename)
        if not path.is_file():
            return f"Error: File {filename} not found."
        return path.read_text(encoding="utf-8")
    except (ValueError, IOError) as e:
        return f"Error: {e}"

def list_files() -> list[str]:
    """
    Lists all files in the storage directory.
    This tool is for discovering what files are available.
    """
    if not STORAGE_DIR.is_dir():
        return []
    return [f.name for f in STORAGE_DIR.glob("*") if f.is_file()]

def delete_file(filename: str) -> str:
    """
    Deletes a specified file from the storage directory.
    This tool is for removing files.
    """
    try:
        path = _get_safe_path(filename)
        if not path.is_file():
            return f"Error: File {filename} not found."
        path.unlink()
        return f"File {filename} deleted successfully."
    except (ValueError, IOError) as e:
        return f"Error: {e}"

# --- Tool Registry ---
# A dictionary that maps tool names to their function objects and descriptions.
# This registry is what the MCP server will use to find and understand the available tools.
AVAILABLE_TOOLS = {
    "create_file": {
        "function": create_file,
        "description": "Creates or overwrites a file with given content. Use when asked to write, create, or save a file.",
        "params": {"filename": "string", "content": "string"},
    },
    "read_file": {
        "function": read_file,
        "description": "Reads the content of a specified file. Use when asked to read, open, or get the content of a file.",
        "params": {"filename": "string"},
    },
    "list_files": {
        "function": list_files,
        "description": "Lists all available files. Use when asked to list or show all files.",
        "params": {},
    },
    "delete_file": {
        "function": delete_file,
        "description": "Deletes a specified file. Use when asked to delete or remove a file.",
        "params": {"filename": "string"},
    },
}

