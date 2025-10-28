from .tools import all_tools

if __name__ == "__main__":
    for tool in all_tools:
        print(f"Tool: {tool}")
        print(f"Args: {tool.args}")
        print(f"Description: {tool.description}")
        print("-" * 100)
