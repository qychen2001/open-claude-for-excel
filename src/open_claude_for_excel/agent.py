from argparse import Namespace

from langchain.agents import create_agent
from langchain.agents.middleware import LLMToolSelectorMiddleware, TodoListMiddleware
from langchain_openai import ChatOpenAI

from open_claude_for_excel.tools import all_tools


def create_excel_agent(args: Namespace):
    model = ChatOpenAI(
        model=args.default_model,
        api_key=args.openai_api_key,
        base_url=args.openai_base_url,
        temperature=0.5,
    )
    agent = create_agent(
        model,
        tools=all_tools,
        middleware=[
            TodoListMiddleware(),
            LLMToolSelectorMiddleware(
                model=model,
                max_tools=5,
            ),
        ],
    )
    return agent
