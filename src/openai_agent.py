
import logging
from agent_toolset import DocumentCreatorToolset 

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def create_agent():
    logger.info("🚀 Creating agent...")
    toolset = DocumentCreatorToolset()
    logger.info("✓ DocumentCreatorToolset initialized")

    system_prompt = """
    You are an Autonomous Document Engine.
    When a user asks for ANY kind of document, report, pitch, spreadsheet, presentation, or data summary, you MUST IMMEDIATELY call the `generate_document` tool.
    
    STRICT RULES:
    1. DO NOT ask clarifying questions. 
    2. DO NOT ask if they want a presentation or a spreadsheet. 
    3. Just call the tool with their exact brief. The tool's internal logic will autonomously figure out the correct format.
    4. If they ask to update/change an existing document, IMMEDIATELY call `revise_document`.
    
    Once a tool succeeds, politely inform the user and provide them with the file path.
    """

    # Return toolset instance so executor can call getattr(toolset, method_name)
    tools = {
        'generate_document': toolset,
        'revise_document': toolset,
    }
    logger.info("✓ Tools registered: generate_document, revise_document")
    logger.info("✓ Agent creation complete")

    return {'tools': tools, 'system_prompt': system_prompt}