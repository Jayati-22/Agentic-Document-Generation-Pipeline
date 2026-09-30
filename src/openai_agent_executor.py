
import json
import logging
import inspect

from typing import Any

from a2a.server.agent_execution import AgentExecutor
from a2a.server.agent_execution.context import RequestContext
from a2a.server.events.event_queue import EventQueue
from a2a.server.tasks import TaskUpdater
from a2a.types import AgentCard, TaskState, TextPart, UnsupportedOperationError
from a2a.utils.errors import ServerError
from openai import AsyncOpenAI

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

class OpenAIAgentExecutor(AgentExecutor):
    def __init__(self, card: AgentCard, tools: dict[str, Any], api_key: str, system_prompt: str):
        logger.info(f"🚀 Initializing OpenAIAgentExecutor")
        self._card = card
        self.tools = tools
        logger.info(f"   Card: {card.name}")
        logger.info(f"   Tools registered: {list(tools.keys())}")
        
        self.client = AsyncOpenAI(api_key=api_key)
        self.model = 'gpt-4o'
        self.system_prompt = system_prompt
        logger.info(f"   Model: {self.model}")
        logger.info(f"✓ OpenAIAgentExecutor initialized")

    async def _process_request(self, message_text: str, context: RequestContext, task_updater: TaskUpdater) -> None:
        logger.info(f"📨 _process_request called")
        logger.info(f"   INPUT: message_text = {message_text}")
        
        messages = [
            {'role': 'system', 'content': self.system_prompt},
            {'role': 'user', 'content': message_text},
        ]
        logger.debug(f"   Initial messages prepared: {len(messages)} messages")

        openai_tools = []
        for tool_name, tool_instance in self.tools.items():
            logger.debug(f"   Processing tool: {tool_name}")
            if hasattr(tool_instance, tool_name):
                func = getattr(tool_instance, tool_name)
                schema = self._extract_function_schema(func)
                openai_tools.append({'type': 'function', 'function': schema})
                logger.debug(f"     ✓ Tool registered: {tool_name}")
            else:
                logger.debug(f"     ✗ Tool not found: {tool_name}")

        logger.info(f"   Total tools registered: {len(openai_tools)}")

        max_iterations = 10
        iteration = 0

        while iteration < max_iterations:
            iteration += 1
            logger.info(f"   [ITERATION {iteration}/{max_iterations}]")
            try:
                logger.debug(f"     Calling OpenAI API with {len(openai_tools)} tools...")
                response = await self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    tools=openai_tools if openai_tools else None,
                    tool_choice='auto' if openai_tools else None,
                    temperature=0.1,
                    max_tokens=4000,
                )

                message = response.choices[0].message
                logger.debug(f"     ✓ Got response")
                logger.debug(f"     Response content: {message.content}")
                logger.debug(f"     Tool calls: {len(message.tool_calls) if message.tool_calls else 0}")
                
                messages.append({
                    'role': 'assistant',
                    'content': message.content,
                    'tool_calls': message.tool_calls,
                })

                if message.tool_calls:
                    logger.info(f"     Processing {len(message.tool_calls)} tool call(s)...")
                    for idx, tool_call in enumerate(message.tool_calls):
                        function_name = tool_call.function.name
                        function_args = json.loads(tool_call.function.arguments)

                        logger.info(f"       [{idx+1}] Tool Call: {function_name}")
                        logger.info(f"           Args: {json.dumps(function_args)}")

                        if function_name in self.tools:
                            tool_instance = self.tools[function_name]
                            if hasattr(tool_instance, function_name):
                                method = getattr(tool_instance, function_name)
                                logger.debug(f"           Executing {function_name}...")
                                result = method(**function_args)
                                if inspect.iscoroutine(result):
                                    logger.debug(f"           Awaiting coroutine...")
                                    result = await result
                                logger.info(f"           ✓ Result: {result[:100] if isinstance(result, str) else str(result)[:100]}")
                            else:
                                result = {'error': f'Method {function_name} not found'}
                                logger.error(f"           ✗ Method not found")
                        else:
                            result = {'error': f'Function {function_name} not found'}
                            logger.error(f"           ✗ Function not found in tools")

                        if hasattr(result, 'model_dump'):
                            result_json = json.dumps(result.model_dump())
                        elif isinstance(result, dict):
                            result_json = json.dumps(result)
                        else:
                            result_json = str(result)

                        messages.append({
                            'role': 'tool',
                            'tool_call_id': tool_call.id,
                            'content': result_json,
                        })

                    logger.debug(f"     All tool calls processed, continuing loop...")
                    await task_updater.update_status(
                        TaskState.working,
                        message=task_updater.new_agent_message([TextPart(text='Processing tool calls...')]),
                    )
                    continue

                if message.content:
                    parts = [TextPart(text=message.content)]
                    logger.info(f"     ✓ Final response: {message.content}")
                    logger.debug(f'     Yielding final response: {parts}')
                    await task_updater.add_artifact(parts)
                    await task_updater.complete()
                    logger.info(f"📨 _process_request completed successfully")
                break

            except Exception as e:
                logger.error(f'     ✗ Error in OpenAI API call: {e}', exc_info=True)
                await task_updater.add_artifact([TextPart(text=f'An error occurred: {e!s}')])
                await task_updater.complete()
                logger.error(f"📨 _process_request failed with error")
                break

        if iteration >= max_iterations:
            logger.warning(f"⚠️  Request exceeded maximum iterations ({max_iterations})")
            await task_updater.add_artifact([TextPart(text='Request exceeded maximum iterations.')])
            await task_updater.complete()

    def _extract_function_schema(self, func):
        import inspect
        sig = inspect.signature(func)
        docstring = inspect.getdoc(func) or ''
        lines = docstring.split('\n')
        description = lines[0] if lines else func.__name__
        properties = {}
        required = []

        for param_name, param in sig.parameters.items():
            param_type = 'string'
            if param.annotation != inspect.Parameter.empty:
                if param.annotation == int: param_type = 'integer'
                elif param.annotation == float: param_type = 'number'
                elif param.annotation == bool: param_type = 'boolean'
                elif param.annotation == list: param_type = 'array'
                elif param.annotation == dict: param_type = 'object'

            if param.default == inspect.Parameter.empty:
                required.append(param_name)

            properties[param_name] = {'type': param_type, 'description': f'Parameter {param_name}'}

        return {
            'name': func.__name__,
            'description': description,
            'parameters': {'type': 'object', 'properties': properties, 'required': required},
        }

    async def execute(self, context: RequestContext, event_queue: EventQueue):
        logger.info(f"🔧 execute() called")
        logger.info(f"   context.task_id: {context.task_id}")
        logger.info(f"   context.context_id: {context.context_id}")
        logger.info(f"   context.message: {context.message}")
        
        updater = TaskUpdater(event_queue, context.task_id, context.context_id)
        if not context.current_task:
            logger.debug(f"   No current task, submitting...")
            await updater.submit()
        await updater.start_work()

        message_text = ''
        for part in context.message.parts:
            if isinstance(part.root, TextPart):
                message_text += part.root.text

        logger.info(f"   Extracted message_text: {message_text}")
        logger.info(f"   Calling _process_request...")
        await self._process_request(message_text, context, updater)

    async def cancel(self, context: RequestContext, event_queue: EventQueue):
        logger.warning(f"⚠️  cancel() called")
        raise ServerError(error=UnsupportedOperationError())