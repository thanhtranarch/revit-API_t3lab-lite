"""
T3LabAI Command Registry
Manages available commands and their statistics
"""


class T3LabAIRegistry:
    """Registry for MCP commands"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(T3LabAIRegistry, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return

        self._commands = {}
        self._ai_enhanced_commands = set()
        self._command_stats = {}
        self._initialized = True

    def register_command(self, name, handler, ai_enhanced=False):
        """Register a new command"""
        self._commands[name] = handler
        if ai_enhanced:
            self._ai_enhanced_commands.add(name)
        self._command_stats[name] = 0

    def get_command_statistics(self):
        """Get command statistics"""
        total_executions = sum(self._command_stats.values())
        most_used = max(self._command_stats, key=self._command_stats.get) if self._command_stats else None

        return {
            'total_commands': len(self._commands),
            'ai_enhanced_count': len(self._ai_enhanced_commands),
            'most_used': most_used,
            'total_executions': total_executions
        }

    def get_available_commands(self):
        """Get list of available command names"""
        return list(self._commands.keys())

    def get_ai_enhanced_commands(self):
        """Get list of AI-enhanced command names"""
        return list(self._ai_enhanced_commands)


# Singleton instance
t3labai_registry = T3LabAIRegistry()
