"""
Button State Management for AI Connection Panel
Handles visual state updates for connection buttons
"""

import os

# Global state for tracking connection status
_connection_state = {
    'connected': False,
    'error': False,
    'last_error': None
}


def update_connection_buttons(connected=False, error=False, error_message=None):
    """
    Update the visual state of connection buttons.

    In PyRevit, button states are typically managed through:
    - Different icon files (icon_on.png, icon_off.png)
    - The button's toggle state

    Args:
        connected: Whether the server is connected
        error: Whether there is an error
        error_message: The error message if any
    """
    global _connection_state

    _connection_state['connected'] = connected
    _connection_state['error'] = error
    _connection_state['last_error'] = error_message if error else None

    # Note: In a full implementation, this would use pyrevit's
    # script.toggle_icon() or similar methods to update button visuals
    # The actual icon switching happens based on the server state
    # when the ribbon is refreshed


def get_connection_state():
    """Get the current connection state"""
    return _connection_state.copy()


def get_last_error():
    """Get the last error message"""
    return _connection_state.get('last_error')


def is_connected():
    """Check if currently connected"""
    return _connection_state.get('connected', False)


def has_error():
    """Check if there is an error"""
    return _connection_state.get('error', False)
