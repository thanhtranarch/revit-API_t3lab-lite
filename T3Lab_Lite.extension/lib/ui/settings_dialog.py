"""
Settings Dialog for T3LabAI MCP
Provides UI for configuring API keys and server settings
"""

from pyrevit import forms


def show_settings_dialog():
    """
    Show the settings dialog for T3LabAI configuration.

    This dialog allows users to:
    - Configure server port
    - Add/remove MCP providers
    - Set API keys for providers
    """
    from config.settings import get_settings

    settings = get_settings()
    server_config = settings.get_server_config()

    # Simple settings dialog using pyrevit forms
    options = [
        "Server Port: {}".format(server_config.get('port', 8080)),
        "Configure API Keys",
        "Manage Providers",
        "Reset to Defaults"
    ]

    selected = forms.SelectFromList.show(
        options,
        title="T3LabAI Settings",
        width=400,
        height=300,
        button_name="Select"
    )

    if selected:
        if "Server Port" in selected:
            _configure_port(settings)
        elif "API Keys" in selected:
            _configure_api_keys(settings)
        elif "Providers" in selected:
            _manage_providers(settings)
        elif "Reset" in selected:
            _reset_settings(settings)


def _configure_port(settings):
    """Configure server port"""
    new_port = forms.ask_for_string(
        default=str(settings.get_server_config().get('port', 8080)),
        prompt="Enter server port:",
        title="Server Port"
    )

    if new_port:
        try:
            port = int(new_port)
            if 1024 <= port <= 65535:
                settings._settings['server']['port'] = port
                settings.save_settings()
                forms.alert("Port updated to {}".format(port), title="Settings Saved")
            else:
                forms.alert("Port must be between 1024 and 65535", title="Invalid Port")
        except ValueError:
            forms.alert("Please enter a valid number", title="Invalid Input")


def _configure_api_keys(settings):
    """Configure API keys"""
    providers = ["Claude", "OpenAI", "Custom"]

    selected = forms.SelectFromList.show(
        providers,
        title="Select Provider",
        button_name="Configure"
    )

    if selected:
        current_key = settings.get_api_key(selected) or ""
        masked_key = "****" + current_key[-4:] if len(current_key) > 4 else ""

        new_key = forms.ask_for_string(
            default=masked_key,
            prompt="Enter API key for {}:".format(selected),
            title="API Key"
        )

        if new_key and not new_key.startswith("****"):
            settings.set_api_key(selected, new_key)
            forms.alert("API key saved for {}".format(selected), title="Settings Saved")


def _manage_providers(settings):
    """Manage MCP providers"""
    forms.alert(
        "Provider management coming soon.\n\n"
        "Current providers: {}".format(len(settings.get_enabled_providers())),
        title="Providers"
    )


def _reset_settings(settings):
    """Reset to default settings"""
    if forms.alert(
        "Are you sure you want to reset all settings to defaults?",
        title="Reset Settings",
        yes=True,
        no=True
    ):
        settings._settings = settings._get_default_settings()
        settings.save_settings()
        forms.alert("Settings reset to defaults", title="Settings Reset")
