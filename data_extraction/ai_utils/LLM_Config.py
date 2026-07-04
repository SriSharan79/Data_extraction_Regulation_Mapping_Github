import os
from .file_manager import ALR_main_folder
import json
# Model repo id 
# model_repo_id = "openai/gpt-oss-20b"
# model_repo_id = "meta-llama/Llama-3.2-3B-Instruct"
model_repo_id = "meta-llama/Llama-3.1-8B-Instruct"
# model_repo_id = "meta-llama/Llama-3.1-8B"
# model_repo_id = "meta-llama/Llama-3.2-1B-Instruct"

# base path where the models were stored
base_path = "/localdata/user/kata_du/LLM Models"
# Adjust the path to your required model directory
local_model_dir = os.path.join(base_path, "00_LLM_model", model_repo_id)

API_keys_config= os.path.join(ALR_main_folder, "API_keys_config.json")

# Map the human-readable key type to its environment-variable name.
KEY_ENV_NAMES = {
    'DLR Ollama': "Ollama_DLR_API_Key",
    'BlaBla Door': "BlaBla_API_Key",
}


def _load_config():
    """Return the persisted API-key config as a dict (empty if none/invalid)."""
    if os.path.exists(API_keys_config):
        try:
            with open(API_keys_config, 'r') as file:
                return json.load(file)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _load_persisted_keys_into_env():
    """
    On import, copy any persisted keys into os.environ (unless already set in
    the environment), so stored keys behave like env vars across launches.
    """
    config_data = _load_config()
    for key_type, env_name in KEY_ENV_NAMES.items():
        if not os.getenv(env_name):
            value = config_data.get(key_type)
            if value:
                os.environ[env_name] = value


_load_persisted_keys_into_env()


def get_stored_api_key(API_Key_type):
    """
    Return the API key for a service from the environment (or persisted config),
    or None if it has not been set. Never prompts.
    """
    env_name = KEY_ENV_NAMES.get(API_Key_type)
    if env_name and os.getenv(env_name):
        return os.getenv(env_name)
    return _load_config().get(API_Key_type) or None


def set_api_key(API_Key_type, api_key):
    """
    Store an API key: set it in os.environ for the current session AND persist
    it to the config file (merging, not overwriting other keys) so it is loaded
    on the next launch.
    """
    api_key = (api_key or "").strip()
    if not api_key:
        return None

    env_name = KEY_ENV_NAMES.get(API_Key_type)
    if env_name:
        os.environ[env_name] = api_key

    config_data = _load_config()
    config_data[API_Key_type] = api_key
    try:
        with open(API_keys_config, 'w') as file:
            json.dump(config_data, file, indent=4)
    except OSError as e:
        print(f"Warning: could not persist API key to {API_keys_config}: {e}")
    return api_key


def get_api_key(API_Key_type):
    """
    CLI helper: return the stored key, or prompt for it on the console and
    persist it. UI code should use get_stored_api_key/set_api_key instead.
    """
    existing = get_stored_api_key(API_Key_type)
    if existing:
        return existing

    print(f"{API_Key_type} API Key not found.")
    api_key = input(f"Please enter your {API_Key_type} API Key: ")
    return set_api_key(API_Key_type, api_key)


def check_api_key(API_Key_type):
    """Return the stored key, prompting on the console (CLI) only if missing."""
    return get_api_key(API_Key_type)


# System Prompts

BLABLADOR_BASE_URL = "https://api.helmholtz-blablador.fz-juelich.de/v1"
PREFERRED_BLABLADOR_MODELS = [
    "GPT-OSS-120b", "Llama-3.1-70B-Instruct", "Llama-3.1-8B-Instruct",
    "Mistral-Large-Instruct-2407", "Mistral-7B-Instruct-v0.3"
]

# DLR Ollama OpenAI-compatible endpoint (base; specific paths are appended).
OLLAMA_BASE_URL = "http://ollama.nimbus.dlr.de/api"

# Default models used unless the user selects a different one at runtime.
# Keep these as the previously hard-coded values so behaviour is unchanged
# out of the box; interactive selection can override them per session.
DEFAULT_BLABLADOR_MODEL = "01 - GPT-OSS-120b - an open model released by OpenAI in August 2025"
DEFAULT_OLLAMA_MODEL = "gpt-oss:20b"

if __name__ == "__main__":
    
    Ollama_DLR_API_Key=check_api_key('DLR Ollama')
    BlaBla_API_Key = check_api_key('BlaBla Door')
    print (BlaBla_API_Key)
    print (Ollama_DLR_API_Key)
