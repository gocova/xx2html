from copy import deepcopy
from os import path
from pathlib import Path
import json
import logging


def initialize(
    app_config_path: str,
    config_filename: str,
    default_config: dict,
    system_config: dict,
):
    """Load app config from disk, merge defaults/system values, and return it."""
    logging.info("Application: Initializing...")

    logging.info(f"Application: Using config path |> {app_config_path}")
    Path(app_config_path).mkdir(parents=True, exist_ok=True)

    current_config = deepcopy(default_config)
    config_file_path = path.join(app_config_path, config_filename)

    if not path.exists(config_file_path):
        with open(config_file_path, "w", encoding="utf-8") as config_file:
            json.dump(current_config, config_file)
    else:
        with open(config_file_path, "r", encoding="utf-8") as config_file:
            user_config = json.load(config_file)
        if isinstance(user_config, dict):
            current_config.update(user_config)
    current_config.update(system_config)
    logging.info(f"Application: Current config |> {current_config}")
    return current_config
