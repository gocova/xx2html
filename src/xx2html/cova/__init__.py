from os import path
from pathlib import Path
import json

from kivy.logger import Logger


def initialize(
    app_config_path: str,
    config_filename: str,
    default_config: dict,
    system_config: dict,
):
    Logger.info("Application: Initializing...")

    Logger.info(f"Application: Using config path |> {app_config_path}")
    Path(app_config_path).mkdir(parents=True, exist_ok=True)

    current_config = default_config
    config_file_path = path.join(app_config_path, config_filename)

    if not path.exists(config_file_path):
        with open(config_file_path, "w") as config_file:
            json.dump(current_config, config_file)
    else:
        with open(config_file_path, "r") as config_file:
            user_config = json.load(config_file)
        if isinstance(user_config, dict):
            current_config.update(user_config)
    current_config.update(system_config)
    Logger.info(f"Application: Current config |> {current_config}")

