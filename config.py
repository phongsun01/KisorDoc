import json
import os
from pathlib import Path

from dotenv import load_dotenv
from pydantic import BaseModel, Field


class AppConfig(BaseModel):
    ProjectPath: str
    OnlineMode: str = "Disable"
    DataFolder: str = "1. Data"
    TemplateFolder: str = "2. Templates"
    FileFolder: str = "3. Files"
    DataSheet: str = "GoiThau"
    CloseWord: str = "false"
    TaskManagerProcess: str = "WINWORD.exe"
    AgentPath: str = ""
    ExceptionSheet: str = "S."
    AppName: str = "KisorDoc-AI"

    @property
    def data_path(self) -> Path:
        return Path(self.ProjectPath) / self.DataFolder

    @property
    def template_path(self) -> Path:
        return Path(self.ProjectPath) / self.TemplateFolder

    @property
    def output_path(self) -> Path:
        return Path(self.ProjectPath) / self.FileFolder


def load_config() -> AppConfig:
    load_dotenv()

    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise EnvironmentError("LOCALAPPDATA environment variable not set")
    config_file = Path(local_app_data) / "UiPathProjectConfigs" / "Config-5.txt"
    if not config_file.exists():
        raise FileNotFoundError(f"Config file not found: {config_file}")
    with open(config_file, encoding="utf-8") as f:
        data = json.load(f)

    project_path = os.environ.get("PROJECT_PATH")
    if project_path:
        data["ProjectPath"] = project_path

    return AppConfig(**data)
