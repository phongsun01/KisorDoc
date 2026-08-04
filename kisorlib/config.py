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
    ExcelToWordWidthFactor: int = 90
    FileRetryDelay: float = 2.0
    FileMaxRetries: int = 3
    DanhMucFile: str = "DanhMuc"

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

    env_mapping = {
        "PROJECT_PATH": "ProjectPath",
        "ONLINE_MODE": "OnlineMode",
        "DATA_FOLDER": "DataFolder",
        "TEMPLATE_FOLDER": "TemplateFolder",
        "FILE_FOLDER": "FileFolder",
        "DATA_SHEET": "DataSheet",
        "CLOSE_WORD": "CloseWord",
        "TASK_MANAGER_PROCESS": "TaskManagerProcess",
        "AGENT_PATH": "AgentPath",
        "EXCEPTION_SHEET": "ExceptionSheet",
        "APP_NAME": "AppName",
        "EXCEL_TO_WORD_WIDTH_FACTOR": "ExcelToWordWidthFactor",
        "FILE_RETRY_DELAY": "FileRetryDelay",
        "FILE_MAX_RETRIES": "FileMaxRetries",
        "DANH_MUC_FILE": "DanhMucFile",
    }

    data = {}
    for env_key, config_key in env_mapping.items():
        val = os.environ.get(env_key)
        if val is not None:
            if config_key == "ExcelToWordWidthFactor":
                try:
                    data[config_key] = int(val)
                except ValueError:
                    data[config_key] = 90
            elif config_key == "FileRetryDelay":
                try:
                    data[config_key] = float(val)
                except ValueError:
                    data[config_key] = 2.0
            elif config_key == "FileMaxRetries":
                try:
                    data[config_key] = int(val)
                except ValueError:
                    data[config_key] = 3
            else:
                data[config_key] = val

    if "ProjectPath" not in data:
        raise ValueError("PROJECT_PATH is required in .env file")

    return AppConfig(**data)
