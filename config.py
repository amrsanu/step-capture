# config.py
import configparser


class Config:
    @staticmethod
    def load_config(file_path="config.ini"):
        config = configparser.ConfigParser()
        config.read(file_path)
        return config

    @staticmethod
    def save_config(config, file_path="config.ini"):
        with open(file_path, "w") as config_file:
            config.write(config_file)
