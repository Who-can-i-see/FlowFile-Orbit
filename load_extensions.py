from os import listdir
import json

class load_extensions:
    def __init__(self):
        # 筛选插件配置文件
        self.extensions = [extension for extension in listdir('extensions') if extension.endswith('.json')]
        self.extensions_config = dict()
    
    def load_extensions(self):
        # 加载插件配置文件
        for extension in self.extensions:
            with open(f'extensions/{extension}', 'r', encoding='utf-8') as f:
                self.extensions_config[extension] = json.load(f)
        
        return self.extensions_config
