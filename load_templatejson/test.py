import json

json_path = 'E:\wordtest\word_win32\load_templatejson\word_template_config (1).json'



with open(json_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

print(config['section']['differentOddEven'])