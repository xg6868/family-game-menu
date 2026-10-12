import base64, os

assets_dir = 'menu_assets'
themes = {
    'sci_fi_night': 'sci_fi_night.jpg',
    'sci_fi_space': 'sci_fi_space.jpg',
    'sci_fi_cyber': 'sci_fi_cyber.jpg',
    'sci_fi_mecha': 'sci_fi_mecha.jpg',
    'sci_fi_matrix': 'sci_fi_matrix.jpg',
    'cat_white': 'cat_white.jpg',
    'cat_orange': 'cat_orange.jpg',
    'cat_black': 'cat_black.jpg',
    'cat_pink': 'cat_pink.jpg',
    'cat_green': 'cat_green.jpg',
}

result = {}
for key, fname in themes.items():
    path = os.path.join(assets_dir, fname)
    with open(path, 'rb') as f:
        data = base64.b64encode(f.read()).decode()
    result[key] = data
    print(f'{key}: {len(data)} chars')

with open('themes_3a_data.py', 'w', encoding='utf-8') as f:
    f.write('# 3A\u6e38\u620f\u58c1\u7eb8base64\u6570\u636e\n')
    f.write('THEME_3A_DATA = {\n')
    for key, data in result.items():
        f.write('    "' + key + '": "' + data + '",\n')
    f.write('}\n')

print('done')
