import os
path = r'D:\2022'
folders = os.listdir(path)
for folder in folders:
    if folder.endswith('.rar'):
        folders.remove(folder)

print(folders)
for folder in folders:
    full_path = os.path.join(path,folder)
    print(full_path)
    subdir = os.listdir(full_path)
    for file in subdir:
        if file.endswith('.xlsx'):
            print(file)
