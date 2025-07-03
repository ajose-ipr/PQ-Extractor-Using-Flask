import os

# Folders to skip entirely
IGNORED_DIRS = {'.git', '__pycache__', 'env', 'venv', '.idea', '.vscode'}

# Optionally read .gitignore entries
def load_gitignore(root_path):
    gitignore_path = os.path.join(root_path, '.gitignore')
    ignored_files = set()
    if os.path.exists(gitignore_path):
        with open(gitignore_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    ignored_files.add(os.path.normpath(line))
    return ignored_files

def generate_directory_tree(start_path='.', indent='', ignored_files=None):
    tree_output = ''
    items = sorted(os.listdir(start_path))
    items = [item for item in items if item not in IGNORED_DIRS]

    for index, item in enumerate(items):
        path = os.path.join(start_path, item)
        rel_path = os.path.relpath(path)

        # Skip if explicitly ignored in .gitignore
        if ignored_files and rel_path in ignored_files:
            continue

        is_last = index == len(items) - 1
        prefix = '└── ' if is_last else '├── '
        tree_output += indent + prefix + item + '\n'

        if os.path.isdir(path):
            extension = '    ' if is_last else '│   '
            tree_output += generate_directory_tree(path, indent + extension, ignored_files)

    return tree_output

if __name__ == '__main__':
    root_dir = os.getcwd()
    print(f"Generating directory tree for: {root_dir}")

    ignored_files = load_gitignore(root_dir)

    tree = f"Directory tree for: {root_dir}\n\n"
    tree += generate_directory_tree(root_dir, ignored_files=ignored_files)

    with open("directory_tree.txt", "w", encoding="utf-8") as f:
        f.write(tree)

    print("Saved to directory_tree.txt ✅")
    print(tree)